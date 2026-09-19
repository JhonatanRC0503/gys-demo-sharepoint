"""Aprovisiona (idempotente, vía PUT) el pipeline de indexación SharePoint -> Azure AI Search:

  1. Data source (type=sharepoint) apuntando al SITIO, filtrado a la carpeta de la demo.
  2. Índice con campo de contenido + vector (búsqueda híbrida: texto BM25 + vector, con
     configuración semántica opcional).
  3. Skillset que genera el embedding del contenido con el modelo de Azure OpenAI
     (text-embedding-3-small) usando la identidad administrada del servicio de Search.
  4. Indexer que conecta todo, con schedule automático cada 5 minutos (near-real-time).

Ejecutar: python -m backend.scripts.provision_search
Requiere: az login (se usa AzureCliCredential para autenticar contra Search y no exponer keys).
"""
from ..common.config import settings
from ..common.search_rest import put, post


def build_connection_string() -> str:
    parts = [
        f"SharePointOnlineEndpoint={settings.sharepoint_site_url}",
        f"ApplicationId={settings.client_id}",
    ]
    if settings.client_secret:
        parts.append(f"ApplicationSecret={settings.client_secret}")
    parts.append(f"TenantId={settings.tenant_id}")
    return ";".join(parts)


def create_datasource() -> None:
    include_folder = f"https://{settings.sharepoint_tenant_host}/{settings.sharepoint_folder_path}"
    body = {
        "name": settings.search_datasource_name,
        "type": "sharepoint",
        "credentials": {"connectionString": build_connection_string()},
        "container": {
            "name": "useQuery",
            "query": f"includeFolder={include_folder}",
        },
    }
    put(f"datasources/{settings.search_datasource_name}", body)
    print(f"Data source '{settings.search_datasource_name}' creado/actualizado.")


def create_index() -> None:
    body = {
        "name": settings.search_index_name,
        "fields": [
            {"name": "id", "type": "Edm.String", "key": True, "searchable": False, "filterable": True},
            {"name": "content", "type": "Edm.String", "searchable": True, "filterable": False, "retrievable": True},
            # 'title'/'url' con estos nombres exactos: el tool de Azure AI Search del agente
            # los usa por convención para renderizar la cita (nombre real + link), en vez de 'doc_0'.
            {"name": "title", "type": "Edm.String", "searchable": True, "filterable": True, "retrievable": True},
            {"name": "url", "type": "Edm.String", "filterable": True, "retrievable": True},
            {
                "name": "content_vector",
                "type": "Collection(Edm.Single)",
                "searchable": True,
                "retrievable": False,
                "dimensions": settings.aoai_embedding_dimensions,
                "vectorSearchProfile": "vector-profile",
            },
            {"name": "metadata_spo_item_name", "type": "Edm.String", "searchable": True, "filterable": True, "retrievable": True},
            {"name": "metadata_spo_item_weburi", "type": "Edm.String", "filterable": True, "retrievable": True},
            {"name": "metadata_spo_item_path", "type": "Edm.String", "filterable": True, "retrievable": True},
            {"name": "metadata_spo_item_content_type", "type": "Edm.String", "filterable": True, "facetable": True, "retrievable": True},
            {"name": "metadata_spo_item_last_modified", "type": "Edm.DateTimeOffset", "filterable": True, "sortable": True, "retrievable": True},
        ],
        "vectorSearch": {
            "algorithms": [{"name": "hnsw-config", "kind": "hnsw"}],
            "profiles": [
                {
                    "name": "vector-profile",
                    "algorithm": "hnsw-config",
                    "vectorizer": "openai-vectorizer",
                }
            ],
            "vectorizers": [
                {
                    "name": "openai-vectorizer",
                    "kind": "azureOpenAI",
                    "azureOpenAIParameters": {
                        "resourceUri": settings.aoai_endpoint,
                        "deploymentId": settings.aoai_embedding_deployment,
                        "modelName": settings.aoai_embedding_deployment,
                    },
                }
            ],
        },
        "semantic": {
            "configurations": [
                {
                    "name": "semantic-config",
                    "prioritizedFields": {
                        "titleField": {"fieldName": "metadata_spo_item_name"},
                        "prioritizedContentFields": [{"fieldName": "content"}],
                    },
                }
            ]
        },
    }
    put(f"indexes/{settings.search_index_name}", body)
    print(f"Índice '{settings.search_index_name}' creado/actualizado (búsqueda híbrida + semántica).")


def create_skillset() -> None:
    body = {
        "name": settings.search_skillset_name,
        "skills": [
            {
                "@odata.type": "#Microsoft.Skills.Text.AzureOpenAIEmbeddingSkill",
                "name": "embed-content",
                "context": "/document",
                "resourceUri": settings.aoai_endpoint,
                "deploymentId": settings.aoai_embedding_deployment,
                "modelName": settings.aoai_embedding_deployment,
                "dimensions": settings.aoai_embedding_dimensions,
                "inputs": [{"name": "text", "source": "/document/content"}],
                "outputs": [{"name": "embedding", "targetName": "content_vector"}],
            }
        ],
    }
    put(f"skillsets/{settings.search_skillset_name}", body)
    print(f"Skillset '{settings.search_skillset_name}' creado/actualizado.")


def create_indexer() -> None:
    body = {
        "name": settings.search_indexer_name,
        "dataSourceName": settings.search_datasource_name,
        "targetIndexName": settings.search_index_name,
        "skillsetName": settings.search_skillset_name,
        "parameters": {
            "configuration": {
                "indexedFileNameExtensions": ".pdf,.docx,.doc,.pptx,.xlsx,.txt",
                "dataToExtract": "contentAndMetadata",
                "failOnUnsupportedContentType": False,
            }
        },
        # Reindexado automático cada 5 minutos (incremental: solo detecta cambios/nuevos/borrados)
        "schedule": {"interval": "PT5M"},
        "fieldMappings": [
            {
                "sourceFieldName": "metadata_spo_site_library_item_id",
                "targetFieldName": "id",
                "mappingFunction": {"name": "base64Encode"},
            },
            {"sourceFieldName": "metadata_spo_item_name", "targetFieldName": "title"},
            {"sourceFieldName": "metadata_spo_item_weburi", "targetFieldName": "url"},
        ],
        "outputFieldMappings": [
            {"sourceFieldName": "/document/content_vector", "targetFieldName": "content_vector"},
        ],
    }
    put(f"indexers/{settings.search_indexer_name}", body)
    print(f"Indexer '{settings.search_indexer_name}' creado/actualizado (schedule cada 5 min).")


def run_now() -> None:
    post(f"indexers/{settings.search_indexer_name}/run")
    print("Ejecución manual del indexer disparada.")


if __name__ == "__main__":
    create_datasource()
    create_index()
    create_skillset()
    create_indexer()
    run_now()
