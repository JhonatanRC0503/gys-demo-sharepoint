"""Actualiza el agente ya creado en Foundry ("agent-sharepoint") para que use el índice
de Azure AI Search como herramienta de conocimiento, con búsqueda híbrida (texto + vector).

Requiere que exista una conexión de proyecto hacia el servicio de Azure AI Search. Si no
existe, se crea automáticamente (keyless, vía RBAC ya asignado al hub de Foundry).

Ejecutar: python -m backend.agent.create_agent
"""
from azure.ai.projects import AIProjectClient
from azure.ai.projects.models import (
    AISearchIndexResource,
    AzureAISearchQueryType,
    AzureAISearchTool,
    AzureAISearchToolResource,
    PromptAgentDefinition,
)
from azure.identity import AzureCliCredential

from ..common.config import settings

INSTRUCTIONS = """Eres un asistente corporativo especializado en documentación interna
almacenada en SharePoint. Tu objetivo principal es responder las dudas de los usuarios
utilizando exclusivamente la información obtenida a través de la herramienta de Azure AI
Search (que refleja el contenido de la biblioteca de SharePoint).

Reglas:
- Si la respuesta no está en los documentos indexados, dilo explícitamente; no inventes.
- Cita siempre el documento fuente usando el formato [message_idx:search_idx†source].
- Sé conciso y preciso, priorizando la información más reciente si hay conflictos.
"""


def ensure_search_connection(project: AIProjectClient):
    try:
        return project.connections.get(settings.foundry_search_connection_name)
    except Exception:
        print(
            f"Conexión '{settings.foundry_search_connection_name}' no encontrada. "
            "Créala primero (ver README / doc: Foundry portal > Management center > "
            "Connected resources > Azure AI Search, autenticación keyless/AAD)."
        )
        raise


def main() -> None:
    credential = AzureCliCredential()
    project = AIProjectClient(endpoint=settings.foundry_project_endpoint, credential=credential)

    connection = ensure_search_connection(project)

    agent = project.agents.create_version(
        agent_name=settings.foundry_agent_name,
        definition=PromptAgentDefinition(
            model=settings.foundry_model_deployment,
            instructions=INSTRUCTIONS,
            tools=[
                AzureAISearchTool(
                    azure_ai_search=AzureAISearchToolResource(
                        indexes=[
                            AISearchIndexResource(
                                project_connection_id=connection.id,
                                index_name=settings.search_index_name,
                                # Híbrido (texto + vector). Cambia a VECTOR_SEMANTIC_HYBRID
                                # para sumar el re-ranker semántico (ya habilitado en el índice).
                                query_type=AzureAISearchQueryType.VECTOR_SIMPLE_HYBRID,
                            )
                        ]
                    )
                )
            ],
        ),
        description="Agente lector documentario de SharePoint (búsqueda híbrida).",
    )
    print(f"Agente actualizado: id={agent.id} name={agent.name} version={agent.version}")


if __name__ == "__main__":
    main()
