"""Indexa la biblioteca de SharePoint directo hacia Azure AI Search, usando prepdocslib
para parsear (Azure AI Document Intelligence), trocear con overlap (chunking) y generar
embeddings (Azure OpenAI). No se usa Azure Blob Storage: cada chunk referencia
directamente la URL del documento en SharePoint (`sharepointUrl`).

Es incremental: usa Microsoft Graph delta query (estado guardado en
backend/.state/sharepoint_delta.json) para procesar solo archivos nuevos/modificados,
y elimina del índice los que fueron borrados en SharePoint.

Ejecutar: python -m backend.scripts.index_sharepoint
"""
import asyncio
import logging
from pathlib import Path
from urllib.parse import urlparse

from azure.identity.aio import ClientSecretCredential

from ..common.config import settings
from ..prepdocslib.searchmanager import SearchManager
from ..prepdocslib.servicesetup import (
    OpenAIHost,
    build_file_processors,
    setup_embeddings_service,
    setup_openai_client,
    setup_search_info,
)
from ..prepdocslib.sharepointlistfilestrategy import SharePointListFileStrategy
from ..prepdocslib.textprocessor import process_text

logger = logging.getLogger("scripts")

STATE_PATH = Path(__file__).resolve().parent.parent / ".state" / "sharepoint_delta.json"
FIELD_NAME_EMBEDDING = "embedding3"


def _aoai_service_name() -> str:
    """El SDK de prepdocslib arma el endpoint a partir del NOMBRE del recurso, no de la URL completa."""
    return urlparse(settings.aoai_endpoint).hostname.split(".")[0]


async def _build_search_manager(azure_credential: ClientSecretCredential) -> SearchManager:
    search_info = setup_search_info(
        search_service=settings.search_service_name,
        index_name=settings.search_index_name,
        azure_credential=azure_credential,
        search_key=settings.search_admin_key or None,
    )
    openai_client, azure_openai_endpoint = setup_openai_client(
        openai_host=OpenAIHost.AZURE,
        azure_credential=azure_credential,
        azure_openai_service=_aoai_service_name(),
    )
    embeddings = setup_embeddings_service(
        openai_host=OpenAIHost.AZURE,
        open_ai_client=openai_client,
        emb_model_name=settings.aoai_embedding_deployment,
        emb_model_dimensions=settings.aoai_embedding_dimensions,
        azure_openai_deployment=settings.aoai_embedding_deployment,
        azure_openai_endpoint=azure_openai_endpoint,
    )
    return SearchManager(
        search_info=search_info,
        search_analyzer_name=None,
        use_acls=False,
        use_parent_index_projection=False,
        embeddings=embeddings,
        field_name_embedding=FIELD_NAME_EMBEDDING,
        search_images=False,  # solo texto, sin indexar imágenes/figuras
    )


async def run_sharepoint_sync() -> dict:
    """Sincroniza (incremental) la biblioteca de SharePoint con el índice de Azure AI Search."""
    azure_credential = ClientSecretCredential(
        tenant_id=settings.tenant_id, client_id=settings.client_id, client_secret=settings.client_secret
    )
    try:
        search_manager = await _build_search_manager(azure_credential)
        await search_manager.create_index()

        file_processors = build_file_processors(
            azure_credential=azure_credential,
            document_intelligence_service=settings.document_intelligence_service,
            document_intelligence_key=settings.document_intelligence_key or None,
        )

        list_strategy = SharePointListFileStrategy(
            tenant_id=settings.tenant_id,
            client_id=settings.client_id,
            client_secret=settings.client_secret,
            site_url=settings.sharepoint_site_url,
            folder_path=settings.sharepoint_folder_path,
            state_path=STATE_PATH,
        )

        processed = 0
        skipped = 0
        async for file in list_strategy.list():
            try:
                processor = file_processors.get(file.file_extension().lower())
                if processor is None:
                    logger.info("Sin parser para '%s', se omite.", file.filename())
                    skipped += 1
                    continue
                pages = [page async for page in processor.parser.parse(content=file.content)]
                sections = process_text(pages, file, processor.splitter, category=None)
                if sections:
                    # Se borra primero para no dejar chunks obsoletos si el documento cambió de tamaño.
                    await search_manager.remove_content(file.filename())
                    await search_manager.update_content(sections, url=file.url)
                    processed += 1
            finally:
                file.close()

        removed = 0
        for filename in list_strategy.deleted_filenames:
            await search_manager.remove_content(filename)
            removed += 1

        return {"processed": processed, "removed": removed, "skipped": skipped}
    finally:
        await azure_credential.close()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    print(asyncio.run(run_sharepoint_sync()))
