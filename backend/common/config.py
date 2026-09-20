"""Configuración compartida cargada desde variables de entorno (.env)."""
import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")


def _require(name: str) -> str:
    value = os.environ.get(name, "")
    if not value:
        raise RuntimeError(f"Falta la variable de entorno {name}. Revisa tu archivo .env")
    return value


class Settings:
    # Entra ID app (SharePoint indexer)
    tenant_id = _require("AZURE_TENANT_ID")
    client_id = _require("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET", "")

    # SharePoint
    sharepoint_site_url = _require("SHAREPOINT_SITE_URL")
    sharepoint_tenant_host = _require("SHAREPOINT_TENANT_HOST")
    sharepoint_folder_path = _require("SHAREPOINT_FOLDER_PATH")

    # Azure AI Search
    search_service_name = _require("SEARCH_SERVICE_NAME")
    search_endpoint = _require("SEARCH_ENDPOINT")
    search_index_name = os.environ.get("SEARCH_INDEX_NAME", "idx-sharepoint-demo")
    search_admin_key = os.environ.get("SEARCH_ADMIN_KEY", "")

    # Indexación automática (API en background)
    index_sync_interval_minutes = int(os.environ.get("INDEX_SYNC_INTERVAL_MINUTES", "5"))

    # Azure AI Document Intelligence (parsing/chunking vía prepdocslib)
    document_intelligence_service = _require("AZURE_DOCUMENTINTELLIGENCE_SERVICE")
    document_intelligence_key = os.environ.get("AZURE_DOCUMENTINTELLIGENCE_KEY", "")

    # Azure OpenAI (embeddings)
    aoai_endpoint = _require("AZURE_OPENAI_ENDPOINT")
    aoai_embedding_deployment = os.environ.get("AZURE_OPENAI_EMBEDDING_DEPLOYMENT", "text-embedding-3-large")
    aoai_embedding_dimensions = int(os.environ.get("AZURE_OPENAI_EMBEDDING_DIMENSIONS", "3072"))

    # Foundry
    foundry_project_endpoint = _require("FOUNDRY_PROJECT_ENDPOINT")
    foundry_model_deployment = os.environ.get("FOUNDRY_MODEL_DEPLOYMENT", "gpt-4.1-mini")
    foundry_agent_name = os.environ.get("FOUNDRY_AGENT_NAME", "agent-sharepoint")
    foundry_search_connection_name = _require("FOUNDRY_SEARCH_CONNECTION_NAME")


settings = Settings()
