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
    search_endpoint = _require("SEARCH_ENDPOINT")
    search_api_version = os.environ.get("SEARCH_API_VERSION", "2026-08-01-preview")
    search_datasource_name = os.environ.get("SEARCH_DATASOURCE_NAME", "sharepoint-demo-datasource")
    search_index_name = os.environ.get("SEARCH_INDEX_NAME", "idx-sharepoint-demo")
    search_skillset_name = os.environ.get("SEARCH_SKILLSET_NAME", "sharepoint-demo-skillset")
    search_indexer_name = os.environ.get("SEARCH_INDEXER_NAME", "sharepoint-demo-indexer")
    search_admin_key = os.environ.get("SEARCH_ADMIN_KEY", "")

    # Azure OpenAI (embeddings)
    aoai_endpoint = _require("AZURE_OPENAI_ENDPOINT")
    aoai_embedding_deployment = os.environ.get("AZURE_OPENAI_EMBEDDING_DEPLOYMENT", "text-embedding-3-small")
    aoai_embedding_dimensions = int(os.environ.get("AZURE_OPENAI_EMBEDDING_DIMENSIONS", "1536"))

    # Foundry
    foundry_project_endpoint = _require("FOUNDRY_PROJECT_ENDPOINT")
    foundry_model_deployment = os.environ.get("FOUNDRY_MODEL_DEPLOYMENT", "gpt-4.1-mini")
    foundry_agent_name = os.environ.get("FOUNDRY_AGENT_NAME", "agent-sharepoint")
    foundry_search_connection_name = _require("FOUNDRY_SEARCH_CONNECTION_NAME")

    # Webhook / API
    graph_webhook_client_state = os.environ.get("GRAPH_WEBHOOK_CLIENT_STATE", "")
    public_webhook_base_url = os.environ.get("PUBLIC_WEBHOOK_BASE_URL", "")


settings = Settings()
