"""Cliente REST reutilizable para Azure AI Search (control plane: datasources, index, skillset, indexer).

Usa Microsoft Entra (AzureCliCredential) para autenticar, evitando manejar admin keys.
Si SEARCH_ADMIN_KEY está definido en el entorno, se usa como alternativa (api-key header).
"""
import requests
from azure.identity import AzureCliCredential

from .config import settings

_credential = AzureCliCredential()
_SCOPE = "https://search.azure.com/.default"


def _headers() -> dict:
    if settings.search_admin_key:
        return {"Content-Type": "application/json", "api-key": settings.search_admin_key}
    token = _credential.get_token(_SCOPE).token
    return {"Content-Type": "application/json", "Authorization": f"Bearer {token}"}


def _url(path: str) -> str:
    separator = "&" if "?" in path else "?"
    return f"{settings.search_endpoint}/{path}{separator}api-version={settings.search_api_version}"


def put(path: str, body: dict) -> requests.Response:
    resp = requests.put(_url(path), headers=_headers(), json=body, timeout=60)
    if not resp.ok:
        raise RuntimeError(f"PUT {path} failed [{resp.status_code}]: {resp.text}")
    return resp


def post(path: str, body: dict | None = None) -> requests.Response:
    resp = requests.post(_url(path), headers=_headers(), json=body or {}, timeout=60)
    if not resp.ok:
        raise RuntimeError(f"POST {path} failed [{resp.status_code}]: {resp.text}")
    return resp


def get(path: str) -> requests.Response:
    resp = requests.get(_url(path), headers=_headers(), timeout=60)
    if not resp.ok:
        raise RuntimeError(f"GET {path} failed [{resp.status_code}]: {resp.text}")
    return resp
