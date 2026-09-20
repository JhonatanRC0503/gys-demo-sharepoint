"""ListFileStrategy que obtiene archivos directamente de una biblioteca de SharePoint
vía Microsoft Graph, sin pasar por Azure Blob Storage. Usa Graph delta query para que,
en corridas posteriores, solo se descarguen/procesen los archivos nuevos o modificados
(y se reporten los eliminados) en vez de recorrer toda la biblioteca cada vez.
"""
import io
import json
import logging
from collections.abc import AsyncGenerator
from pathlib import Path
from typing import Optional
from urllib.parse import unquote, urlparse

import aiohttp
from azure.identity.aio import ClientSecretCredential

from .listfilestrategy import File, ListFileStrategy

logger = logging.getLogger("scripts")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"


class SharePointListFileStrategy(ListFileStrategy):
    """Lista (y descarga en memoria) los archivos de una carpeta de una biblioteca de
    SharePoint, usando Graph delta para procesar solo lo nuevo/modificado en cada corrida.
    Los nombres de archivos borrados desde la última corrida quedan en `deleted_filenames`
    después de agotar el generador de `list()`.
    """

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        site_url: str,
        folder_path: str,
        state_path: Path,
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_url = site_url.rstrip("/")
        self.folder_path = folder_path.strip("/")
        self.state_path = state_path
        self.deleted_filenames: list[str] = []

    def _load_delta_link(self) -> Optional[str]:
        if self.state_path.exists():
            try:
                return json.loads(self.state_path.read_text()).get("delta_link")
            except (json.JSONDecodeError, OSError):
                return None
        return None

    def _save_delta_link(self, delta_link: str) -> None:
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        self.state_path.write_text(json.dumps({"delta_link": delta_link}))

    async def _site_id(self, session: aiohttp.ClientSession) -> str:
        parsed = urlparse(self.site_url)
        site_path = parsed.path.lstrip("/")
        async with session.get(f"{GRAPH_BASE}/sites/{parsed.netloc}:/{site_path}") as resp:
            resp.raise_for_status()
            data = await resp.json()
            return data["id"]

    def _library_and_relative_path(self) -> tuple[Optional[str], str]:
        """SHAREPOINT_FOLDER_PATH puede venir como 'sites/<site>/<biblioteca>/<carpeta>' (formato
        heredado del data source del indexer nativo) o ya relativo como '<biblioteca>/<carpeta>'.
        Devuelve (nombre_de_biblioteca_o_None, ruta_relativa_dentro_de_esa_biblioteca).
        """
        site_prefix = urlparse(self.site_url).path.strip("/")  # ej: sites/Pruebascofide
        remaining = self.folder_path
        if site_prefix and remaining.lower().startswith(f"{site_prefix.lower()}/"):
            remaining = remaining[len(site_prefix) + 1 :]
        if not remaining:
            return None, ""
        library_name, _, relative_path = remaining.partition("/")
        return library_name, relative_path

    async def _drive_id(self, session: aiohttp.ClientSession, site_id: str, library_name: Optional[str]) -> str:
        if library_name:
            async with session.get(f"{GRAPH_BASE}/sites/{site_id}/drives") as resp:
                resp.raise_for_status()
                data = await resp.json()
            target = library_name.lower()
            for drive in data.get("value", []):
                # El "name" para mostrar puede diferir del slug en la URL (ej. "sharepoint-demo" vs
                # "sharepointdemo"); comparamos contra ambos para no depender de que coincidan.
                web_url_slug = unquote(urlparse(drive.get("webUrl", "")).path.rstrip("/").rsplit("/", 1)[-1]).lower()
                display_name = drive.get("name", "").lower()
                if target in (web_url_slug, display_name, display_name.replace(" ", "-"), display_name.replace(" ", "")):
                    return drive["id"]
            logger.warning(
                "No se encontró la biblioteca '%s' entre los drives del sitio; se usa la biblioteca por defecto.",
                library_name,
            )
        async with session.get(f"{GRAPH_BASE}/sites/{site_id}/drive") as resp:
            resp.raise_for_status()
            data = await resp.json()
            return data["id"]

    async def _folder_item_id(self, session: aiohttp.ClientSession, drive_id: str, relative_path: str) -> str:
        path = f"{GRAPH_BASE}/drives/{drive_id}/root" if not relative_path else f"{GRAPH_BASE}/drives/{drive_id}/root:/{relative_path}"
        async with session.get(path) as resp:
            resp.raise_for_status()
            data = await resp.json()
            return data["id"]

    async def _download(self, session: aiohttp.ClientSession, drive_id: str, item_id: str) -> bytes:
        async with session.get(f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content") as resp:
            resp.raise_for_status()
            return await resp.read()

    async def list(self) -> AsyncGenerator[File, None]:
        self.deleted_filenames = []
        credential = ClientSecretCredential(
            tenant_id=self.tenant_id, client_id=self.client_id, client_secret=self.client_secret
        )
        try:
            token = await credential.get_token(GRAPH_SCOPE)
            headers = {"Authorization": f"Bearer {token.token}"}
            async with aiohttp.ClientSession(headers=headers) as session:
                site_id = await self._site_id(session)
                library_name, relative_path = self._library_and_relative_path()
                drive_id = await self._drive_id(session, site_id, library_name)
                folder_item_id = await self._folder_item_id(session, drive_id, relative_path)

                delta_link = self._load_delta_link()
                url = delta_link or f"{GRAPH_BASE}/drives/{drive_id}/items/{folder_item_id}/delta"

                new_delta_link = None
                while url:
                    async with session.get(url) as resp:
                        resp.raise_for_status()
                        data = await resp.json()
                    for item in data.get("value", []):
                        if "file" not in item:
                            continue  # carpetas: sin contenido que indexar
                        if item.get("deleted"):
                            name = item.get("name")
                            if name:
                                self.deleted_filenames.append(name)
                            else:
                                logger.warning("Item eliminado sin nombre en delta, no se puede quitar del índice.")
                            continue
                        logger.info("Descargando '%s' desde SharePoint", item["name"])
                        content = await self._download(session, drive_id, item["id"])
                        file_io = io.BytesIO(content)
                        file_io.name = item["name"]
                        yield File(content=file_io, acls={}, url=item.get("webUrl"))
                    url = data.get("@odata.nextLink")
                    if "@odata.deltaLink" in data:
                        new_delta_link = data["@odata.deltaLink"]

                if new_delta_link:
                    self._save_delta_link(new_delta_link)
        finally:
            await credential.close()
