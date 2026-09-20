"""API FastAPI del proyecto (SharePoint + Azure AI Search).

La indexación se hace con prepdocslib: descarga directo de SharePoint (Microsoft Graph),
parsing/chunking con Azure AI Document Intelligence y embeddings con Azure OpenAI. Sin
Azure Blob Storage: cada chunk referencia la URL del documento en SharePoint.

Además de disparar la sincronización manualmente, el propio proceso corre un loop en
background que revisa SharePoint cada INDEX_SYNC_INTERVAL_MINUTES (default 5) y solo
procesa archivos nuevos/modificados/borrados (incremental, vía Graph delta).

Endpoints:
  GET  /health
  POST /index/sharepoint         -> dispara una sincronización manual (bloquea si ya hay una corriendo)
  GET  /index/sharepoint/status  -> estado de la última corrida (manual o automática)

Ejecutar: uvicorn backend.api.main:app --reload --port 8000
"""
import asyncio
import logging
from contextlib import asynccontextmanager
from datetime import datetime, timezone

from fastapi import FastAPI, HTTPException

from ..common.config import settings
from ..scripts.index_sharepoint import run_sharepoint_sync

logger = logging.getLogger("scripts")

_sync_lock = asyncio.Lock()
_last_sync: dict = {"started_at": None, "finished_at": None, "success": None, "result": None, "error": None}


async def _run_sync_locked(trigger: str) -> dict:
    if _sync_lock.locked():
        raise HTTPException(status_code=409, detail="Ya hay una sincronización en curso.")
    async with _sync_lock:
        _last_sync["started_at"] = datetime.now(timezone.utc).isoformat()
        _last_sync["trigger"] = trigger
        try:
            result = await run_sharepoint_sync()
            _last_sync.update(success=True, result=result, error=None)
        except Exception as exc:  # noqa: BLE001 - se registra y se refleja en /status
            logger.exception("Fall\u00f3 la sincronizaci\u00f3n de SharePoint (trigger=%s)", trigger)
            _last_sync.update(success=False, result=None, error=str(exc))
        finally:
            _last_sync["finished_at"] = datetime.now(timezone.utc).isoformat()
        return _last_sync


async def _background_sync_loop() -> None:
    interval_seconds = settings.index_sync_interval_minutes * 60
    while True:
        await asyncio.sleep(interval_seconds)
        if _sync_lock.locked():
            logger.info("Se omite la corrida programada: ya hay una sincronizaci\u00f3n en curso.")
            continue
        await _run_sync_locked(trigger="schedule")


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(_background_sync_loop())
    try:
        yield
    finally:
        task.cancel()


app = FastAPI(title="SharePoint Indexer API", version="1.0.0", lifespan=lifespan)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/index/sharepoint")
async def index_sharepoint():
    """Revisa SharePoint por archivos nuevos/modificados/borrados y sincroniza el índice."""
    result = await _run_sync_locked(trigger="manual")
    return {"success": result["success"], **(result["result"] or {}), "error": result["error"]}


@app.get("/index/sharepoint/status")
def index_sharepoint_status():
    """Estado de la última sincronización (manual o disparada por el schedule automático)."""
    return {
        "sync_interval_minutes": settings.index_sync_interval_minutes,
        "running": _sync_lock.locked(),
        **_last_sync,
    }
