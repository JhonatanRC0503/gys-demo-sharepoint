"""API FastAPI para disparar/monitorear la indexación de SharePoint y recibir
notificaciones de cambios de Microsoft Graph (near real-time).

Endpoints:
  GET  /health
  POST /reindex           -> dispara una corrida manual del indexer (equivalente al botón "reindexar")
  GET  /reindex/status    -> estado de la última corrida
  POST /webhooks/sharepoint -> receptor de notificaciones de Microsoft Graph (subscriptions)
  GET  /webhooks/sharepoint -> validación de la suscripción (Graph hace un GET con validationToken)

Ejecutar: uvicorn backend.api.main:app --reload --port 8000
"""
from fastapi import FastAPI, Request, Response, HTTPException

from ..common.config import settings
from ..common.search_rest import get, post

app = FastAPI(title="SharePoint Indexer API", version="1.0.0")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/reindex")
def reindex():
    """Dispara manualmente una corrida del indexer (uso: botón de UI, cron externo, etc.)."""
    post(f"indexers/{settings.search_indexer_name}/run")
    return {"triggered": True, "indexer": settings.search_indexer_name}


@app.get("/reindex/status")
def reindex_status():
    resp = get(f"indexers/{settings.search_indexer_name}/status")
    data = resp.json()
    last_result = data.get("lastResult") or {}
    return {
        "status": data.get("status"),
        "lastResultStatus": last_result.get("status"),
        "itemsProcessed": last_result.get("itemsProcessed"),
        "itemsFailed": last_result.get("itemsFailed"),
        "errorMessage": last_result.get("errorMessage"),
        "startTime": last_result.get("startTime"),
        "endTime": last_result.get("endTime"),
    }


@app.get("/webhooks/sharepoint")
def validate_subscription(validationToken: str | None = None):
    """Microsoft Graph valida la suscripción con un GET que incluye ?validationToken=...
    Hay que devolverlo tal cual, como texto plano, con 200 OK."""
    if validationToken is None:
        raise HTTPException(status_code=400, detail="Falta validationToken")
    return Response(content=validationToken, media_type="text/plain", status_code=200)


@app.post("/webhooks/sharepoint")
async def receive_notification(request: Request):
    """Recibe notificaciones de cambio (Microsoft Graph change notifications) sobre la
    biblioteca de SharePoint y dispara una corrida incremental del indexer.

    Seguridad: se valida el 'clientState' compartido para confirmar que la notificación
    viene de la suscripción que creamos (ver scripts/create_graph_subscription.py).
    """
    payload = await request.json()
    notifications = payload.get("value", [])

    for notification in notifications:
        client_state = notification.get("clientState")
        if settings.graph_webhook_client_state and client_state != settings.graph_webhook_client_state:
            # Notificación no confiable: se ignora en vez de disparar reindexado.
            continue

    if notifications:
        try:
            post(f"indexers/{settings.search_indexer_name}/run")
        except RuntimeError as exc:
            # 409 = el indexer ya está corriendo (p.ej. por el schedule); no es un error real.
            if "409" not in str(exc):
                raise

    # Graph requiere 202 Accepted rápido (<10s) para no reintentar/desactivar la suscripción.
    return Response(status_code=202)
