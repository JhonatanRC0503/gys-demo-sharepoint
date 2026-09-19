"""Crea (o renueva) una suscripción de Microsoft Graph a cambios en la biblioteca de
SharePoint, para notificar a nuestro webhook (api/main.py) casi en tiempo real y así
disparar un reindexado incremental sin esperar al schedule de 5 minutos.

Requisitos:
  - PUBLIC_WEBHOOK_BASE_URL debe ser una URL HTTPS públicamente alcanzable (Graph la valida
    en el momento de crear la suscripción). Para probar en local, exponer con ngrok/devtunnel:
        devtunnel host -p 8000 --allow-anonymous
    o
        ngrok http 8000
    y usar esa URL pública como PUBLIC_WEBHOOK_BASE_URL.
  - La app de Entra ya tiene Files.Read.All + Sites.Read.All (Application), suficientes
    para suscribirse a cambios del drive (biblioteca) vía client credentials.
  - Las suscripciones de Graph expiran (para drives, máx. ~4230 minutos ~ 2.9 días). Hay que
    renovarlas antes de que expiren (ver renew_graph_subscription mas abajo) con un cron/job.

Ejecutar: python -m backend.scripts.create_graph_subscription
"""
import secrets
import urllib.parse
from datetime import datetime, timedelta, timezone

import requests
from azure.identity import ClientSecretCredential

from ..common.config import settings

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
MAX_EXPIRATION_MINUTES = 4230  # límite de Graph para el recurso "drive"


def _get_token() -> str:
    credential = ClientSecretCredential(
        tenant_id=settings.tenant_id,
        client_id=settings.client_id,
        client_secret=settings.client_secret,
    )
    return credential.get_token(GRAPH_SCOPE).token


def _get_site_id(token: str) -> str:
    parsed = urllib.parse.urlparse(settings.sharepoint_site_url)
    site_path = parsed.path.lstrip("/")  # ej: sites/Pruebascofide
    url = f"https://graph.microsoft.com/v1.0/sites/{parsed.netloc}:/{site_path}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    return resp.json()["id"]


def _get_default_drive_id(token: str, site_id: str) -> str:
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    resp.raise_for_status()
    return resp.json()["id"]


def create_subscription() -> dict:
    if not settings.public_webhook_base_url:
        raise RuntimeError(
            "Define PUBLIC_WEBHOOK_BASE_URL en .env con una URL HTTPS pública "
            "(ngrok/devtunnel/Azure) antes de crear la suscripción."
        )
    if not settings.client_secret:
        raise RuntimeError("Define AZURE_CLIENT_SECRET en .env (client credentials flow).")

    token = _get_token()
    site_id = _get_site_id(token)
    drive_id = _get_default_drive_id(token, site_id)

    client_state = settings.graph_webhook_client_state or secrets.token_urlsafe(32)
    expiration = datetime.now(timezone.utc) + timedelta(minutes=MAX_EXPIRATION_MINUTES)

    body = {
        "changeType": "updated",
        "notificationUrl": f"{settings.public_webhook_base_url.rstrip('/')}/webhooks/sharepoint",
        "resource": f"/drives/{drive_id}/root",
        "expirationDateTime": expiration.isoformat(),
        "clientState": client_state,
    }
    resp = requests.post(
        "https://graph.microsoft.com/v1.0/subscriptions",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
        timeout=30,
    )
    if not resp.ok:
        raise RuntimeError(f"No se pudo crear la suscripción: {resp.status_code} {resp.text}")

    data = resp.json()
    print(f"Suscripción creada: id={data['id']} expira={data['expirationDateTime']}")
    if not settings.graph_webhook_client_state:
        print(
            f"IMPORTANTE: guarda este clientState en tu .env como GRAPH_WEBHOOK_CLIENT_STATE "
            f"para validar notificaciones futuras:\n  {client_state}"
        )
    return data


if __name__ == "__main__":
    create_subscription()
