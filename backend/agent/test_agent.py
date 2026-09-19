"""Prueba rápida por consola: envía una pregunta al agente "agent-sharepoint" y
muestra la respuesta con sus citas (fuente SharePoint).

Ejecutar: python -m backend.agent.test_agent "¿Qué dice el documento de Carolina Arias Mejia?"
"""
import sys

from azure.ai.projects import AIProjectClient
from azure.identity import AzureCliCredential

from ..common.config import settings


def main() -> None:
    question = " ".join(sys.argv[1:]) or "¿Qué documentos tienes disponibles?"

    credential = AzureCliCredential()
    project = AIProjectClient(endpoint=settings.foundry_project_endpoint, credential=credential)
    openai_client = project.get_openai_client()

    response = openai_client.responses.create(
        input=question,
        tool_choice="required",
        extra_body={"agent_reference": {"name": settings.foundry_agent_name, "type": "agent_reference"}},
    )

    print("Respuesta:\n", response.output_text)
    for item in response.output:
        if getattr(item, "type", None) == "message":
            for content in item.content:
                for annotation in getattr(content, "annotations", None) or []:
                    if getattr(annotation, "type", None) == "url_citation":
                        print(f"Cita: {annotation.url}")


if __name__ == "__main__":
    main()
