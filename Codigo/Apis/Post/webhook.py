import requests
import logging
import os
#from textwrap import dedent
from jinja2 import Environment, FileSystemLoader

# --- Variables de Entorno ---
url_n8n_base = os.getenv("url_n8n_base")
puerto_n8n = os.getenv("puerto_n8n")

if puerto_n8n:
    url_n8n_base = f"{url_n8n_base}:{puerto_n8n}"

webhook_correo = os.getenv("webhook_correo")

url_n8n_correo = f"{url_n8n_base}{webhook_correo}"

para_venv = os.getenv("para_wc")
para_lista = para_venv.split(",") if para_venv else []
copia_venv = os.getenv("copia_wc")
copias_lista = copia_venv.split(",") if copia_venv else []

ruta_plantilla = "/app/Codigo/Plantillas/Correo"
env = Environment(loader=FileSystemLoader(ruta_plantilla))

def enviar_error_general(ctx,palabra_clave,detalle_ramos):

    logging.info("-----------------------------")

    template = env.get_template("error.html")

    html = template.render(
        titulo=f"⚠ Problemas en la {palabra_clave}",
        cliente=ctx.cliente,
        ruc=ctx.ruc,
        detalle_ramos=detalle_ramos
    )

    polizas = ", ".join(
        str(x["poliza"])
        for x in detalle_ramos
    )

    payload = {
        "Para": para_lista,
        "Copia": copias_lista,
        "Asunto": f"Error en la {palabra_clave} - Pólizas: {polizas}",
        "Mensaje": html
    }

    try:
        response = requests.post(url_n8n_correo,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Notificación enviada al equipo Jishu")
        else:
            logging.error(f"❌ Problemas en el envio de notificación al equipo Jishu - {response.status_code} - {response.text}")

    except Exception as e:
        logging.error(f"❌ Error enviando la notificación por el webhook, Motivo : {e}")
