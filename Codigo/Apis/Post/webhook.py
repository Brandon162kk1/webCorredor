import requests
import logging
import os
import base64
#from textwrap import dedent
from jinja2 import Environment, FileSystemLoader

# --- Variables de Entorno ---
service_n8n = os.getenv("service_n8n")

webhook_wsp = os.getenv("webhook_wsp")
url_n8n_wsp = f"{service_n8n}{webhook_wsp}"

# url_n8n_base = os.getenv("url_n8n_base")
# puerto_n8n = os.getenv("puerto_n8n")

# if puerto_n8n:
#     url_n8n_base = f"{url_n8n_base}:{puerto_n8n}"

webhook_correo = os.getenv("webhook_correo")
url_n8n_correo = f"{service_n8n}{webhook_correo}"

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

def enviar_msj_wsp(ramo,tipo,archivo=None,proceso=None,cliente=None):

    logging.info("-----------------------------")
    logging.info(f"⌛ Enviando Notificación por WhatsApp")

    # Intentar usar el celular del ejecutivo
    celular = ramo.celular if ramo.celular else os.getenv("celular_emergencia")
    
    if not celular or str(celular).strip().lower() == "none":
        logging.error("⚠️ No se pudo enviar WhatsApp: Teléfono del ejecutivo no está definido")
        return

    telefono = str(celular).strip()

    if not telefono.startswith("51"):
        telefono = "51" + telefono

    payload = {
        "instancia": f"{os.getenv('instancia')}",
        "telefono": telefono
    }

    if tipo == "notificacion":
        payload["tipo"] = f"sendText"
        payload["mensaje"] = f"""Reenviar el codigo de 6 digitos para continuar con la {proceso.lower()} de {cliente}"""

    elif tipo == "documento":

        if not archivo:
            logging.error("⚠️ No se recibió el archivo")
            return

        if not os.path.exists(archivo):
            logging.error(f"❌ No existe el archivo: {archivo}")
            return


        try:
            with open(archivo, "rb") as f:
                archivo_base64 = base64.b64encode(f.read()).decode("utf-8")

            payload["archivo"] = archivo_base64
            payload["nombreArchivo"] = os.path.basename(archivo)
            payload["mimetype"] = "application/pdf"
            payload["tipo"] = f"sendMedia"

            payload["mensaje"] = f"""📋 ."""

        except Exception as e:
            logging.error(f"❌ Error convirtiendo .jpeg a Base64: {e}")
            return

    else:
        logging.error(f"❌ Tipo de mensaje no soportado: {tipo}")
        return

    try:
        response = requests.post(url_n8n_wsp,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Notificación enviada por Evolution API a {telefono}")
        else:
            logging.error(f"⚠️ Problemas en el envio de notificación a Evolution API - {response.status_code} - {response.text}")

    except Exception as e:
        logging.error(f"❌ Error enviando la notificación por el webhook, Motivo : {e}")
