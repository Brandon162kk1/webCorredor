import requests
import logging
import base64
import os

# --- Variables de Entorno ---
url_n8n_error = os.getenv("url_n8n_error")

def enviar_error_consolidado(cliente, proceso, detalle_final, ruta_carpeta):

    imagenes_payload = []

    try:
        # 🔎 Buscar solo archivos que empiecen con ERROR_ y sean .png
        for archivo in os.listdir(ruta_carpeta):

            if archivo.startswith("ERROR_") and archivo.lower().endswith(".png"):

                ruta_completa = os.path.join(ruta_carpeta, archivo)

                with open(ruta_completa, "rb") as f:
                    imagen_base64 = base64.b64encode(f.read()).decode("utf-8")

                imagenes_payload.append({
                    "imagen_nombre": archivo,
                    "imagen_base64": imagen_base64
                })

    except Exception as e:
        logging.error(f"❌ Error leyendo imágenes de la carpeta: {e}")
        return

    # 📦 Payload consolidado
    payload = {
        "Cliente": cliente,
        "Proceso": proceso,
        "Tipo de Error": "Errores Por Ramo",
        "Detalle de Error": detalle_final,
        "Imagenes": imagenes_payload if imagenes_payload else None
    }

    try:
        response = requests.post(url_n8n_error, json=payload, timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info("✅ Notificación consolidada enviada al equipo Jishu")
        else:
            logging.error(
                f"❌ Problemas en el envío al equipo Jishu "
                f"- {response.status_code} - {response.text}"
            )

    except Exception as e:
        logging.error(
            f"❌ Error enviando la notificación por el webhook. Motivo: {e}"
        )

def enviar_error_interno(cliente,proceso,ctx_ramo,palabra_clave,error,detalle_error,nombreImagen):
    
    with open(nombreImagen, "rb") as f:
        imagen_base64 = base64.b64encode(f.read()).decode("utf-8")

    tramas = []

    if ctx_ramo.trama:
        tramas.append(ctx_ramo.trama)

    if ctx_ramo.trama_97:
        tramas.append(ctx_ramo.trama_97)

    payload = {

        "Cliente": cliente,
        "Poliza": ctx_ramo.poliza,
        "Fecha Inicio": ctx_ramo.f_inicio,
        "Fecha Fin ": ctx_ramo.f_fin,
        "Trama": tramas if tramas else None,
        "Sede": ctx_ramo.sede,
        "Compania": ctx_ramo.compania,
        "Tipo de Error": error,
        "Detalle de Error": str(detalle_error),
        "Proceso": proceso,
        "Palabra Clave": palabra_clave,
        "imagen_nombre": nombreImagen,
        "imagen_base64": imagen_base64
    }

    try:
        response = requests.post(url_n8n_error,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Notificación enviada al equipo Jishu")
        else:
            logging.error(f"❌ Problemas en el envio de notificacion al equipo Jishu - {response.status_code} - {response.text}")

    except Exception as e:
        logging.error(f"❌ Error enviando la notificacion por el webhook, Motivo : {e}")