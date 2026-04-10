import requests
import logging
import os
from LinuxDebian.Carpetas.rutas import obtener_imagenes_error

# --- Variables de Entorno ---
url_n8n_error = os.getenv("url_n8n_error")
url_n8n_anular_pdf = os.getenv("url_n8n_anular_pdf")

def enviar_error_interno(cliente,proceso,ctx_ramo,palabra_clave,error,detalle_error,ruta_carpeta,const):
    
    imagenes = obtener_imagenes_error(ruta_carpeta, const)

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
        "Trama": ' y '.join(tramas) if tramas else None,
        "Sede": ctx_ramo.sede,
        "Compania": ctx_ramo.compania,
        "Tipo de Error": error,
        "Detalle de Error": str(detalle_error),
        "Proceso": proceso,
        "Palabra Clave": palabra_clave,
        "imagen_nombre": f"Error_{const}.png",
        "imagen_base64": imagenes[0] if imagenes else None
    }

    try:
        response = requests.post(url_n8n_error,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Notificación enviada al equipo Jishu")
        else:
            logging.error(f"❌ Problemas en el envio de notificación al equipo Jishu - {response.status_code} - {response.text}")

    except Exception as e:
        logging.error(f"❌ Error enviando la notificación por el webhook, Motivo : {e}")
