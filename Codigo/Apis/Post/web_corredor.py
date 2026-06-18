import requests
import logging
import os

from LinuxDebian.Carpetas.rutas import obtener_imagenes_error_para_nota

# --- Variables de Entorno ---
API_KEY = os.getenv("API_KEY")
API_BASE_URL = os.getenv("API_BASE_URL") 
PUERTO_HOST = os.getenv("PUERTO_HOST",API_BASE_URL)

#-- Header Global para autenticación (si es necesario) --
headers = {
    "X-Api-Key": API_KEY
}

def enviar_nota_movimiento(id_movimiento,detalle_error,correo,ruta_carpeta,const):

    url = f"{API_BASE_URL}/api/movimiento/registrar-nota"

    imagenes = obtener_imagenes_error_para_nota(ruta_carpeta, const)

    payload = {
        "MovimientoId": id_movimiento,
        "Contenido": str(detalle_error),
        "UsuarioNombre": correo
    }

    if imagenes:
        payload["ImagenBase64"] = imagenes

    try:
        response = requests.post(url,json=payload,headers=headers,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Nota registrada correctamente para el movimiento {id_movimiento}")
        else:
            logging.error(f"❌ Problemas en el registro de notas | Movimiento {id_movimiento} " f"| Status {response.status_code} | Resp {response.text}")

    except Exception as e:
        logging.error(f"❌ Error conectando al API para registrar notas | Movimiento {id_movimiento} | Motivo: {e}")
