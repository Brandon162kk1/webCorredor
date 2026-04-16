import requests
import logging
import os
import json
from dotenv import load_dotenv
from LinuxDebian.Carpetas.rutas import obtener_imagenes_error

load_dotenv("/app/variables.env")

# --- Variables de Entorno ---
API_KEY = os.getenv("API_KEY")
API_BASE_URL = os.getenv("API_BASE_URL") 
PUERTO_HOST = os.getenv("PUERTO_HOST",API_BASE_URL)

#-- Header Global para autenticación (si es necesario) --
headers = {
    "X-Api-Key": API_KEY
}

def enviar_aviso_captcha(ramo):
    url = f"{API_BASE_URL}/api/plantilla-captcha/movimiento/{ramo.id_poliza}"
    payload = {
      "captchaRequerido": True
    }
    try:
        response = requests.put(url,json=payload,headers=headers,timeout=30)
        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Aviso de captcha enviado correctamente al Movimiento {ramo.id_poliza}")
            return True
        else:
            logging.error(f"❌ Error enviando aviso de captcha " f"| Status {response.status_code} | Resp {response.text}")
            return False
    except Exception as e:
        logging.error(f"❌ Error conectando al API para enviar aviso de captcha | {e}")
        return False

def enviar_puerto(id_movimiento,puerto):


    url = f"{API_BASE_URL}/api/movimiento/{id_movimiento}/puerto-host"

    payload = {
        "puertoHost": f"{PUERTO_HOST}:{puerto}"
    }

    try:
        response = requests.put(url,json=payload,headers=headers,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Puerto enviado registrado correctamente | Movimiento {id_movimiento}")
            return True
        else:
            logging.error(f"❌ Error registrando puerto | Movimiento {id_movimiento} " f"| Status {response.status_code} | Resp {response.text}")
            return False

    except Exception as e:
        logging.error(f"❌ Error conectando al API para registrar puerto | Movimiento {id_movimiento} | {e}")
        return False
 
def enviar_estaca(id_movimiento, ramo, afirmacion_constancia,afirmacion_proforma):

    url = f"{API_BASE_URL}/api/movimiento/{id_movimiento}/estaca-ramo"

    payload = {
        "ramo": ramo,
        "estaca": afirmacion_constancia,
        "estacaProforma": afirmacion_proforma
    }

    try:
        response = requests.put(url,json=payload,headers=headers,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Registro actualizado correctamente | Movimiento {id_movimiento} | Ramo {ramo}")
        else:
            logging.error(f"❌ Error actualizando registro | Movimiento {id_movimiento} " f"| Status {response.status_code} | Resp {response.text}")

    except Exception as e:
        logging.error(f"❌ Error conectando al API para actualizar | Movimiento {id_movimiento} | Ramo {ramo} | {e}")

def enviar_error_movimiento(id_movimiento, ramo, error, detalle_error,ruta_carpeta,const):

    url = f"{API_BASE_URL}/api/movimiento/{id_movimiento}/error"

    imagenes = obtener_imagenes_error(ruta_carpeta, const)

    payload = {
        "ramo": ramo,
        "error": error,
        "detalleError": str(detalle_error)
    }

    if imagenes:
        payload["imagenes"] = imagenes

    try:
        response = requests.put(url,json=payload,headers=headers,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Error registrado correctamente | Movimiento {id_movimiento} | Ramo {ramo}")
        else:
            logging.error(f"❌ Problemas en el registro de error | Movimiento {id_movimiento} " f"| Status {response.status_code} | Resp {response.text}")

    except Exception as e:
        logging.error(f"❌ Error conectando al API para registrar 'errores' | Movimiento {id_movimiento} | Ramo {ramo} | {e}")

def enviar_documentos(id_movimiento,ruta_pdf,ramo,tipoDocumento):
   
    url = f"{API_BASE_URL}/api/movimiento/{id_movimiento}/subir-documentos"

    try:

        nombre_archivo_form = "documento_0"

        documentos_json = [
            {
                "tipoDocumento": tipoDocumento,
                "ramo": ramo.capitalize(), 
                "nombreArchivo": nombre_archivo_form
            }
        ]

        with open(ruta_pdf, "rb") as f:
            files = {
                nombre_archivo_form: (
                    os.path.basename(ruta_pdf),
                    f,
                    "application/pdf"
                )
            }

            data = {
                "documentosJson": json.dumps(documentos_json)
            }

            response = requests.put(
                url,
                files=files,
                headers=headers,
                data=data,
                timeout=30
            )

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ {tipoDocumento} enviado correctamente | Movimiento {id_movimiento}")
        else:
            logging.error(
                f"❌ Error enviando {tipoDocumento} | Movimiento {id_movimiento} "
                f"| Status {response.status_code} | Resp {response.text}"
            )

    except Exception as e:
        logging.error(f"❌ Error enviando documento | Movimiento {id_movimiento} | {e}")
