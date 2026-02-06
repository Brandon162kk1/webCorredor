import requests
import logging
import os
import json

#base_url = "https://webcorredor.azurewebsites.net/api/movimiento/"
base_url = "http://192.168.0.201:64328/api/movimiento/"

def enviar_puerto(id_movimiento,puerto):


    url = f"{base_url}{id_movimiento}/puerto-host"

    payload = {
        #"puertoHost": f"https://webcorredor.azurewebsites.net:{puerto}"
        "puertoHost": f"http://192.168.0.201:{puerto}"
    }

    try:
        response = requests.put(url,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Puerto enviado registrado correctamente | Movimiento {id_movimiento}")
            return True
        else:
            logging.error(f"❌ Error registrando puerto | Movimiento {id_movimiento} " f"| Status {response.status_code} | Resp {response.text}")
            return False

    except Exception as e:
        #logging.exception(f"❌ Error conectando al API para registrar puerto | Movimiento {id_movimiento}")
        logging.error(f"❌ Error conectando al API para registrar puerto | Movimiento {id_movimiento} | {e}")
        return False
 
def enviar_estaca(id_movimiento, ramo, afirmacion_constancia,afirmacion_proforma):
    url = f"{base_url}{id_movimiento}/estaca-ramo"

    payload = {
        "ramo": ramo,
        "estaca": afirmacion_constancia,
        "estacaProforma": afirmacion_proforma
    }

    try:
        response = requests.put(url,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Registro actualizado correctamente | Movimiento {id_movimiento} | Ramo {ramo}")
            return True
        else:
            logging.error(f"❌ Error actualizando registro | Movimiento {id_movimiento} " f"| Status {response.status_code} | Resp {response.text}")
            return False

    except Exception as e:
        #logging.exception(f"❌ Error conectando al API para registrar estaca | Movimiento {id_movimiento} | Ramo {ramo}")
        logging.error(f"❌ Error conectando al API para actualizar | Movimiento {id_movimiento} | Ramo {ramo} | {e}")
        return False

def enviar_error_movimiento(id_movimiento, ramo, error, detalle_error):

    url = f"{base_url}{id_movimiento}/error"

    payload = {
        "ramo": ramo,
        "error": error,
        "detalleError": str(detalle_error)
    }

    try:
        response = requests.put(url,json=payload,timeout=30)

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ Error registrado correctamente | Movimiento {id_movimiento} | Ramo {ramo}")
            return True
        else:
            logging.error(f"❌ Problemas en el registro de error | Movimiento {id_movimiento} " f"| Status {response.status_code} | Resp {response.text}")
            return False

    except Exception as e:
        #logging.exception(f"❌ Error conectando al API para registrar 'errores' | Movimiento {id_movimiento} | Ramo {ramo}")
        logging.error(f"❌ Error conectando al API para registrar 'errores' | Movimiento {id_movimiento} | Ramo {ramo} | {e}")
        return False

def enviar_documentos(id_movimiento,ruta_pdf,ramo,tipoDocumento):
   
    url = f"{base_url}{id_movimiento}/subir-documentos"

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
                data=data,
                timeout=30
            )

        if response.status_code in (200, 201, 204):
            logging.info(f"✅ {tipoDocumento} enviado correctamente | Movimiento {id_movimiento}")
            return True
        else:
            logging.error(
                f"❌ Error enviando {tipoDocumento} | Movimiento {id_movimiento} "
                f"| Status {response.status_code} | Resp {response.text}"
            )
            return False

    except Exception as e:
        #logging.exception(f"❌ Error conectando al API al enviar {tipoDocumento} : {id_movimiento}")
        logging.error(f"❌ Error enviando documento | Movimiento {id_movimiento} | {e}")
        return False
