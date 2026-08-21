import requests
import time
import logging

def codigo_compania(url, api_key, max_retries=15):

    headers = {
        "x-api-key": f"{api_key}"
    }

    intentos = 0
    while intentos < max_retries:
        try:
            resp = requests.get(f"{url}", headers=headers, timeout=10)

            if resp.status_code == 200:
                codigo = resp.json()["codigo"]
                logging.info(f"✅ Código recibido: {codigo}")
                return codigo
            else:
                logging.warning(f"⚠️ Intento {intentos + 1}/{max_retries} - Código de estado inesperado: {resp.status_code}")
        except Exception as e:
            logging.warning(f"⚠️ Intento {intentos + 1}/{max_retries} - Error de conexión: {e}")

        intentos += 1
        time.sleep(2)

    raise RuntimeError(f"❌ Se excedieron los {max_retries} reintentos para obtener el código de la compañía desde: {url}")
