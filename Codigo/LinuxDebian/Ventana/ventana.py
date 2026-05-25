# -- Imports --
import time
import subprocess
import logging
import os
#-- Froms --
from Apis.Put.web_corredor import enviar_puerto

#-------- Interacción Humana --------
def bloquear_interaccion():
    subprocess.run(["x11vnc", "-remote", "viewonly"], check=False)
    logging.info("🔒 Interacción humana BLOQUEADA (VNC view-only)")

def desbloquear_interaccion():
    subprocess.run(["x11vnc", "-remote", "noviewonly"], check=False)
    logging.info("✋ Interacción humana HABILITADA")

#-------- Esperas Avanzadas --------
def esperar_archivos_nuevos(directorio, archivos_antes, extension, cantidad, timeout=180):
    """
    Espera archivos nuevos con determinada extensión.
    extension: ".zip", ".pdf", ".xlsx", etc.
    """

    inicio = time.time()

    while time.time() - inicio < timeout:
        actuales = set(os.listdir(directorio))
        nuevos = actuales - archivos_antes

        # Filtrar por extensión (case insensitive)
        nuevos = {
            f for f in nuevos
            if f.lower().endswith(extension.lower())
        }

        if len(nuevos) >= cantidad:

            # Validar que no estén en descarga (.crdownload)
            archivos_validos = []
            for f in nuevos:
                ruta = os.path.join(directorio, f)
                if not os.path.exists(ruta + ".crdownload"):
                    archivos_validos.append(ruta)

            if len(archivos_validos) >= cantidad:
                return archivos_validos

        time.sleep(1)

    return None

def enviar_puerto_por_ramos(RAMOS, puerto):
    for r in RAMOS:

        ctx_ramo = r["ctx"]

        if not ctx_ramo.id_poliza:
            continue

        if ctx_ramo.activo:
            continue

        logging.info(f"⌛ Enviando puerto {puerto} al Id -> {ctx_ramo.id_poliza} de '{r['ramo']}'")

        if not enviar_puerto(ctx_ramo.id_poliza, puerto):
            raise Exception(f"No se pudo enviar puerto {puerto}")

        time.sleep(1)

        return True

    logging.info("ℹ️ No hubo ramos disponibles para enviar puerto")
    return False
