# -- Imports --
import time
import subprocess
import logging
import os
#-- Froms --

#-------- Interacción Humana --------
def bloquear_interaccion():
    subprocess.run(["x11vnc", "-remote", "viewonly"], check=False)
    logging.info("🔒 Interacción humana BLOQUEADA (VNC view-only)")

def desbloquear_interaccion():
    subprocess.run(["x11vnc", "-remote", "noviewonly"], check=False)
    logging.info("✋ Interacción humana HABILITADA")

#-------- Esperas Avanzadas --------
def esperar_archivos_nuevos(directorio, archivos_antes, extension, cantidad, timeout=60):
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
