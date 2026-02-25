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

#-------- Esperas --------
def esperar_descarga_por_nombre(ruta_archivo, timeout=60):
    tiempo_inicio = time.time()

    while True:
        # Si ya existe y no hay archivo temporal
        if os.path.exists(ruta_archivo):
            #if not os.path.exists(ruta_archivo + ".crdownload"):
                return True

        if time.time() - tiempo_inicio > timeout:
            return False

        time.sleep(1)

def esperar_descarga_por_evento(directorio, timeout=60):
    inicio = time.time()

    while True:
        # Buscar archivos temporales de Chrome
        temporales = [
            f for f in os.listdir(directorio)
            if f.endswith(".crdownload")
        ]

        if not temporales:
            return True  # Ya no hay descargas activas

        if time.time() - inicio > timeout:
            return False

        time.sleep(1)

def esperar_ultima_descarga(directorio, timeout=60):
    tiempo_inicio = time.time()

    while True:

        extensiones_ignorar = (".crdownload", ".txt")

        archivos = [
            os.path.join(directorio, f)
            for f in os.listdir(directorio)
            if not f.endswith(extensiones_ignorar)
        ]

        if archivos:
            # Tomar el archivo más reciente
            ultimo = max(archivos, key=os.path.getctime)
            return ultimo

        if time.time() - tiempo_inicio > timeout:
            return None

        time.sleep(1)

# def esperar_archivos_nuevos(directorio, archivos_antes, cantidad_esperada, timeout=60):
#     inicio = time.time()

#     while time.time() - inicio < timeout:
#         actuales = set(os.listdir(directorio))
#         nuevos = actuales - archivos_antes

#         # Solo PDFs
#         nuevos = {f for f in nuevos if f.endswith(".pdf")}

#         if len(nuevos) >= cantidad_esperada:
#             rutas = [os.path.join(directorio, f) for f in nuevos]
#             return rutas

#         time.sleep(1)

#     return None

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
