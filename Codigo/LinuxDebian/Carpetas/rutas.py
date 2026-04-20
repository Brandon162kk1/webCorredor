# -*- coding: utf-8 -*-
# -- Imports --
import os
import logging
import base64
#---- Froms ---
from Tiempo.fechas_horas import get_timestamp,get_dia,get_mes
from io import StringIO

#---- Carpetas Descargas -----------
nombre_carpeta_descargas = "Downloads"
download_path = f"/app/{nombre_carpeta_descargas}"

def obtener_imagenes_error(ruta_carpeta, const):

    imagenes_payload = []

    try:
        for archivo in os.listdir(ruta_carpeta):

            if archivo.startswith(f"ERROR_{const}") and archivo.lower().endswith(".png"):

                ruta_completa = os.path.join(ruta_carpeta, archivo)

                with open(ruta_completa, "rb") as f:
                    imagen_base64 = base64.b64encode(f.read()).decode("utf-8")

                imagenes_payload.append(imagen_base64)

    except Exception as e:
        logging.error(f"❌ Error leyendo imágenes de la carpeta: {e}")

    return imagenes_payload

def armar_ruta_archivos(tipo_proceso, ba_codigo, bb_codigo, compania_BA, compania_BB,ctx, poliza1, poliza2, poliza3):

    log_buffer = StringIO()
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[logging.StreamHandler(log_buffer)],
        force=True
    )

    id_poliza = obtener_id_poliza(ctx)

    prefijo = "PRUEBAS_" if ctx.entorno == "LOCAL" else ""

    mapa_procesos = {
        "IN": "Inclusiones",
        "RE": "Renovaciones"
    }

    nombre_base = mapa_procesos.get(tipo_proceso)
    nombre_carpeta_dia = f"{prefijo}{nombre_base}_{get_dia()}_{get_mes()}"

    carpeta_dia = os.path.join(download_path, nombre_carpeta_dia)

    nom_subcarpeta= f"{id_poliza}_{tipo_proceso}_{ba_codigo}_{compania_BA}_{bb_codigo}_{compania_BB}_{poliza1}_{poliza2}_{poliza3}_{get_timestamp()}"
    subcarpeta_inclusion = os.path.join(carpeta_dia, nom_subcarpeta)

    ruta_archivos_x_inclu = f"/app/Downloads/{nombre_carpeta_dia}/{nom_subcarpeta}"

    ruta_log = os.path.join(subcarpeta_inclusion, f"log_{get_timestamp()}.txt")

    os.makedirs(carpeta_dia, exist_ok=True)
    os.makedirs(subcarpeta_inclusion, exist_ok=True)

    root_logger = logging.getLogger()
    root_logger.handlers.clear()

    file_handler = logging.FileHandler(ruta_log, mode="a", encoding="utf-8")
    file_handler.setFormatter(logging.Formatter("%(message)s"))
    root_logger.addHandler(file_handler)

    log_buffer.seek(0)
    for line in log_buffer.readlines():
        logging.info(line.strip())

    return ruta_archivos_x_inclu

def obtener_id_poliza(ctx):
    for attr in ["salud", "pension", "vida"]:
        obj = getattr(ctx, attr, None)
        if obj:
            id_poliza = getattr(obj, "id_poliza", None)
            if id_poliza:  # válido (no None, no vacío)
                return id_poliza

    raise ValueError("No se encontró identificador de la solicitud")
