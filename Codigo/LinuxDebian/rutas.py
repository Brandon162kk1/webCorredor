# -*- coding: utf-8 -*-
# -- Imports --
import os
#import docker
import logging
import logging
#---- Froms ---
from Tiempo.fechas_horas import get_timestamp,get_dia,get_mes
from io import StringIO

#---- Carpetas Descargas -----------
nombre_carpeta_descargas = "Downloads"
download_path = f"/app/{nombre_carpeta_descargas}"

def armar_ruta_archivos(tipo_proceso, ba_codigo, bb_codigo, compania_BA, compania_BB, poliza1, poliza2, poliza3):

    # --- 👇 CREAR UN BUFFER NUEVO POR CADA CORREO ---
    log_buffer = StringIO()
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[logging.StreamHandler(log_buffer)],
        force=True
    )

    if tipo_proceso == 'IN':
        nombre_carpeta_dia = f"Inclusiones_{get_dia()}_{get_mes()}"
    else:
        nombre_carpeta_dia = f"Renovaciones_{get_dia()}_{get_mes()}"

    carpeta_dia = os.path.join(download_path, nombre_carpeta_dia)

    nom_subcarpeta= f"{tipo_proceso}_{ba_codigo}_{compania_BA}_{bb_codigo}_{compania_BB}_{poliza1}_{poliza2}_{poliza3}_{get_timestamp()}"
    subcarpeta_inclusion = os.path.join(carpeta_dia, nom_subcarpeta)

    ruta_archivos_x_inclu = f"/app/Downloads/{nombre_carpeta_dia}/{nom_subcarpeta}"

    ruta_log = os.path.join(subcarpeta_inclusion, f"log_{get_timestamp()}.txt")

    os.makedirs(carpeta_dia, exist_ok=True)
    os.makedirs(subcarpeta_inclusion, exist_ok=True)

    # Logger definitivo solo para este correo
    root_logger = logging.getLogger()
    root_logger.handlers.clear()

    file_handler = logging.FileHandler(ruta_log, mode="a", encoding="utf-8")
    file_handler.setFormatter(logging.Formatter("%(message)s"))
    root_logger.addHandler(file_handler)

    # --- Pasar logs temporales al archivo ---
    log_buffer.seek(0)
    for line in log_buffer.readlines():
        logging.info(line.strip())

    return ruta_archivos_x_inclu