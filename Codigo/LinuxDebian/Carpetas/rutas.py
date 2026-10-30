# -*- coding: utf-8 -*-
# -- Imports --
import os,time
import logging
import base64
#---- Froms ---
from Tiempo.fechas_horas import get_timestamp,get_dia,get_mes
from io import StringIO

#---- Carpetas Descargas -----------
nombre_carpeta_descargas = "Downloads"
download_path = f"/app/{nombre_carpeta_descargas}"

import io
from PIL import Image

def optimizar_imagen_base64(ruta_archivo, max_width=1280, max_bytes=600000):
    try:
        tamano_orig = os.path.getsize(ruta_archivo)
        if tamano_orig <= max_bytes:
            with open(ruta_archivo, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")

        # Si excede el tamaño seguro para correos HTML (p. ej. 1.5MB), la optimizamos
        with Image.open(ruta_archivo) as img:
            # Convertir a RGB por si es RGBA y guardamos en JPEG
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")

            # Redimensionar si es muy ancha
            if img.width > max_width:
                proporcion = max_width / float(img.width)
                alto_nuevo = int(float(img.height) * proporcion)
                img = img.resize((max_width, alto_nuevo), Image.Resampling.LANCZOS)

            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=80, optimize=True)
            data_opt = buffer.getvalue()
            logging.info(f"⚡ Imagen optimizada para correo: {os.path.basename(ruta_archivo)} ({tamano_orig} -> {len(data_opt)} bytes)")
            return base64.b64encode(data_opt).decode("utf-8")
    except Exception as e:
        logging.warning(f"⚠️ No se pudo optimizar imagen con PIL ({e}), enviando original")
        with open(ruta_archivo, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")

def obtener_imagenes_para_correo_general(ruta_carpeta, nombre_error):

    try:
        prefijos = (nombre_error,) if isinstance(nombre_error, str) else tuple(nombre_error)
        archivos = os.listdir(ruta_carpeta) if os.path.exists(ruta_carpeta) else []

        for archivo in archivos:

            if archivo.lower().endswith(".png") and any(archivo.startswith(f"ERROR_{pref}") for pref in prefijos):

                ruta_completa = os.path.join(
                    ruta_carpeta,
                    archivo
                )

                data_b64 = optimizar_imagen_base64(ruta_completa)
                return data_b64

        logging.warning(f"⚠️ No se encontró evidencia visual con prefijo(s) {prefijos} en {ruta_carpeta}")

    except Exception as e:

        logging.error(f"❌ Error leyendo imagen: {e}")

    return None

def obtener_imagenes_error_para_nota(ruta_carpeta, const):

    imagenes_payload = []

    try:
        prefijos = (const,) if isinstance(const, str) else tuple(const)

        for archivo in os.listdir(ruta_carpeta):

            if archivo.lower().endswith(".png") and any(archivo.startswith(f"ERROR_{pref}") for pref in prefijos):

                ruta_completa = os.path.join(ruta_carpeta, archivo)
                imagen_base64 = optimizar_imagen_base64(ruta_completa)
                if imagen_base64:
                    imagenes_payload.append(imagen_base64)

    except Exception as e:
        logging.error(f"❌ Error leyendo imágenes de la carpeta: {e}")

    return imagenes_payload

def armar_ruta_archivos(tipo_proceso, ba_codigo, bb_codigo, compania_BA, compania_BB,ctx, poliza1, poliza2, poliza3,entorno):

    log_buffer = StringIO()
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        handlers=[logging.StreamHandler(log_buffer)],
        force=True
    )

    id_poliza = obtener_id_poliza(ctx)

    prefijo = "PRUEBAS_" if not entorno else ""

    mapa_procesos = {
        "IN": "Inclusiones",
        "RE": "Renovaciones",
        "NE": "Neteos"
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

def contar_archivos(ruta):
    return len([
        f for f in os.listdir(ruta)
        if os.path.isfile(os.path.join(ruta, f))
    ])

def asignar_error_trama(ctx_ramo):
    tipo = "Fallas en la Trama con Azure"
    detalle = "Problemas al obtener la trama, comunicate con el administrador"

    if ctx_ramo.ramo == "VIDALEY":
        return None, None, tipo, detalle

    return tipo, detalle, None, None

def descargar_trama(driver, ruta, nombre, ruta_descargas):
    antes = contar_archivos(ruta_descargas)

    logging.info(f"⌛ Descargando {nombre}")
    driver.get(ruta)
    time.sleep(2)

    despues = contar_archivos(ruta_descargas)

    return despues == antes + 1
