# -*- coding: utf-8 -*-
# -- Froms ---
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# -- Imports --
import os
import logging
import glob
import pdfplumber
import unicodedata
import re
import time
import random

def obtener_monto_codigo_endoso_pos_vl(endoso_pdf):

    with pdfplumber.open(endoso_pdf) as pdf:
        texto_total = ""
        for pagina in pdf.pages:
            extraido = pagina.extract_text()
            if extraido:
                texto_total += extraido + "\n"

        if not texto_total.strip():
            logging.warning("⚠️ No se extrajo texto del PDF.")
            return None, None

        # Normalizar caracteres extraños
        texto_total = texto_total.replace("œ", "u")
        texto_total = unicodedata.normalize("NFKD", texto_total)

        # Buscar monto con "S/" y número con "Proforma Nro."
        match = re.search(
            r"S/\s*([\d.,]+).*?Proforma\s*(?:Nro\.?|N°)?\s*(\d+)",
            texto_total,
            re.IGNORECASE | re.DOTALL
        )
        if match:
            monto = match.group(1).replace(",", "")  # limpia comas si hubiera
            proforma = match.group(2)
            return monto, proforma

    return None, None

def obtener_codigos_cuota_positiva_sctr(endoso_pdf):
    
    codigos = []

    with pdfplumber.open(endoso_pdf) as pdf:
        texto_total = ""
        for pagina in pdf.pages:
            extraido = pagina.extract_text()
            if extraido:
                texto_total += extraido + "\n"

        if not texto_total.strip():
            logging.warning("⚠️ No se extrajo texto del PDF.")
            return []

        # Normaliza caracteres raros
        texto_total = texto_total.replace("œ", "u")
        texto_total = unicodedata.normalize("NFKD", texto_total)

        # Buscar todos los números de proforma
        codigos = re.findall(r"Numero\s*de\s*Proforma\s*:\s*(\d+)", texto_total, re.IGNORECASE)

        # Eliminar duplicados manteniendo el orden original
        codigos_unicos = list(dict.fromkeys(codigos))

        # logging.info("📌 Códigos de proforma encontrados:", codigos_unicos)

    return codigos_unicos

def obtener_montos_endoso_positiva_sctr(endoso_pdf):

    texto_total = ""

    with pdfplumber.open(endoso_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_total += texto + "\n"

    if not texto_total.strip():
        logging.warning("⚠️ No se extrajo texto del PDF.")
        return [], []

    # Buscar Prima Comercial + IGV
    prima = re.findall(r"Prima\s+Comercial\s*\+?\s*IGV\s*S/?\s*([\d,]+\.\d{2})", texto_total, re.IGNORECASE) # PRIMA SALE PARA PENSION SEGUN PDF

    # Buscar Total : S/ monto
    total = re.findall(r"Total\s*:\s*S/?\s*([\d,]+\.\d{2})", texto_total, re.IGNORECASE) # TOTAL SALE PARA SALUD SEGUN PDF

    # Eliminar duplicados
    primas_unicas = list(set(prima))
    totales_unicos = list(set(total))

    return primas_unicas, totales_unicos

def obtener_ultimo_archivo_descargado_x_identificador(identificador,ruta_archivos_x_inclu):

    patron = os.path.join(ruta_archivos_x_inclu, f"{identificador}.pdf")
    archivos = glob.glob(patron)

    if not archivos:
        raise FileNotFoundError(f"No se encontró el documento {identificador}.pdf")

    # Ordenar por fecha de modificación (último primero)
    archivos.sort(key=os.path.getmtime, reverse=True)
    return archivos[0]  # El más reciente

def escribir_lento(elemento, texto, min_delay, max_delay):
    """Envía texto carácter por carácter con retrasos aleatorios."""
    for letra in texto:
        elemento.send_keys(letra)
        time.sleep(random.uniform(min_delay, max_delay))

def mover_y_hacer_click_simple(driver, elemento, steps=6, pause_between=0.06):
    """
    Mueve el mouse en 'steps' pasos hacia el centro del elemento y hace click.
    driver: tu instancia de webdriver
    elemento: WebElement destino
    """
    action = ActionChains(driver)
    # asegurarnos que el elemento esté visible en pantalla
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
    time.sleep(random.uniform(0.15, 0.45))

    # posiciona el mouse sobre el elemento (move_to_element genera mouseover)
    action.move_to_element(elemento).pause(random.uniform(0.05, 0.18)).perform()

    # pequeños movimientos aleatorios alrededor antes del click
    for _ in range(steps):
        offset_x = random.randint(-6, 6)
        offset_y = random.randint(-6, 6)
        action.move_by_offset(offset_x, offset_y).pause(pause_between)
    # volver al elemento y click
    action.move_to_element(elemento).pause(random.uniform(0.08, 0.2)).click().perform()

def validar_pagina(driver):

    asunto = ""

    try:

        # Validar si aparece el mensaje de error en el body
        if "The requested URL was rejected. Please consult with your administrator." in driver.page_source:
            asunto = "Página web de La Positiva fuera de Servicio"
            return False,asunto

        # Validar si aparece el campo de usuario
        user_field = WebDriverWait(driver,5).until(EC.presence_of_element_located((By.NAME, "txtUsuario")))
        if user_field:
            asunto = "Página web de La Positiva fuera de Servicio"
            return False,asunto

    except TimeoutException:
        return True,asunto

def validardeuda(driver,wait):

    try:

        WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.ID, "divAlertas")))
        logging.warning("⚠️ Apareció el modal de advertencia.")

        mensaje_deuda = wait.until(EC.visibility_of_element_located((By.ID, "spMensaje")))
        deuda_advertencia = mensaje_deuda.get_attribute("innerText").strip()

        logging.warning(f"📄 Mensaje de advertencia detectado:\n{deuda_advertencia}")
        return deuda_advertencia

    except TimeoutException:
        return False

# Ruta del directorio compartido entre contenedores
SYNC_DIR = "/app/sync"
# Asegurar que exista dentro del volumen (solo la primera vez)
os.makedirs(SYNC_DIR, exist_ok=True)
# Archivo de lock compartido
LOCK_FILE = os.path.join(SYNC_DIR, "session.lock")

def acquire_lock():
    try:
        # Intentar crear el archivo de lock
        fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_RDWR)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
        return True
    except FileExistsError:
        # Ya existe => alguien más tiene el lock
        return False

def release_lock():
    try:
        os.remove(LOCK_FILE)
        logging.info("🔓 Lock liberado.")
    except FileNotFoundError:
        pass

def wait_for_lock():
    logging.info("🔒 Esperando que se libere el lock...")
    while True:
        if not os.path.exists(LOCK_FILE):
            if acquire_lock():
                logging.info("✅ Lock adquirido.")
                return True
        time.sleep(5)  # espera 5 segundos antes de volver a intentar