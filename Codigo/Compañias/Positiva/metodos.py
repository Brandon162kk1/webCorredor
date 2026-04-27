# -*- coding: utf-8 -*-
# -- Froms ---
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException,StaleElementReferenceException
from LinuxDebian.Ventana.ventana import esperar_archivos_nuevos
# -- Imports --
import os
import logging
import time
import random
import pdfplumber
import re

def descargar_documento_por_codigo(driver,wait,codigo_documento,prefijo,ramo,ruta_archivos_x_inclu):

    span = wait.until(EC.visibility_of_element_located((
        By.XPATH, f"//span[contains(text(),'{codigo_documento}')]"
    )))

    bloque = span.find_element(By.XPATH, "./ancestor::div[1]")

    lupa = bloque.find_element(
        By.XPATH,
        f".//img[contains(@data-nropolizasalud,'{ramo.poliza}') or contains(@data-nropolizapension,'{ramo.poliza}')]"
    )

    #lupa = span_numero.find_element(By.XPATH,f".//ancestor::tr//img[contains(@data-nropolizasalud,'{list_polizas[0]}') or contains(@data-nropolizapension,'{list_polizas[0]}')]")
    #lupa = (By.XPATH, selector_xpath)
    error_btn = (By.ID, "btnAceptarError")

    resultadol = wait.until(
        EC.any_of(
            EC.element_to_be_clickable(lupa),
            EC.element_to_be_clickable(error_btn)
        )
    )

    if resultadol.get_attribute("id") == "btnAceptarError":
        raise Exception(f"Advertencia detectada. Código: {codigo_documento}")

    for _ in range(3):
        try:
            resultadol.click()
            logging.info(f"🖱️ Clic en la lupa {codigo_documento}")
            break
        except StaleElementReferenceException:
            resultado = wait.until(EC.element_to_be_clickable(lupa))

    wait.until(EC.visibility_of_element_located((By.ID, "divPanelPDFMaster")))
    logging.info("📄 Panel PDF visible")

    boton_guardar = wait.until(EC.element_to_be_clickable((By.ID, "btnDescargarConstanciaM")))

    archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

    driver.execute_script("arguments[0].click();", boton_guardar)
    logging.info(f"🖱️ Clic en Descargar Constancia")

    archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu, archivos_antes, ".pdf", cantidad=1)

    if archivo_nuevo:
        ruta_final = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
        os.rename(archivo_nuevo[0], ruta_final)
        logging.info(f"🔄 Constancia renombrada")
    else:
        raise Exception(f"No se descargó constancia, buscar en la compania con el código '{codigo_documento}'")

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

    page = driver.page_source

    if "The requested URL was rejected. Please consult with your administrator." in page:
        return False, "Página web de La Positiva fuera de Servicio"

    if "404 - File or directory not found." in page:
        return False, "Página 404 - Archivo o directorio no encontrado"

    overlay = (By.ID, "ID_MODAL_PROCESS")
    user_field_1 = (By.NAME, "txtUsuario")
    user_field_2 = (By.NAME, "username")

    try:
        WebDriverWait(driver, 5).until(
            EC.any_of(
                EC.visibility_of_element_located(overlay),
                EC.presence_of_element_located(user_field_1),
                EC.presence_of_element_located(user_field_2)
            )
        )

        loader = driver.find_elements(*overlay)
        if loader and loader[0].is_displayed():
            return False, "La página está demorando demasiado en cargar"

        if driver.find_elements(*user_field_1):
            return False, "Redireccionó a otra página (login)"

        if driver.find_elements(*user_field_2):
            return False, "Redireccionó a otra página (login)"

        return True, asunto

    except TimeoutException:
        return True, asunto

def obtener_tramite_pdf(ruta_pdf):

    try:

        texto = ""

        with pdfplumber.open(ruta_pdf) as pdf:
            for page in pdf.pages:
                texto += page.extract_text() or ""

        # Normalizar texto (muy importante en PDFs)
        texto = " ".join(texto.split())

        # 1️⃣ Validar si existe Nota:
        if "Nota:" not in texto:
            return False

        # 2️⃣ Buscar fecha + número de trámite
        patron = r"\d{1,2}\s+de\s+\w+\s+de\s+\d{4}\s+(\d+)"

        match = re.search(patron, texto)

        if match:
            return match.group(1)  # número de trámite
        else:
            return None

    except Exception as e:
        logging.error(f"❌ Error leyendo PDF: {e}")
        return None

def leer_pdf(ruta_pdf):
    texto_completo = ""

    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo += texto + "\n"

    # Buscar la primera aparición de "Celda"
    indice = texto_completo.find("Celda")

    if indice != -1:
        return texto_completo[indice:]
    else:
        return ""  # No encontró la palabra

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
