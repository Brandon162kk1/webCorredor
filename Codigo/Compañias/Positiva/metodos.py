# -*- coding: utf-8 -*-
# -- Froms ---
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException,StaleElementReferenceException,ElementClickInterceptedException
from LinuxDebian.Ventana.ventana import esperar_archivos_nuevos
from Tiempo.fechas_horas import time_espera_alea
# -- Imports --
import os
import logging
import time
import random
import pdfplumber
import re
import shutil

def validar_modal_error(driver, elemento):
    if not isinstance(elemento, bool) and elemento.get_attribute("id") == "divAlertaErrorGeneral":
        mensaje = driver.find_element(
            By.XPATH,
            "//div[@id='divAlertaErrorGeneral']//p/span[2]"
        ).text.strip()
        raise Exception(mensaje)

def descargar_documento_por_codigo(driver,wait,codigo_documento,palabra_clave,tipo_mes,ba_codigo,list_polizas,ramo,ruta_archivos_x_inclu):

    # Esperar que aparezca el código
    span_xpath = f"//span[normalize-space()='{codigo_documento}']"

    wait.until(
        EC.visibility_of_element_located((By.XPATH, span_xpath))
    )

    logging.info(f"✅ Span encontrado: {codigo_documento}")

    # XPath relativo de la lupa
    if len(list_polizas) == 1:
        if ba_codigo == "1":
            xpath_lupa_relativo = (
                f"./following-sibling::span[1]"
                f"//img[@data-nropolizasalud='{ramo.poliza}']"
            )
        else:
            xpath_lupa_relativo = (
                f"./following-sibling::span[1]"
                f"//img[@data-nropolizapension='{ramo.poliza}']"
            )
    else:
        xpath_lupa_relativo = (
            f"./following-sibling::span[1]"
            f"//img[@data-nropolizasalud='{list_polizas[0]}' "
            f"and @data-nropolizapension='{list_polizas[1]}']"
        )

    error_btn = (By.ID, "btnAceptarError")

    def obtener_lupa():
        """
        Siempre vuelve a localizar el span y la lupa.
        Así evitamos referencias stale.
        """
        span = driver.find_element(By.XPATH, span_xpath)

        lupa = span.find_element(By.XPATH, xpath_lupa_relativo)
        logging.info(lupa.get_attribute("outerHTML"))

        return span.find_element(By.XPATH, xpath_lupa_relativo)

    try:
        wait.until(lambda d: obtener_lupa().is_displayed())
    except TimeoutException:
        raise Exception(
            f"Problemas en la compañía, buscar y descargar los documentos: {codigo_documento}"
        )

    error_btn = (By.ID, "btnAceptarError")

    for intento in range(3):

        try:

            errores = driver.find_elements(*error_btn)

            if errores and errores[0].is_displayed():
                raise Exception(
                    f"Advertencia detectada. Código de la {palabra_clave}: {codigo_documento}"
                )

            lupa = obtener_lupa()
            wait.until(lambda d: lupa.is_displayed() and lupa.is_enabled())

            driver.execute_script("arguments[0].click();", lupa)

            logging.info(f"🖱️ Clic en la lupa {codigo_documento}")
            break

        except (StaleElementReferenceException,ElementClickInterceptedException):

            logging.warning(
                f"⚠️ DOM actualizado. Reintentando ({intento+1}/3)..."
            )

    else:
        raise Exception(
            f"No fue posible hacer clic en la lupa del código {codigo_documento}"
        )

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
        raise Exception(f"No se descargó constancia, buscar en la compañía con el código '{codigo_documento}'")

    if tipo_mes == 'MA':

        driver.switch_to.frame("ifContenedorPDFMaster")

        boton_embebido = wait.until(EC.element_to_be_clickable((By.ID, "open-button")))

        archivos_antes_2 = set(os.listdir(ruta_archivos_x_inclu))

        driver.execute_script("arguments[0].click();", boton_embebido)
        logging.info(f"🖱️ Clic en Descargar Endoso")

        endoso = esperar_archivos_nuevos(ruta_archivos_x_inclu, archivos_antes_2, ".pdf", cantidad=1)

        if endoso:
            logging.info(f"✅ Endoso descargado exitosamente")
            ruta_final = os.path.join(ruta_archivos_x_inclu, f"endoso_{ramo.poliza}.pdf")
            os.rename(endoso[0], ruta_final)
            logging.info("🔄 Endoso renombrado")
        else:
            raise Exception("No se descargó el endoso")

        driver.switch_to.default_content()

    if len(list_polizas) == 2:

        ruta_salud = os.path.join(ruta_archivos_x_inclu, f"{list_polizas[0]}.pdf")
        ruta_pension = os.path.join(ruta_archivos_x_inclu, f"{list_polizas[1]}.pdf")

        if os.path.exists(ruta_pension):
            shutil.copy2(ruta_pension, ruta_salud)
            logging.info(f"📄 Copia creada como '{list_polizas[0]}.pdf'")

        if tipo_mes == 'MA':
            ruta_endoso_salud = os.path.join(ruta_archivos_x_inclu, f"endoso_{list_polizas[0]}.pdf")
            ruta_endoso_pension = os.path.join(ruta_archivos_x_inclu, f"endoso_{list_polizas[1]}.pdf")

            if os.path.exists(ruta_endoso_pension):
                shutil.copy2(ruta_endoso_pension, ruta_endoso_salud)
                logging.info(f"📄 Copia creada como 'endoso_{list_polizas[0]}.pdf'")

    btn_cancelar_boton = wait.until(EC.element_to_be_clickable((By.ID, "btnPDFCancelarM")))
    btn_cancelar_boton.click()
    logging.info("✅ Cerrando panel de documentos")

    try:
        # wait.until(EC.invisibility_of_element_located((By.ID, "divPanelPDFMaster")))
        # logging.info("📴 Panel PDF cerrado correctamente")

        modal_pdf = (By.ID, "divPanelPDFMaster")
        error_locator2 = (By.XPATH, "//h3[contains(text(),'Actualmente estamos presentando problemas, por favor')]")
        resultadof = wait.until(
            EC.any_of(
                EC.invisibility_of_element_located(modal_pdf),
                EC.presence_of_element_located(error_locator2)
            )
        )

        if resultadof.get_attribute("id") == "divPanelPDFMaster":
            logging.info("📴 Panel PDF cerrado correctamente")
        else:
            pass

    except TimeoutException:
        pass

    logging.info(f"✅ {palabra_clave} realizada exitosamente")

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
    time_espera_alea(0.15,0.45)

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

    # Error de Chrome: conexión reseteada
    if "ERR_CONNECTION_RESET" in page:
        return False, "Chrome: ERR_CONNECTION_RESET"

    # Otros errores de Chrome
    errores_chrome = [
        "ERR_CONNECTION_TIMED_OUT",
        "ERR_NAME_NOT_RESOLVED",
        "ERR_CONNECTION_REFUSED",
        "ERR_INTERNET_DISCONNECTED",
        "ERR_EMPTY_RESPONSE",
    ]

    for error in errores_chrome:
        if error in page:
            return False, f"Chrome: {error}"

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
