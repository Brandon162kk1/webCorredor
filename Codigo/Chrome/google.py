from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from Tiempo.fechas_horas import get_timestamp
import logging
import os
import json

def tomar_capturar(driver, ruta, prefijo):
    nombre = f"{prefijo}_{get_timestamp()}.png"
    ruta_completa = os.path.join(ruta, nombre)
    driver.save_screenshot(ruta_completa)
    return ruta_completa

def abrirDriver(ruta_archivos_x_inclu):

    # Configurar opciones de Chrome
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--no-sandbox') 
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument("--window-size=1920,1080")  
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")

    # Importante para imprimir sin Dialogo
    chrome_options.add_argument("--kiosk-printing")

    # Configuracion de impresion a PDF predeterminada
    app_state = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": ""
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "isHeaderFooterEnabled": False,
        "isCssBackgroundEnabled": True
    }

    prefs = {
        # Carpeta para descargas regulares
        "download.default_directory": ruta_archivos_x_inclu,
        # Carpeta obligatoria para 'Guardar como PDF' de la ventana de impresion
        "savefile.default_directory": ruta_archivos_x_inclu,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
        # Serializar appState como string JSON valido
        "printing.print_preview_sticky_settings.appState": json.dumps(app_state)
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    # Indicar ruta correcta de Chrome
    chrome_options.binary_location = "/usr/bin/google-chrome"

    # Usar ChromeDriver
    try:
        logging.info("🟡 Iniciando ChromeDriver")
        service = Service("/usr/local/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=chrome_options)
        logging.info("🟢 ChromeDriver iniciado correctamente.")
    except Exception as e:
        logging.error(f"❌ Error al iniciar ChromeDriver: {e}")
        raise

    return driver
