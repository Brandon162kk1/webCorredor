# -*- coding: utf-8 -*-
# -- Froms ---
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
# -- Imports --
import logging

def abrirDriver(ruta_archivos_x_inclu):

    #-----------Configurar opciones de Chrome y la carpeta de descargas
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--no-sandbox') 
    chrome_options.add_argument('--disable-popup-blocking')
    chrome_options.add_argument("--window-size=1920,1080")  
    chrome_options.add_argument("--disable-dev-shm-usage") 
    chrome_options.add_argument("--disable-gpu")

    prefs = {
        "download.default_directory": ruta_archivos_x_inclu,
        "download.prompt_for_download": False, 
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,  
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.exit_type": "None",
        "download.extensions_to_open": ""

    }
    chrome_options.add_experimental_option("prefs", prefs)

    # ✅ Evitar creación de perfiles nuevos
    chrome_options.add_argument(f"--user-data-dir=/tmp/chrome-profile")
    #-----------------------

    try:
        logging.info("🟡 Iniciando ChromeDriver con webdriver_manager")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        logging.info("🟢 ChromeDriver iniciado correctamente.")

    except Exception as e:
        logging.info(f"❌ Error al iniciar ChromeDriver: {e}")

    return driver
