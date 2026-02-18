# -- Imports --
import time
import subprocess
import logging
#-- Froms --
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

def descargar_documento(driver,boton_descarga,nombre_documento,impresion,pestaña):
    
    try:

        driver.execute_script("arguments[0].click();", boton_descarga)
        logging.info("🖱 Clic con JS en el botón de descarga")

        logging.info("⌛ Esperando la ventana descarga de Linux Debian")

        if not esperar_ventana("Save File"):
            raise Exception("No apareció la ventana 'Save File'")

        if impresion:
            time.sleep(3)
            subprocess.run(["xdotool", "key", "Return"])
            logging.info(f"🖱️ Clic en 'Save'")
            time.sleep(3)

        subprocess.run(["xdotool", "search", "--name", "Save File", "windowactivate", "windowfocus"])
        logging.info("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
        subprocess.run(["xdotool", "type","--delay", "100", nombre_documento])
        logging.info("📄 Se cambio el nombre del documento")
        time.sleep(2)
        subprocess.run(["xdotool", "key", "Return"])
        logging.info("🖱️ Clic en 'Save' para confirmar la descarga")
        time.sleep(2)

        if impresion:
            try:
                time.sleep(3)
                aceptar_btn = WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID, "CloseModal")))
                aceptar_btn.click()
                logging.info(f"🖱️ Clic en 'Aceptar' para cerrar el Modal")
            except TimeoutException:
                pass

        return True

    except Exception as ex:
        logging.error(f"❌ Error durante el flujo de descarga: {ex}")
        if pestaña:
            driver.close()
        return False

def subir_trama(boton,ruta_trama):
    
    try:

        boton.click()
        logging.info("🖱️ Clic en el botón 'Subir'")

        logging.info("⌛ Esperando la ventana 'Open File' de Linux Debian")

        ventana_id = esperar_ventana("Open File")
        if not ventana_id:
            raise Exception("No apareció la ventana 'Open File'")

        ids = subprocess.check_output(
            ["xdotool", "search", "--onlyvisible", "--name", "."]
        ).decode().split()

        for win in ids:
            name = subprocess.check_output(
                ["xdotool", "getwindowname", win]
            ).decode().strip()
            logging.info(str(win) + " -> " + name)

        logging.info(f"Id Ventana: {ventana_id}")
        # Activar la ventana específica
        subprocess.run(["xdotool", "windowactivate", "--sync", ventana_id])
        logging.info("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
        time.sleep(0.5)
        subprocess.run(["xdotool", "type","--delay", "200", ruta_trama])
        logging.info("📄 Se cambio el nombre del documento")
        subprocess.run(["xdotool", "key", "--clearmodifiers", "Alt+o"])
        subprocess.run(["xdotool", "key", "--window", ventana_id, "Return"])
        logging.info("🖱️ Clic en Open")

        if not esperar_que_cierre("Open File"):
            raise Exception("La ventana sigue abierta después del Enter")
        
        return True

    except Exception as ex:
        logging.error(f"❌ Error durante el flujo de subida: {ex}")
        return False

def obtener_id_ventana(nombre):
    result = subprocess.run(
        ["xdotool", "search", "--name", nombre],
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True
    )
    if result.stdout.strip():
        return result.stdout.strip().split("\n")[0]  # Tomamos la primera
    return None

def esperar_ventana(nombre, timeout=10):
    inicio = time.time()
    while time.time() - inicio < timeout:
        ventana_id = obtener_id_ventana(nombre)
        if ventana_id:
            return ventana_id
        time.sleep(0.5)
    return None

def esperar_que_cierre(nombre, timeout=10):
    inicio = time.time()
    while time.time() - inicio < timeout:
        if not obtener_id_ventana(nombre):
            return True
        time.sleep(0.5)
    return False

def descargar_trama():

    try:


        ventana_id = esperar_ventana("Save File")
        if not ventana_id:
            raise Exception("No apareció la ventana 'Save File'")

        ids = subprocess.check_output(
            ["xdotool", "search", "--onlyvisible", "--name", "."]
        ).decode().split()

        for win in ids:
            name = subprocess.check_output(
                ["xdotool", "getwindowname", win]
            ).decode().strip()
            logging.info(str(win) + " -> " + name)

        logging.info(f"Id Ventana: {ventana_id}")
        # Activar la ventana específica
        subprocess.run(["xdotool", "windowactivate", "--sync", ventana_id])
        logging.info("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
        time.sleep(0.5)
        subprocess.run(["xdotool", "key", "--clearmodifiers", "Alt+o"])
        subprocess.run(["xdotool", "key", "--window", ventana_id, "Return"])
        logging.info("🖱️ Clic en Save")

        if not esperar_que_cierre("Save File"):
            raise Exception("La ventana sigue abierta después del Enter")
        # #-----------------------------------------
        # ventana_id = esperar_ventana("Save File")

        # if not ventana_id:
        #     raise Exception("No apareció la ventana 'Save File'")

        # # Activar la ventana específica
        # subprocess.run(["xdotool", "windowactivate", "--sync", ventana_id])
        # time.sleep(0.5)

        # subprocess.run(["xdotool", "key", "--clearmodifiers", "Alt+o"])
        # subprocess.run(["xdotool", "key", "--window", ventana_id, "Return"])
        # logging.info("🖱️ Clic en 'Save'")

        # # Esperar que se cierre
        # if not esperar_que_cierre("Save File"):
        #     raise Exception("La ventana sigue abierta después del Enter")

        return True

    except Exception as ex:
        logging.error(f"❌ Error durante la descarga: {ex}")
        return False
