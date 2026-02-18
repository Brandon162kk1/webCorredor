#   --- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from Tiempo.fechas_horas import get_timestamp
#   --- Imports ----
import time
import os
import logging

def procesar_solicitud_san_protecta_vl(driver,wait,ruc_empresa,tipo_proceso,ruta_archivos_x_inclu,palabra_clave,ramo):

    try:

        # ------------------ Inicio del Flujo de Automatización ------------------
        driver.get("https://plataformadigital.protectasecurity.pe/ecommerce/extranet/login")
        logging.info("--- Ingresando a Protecta Seguros🌐 ---")

        user_input = wait.until(EC.presence_of_element_located((By.ID, "username")))
        user_input.clear()
        user_input.send_keys(ramo.usuario)
        logging.info("⌨️ Digitando el Username")
        
        pass_input = wait.until(EC.presence_of_element_located((By.ID, "password")))
        pass_input.clear()
        pass_input.send_keys(ramo.clave)
        logging.info("⌨️ Digitando el Password")

        logging.info("🧩 Resuelve el CAPTCHA manualmente y clic en 'Ingresar'.")

        #desbloquear_interaccion()
        driver.save_screenshot(os.path.join(ruta_archivos_x_inclu,f"captcha_{get_timestamp()}.png"))
        wait_humano = WebDriverWait(driver,300)
        wait_humano.until(EC.presence_of_element_located((By.XPATH, "//span[text()='VIDA LEY']")))

        #bloquear_interaccion()

        logging.info("✅ Login exitoso detectado (Cerrar sesión visible)")
        logging.info("🚀 Continuando flujo automáticamente")

        # ingresar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Ingresar')]")))
        # ingresar_btn.click()
        # logging.info("🖱️ Clic en 'Ingresar'")

        span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='VIDA LEY']")))
        span_element.click()
        logging.info("🖱️ Clic en 'Vida Ley'")

        polizas_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Pólizas']")))
        polizas_element.click()
        logging.info("🖱️ Clic en 'Pólizas'")

        menu = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[@href='/ecommerce/extranet/policy-transactions-all' and not(@hidden)]"
            ))
        )

        driver.execute_script("arguments[0].click();", menu)
        logging.info("🖱️ Clic en 'Consulta de transacciones'")

        time.sleep(3)

        # Encuentra el elemento <select>
        dropdown = Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='typeTransac']"))))

        if tipo_proceso == 'IN':
            dropdown.select_by_value("2")
        else:
            dropdown.select_by_value("4")
        logging.info(f"🖱️ Clic en '{palabra_clave}'")

        time.sleep(2)

        # # -- Por defecto --
        # dropdown2 = Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='branch']"))))
        # dropdown2.select_by_value("73") # -- > VIDA LEY
        # logging.info(f"🖱️ Clic en 'Vida Ley'")
        # time.sleep(2)

        # # -- Por defecto --
        # dropdown3 = Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='product']"))))
        # dropdown3.select_by_value("1")
        # logging.info(f"🖱️ Clic en 'Seguro de Vida Ley Trabajadores'")
        #time.sleep(2)

        input_element = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@formcontrolname='policy']")))
        input_element.clear()
        input_element.send_keys(ramo.poliza)
        logging.info("⌨️ Digitando la Póliza")
        time.sleep(2)

        dropdown4 = Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='contractorDocumentType']"))))
        dropdown4.select_by_value("1")
        logging.info(f"🖱️ Clic en 'RUC'")
        time.sleep(2)

        input_num_contra = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@formcontrolname='contractorDocumentNumber']")))
        input_num_contra.clear()
        input_num_contra.send_keys(ruc_empresa)
        logging.info("⌨️ Digitando RUC")
        time.sleep(2)

        buscar_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Buscar']]")))
        buscar_button.click()
        logging.info(f"🖱️ Clic en 'Buscar'")
        time.sleep(5)

        # Seleccionar la póliza haciendo clic en el radio button correspondiente
        radio_seleccion = wait.until(EC.element_to_be_clickable((By.ID,f"policy{ramo.poliza}")))
        radio_seleccion.click()
        logging.info(f"🖱️ Clic en la Póliza encontrada")
        time.sleep(3)
        
        if tipo_proceso == 'IN':
            label_tipo_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Incluir')]")))
        else:
            label_tipo_proceso  = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Renovar')]")))

        # Clic en el label (esto activa el input radio asociado)
        label_tipo_proceso.click()
        logging.info(f"🖱️ Clic en {palabra_clave}")
        
        # Esperar a que cargue la nueva página de adjuntar el archivo
        #time.sleep(8)

        file_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
        # Hacer scroll hasta el input
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", file_input)

        ruta_trama_protecta = f"{ruta_archivos_x_inclu}/{ramo.poliza}.xlsm"
        file_input.send_keys(ruta_trama_protecta)
        logging.info(f" ✅ Trama '{ramo.poliza}.xlsm' subida para validar' ")

        # Click en Validar
        boton_validar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Validar']]")))
        boton_validar.click()
        logging.info(f"🖱️ Clic en Validar Trama")

        # Aparece un Modal para confirmar
        btn_ok = wait.until(EC.element_to_be_clickable((By.XPATH,"//button[contains(@class,'swal2-confirm') and normalize-space()='OK']")))
        driver.execute_script("arguments[0].click();", btn_ok)
        logging.info(f"🖱️ Clic en Aceptar la Validación")

        boton_procesar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[span[text()='PROCESAR']]")))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_procesar)
        #wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='PROCESAR']]"))).click()

    except Exception as e:
        logging.warning(f"Error en {e}")
    finally:
        input("Esperar")