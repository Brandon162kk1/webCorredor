#   --- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
#   --- Imports ----
import time
import os
import logging

#----- Variables de Entorno -------
urlRimacCorredores = os.getenv("urlRimacCorredores")
remitente = os.getenv("remitente")
passwordCorredores = os.getenv("passwordCorredores")

def realizar_solicitud_corredores(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,
                                  palabra_clave,ejecutivo_responsable,fVigenciaInicio,fVigenciaFin,ba_codigo,
                                  nombre_cliente,ruc_empresa):

    driver.get(urlRimacCorredores) 
    logging.info("⌛ Ingresando a la URL.")
 
    time.sleep(1)
 
    correo_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Correo electrónico']")))
    correo_input.click()
    correo_input.send_keys(remitente)
    logging.info("⌨️ Digitando el Correo.")
 
    time.sleep(1.2)
 
    password_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Contraseña']")))
    password_input.click()
    password_input.send_keys(passwordCorredores)
    logging.info("⌨️ Digitando el Password")
 
    time.sleep(2.1)
 
    btn_ingresar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Ingresar']]")))
    btn_ingresar.click()
    logging.info("🖱️ Clic en 'Ingresar'")
    time.sleep(3.5)
 
    # try:
    #     driver.find_element(By.ID, "recaptcha-anchor").click()
 
    # except Exception:
    #     logging.info("❌ No se encontró el reCAPTCHA, continuando...")
 
 
    # iframe_tag = driver.find_element(By.XPATH, "//iframe[contains(@src, 'recaptcha')]")
    # sitekey_url = iframe_tag.get_attribute("src")
    # parsed = urlparse.urlparse(sitekey_url)
    # sitekey = urlparse.parse_qs(parsed.query)["k"][0]
    # logging.info(f"✅ Sitekey encontrada: {sitekey}")
    # resp = requests.post("http://2captcha.com/in.php", data={
    #     "key": API_KEY,
    #     "method": "userrecaptcha",
    #     "googlekey": sitekey,
    #     "pageurl": driver.current_url
    # })
    # if not resp.text.startswith("OK|"):
    #     raise Exception(f"❌ Error al enviar captcha a 2Captcha: {resp.text}")
    # captcha_id = resp.text.split('|')[1]
    # logging.info(f"⌛ Captcha enviado a 2Captcha. ID: {captcha_id}")
    # result_url = f"http://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}"
    # token = None
    # for i in range(30):
    #     time.sleep(5)
    #     res = requests.get(result_url)
    #     if "OK" in res.text:
    #         token = res.text.split('|')[1]
    #         logging.info("✅ Captcha resuelto por 2Captcha")
    #         break
    # if not token:
    #     raise Exception("❌ No se pudo resolver el captcha")
    # # 5. Inyectar el token en el campo oculto del captcha
    # driver.execute_script('document.getElementById("g-recaptcha-response").value = arguments[0];', token)
 
    # boton_ingresar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Ingresar' and @aria-disabled='false']")))
    # driver.execute_script("arguments[0].click();", boton_ingresar)
    # logging.info("✅ Click en botón 'Ingresar'")
 
    codigo_path = "/codigo_rimac_Web/codigo.txt"
 
    logging.info("⏳ Esperando código...")
    while not os.path.exists(codigo_path):
        time.sleep(2)
    with open(codigo_path, "r") as f:
        codigo = f.read().strip()
    logging.info(f"✅ Código recibido desde volumen: {codigo}")
 
    primer_input = wait.until(EC.element_to_be_clickable((By.XPATH, "(//input[@maxlength='1'])[1]")))
    primer_input.click()
    primer_input.send_keys(codigo)
    logging.info(f"⌨️ Código {codigo} digitado correctamente en TOKEN")

    # --- Eliminar el archivo después de usarlo ---
    try:
        os.remove(codigo_path)
        logging.info("🧹 Archivo codigo.txt eliminado del volumen.")
    except FileNotFoundError:
        logging.warning("⚠️ No se encontró codigo.txt al intentar eliminarlo (ya fue borrado).")
    except Exception as e:
        logging.error(f"❌ Error al eliminar codigo.txt: {e}")

    btn_ingresar2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Ingresar' and @aria-disabled='false']")))
    btn_ingresar2.click()
    logging.info("🖱️ Clic en 'Ingresar' Luego del TOKEN.")
 
    time.sleep(4)
 
    try:

        #Corregir pq hizo Vida Ley cuando debio ser SCTR

        fechaWC = "01-01-2026"
 
        operacion = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Operación']")))
        operacion.click()
        logging.info("🖱️ Clic en Operación.")
 
        time.sleep(1.2)
 
        nuevo_gestion = wait.until(EC.element_to_be_clickable((By.XPATH, "//li//span[normalize-space()='Nuevo gestión']")))
        nuevo_gestion.click()
        logging.info("🖱️ Clic en Nuevo gestión.")
 
        time.sleep(1.2)

        try:
            label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[text()='Número de RUC']")))
            label.click()
            logging.info("🖱️ Clic en el Label 'Número de RUC'.")
 
            # input_ruc = wait.until(EC.visibility_of_element_located((By.ID, "1439f3dabc2c2d4f23e3d6ddcad2f9f9_input_doc_ruc")))
            # input_ruc.click()
            # logging.info("🖱️ Clic en el Input 'Número de RUC'.")
            label.clear()
            label.send_keys(ruc_empresa)
            logging.info(f"⌨️ RUC ingresado: {ruc_empresa}")
        except Exception:
            driver.switch_to.active_element.clear()
            driver.switch_to.active_element.send_keys(ruc_empresa)
            logging.info(f"⌨️ RUC ingresado: {ruc_empresa} SEGUNDO!!!!!!!!!!")
 
        time.sleep(1.2)
 
        boton_buscar = driver.find_element(By.XPATH, "//button[normalize-space()='BUSCAR']")
        boton_buscar.click()
        logging.info("🖱️ Clic en el Botón 'BUSCAR'.")
 
        time.sleep(3.5)
 
        boton_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[normalize-space()='SELECCIONAR CLIENTE']]")))
        boton_select.click()
        logging.info("🖱️ Clic en el Botón 'SELECCIONAR CLIENTE'.")
 
        time.sleep(1.2)
 
        situacion = "Inclusiones" if tipo_proceso == "IN" else "Renovación"
        
        card_inc_ren = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'icon-card')][.//h3[normalize-space()='{situacion}']]")))
        card_inc_ren.click()
        logging.info(f"🖱️ Clic en el Botón '{palabra_clave}'.")
 
        time.sleep(2)
 
        logging.info(f"Proceso seleccionado: Nueva {palabra_clave}")
 
        time.sleep(2.1)
 
        input_poliza = wait.until(EC.visibility_of_element_located((By.ID, "nroPoliza")))
        input_poliza.click()
        logging.info("🖱️ Clic en el input 'nroPoliza' ")
        input_poliza.send_keys(list_polizas[0])
        logging.info(f"⌨️ Póliza ingresada: {list_polizas[0]}.")
 
        time.sleep(1.2)
        boton_buscar = wait.until(EC.element_to_be_clickable((By.XPATH,"//button[normalize-space()='Buscar']")))
        boton_buscar.click()
        logging.info("🖱️ Clic en el Botón 'BUSCAR'.")
 
        time.sleep(1)

        situacion2 = "SELECCIONAR" if tipo_proceso == "IN" else "RENOVAR"
 
        boton_seleccionar = wait.until(EC.element_to_be_clickable((By.XPATH, f"//button[normalize-space()='{situacion2}']")))
        driver.execute_script("arguments[0].click();", boton_seleccionar)
        logging.info(f"🖱️ Clic en el Botón '{situacion2}' de la Póliza.")

 
        time.sleep(2.1)
        # dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.dropdown-container")))
        # driver.execute_script("arguments[0].click();", dropdown)
        # logging.info("🖱️ Clic en el Dropdown de 'Fecha de Inicio de Vigencia'.")
 
 
        #ESTO ES TEST NOMAS
        dropdown_arrow = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.dropdown-container div.dropdown-arrow")))
        driver.execute_script("arguments[0].click();", dropdown_arrow)
        logging.info("🖱️ Clic en la flecha del dropdown")
 
        time.sleep(1.2)
 
        item = wait.until(EC.element_to_be_clickable((By.XPATH,f"//div[contains(@class,'dropdown-item') and normalize-space()='{fechaWC}']")))
        driver.execute_script("arguments[0].click();", item)
        logging.info(f"🖱️ Fecha seleccionada: {fechaWC}")
 
        time.sleep(1.2)
 
        boton_registrar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Registrar']")))
        driver.execute_script("arguments[0].click();", boton_registrar)
        logging.info("🖱️ Click en el botón 'Registrar'")
 
        time.sleep(2.1)
 
        input_file = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.file-uploader input[type='file']")))
        ruta_trama_xls_rimacWC = f"{ruta_archivos_x_inclu}/TramaWC.xlsx"
        input_file.send_keys(ruta_trama_xls_rimacWC)
 
        logging.info("✅ Trama cargada")
 
        input("ESPERA")
 
        boton_listo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='LISTO']")))
        driver.execute_script("arguments[0].click();", boton_listo)
        logging.info("🖱️ Click en el botón 'LISTO'")
       
        input("Esperar FINAL :D")
    except Exception as e:
        raise Exception(f"Error en el flujo de selección de póliza: {list_polizas[0]}, {e}")
 
    input("Espera")