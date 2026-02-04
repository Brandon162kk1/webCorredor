#   --- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from LinuxDebian.ventana import desbloquear_interaccion
from Tiempo.fechas_horas import get_timestamp
#   --- Imports ----
import time
import os
import logging

#----- Variables de Entorno -------
urlRimacSAS = os.getenv("urlRimacSAS")
usernameRimac = os.getenv("usernameRimac")
passwordRimac = os.getenv("passwordRimac")

def realizar_solicitud_SAS(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,ejecutivo_responsable,
                           ba_codigo,bb_codigo,nombre_cliente,ruc_empresa,ramo):

    desbloquear_interaccion()

    driver.get(urlRimacSAS) 
    logging.info("🔐 Iniciando sesión en RIMAC SAS...")
 
    user_input = wait.until(EC.element_to_be_clickable((By.ID, "CODUSUARIO")))
    user_input.clear()
    user_input.send_keys(usernameRimac)
    logging.info("⌨️ Usuario ha sido Digitando.")
 
    pass_input = wait.until(EC.element_to_be_clickable((By.ID, "CLAVE")))
    pass_input.clear()
    pass_input.send_keys(passwordRimac)
    logging.info("⌨️ Password ha sido Digitado")
 
    ingresar_btn = wait.until(EC.element_to_be_clickable((By.ID, "btningresar")))
    driver.execute_script("arguments[0].click();", ingresar_btn)
    logging.info("🖱️ Clic en 'Ingresar'.")

    # --- Esperar código ---
    codigo_path = "/codigo_rimac_SAS/codigo.txt"
 
    logging.info("⏳ Esperando código...")
    while not os.path.exists(codigo_path):
        time.sleep(2)
 
    with open(codigo_path, "r") as f:
        codigo = f.read().strip()
 
    logging.info(f"✅ Código recibido desde volumen: {codigo}")

    # --- Escribir en INPUT TOKEN ---
    logging.info("⌛ Buscando input TOKEN...")
 
    token_input = wait.until(EC.element_to_be_clickable((By.ID, "TOKEN")))
    token_input.clear()
    token_input.send_keys(codigo)
 
    logging.info(f"⌨️ Código {codigo} digitado correctamente en TOKEN")
 
    ingresar_btn2 = wait.until(EC.element_to_be_clickable((By.ID, "btningresar")))
    driver.execute_script("arguments[0].click();", ingresar_btn2)
    logging.info("🖱️ Clic en 'Ingresar' Luego del TOKEN.")
 
    input("Esperar")
 
    try:
        
        actions = ActionChains(driver)
        span_transacciones = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Transacciones']")))
        actions.double_click(span_transacciones).perform()
 
        logging.info("🖱️ Doble clic realizado en 'Transacciones'")
 
        span_emision = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Emisión Vida Ley']")))
        actions.double_click(span_emision).perform()
 
        logging.info("🖱️ Doble clic realizado en 'Emisión Vida Ley'")
 
        span_mantenimiento = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Mantenimiento Póliza']")))
        span_mantenimiento.click()
 
        logging.info("🖱️ Clic realizado en 'Mantenimiento Póliza'")
 
        boton_buscar = wait.until(EC.element_to_be_clickable((By.ID, "ext-gen306")))
        boton_buscar.click()
 
        logging.info("🖱️ Clic realizado en el botón Buscar")

        input_tipo_doc = wait.until(EC.element_to_be_clickable((By.ID, "ext-comp-1102")))
        input_tipo_doc.click()
 
        input_tipo_doc.send_keys("RUC")
        input_tipo_doc.send_keys(Keys.ENTER)
        logging.info("⌨️ Tipo de documento 'RUC' ingresado. Y Enter")
 
        input_num_doc = wait.until(EC.element_to_be_clickable((By.ID, "ext-comp-1103")))
        input_num_doc.click()
        logging.info("🖱️ Clic en el input 'Número de documento'")
 
        input_num_doc.send_keys(ruc_empresa)
        input_num_doc.send_keys(Keys.ENTER)
        logging.info(f"⌨️ Número de documento '{ruc_empresa}' ingresado. Y Enter")
 
        grid_body = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#ext-gen332 .x-grid3-body")))
 
        filas = grid_body.find_elements(By.CSS_SELECTOR, ".x-grid3-row")
 
        encontrado = False
 
        for fila in filas:
            try:
 
                num_doc_cell = fila.find_element(By.CSS_SELECTOR, "div.x-grid3-col-4")
                texto = num_doc_cell.text.strip()
 
                if texto == ruc_empresa:
                    logging.info(f"✔ Encontrado número: {texto}, haciendo clic en la fila.")
                    fila.click()
                    encontrado = True
                    break
 
            except:
                continue
 
        if not encontrado:
            logging.info("❌ No se encontró ese número de documento en la tabla.")
 
        btn_seleccionar = wait.until(EC.element_to_be_clickable((By.ID, "ext-gen320")))
        driver.execute_script("arguments[0].click();", btn_seleccionar)
    
        logging.info("✔ Botón 'Seleccionar' clickeado correctamente")
 
        btn_buscar = wait.until(EC.element_to_be_clickable((By.ID, "ext-gen158")))
        driver.execute_script("arguments[0].click();", btn_buscar)
 
        logging.info("✔ Botón 'Buscar' clickeado correctamente")
 
        filas = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".x-grid3-body .x-grid3-row"))) 
    
        encontrado = False 
    
        for fila in filas:
           try: 
               poliza = fila.find_element(By.CSS_SELECTOR, ".x-grid3-col-5").text.strip()
               estado = fila.find_element(By.CSS_SELECTOR, ".x-grid3-col-10").text.strip() 
               logging.info(f"Póliza: {poliza}" "|" f"Estado: {estado}") 
               if ramo.poliza in poliza and estado.lower() == "activo": 
                   fila.click() 
                   logging.info("✔ Click en la fila correcta") 
                   encontrado = True 
                   break
 
           except Exception as e: logging.info("Error leyendo fila:", e)
 
        if not encontrado:
            logging.info(f"❌ No se encontró póliza activa que contenga:{ramo.poliza}")
 
        btn_editar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.tb-edit")))
        btn_editar.click()
 
        logging.info("✔️ Click en botón 'Editar'")
 
        btn_movimiento = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Movimientos')]")))
        btn_movimiento.click()
        logging.info("✔ Botón 'Movimientos' clickeado correctamente.")
 
        inp_tipo_mov = wait.until(EC.element_to_be_clickable((By.XPATH, "//fieldset[@id='ext-comp-1389']//input[@size='24']")))
        inp_tipo_mov.click()
        inp_tipo_mov.clear()
        inp_tipo_mov.send_keys("IN")
        inp_tipo_mov.send_keys(Keys.ARROW_DOWN)
        inp_tipo_mov.send_keys(Keys.ENTER)
 
        logging.info("⌨️ Escrito 'IN' y presionado ENTER.")
 
        btn_continuar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.tb-go")))
        btn_continuar.click()
        logging.info("✔️ Click en 'Continuar'")
 
        input_file = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.x-form-file[type='file']")))
        file_path = os.path.abspath(os.path.join(ruta_archivos_x_inclu, "tramavidaleytestsas.xlsx"))
        input_file.send_keys(file_path)
 
        logging.info(f"📤 Archivo cargado correctamente: {file_path}")
 
        # Esperar a que existan los 4 botones
        botones = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "button.tb-save")))
 
        if len(botones) < 4:
            raise Exception(f"❌ Solo se encontraron {len(botones)} botones .tb-save, no 4.")
 
        continuar_btn = botones[3]
 
        wait.until(EC.element_to_be_clickable((By.XPATH, f"(//button[contains(@class,'tb-save')])[4]")))
        driver.execute_script("arguments[0].click();", continuar_btn)
 
        logging.info("✔️ Click en boton CONTINUAR.")
 
        try:
 
            estado_input = wait.until(EC.presence_of_element_located((By.ID, "ext-comp-2271")))
            estado = estado_input.get_attribute("value").strip()
 
            mensaje_area = driver.find_element(By.ID, "ext-comp-2272")
            mensaje = mensaje_area.get_attribute("value").strip()
 
            logging.info(f"📌 Estado detectado: {estado}")
            logging.info(f"📌 Mensaje detectado: {mensaje}")
 
            # Si el estado es ERROR → detener flujo
            if estado.upper() == "ERROR":
                texto_error = f"Estado: {estado}\nMensaje: {mensaje}\n"
 
                # Guardar archivo
                with open("errores_trama.txt", "a", encoding="utf-8") as f:
                    f.write(texto_error + "\n")
 
                # Captura de pantalla
                ruta_imagen = os.path.join(ruta_archivos_x_inclu, f"{get_timestamp()}.png")
                driver.save_screenshot(ruta_imagen)
                logging.info("📸 Captura guardada en:", ruta_imagen)
 
                logging.info("🛑 Proceso detenido por ERROR en validación de trama.")
                return False
 
            logging.info("✔️ No se detectó ERROR en el formulario ExtJS.")
 
        except Exception as e:
            logging.info("❌ No apareció SweetAlert, Angular ni formulario ExtJS:", e)
 
    except Exception as e:
        logging.info(f" ❌ Error en el flujo de selección de póliza en RIMAC SAS: {e}")
        raise Exception(e)
 
    input("Esperar")