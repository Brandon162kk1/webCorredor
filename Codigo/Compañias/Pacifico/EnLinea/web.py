#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from Tiempo.fechas_horas import get_timestamp
#---- Import ---
import os
import logging
import time
import subprocess

# --- Variables de Entorno ---
url_pacifico = os.getenv("url_pacifico")
correo = os.getenv("remitente")
password = os.getenv("passwordCorreo")

def realizar_solicitud_vl(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,ejecutivo_responsable,fVigenciaInicio,fVigenciaFin,bab_codigo):
    
    numero_poliza = list_polizas[0]
    tipoError = ""
    detalleError = ""

    driver.get(url_pacifico) 
    logging.info("⌛ Cargando la Web de Pacifico")
           
    mi_portafolio = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Mi portafolio')]")))
    mi_portafolio.click()
    logging.info("🖱️ Clic en Portafolio.")

    somos_corredores = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Somos Corredores')]")))
    somos_corredores.click()
    logging.info("🖱️ Clic en Somos Corredores")

    user_input = wait.until(EC.element_to_be_clickable((By.ID, "i0116")))
    user_input.clear()
    user_input.send_keys(correo)
    logging.info("✅ Digitando el Correo")

    boton_next = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    boton_next.click()
    logging.info("🖱️ Clic en 'Next'.")
        
    pass_input = wait.until(EC.element_to_be_clickable((By.ID, "i0118")))
    pass_input.clear()
    pass_input.send_keys(password)
    logging.info("✅ Digitando el Password")

    ingresar_btn = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    ingresar_btn.click()
    logging.info("🖱️ Se hizo clic en 'Ingresar'.")

    sms_option = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@class='table' and @data-value='OneWaySMS']")))
    driver.execute_script("arguments[0].click();", sms_option)
    logging.info("🖱️ Clic en 'Enviar un mensaje de texto'.")

    codigo_path = "/shared/codigo.txt"

    while not os.path.exists(codigo_path):
        #print("⏳ Esperando que se cree codigo.txt...")
        time.sleep(2)

    with open(codigo_path, "r") as f:
        codigo = f.read().strip()

    logging.info(f"✅ Código recibido desde volumen: {codigo}")

    clave_sms = wait.until(EC.element_to_be_clickable((By.ID, "idTxtBx_SAOTCC_OTC")))
    clave_sms.clear()
    clave_sms.send_keys(codigo)
    logging.info("✅ Digitando el código")

    # --- Eliminar el archivo después de usarlo ---
    try:
        os.remove(codigo_path)
        #logging.info("🧹 Archivo codigo.txt eliminado del volumen.")
    except FileNotFoundError:
        logging.warning("⚠️ No se encontró codigo.txt al intentar eliminarlo (ya fue borrado).")
    except Exception as e:
        logging.error(f"❌ Error al eliminar codigo.txt: {e}")

    ingresar_btn = wait.until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue")))
    ingresar_btn.click()
    logging.info("🖱️ Verificando el código.")

    try:
        boton_conf = WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        boton_conf.click()
        logging.info("🖱️ Clic en 'Yes'.")
    except TimeoutException:
        ruta_imagen2 = os.path.join(ruta_archivos_x_inclu,f"{get_timestamp()}.png")
        driver.save_screenshot(ruta_imagen2)
        pass

    for i in range(2):
        try:
            btn_aceptar = WebDriverWait(driver,1.5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.pga-alert-close")))
            btn_aceptar.click()
            time.sleep(3)
            logging.info("🖱️ Se detectó modal y se hizo clic en 'Aceptar'.")
            break
        except TimeoutException:
            pass

    poliza_input = wait.until(EC.presence_of_element_located((By.ID, "inputBuscador")))
    poliza_input.clear()
    poliza_input.send_keys(numero_poliza)
    logging.info(f"✅ Póliza ingresada: {numero_poliza}")

    buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "busqueda")))
    buscar_btn.click()
    logging.info("🖱️ Se hizo clic en 'Buscar'.")

    wait.until(EC.visibility_of_element_located((By.ID,"tablaPoliza")))
    logging.info("⌛ Esperando que cargue la tabla...")

    # select_elem = wait.until(EC.presence_of_element_located((By.NAME, "tablaPoliza_length")))
    # Select(select_elem).select_by_value("100")
    # print("🖱️ Se seleccionó '100' registros para mostrar más filas.")

    try:
        wait.until(lambda d: len(d.find_elements(By.XPATH, "//table[@id='tablaPoliza']//tr")) > 1)
        logging.info("✅ La tabla tiene al menos 1 fila.")
    except TimeoutException:
        logging.warning("⏰ No se cargaron las filas en el tiempo esperado.")

    table = driver.find_element(By.ID, "tablaPoliza")
    rows = table.find_elements(By.XPATH, ".//tbody//tr")

    fila_encontrada = False

    for i, row in enumerate(rows):

        cells = row.find_elements(By.TAG_NAME, "td")

        if len(cells) < 11:
            logging.warning(f"⚠️ Fila {i} ignorada (tiene {len(cells)} columnas)")
            continue
     
        poliza_cia = cells[3].text.strip()
        inicio_vigencia_cia = cells[5].text.strip()
        fin_vigencia_cia = cells[6].text.strip()
        estado_cia = cells[9].text.strip()
       
        if numero_poliza == poliza_cia:

            logging.info(f"✅ Fila encontrada: Poliza='{poliza_cia}', Estado='{estado_cia}', Inicio de Vigencia='{inicio_vigencia_cia}', Fin de Vigencia='{fin_vigencia_cia}'.")
        
            enlace_poliza = cells[3].find_element(By.TAG_NAME, "a")
            wait.until(EC.element_to_be_clickable(enlace_poliza))
            driver.execute_script("arguments[0].click();", enlace_poliza)
            logging.info("🖱️ Clic (por JS) en la póliza")

            fila_encontrada = True

            break

    if not fila_encontrada:
        raise Exception(f"❌ No se encontró la póliza {numero_poliza} en la tabla.")   #Mejorar esta logica

    try:
        wait.until(EC.visibility_of_element_located((By.ID, "ajax-loading")))
    except TimeoutException:
        pass
    logging.info("⌛ Cargando...")

    wait.until(EC.invisibility_of_element_located((By.ID, "ajax-loading")))
    logging.info("⌛ Loader desapareció, continuando...")

    if bab_codigo in ['1', '2', '3']:

        try:
            solicitud_sctr(driver,wait,numero_poliza,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,fVigenciaInicio,fVigenciaFin,bab_codigo)
            return True, tipoError, detalleError
        except Exception as e:
            logging.error(f"❌ Error en Pacifico (SCTR) - {tipo_mes}: {e}")
            return False, f"PACI-SCTR-{tipo_mes}", e

    else:

        try:
            solicitud_vl(driver,wait,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,fVigenciaInicio,fVigenciaFin,bab_codigo)
            return True, tipoError, detalleError
        except Exception as e:
            logging.error(f"❌ Error en Pacifico (VL) - {tipo_mes}: {e}")
            return False, f"PACI-VL-{tipo_mes}", e
