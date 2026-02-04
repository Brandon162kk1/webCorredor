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

def click_descarga_documento(driver,boton_descarga,nombre_documento):
    
    try:

        driver.execute_script("arguments[0].click();", boton_descarga)
        logging.info("✅ Se hizo clic con JS en el botón de descarga.")
        logging.info("⌛ Esperando la ventana descarga de Linux Debian...")
        time.sleep(2)
        subprocess.run(["xdotool", "search", "--name", "Save File", "windowactivate", "windowfocus"])
        logging.info("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
        subprocess.run(["xdotool", "type","--delay", "100", nombre_documento])
        logging.info("📄 Se cambio el nombre del documento")
        time.sleep(2)
        subprocess.run(["xdotool", "key", "Return"])
        logging.info("🖱️ Se presionó 'Enter' para confirmar la descarga.")
        time.sleep(2)
        return True

    except Exception as ex:
        logging.error(f"❌ Error durante el flujo de descarga: {ex}")
        return False

def solicitud_vl(driver,wait,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,fVigenciaInicio,fVigenciaFin,ba_codigo):
    
    tab_endosos = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#idTabGenerarEndosos a")))
    driver.execute_script("arguments[0].click();", tab_endosos)
    logging.info("🖱️ Clic en Generar Endosos")

    try:
        # Intentar detectar el modal de error
        modal_error = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH,"//div[contains(@class,'mensaje-descarga') and contains(@class,'error')]")))
    
        titulo_error = modal_error.find_element(By.TAG_NAME, "h3").text
        logging.error(f"❌ Se detectó un error inesperado en la web: {titulo_error}")

        raise Exception(titulo_error)

    except TimeoutException:
        # No apareció error → continuar flujo normal
        logging.info("✔ No se encontró modal de error inesperado")

    tab_personas = wait.until(EC.element_to_be_clickable((By.ID, "idTabInclusionPersonas")))
    driver.execute_script("arguments[0].click();", tab_personas)
    logging.info("🖱️ Clic en Inclusión de Personas (por li)")

def solicitud_sctr(driver,wait,numero_poliza,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,
                   fVigenciaInicio,fVigenciaFin,ba_codigo):

    # Clic en la pestaña de Gestión
    tab_gestion = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#idTabGestion a")))
    driver.execute_script("arguments[0].click();", tab_gestion)
    logging.info("🖱️ Clic en Gestión de Póliza")

    try:
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='titulo' and contains(., 'Un momento por favor')]")))
        logging.info("⌛ Cargando")
    except:
        logging.info("✅ Ya desapareció antes de que empiece a esperar")

    if tipo_proceso == 'IN':
        btn_incluir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Incluye ahora']")))
        time.sleep(3)
        btn_incluir.click()
        logging.info(f"🖱️ Clic en el boton 'Incluye ahora'.")
    else:
        btn_renovar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Renueva ahora']")))
        time.sleep(3)
        btn_renovar.click()
        logging.info(f"🖱️ Clic en el boton 'Renueva ahora'.")
     
    try:
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='titulo' and contains(., 'Un momento por favor')]")))
        logging.info("⌛ Cargando")
    except:
        logging.info("✅ Ya desapareció antes de que empiece a esperar")

    try:
        span_error = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[@class='titulo' and contains(text(), 'Por el momento no puedes realizar esta operación')]")))
        logging.warning(f"⚠️ Se detectó el mensaje de error: {span_error.text}")
        raise Exception(f"{span_error.text}")
    except TimeoutException:
        logging.info("✅ No apareció el mensaje de error — se continúa con el flujo normal.")
        pass

    time.sleep(3)

    boton_aceptar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Aceptar')]")))
    boton_aceptar.click()
    logging.info("🖱️ Clic en 'Aceptar'.")

    try:
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='titulo' and contains(., 'Un momento por favor')]")))
        logging.info("⌛ Cargando")
    except:
        logging.info("Ya desapareció antes de que empiece a esperar")

    time.sleep(3)

    containers = driver.find_elements(By.CSS_SELECTOR, "div.sctr-pagination-container")
    logging.info(f"Cantidad de contenedores encontrados: {len(containers)}")

    if len(containers) >= 2:
        segundo = containers[1]
        checkbox = segundo.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
        driver.execute_script("arguments[0].click();", checkbox)
        logging.info("Clic en el segundo contenedor para seleccionar la poliza amarrada")
    else:
        logging.info("Solo hay un contenedor.")

    time.sleep(3)

    boton_aceptar2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Aceptar')]")))
    boton_aceptar2.click()
    logging.info("🖱️ Clic en 'Aceptar'.")

    try:
        boton_siguiente = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Siguiente')]")))
        time.sleep(3)
        boton_siguiente.click()
        logging.info("🖱️ Clic en 'Siguiente'.")
    except  TimeoutException:
        logging.info("❌ No salio el boton 'Siguiente'.")

    try:
        wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='titulo' and contains(., 'Un momento por favor')]")))
        logging.info("⌛ Cargando")
    except:
        logging.info("Ya desapareció antes de que empiece a esperar")

    try:
        boton_si_continuar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Sí, continuar')]")))
        time.sleep(3)
        boton_si_continuar.click()
        logging.info("🖱️ Clic en 'Sí, continuar'.")
    except TimeoutException:
        logging.info("✅ No salio un modal, continunado")

    input_file = wait.until(EC.presence_of_element_located((By.ID, "fileInputWorkers")))
    file_path = os.path.abspath(os.path.join(ruta_archivos_x_inclu,f"{list_polizas[0]}.xlsx"))
    input_file.send_keys(file_path)
    logging.info(f"📄 Trama {list_polizas[0]}.xlsx subida.")

    try:
        # Esperar directamente a que aparezca el texto exacto "Archivo validado."
        wait.until(EC.text_to_be_present_in_element((By.XPATH, "//div[contains(text(), 'Archivo validado.')]"),"Archivo validado."))

        # Si llega hasta aquí, el texto realmente apareció
        logging.info(f"✅ La trama {list_polizas[0]}.xlsx fue validada correctamente.")

        boton_siguiente2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Siguiente')]")))
        time.sleep(3)
        boton_siguiente2.click()
        logging.info("🖱️ Clic en 'Siguiente'.")

    except TimeoutException:
        logging.info("⏳ No apareció el mensaje de validación de éxito. Verificando si hay errores...")

        try:
            error_div = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'col-9') and contains(., 'Hemos detectado')]")))

            span_text = error_div.find_element(By.TAG_NAME, "span").text
            strong_text = error_div.find_element(By.TAG_NAME, "strong").text

            logging.warning(f"⚠ {span_text} {strong_text}")

            ruta_imagen = os.path.join(ruta_archivos_x_inclu, f"{get_timestamp()}.png")
            driver.save_screenshot(ruta_imagen)

            logging.info("➡️ Intentando descargar los errores de la Trama...")

            boton_descarga = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Errores de trama')]")))
            if click_descarga_documento(driver,boton_descarga,"errores_trama"):
                logging.info("✅ Detalle de errores descargado correctamente.")
                raise Exception("Trama con errores.")
            else:
                logging.error("❌ Falló la descarga del detalle de errores.")

        except TimeoutException:

            try:
                error_div2 = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class, 'col-9') and contains(., 'Archivo validado.')]")))

                strong_text = error_div2.find_element(By.TAG_NAME, "strong").text
                full_text = error_div2.text

                logging.warning(f"⚠️ {full_text}")

                ruta_imagen = os.path.join(ruta_archivos_x_inclu, f"{get_timestamp()}.png")
                driver.save_screenshot(ruta_imagen)

                textSol = "incluidos" if tipo_proceso == 'IN' else "renovados"
                raise Exception(f"Trama con trabajadores ya {textSol}.")

            except TimeoutException:
                logging.warning("❌ No se encontró ni mensaje de validación ni mensaje de error.")
            
    boton_inc_ren = wait.until(EC.element_to_be_clickable((By.XPATH, f'//button[contains(text(), "{ "Incluir" if tipo_proceso == "IN" else "Renovar" }")]')))
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_inc_ren)
    time.sleep(3)
    boton_inc_ren.click()
    logging.info(f'🖱️ Clic en {"Incluir" if tipo_proceso == "IN" else "Renovar"}')

    time.sleep(3)

    # 1. Esperar que aparezca el contenedor
    contenedor = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.content-documents")))

    # 2. Obtener SOLO los botones dentro del div
    botones = contenedor.find_elements(By.TAG_NAME, "button")

    logging.info("Total botones encontrados:", len(botones))

    # Base común para ambos casos
    mapa_nombres = {
        "constancia": f"{list_polizas[0]}",
        "factura": f"factura{list_polizas[0]}"
    }

    # Cambios solo para "liquidación"
    if ba_codigo == '3':
        mapa_nombres["liquidación"] = f"endoso_{list_polizas[1]}"  # Pensiones
    else:
        mapa_nombres["liquidación"] = f"endoso_{list_polizas[0]}"  # Salud o general

    # 3. Hacerles clic uno por uno
    for boton in botones:
        # Asegurar que es visible y clickeable
        btn = wait.until(EC.element_to_be_clickable(boton))
        texto = boton.text.strip()

        alias = None  # nombre final para guardar archivo

        # Buscar si alguna palabra clave está dentro del texto del botón
        for palabra, nuevo_nombre in mapa_nombres.items():
            if palabra in texto:
                alias = nuevo_nombre
                break  # ya encontraste la coincidencia, no sigas

        # Si no coinciden, poner un nombre genérico
        if not alias:
            driver.save_screenshot(os.path.join(ruta_archivos_x_inclu,f"docDesconocido_{get_timestamp()}.png"))
            raise Exception ("Documento desconocido")

        if click_descarga_documento(driver,btn,alias):
            time.sleep(3)
            ruta_doc_descargado = f"{ruta_archivos_x_inclu}/{alias}.pdf"
            if os.path.exists(ruta_doc_descargado):
                logging.info(f"✅ {alias}.pdf descargado exitosamente")
            else:
                logging.warning(f"⚠️ No se encontró el archivo descargado en {ruta_archivos_x_inclu}")
        else:
            mensaje_descarga = "No se logró descargar la constancia. Verifica manualmente."
            raise Exception (f"{mensaje_descarga}")

    time.sleep(2)
    btn_finalizar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Finalizar']")))
    btn_finalizar.click()
    logging.info("🖱️ Clic en 'Finalizar'.")

    input("--Esperando Pruebas---")
    # te regressa a la pagina de renovacion e inclusion en SCTR

def realizar_solicitud_pacifico(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,
                                ejecutivo_responsable,fVigenciaInicio,fVigenciaFin,bab_codigo):
    

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

    codigo_path = "/codigo/codigo.txt"

    while not os.path.exists(codigo_path):
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
