#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoAlertPresentException,
    StaleElementReferenceException,
    WebDriverException,
    UnexpectedAlertPresentException
)
from Tiempo.fechas_horas import get_timestamp
from LinuxDebian.Ventana.ventana import esperar_archivos_nuevos
from Chrome.google import tomar_capturar
#---- Import ---
import os
import logging
import time

# --- Variables de Entorno ---
url_pacifico = os.getenv("url_pacifico")

def solicitud_vl(driver,wait,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,ba_codigo,ramo):

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

def solicitud_sctr(driver,wait,numero_poliza,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,ba_codigo,ramo):

    tab_gestion = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#idTabGestion a")))
    driver.execute_script("arguments[0].click();", tab_gestion)
    logging.info("🖱️ Clic en Gestión de Póliza")

    if tipo_proceso == 'IN':
        boton = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space()='Incluye ahora']")
            )
        )
        accion = "Incluye ahora"
    else:
        boton = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space()='Renueva ahora']")
            )
        )
        accion = "Renueva ahora"

    driver.execute_script("arguments[0].click();", boton)
    logging.info(f"🖱️ Clic en '{accion}'")

    # Esperar cualquiera de los resultados posibles
    def esperar_resultado(driver):

        # Error
        error = driver.find_elements(By.XPATH,"//span[@class='titulo' and contains(., 'Por el momento no puedes realizar esta operación')]")

        if error and error[0].is_displayed():
            return "ERROR", error[0]

        # Modal con Aceptar
        modal = driver.find_elements(By.CSS_SELECTOR,"div.modal-content-niche")

        if modal and modal[0].is_displayed():
            return "MODAL", modal[0]

        return False

    resultado, elemento = wait.until(esperar_resultado)

    if resultado == "ERROR":

        logging.warning(f"⚠️ Se detectó el mensaje de error: {elemento.text}")
        raise Exception(elemento.text)


    elif resultado == "MODAL":

        boton_aceptar = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//div[contains(@class, 'modal-footer-niche')]//button[normalize-space()='Aceptar']"
                )
            )
        )

        driver.execute_script("arguments[0].click();",boton_aceptar)
        logging.info("🖱️ Clic en 'Aceptar'")

    containers = wait.until(
        lambda driver: driver.find_elements(
            By.CSS_SELECTOR,
            "div.sctr-pagination-container"
        )
    )
    logging.info(f"✅ Cantidad de contenedores encontrados: {len(containers)}")

    if len(containers) >= 2:
        segundo = containers[1]
        checkbox = segundo.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
        driver.execute_script("arguments[0].click();", checkbox)
        logging.info("🖱️ Clic en el segundo contenedor para seleccionar la póliza amarrada")
    else:
        logging.info("✅ Solo hay un contenedor")

    time.sleep(3)

    # Modal con el boton Aceptar
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.modal-content-niche")))
    boton_aceptar2 = wait.until(EC.presence_of_element_located((By.XPATH, ".//div[contains(@class, 'modal-footer-niche')]//button[normalize-space()='Aceptar']")))
    driver.execute_script("arguments[0].click();", boton_aceptar2)
    logging.info("🖱️ Clic en 'Aceptar'")

    def esperar_siguiente_o_continuar(driver):

        # Siguiente
        botones = driver.find_elements(By.XPATH,"//button[normalize-space()='Siguiente']")

        if botones and botones[0].is_displayed() and botones[0].is_enabled():
            return "SIGUIENTE", botones[0]

        # Sí, continuar
        botones = driver.find_elements(By.XPATH,"//button[normalize-space()='Sí, continuar']")

        if botones and botones[0].is_displayed() and botones[0].is_enabled():
            return "CONTINUAR", botones[0]

        # Si no hay ningún botón, pero ya podemos subir el archivo
        inputs = driver.find_elements(By.ID,"fileInputWorkers")

        if inputs and inputs[0].is_displayed():
            return "ARCHIVO", inputs[0]

        return False

    resultado2, elemento2 = wait.until(esperar_siguiente_o_continuar)

    if resultado2 == "SIGUIENTE":

        driver.execute_script("arguments[0].click();", elemento2)
        logging.info("🖱️ Clic en 'Siguiente'")

        # Ahora esperamos qué viene después
        resultado, elemento = wait.until(esperar_siguiente_o_continuar)

    if resultado2 == "CONTINUAR":

        driver.execute_script("arguments[0].click();", elemento2)
        logging.info("🖱️ Clic en 'Sí, continuar'")

    input_file = wait.until(EC.presence_of_element_located((By.ID, "fileInputWorkers")))
    file_path = os.path.abspath(os.path.join(ruta_archivos_x_inclu,f"{list_polizas[0]}.xlsx"))
    input_file.send_keys(file_path)
    logging.info(f"📄 Trama {list_polizas[0]}.xlsx subida")

    def esperar_resultado_validacion(driver):

        # 1. Ya incluidos / renovados
        elementos = driver.find_elements(By.XPATH,"//div[contains(@class, 'col-9') and contains(., 'Archivo validado.')]")

        if elementos:
            return "YA_PROCESADO", elementos[0]

        # 2. Errores de trama
        elementos = driver.find_elements(By.XPATH,"//div[contains(@class, 'col-9') and contains(., 'Hemos detectado')]")

        if elementos:
            return "ERROR", elementos[0]

        # 3. Validación correcta
        elementos = driver.find_elements(By.XPATH,"//div[contains(text(), 'Archivo validado.')]")

        if elementos:
            return "VALIDADO", elementos[0]

        return False

    resultado, elemento = wait.until(esperar_resultado_validacion)

    #resultado, elemento = wait.until(lambda driver: esperar_resultado_validacion(driver))

    if resultado == "VALIDADO":

        logging.info(f"✅ La trama {list_polizas[0]}.xlsx fue validada correctamente")

        boton_siguiente2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Siguiente')]")))
        boton_siguiente2.click()
        logging.info("🖱️ Clic en 'Siguiente'")


    elif resultado == "ERROR":

        span_text = elemento.find_element(By.TAG_NAME, "span").text

        strong_text = elemento.find_element(By.TAG_NAME, "strong").text

        logging.warning(f"⚠️ {span_text} {strong_text}")

        tomar_capturar(driver,ruta_archivos_x_inclu,f"errores")

        # logging.info("⌛ Descargando los errores de la Trama...")

        # boton_descarga = wait.until(
        #     EC.element_to_be_clickable(
        #         (By.XPATH, "//button[contains(., 'Errores de trama')]")
        #     )
        # )

        raise Exception(f"Trama con errores, revisar")

        # # aquí tu lógica de descarga
        # if click_descarga_documento(driver,boton_descarga,"errores_trama"):
        #     logging.info("✅ Detalle de errores descargado correctamente.")
        #     raise Exception("Trama con errores.")
        # else:
        #     logging.error("❌ Falló la descarga del detalle de errores.")

    elif resultado == "YA_PROCESADO":

        full_text = elemento.text
        logging.warning(f"⚠️ {full_text}")
        textSol = "incluidos" if tipo_proceso == "IN" else "renovados"
        archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

        tomar_capturar(driver,ruta_archivos_x_inclu,textSol)

        boton_descarga = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Descargar detalle')]")
            )
        )

        #----------------
        driver.execute_script("arguments[0].click();", boton_descarga)
        logging.info(f"🖱️ Clic en Descargar")

        archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".xlsx",cantidad=1)

        if archivo_nuevo:
            logging.info(f"✅ Documento descargado exitosamente")
        else:
            logging.warning("❌ Falló la descarga del detalle de trabajadores ya procesados")

        #-----------------
        raise Exception(f"Trama con trabajadores ya {textSol}")
       
    boton_inc_ren = wait.until(EC.element_to_be_clickable((By.XPATH, f'//button[contains(text(), "{ "Incluir" if tipo_proceso == "IN" else "Renovar" }")]')))
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_inc_ren)
    time.sleep(2)
    boton_inc_ren.click()
    logging.info(f'🖱️ Clic en {"Incluir" if tipo_proceso == "IN" else "Renovar"}')

    time.sleep(2)

    # Esperar que aparezca el contenedor
    contenedor = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.content-documents")))

    # Obtener SOLO los botones dentro del div
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

        archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

        # Asegurar que es visible y clickeable
        btn = wait.until(EC.element_to_be_clickable(boton))
        texto = boton.text.strip()

        alias = None  # nombre final para guardar archivo

        # Buscar si alguna palabra clave está dentro del texto del botón
        for palabra, nuevo_nombre in mapa_nombres.items():
            if palabra.lower() in texto.lower():
                alias = nuevo_nombre
                break  # ya encontraste la coincidencia, no sigas

        # Si no coinciden, poner un nombre genérico
        if not alias:
            driver.save_screenshot(os.path.join(ruta_archivos_x_inclu,f"docDesconocido_{get_timestamp()}.png"))
            raise Exception ("Documento desconocido")

        driver.execute_script("arguments[0].click();", btn)
        logging.info(f"🖱️ Clic en Descargar")

        archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=1)

        if archivo_nuevo:
            logging.info(f"✅ Documento '{texto}' descargado exitosamente")
            ruta_original = archivo_nuevo[0]
            ruta_final = os.path.join(ruta_archivos_x_inclu, f"{alias}.pdf")
            os.rename(ruta_original, ruta_final)
            logging.info(f"🔄 {texto} renombrado a '{alias}.pdf'")
        else:
            raise Exception(f"No se descargo ningun archivo")

    time.sleep(2)
    btn_finalizar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Finalizar']")))
    btn_finalizar.click()
    logging.info("🖱️ Clic en 'Finalizar'")

def realizar_solicitud_pacifico(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,ruc_empresa,ejecutivo_responsable,bab_codigo,ramo):
 
    numero_poliza = list_polizas[0]
    tipoError = ""
    detalleError = ""

    driver.get(url_pacifico) 
    logging.info("⌛ Cargando la Web de Pacifico Corredores")
           
    mi_portafolio = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Mi portafolio')]")))
    mi_portafolio.click()
    logging.info("🖱️ Clic en Portafolio")

    somos_corredores = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Somos Corredores')]")))
    somos_corredores.click()
    logging.info("🖱️ Clic en Somos Corredores")

    user_input = wait.until(EC.element_to_be_clickable((By.ID, "i0116")))
    user_input.clear()
    user_input.send_keys(ramo.usuario)
    logging.info("✅ Digitando el Correo")

    boton_next = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    boton_next.click()
    logging.info("🖱️ Clic en 'Next'")
        
    pass_input = wait.until(EC.element_to_be_clickable((By.ID, "i0118")))
    pass_input.clear()
    pass_input.send_keys(ramo.clave)
    logging.info("✅ Digitando el Password")

    ingresar_btn = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    ingresar_btn.click()
    logging.info("🖱️ Se hizo clic en 'Ingresar'")

    sms_option = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@class='table' and @data-value='OneWaySMS']")))
    driver.execute_script("arguments[0].click();", sms_option)
    logging.info("🖱️ Clic en 'Enviar un mensaje de texto'")

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
        logging.info("🧹 Archivo codigo.txt eliminado del volumen.")
    except FileNotFoundError:
        logging.warning("⚠️ No se encontró codigo.txt al intentar eliminarlo (ya fue borrado).")
    except Exception as e:
        logging.error(f"❌ Error al eliminar codigo.txt: {e}")

    ingresar_btn = wait.until(EC.element_to_be_clickable((By.ID, "idSubmit_SAOTCC_Continue")))
    ingresar_btn.click()
    logging.info("🖱️ Verificando el código")

    try:
        boton_conf = WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        boton_conf.click()
        logging.info("🖱️ Clic en 'Yes'")
    except TimeoutException:
        ruta_imagen2 = os.path.join(ruta_archivos_x_inclu,f"{get_timestamp()}.png")
        driver.save_screenshot(ruta_imagen2)
        pass

    for i in range(2):
        try:
            btn_aceptar = WebDriverWait(driver,1.5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.pga-alert-close")))
            btn_aceptar.click()
            time.sleep(3)
            logging.info("🖱️ Se detectó modal y se hizo clic en 'Aceptar'")
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

    error = False
    constancia = False
    proforma = False
    tipoError = ""
    detalleError = ""

    try:

        if bab_codigo in ['1', '2', '3']:
            solicitud_sctr(driver,wait,numero_poliza,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,bab_codigo,ramo)
        else:
            solicitud_vl(driver,wait,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,bab_codigo,ramo)

    except WebDriverException as e:

        error = True
        logging.exception(f"⚠️ Error técnico de Selenium | {e}")
        detalleError = "Problemas Técnicos del Agente"

    except Exception as e:
        error = True
        detalleError = str(e)
    finally:

        if error:
            logging.error(f"❌ Error en Pacifico ({'SCTR' if bab_codigo == 4 else 'VL'}) - {tipo_mes}: {detalleError}")
            return constancia,proforma,f"PACI-{'SCTR' if bab_codigo == 4 else 'VL'}-{tipo_mes}",detalleError
        else:
            return True,True,tipoError,detalleError
