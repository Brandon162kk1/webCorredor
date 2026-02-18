#   --- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from Tiempo.fechas_horas import get_timestamp,get_fecha_menos_x_dias
from LinuxDebian.ventana import desbloquear_interaccion,bloquear_interaccion,esperar_archivos_nuevos
from LinuxDebian.OtrosMetodos.metodos import subir_trama
from Compañias.Sanitas.metodos import imagen_a_pdf
#   --- Imports ----
import time
import os
import logging
import zipfile
import re

def procesar_solicitud_san_crecer_vl(driver,wait,tipo_proceso,ruta_archivos_x_inclu,ejecutivo_responsable,palabra_clave,ruc_empresa,tipo_mes,ramo):
 
    tipoError = ""
    detalleError = ""

    driver.get("https://app.crecerseguros.pe/Plataforma/manage/login")
    logging.info("--- Ingresando a Crecer Seguros🌐 ---")

    user_input = wait.until(EC.presence_of_element_located((By.ID, "suser")))
    user_input.clear()
    user_input.send_keys(ramo.usuario)
    logging.info("⌨️ Digitando el Username")
        
    pass_input = wait.until(EC.presence_of_element_located((By.ID, "spassword")))
    pass_input.clear()
    pass_input.send_keys(ramo.clave)
    logging.info("⌨️ Digitando el Password")

    logging.info("🧩 Resuelve el CAPTCHA manualmente y clic en 'Ingresar'.")

    #desbloquear_interaccion()
    ruta_imagen = os.path.join(ruta_archivos_x_inclu,f"captcha_{get_timestamp()}.png")
    driver.save_screenshot(ruta_imagen)
    wait_humano = WebDriverWait(driver,300)
    wait_humano.until(EC.presence_of_element_located((By.XPATH, "//a[contains(normalize-space(),'Cerrar sesión')]")))

    #bloquear_interaccion()

    logging.info("✅ Login exitoso detectado (Cerrar sesión visible)")
    logging.info("🚀 Continuando flujo automáticamente")

    if tipo_proceso == 'IN':

        try:
            inclusion_crecer_vly(driver,wait,ruta_archivos_x_inclu,ramo)
            return False,True,tipoError,detalleError
        except Exception as e:
            logging.error(f"❌ Error en Crecer (VL) - {tipo_mes}: {e}")
            return False,False,f"CREC-VL-{tipo_mes}",e

    elif tipo_proceso == 'RE':

        try:
            renovacion_crecer_vly(driver,wait,ruta_archivos_x_inclu,ejecutivo_responsable,ramo)
            return False,True,tipoError,detalleError
        except Exception as e :
            logging.error(f"❌ Error en Crecer (VL) - {tipo_mes}: {e}")
            return False,False,f"CREC-VL-{tipo_mes}",e

    else:

        try:
            buscar_cuota_y_descargar_constancia(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ramo)
            return True,False,tipoError,detalleError
        except Exception as e :
            logging.error(f"❌ Error en Crecer (VL) - {tipo_mes}: {e}")
            return False,False,f"CREC-VL-{tipo_mes}",e

def inclusion_crecer_vly(driver,wait,ruta_archivos_x_inclu,ramo):

    span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Gestión de Endoso']")))
    logging.info("🔍 Ubicandose en 'Gestion de Endoso'")
    time.sleep(3)

    actions = ActionChains(driver)
    actions.move_to_element(span_element).click().perform()

    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Generar Endoso Inclusión / Exclusión")))
    link.click()
    logging.info("🖱️ Clic en 'Generar Endoso Inclusión / Exclusión'")

    time.sleep(3)

    input_fullname = wait.until(EC.visibility_of_element_located((By.ID, "spolizanumber")))
    input_fullname.clear()
    input_fullname.send_keys(ramo.poliza)
    logging.info(f"✅ Se ingresó la póliza: {ramo.poliza}")
        
    select_ramos = wait.until(EC.visibility_of_element_located((By.ID, "Ramos")))
    select = Select(select_ramos)
    select.select_by_value("2")
    logging.info("✅ Se seleccionó la opcion Vida Ley")

    select_accion = wait.until(EC.visibility_of_element_located((By.ID, "Action")))
    select2= Select(select_accion)
    select2.select_by_value("1")
    logging.info("✅ Se seleccionó la opción Inclusión")

    time.sleep(2)

    boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Buscar Póliza')]")))
    boton.click()
    logging.info("🖱️ Clic en 'Buscar Póliza")

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "overlay")))

    #Esperar que cargue la tabla con mas de 0 filas y esten visibles
    wait.until(EC.visibility_of_element_located((By.XPATH,"//table[contains(@class, 'table-striped') and contains(@class, 'table-bordered')]//tbody/tr")))

    #CONSULTAR SI VA PARECE QUE NO VA
    # campo_fecha = wait.until(EC.presence_of_element_located((By.ID, "dFechaInicio1")))
    # # Hacer clic para abrir el calendario
    # actions = ActionChains(driver)
    # actions.move_to_element(campo_fecha).click().perform()

    # fecha_actual = fecha.strftime('%d/%m/%Y')
    # campo_fecha.send_keys(fecha_actual)
    # logging.info("✅ Se ingresó la fecha actual")

    time.sleep(3)

    input_file = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Cargar excel')]")))

    # Scroll hacia el botón
    driver.execute_script("arguments[0].scrollIntoView(true);", input_file)

    ruta_trama_xls_sanitas_vl = f"{ruta_archivos_x_inclu}/{ramo.poliza}.xls"
    input_file.send_keys(ruta_trama_xls_sanitas_vl)
    logging.info(f"✅ Trama y/o Excel subido para la póliza: '{ramo.poliza}'")

    try:
        # Modal de Endoso Retroactivo
        boton_ok = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'OK')]")))
        driver.screenshot(os.path.join(ruta_archivos_x_inclu, f"error_subidaTramaCrecer_{ramo.poliza}.png"))
        boton_ok.click()
        raise  Exception("❌ Error al subir la Trama/Excel en Crecer Vida Ley.")
    except :
        pass

    time.sleep(5)

    # Encuentra el botón por su texto
    boton_endoso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Generar Endoso')]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", boton_endoso)
    boton_endoso.click()
    logging.info("🖱️ Clic en 'Generar Endoso'")

    time.sleep(10)

    # Se genera el codigo de pago para enviar al cliente ,capturarlo
    driver.save_screenshot(os.path.join(ruta_archivos_x_inclu, f"codigoPagoCrecer_{ramo.poliza}.png"))

def renovacion_crecer_vly(driver,wait,ruta_archivos_x_inclu,correo,ramo):
    
    span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Gestión de Cotización']")))
    logging.info("🔍 Ubicándose en 'Gestión de Cotización'")
    time.sleep(3)

    actions = ActionChains(driver)
    actions.move_to_element(span_element).click().perform()
    logging.info("🖱️ Clic en 'Gestión de Cotización'")

    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Renovación Vida Ley")))
    link.click()
    logging.info("🖱️ Clic en 'Renovación Vida Ley'")

    time.sleep(3)
        
    input_fullname = wait.until(EC.visibility_of_element_located((By.ID, "spolizanumber")))
    input_fullname.clear()
    input_fullname.send_keys(ramo.poliza)
    logging.info(f"✅ Se ingresó la póliza: {ramo.poliza}")

    boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Busca Póliza')]")))
    boton.click()
    logging.info("🖱️ Clic en 'Busca Póliza")

    try:
        div_poliza= WebDriverWait(driver,8).until(EC.visibility_of_element_located((By.ID, "swal2-content")))
        texto_alerta_poliza = div_poliza.text
        logging.warning(f"⚠️ Mensaje de Advertencia encontrado: {texto_alerta_poliza}")
        driver.screenshot(os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}_noExiste.png"))
        raise Exception(texto_alerta_poliza)    
    except TimeoutException:
        pass

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "overlay")))

    boton_subir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Cargar plantilla')]")))

    ruta_trama_crecer_vl_xlsx = f"{ruta_archivos_x_inclu}/{ramo.poliza}.xlsx"

    if not os.path.exists(ruta_trama_crecer_vl_xlsx):
        raise FileNotFoundError(f"La ruta {ruta_trama_crecer_vl_xlsx} no existe")

    if subir_trama(boton_subir,ruta_trama_crecer_vl_xlsx):               
        time.sleep(3)
        logging.info(f"✅ Trama {ramo.poliza}.xls adjuntada")
    else:
        raise Exception(f"No se pudo subir la trama {ramo.poliza}.xlsx")

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "overlay")))

    try:
        # Modal de error al subir la trama
        div_alerta = WebDriverWait(driver,8).until(EC.visibility_of_element_located((By.ID, "swal2-content")))
        texto_alerta = div_alerta.text
        logging.warning(f"⚠️ Mensaje de Error encontrado: {texto_alerta}")
        driver.screenshot(os.path.join(ruta_archivos_x_inclu, f"errorTrama_{ramo.poliza}.png"))
        raise Exception(texto_alerta)
    except TimeoutException:
        pass

    boton_validar = wait.until(EC.visibility_of_element_located((By.XPATH, '//input[@value="Validar"]')))
    driver.execute_script("arguments[0].scrollIntoView(true);", boton_validar)
    boton_validar.click()
    logging.info("🖱️ Clic en 'Validar'")
        
    time.sleep(5)

    try:

        boton_ok = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'swal2-confirm') and normalize-space()='OK']")))
        boton_ok.click()
        logging.info("🖱️ Clic en 'Ok'")

        wait.until(EC.staleness_of(boton_ok))
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        # logging.info(f"URL actual: {driver.current_url}")

        # if "check-amarillo" in driver.page_source:
        #     logging.info("✅ El checkbox existe en el HTML")
        # else:
        #     raise Exception("El checkbox NO está en el HTML")

        # checkbox = driver.find_element(By.ID, "check-amarillo")
        # logging.info(f"Visible: {checkbox.is_displayed()}")
        # logging.info(f"Habilitado: {checkbox.is_enabled()}")

        checkbox = wait.until(EC.presence_of_element_located((By.ID, "check-amarillo")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox)
        driver.execute_script("arguments[0].click();", checkbox)
        logging.info("🖱️ Checkbox marcado correctamente")

        try:
            btn_aceptar = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='Aceptar']")))
            btn_aceptar.click()
            logging.info("🖱️ Clic en 'Aceptar'")
        except TimeoutException:
            pass

        btn_generar = wait.until(EC.element_to_be_clickable((By.ID, "btngenerate")))
        btn_generar.click()
        logging.info("🖱️ Clic en 'Renovar'")

        mensaje_elemento = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "p.mensaje")))
        texto_mensaje = mensaje_elemento.text
        match = re.search(r"número\s+(\d+)", texto_mensaje)

        if not match:
            raise Exception("No se encontró número de cotizacion en la renovación")

        numero_renovacion = match.group(1)
        logging.info(f"✅ Renovación generada correctamente con Número de Cotizacion: {numero_renovacion}")

        btn_cerrar = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='Cerrar']")))
        btn_cerrar.click()
        logging.info("🖱️ Clic en 'Cerrar'")

        try:

            email_input = wait.until(EC.presence_of_element_located((By.ID, "semailpolicy")))
            wait.until(EC.visibility_of_element_located((By.ID, "semailpolicy")))
            email_input.clear()
            email_input.send_keys(correo)
            logging.info("✍️ Correo ingresado correctamente")

            # Esperar botón Guardar Correo
            btn_guardar = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='Guardar Correo']")))
            btn_guardar.click()
            logging.info("🖱️ Clic en Guardar Correo")

        except TimeoutException:

            btn_ok = wait.until(EC.element_to_be_clickable((By.XPATH,"//button[contains(@class,'swal2-confirm') and normalize-space()='OK']")))
            driver.execute_script("arguments[0].click();", btn_ok)
            logging.info("🖱️ Clic en 'Ok'")

            time.sleep(300)

            #Clic en Pago Efectivo me falta
            #------------------------------------

            contenido = wait.until(EC.visibility_of_element_located((By.ID, "swal2-content")))
            texto_cip = contenido.text
            logging.info(f"Mensaje CIP: {texto_cip}")

            btn_ok = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
            driver.execute_script("arguments[0].click();", btn_ok)
            logging.info("🖱️ Clic en 'Ok'")

            wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-popup")))

            # 1️⃣ Esperar que el iframe esté disponible y cambiar contexto
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@src,'pagoefectivo.pe')]")))
            logging.info("✅ Ya estoy dentro del iframe")

            # Esperar que aparezca el texto indicador
            wait.until(EC.visibility_of_element_located((By.XPATH, "//p[contains(text(),'Código de pago')]")))

            # Ahora obtener el número (solo números largos)
            cip_elemento = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[normalize-space()[string-length()>=8 and number(.)=number(.)]]")))

            codigo_cip = cip_elemento.text
            logging.info(f"CIP obtenido: {codigo_cip}")

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",cip_elemento)
            logging.info("Scroll hasta el codigo para tomar captura")

            cip_elemento.screenshot(os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.png"))

            imagen_a_pdf(os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.png"), os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.pdf"))

            driver.save_screenshot(os.path.join(ruta_archivos_x_inclu,"cod_cip.png"))

            imagen_a_pdf(os.path.join(ruta_archivos_x_inclu,"cod_cip.png"),os.path.join(ruta_archivos_x_inclu,f"endoso_{ramo.poliza}.pdf"))

            # Cuando termines
            driver.switch_to.default_content()
            logging.info("Saliendo del iframe")

    except Exception as e:
        logging.error(f"❌ Error durante el proceso de renovación: {e}")
    finally:
        input("Esperar")

def buscar_cuota_y_descargar_constancia(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ramo,cod_pago='343059073'):

    span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Gestión de Cotización']")))
    actions = ActionChains(driver)
    actions.move_to_element(span_element).click().perform()
    logging.info("🖱️ Clic en 'Gestión de Cotización'")
    time.sleep(3)

    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Bandeja de Evaluación")))
    link.click()
    logging.info("🖱️ Clic en 'Bandeja de Evaluación'")
    time.sleep(3)

    input_fullname = wait.until(EC.visibility_of_element_located((By.ID, "sfullname")))
    input_fullname.clear()
    input_fullname.send_keys(ruc_empresa)
    logging.info(f"🖱️ Se ingresó el RUC '{ruc_empresa}'")

    input_fecha = driver.find_element(By.ID, "solicitud_fecha_inicio")
    driver.execute_script("arguments[0].removeAttribute('readonly')", input_fecha)
    input_fecha.clear()
    input_fecha.send_keys(get_fecha_menos_x_dias(7))
    logging.info("🖱️ Se ingresó un mes antes de la fecha de inicio de la vigencia")

    time.sleep(3)

    input_poliza = wait.until(EC.visibility_of_element_located((By.XPATH,"//label[normalize-space()='Póliza']/following-sibling::input")))
    input_poliza.clear()
    input_poliza.send_keys(ramo.poliza)
    logging.info(f"🖱️ Se ingresó la Póliza '{ramo.poliza}'")

    select_element = wait.until(EC.visibility_of_element_located((By.ID, "Estados")))
    select = Select(select_element)
    select.select_by_value("0")
    logging.info("🖱️ Se selecciono la opción 'Todos'")

    boton_consultar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Consultar')]")))
    driver.execute_script("arguments[0].click();", boton_consultar)
    logging.info("🖱️ Clic en Consultar")
    time.sleep(5)

    wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

    body = driver.find_element(By.TAG_NAME, "body")
    ActionChains(driver).move_to_element(body).click().perform()
    logging.info("🖱️ Clic en un lugar vacío de la página para asegurar la interacción")

    select_elem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "select.select2.form-control.input-sm")))
    Select(select_elem).select_by_value("20")
    logging.info("🖱️ Se seleccionó '20' registros para mostrar más filas")
    time.sleep(5)

    wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

    try:
        table = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.table.table-striped.table-bordered")))
        rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "table.table.table-striped.table-bordered tbody tr")))
        logging.info(f"✅ ¡Tabla cargada con {len(rows)} filas!")
    except TimeoutException:
        driver.save_screenshot(os.path.join(ruta_archivos_x_inclu,f"{ramo.poliza}_{get_timestamp()}.png"))
        raise Exception("❌ La tabla no tiene filas")

    fila_encontrada = False

    logging.info("-------------------------------------------------")
    logging.info(f"🔍 Buscando el código de pago '{cod_pago}'")

    for row in rows:

        cells = row.find_elements(By.TAG_NAME, "td")

        if len(cells) < 15:
            continue

        codigo_pago = cells[15].text.strip()
        estado_pago = cells [16].text.strip()    

        if codigo_pago == cod_pago :

            fila_encontrada = True
            logging.info("✅ Fila encontrada")

            if estado_pago.lower() == "pagado":

                if len(cells) > 18:

                    constancias_zip = cells[9]

                    try:

                        link = constancias_zip.find_element(By.TAG_NAME, "a")
                        driver.execute_script("arguments[0].click();", link)
                        logging.info("🖱️ Clic con JS en el link")

                        wait.until(EC.invisibility_of_element_located((By.XPATH, "//ngx-spinner")))

                        archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

                        btn_descargar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Descargar Póliza y Constancia']")))
                        btn_descargar.click()
                        logging.info("🖱️ Clic en Descargar Póliza y Constancia")

                        zips = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,extension=".zip",cantidad_esperada=1,timeout=90)

                        if not zips:
                            raise Exception("No se descargó ningún ZIP")

                        ruta_zip = zips[0]
                        logging.info(f"✅ ZIP detectado: {ruta_zip}")

                        #try:

                        if not os.path.exists(ruta_zip):
                            raise Exception(f"No existe el ZIP descargado: {ruta_zip}")

                        with zipfile.ZipFile(ruta_zip, 'r') as z:

                            # Filtrar PDFs que contengan la palabra "Constancia"
                            pdf_constancia = [
                                f for f in z.namelist()
                                if f.lower().endswith(".pdf") and "constancia" in f.lower()
                            ]

                            if not pdf_constancia:
                                raise Exception("No se encontró ningún PDF de Constancia dentro del ZIP")

                            # Si hubiera más de uno, tomamos el primero que coincida
                            pdf_interno = pdf_constancia[0]

                            nuevo_nombre_pdf = f"{ramo.poliza}.pdf"
                            ruta_pdf_salida = os.path.join(ruta_archivos_x_inclu, nuevo_nombre_pdf)

                            # Extraer directamente el archivo
                            with z.open(pdf_interno) as archivo_origen, open(ruta_pdf_salida, "wb") as archivo_destino:
                                archivo_destino.write(archivo_origen.read())

                        logging.info(f"✅ PDF de Constancia extraído correctamente en: {ruta_pdf_salida}")

                        # finally:
                        #     if os.path.exists(ruta_zip):
                        #         os.remove(ruta_zip)
                        #         logging.info("🧹 Archivo ZIP eliminado correctamente")
                        #     else:
                        #         raise Exception("El archivo ZIP no existe")


                    except Exception as e:
                        raise Exception(f"No se pudo descargar la carpeta ZIP, Cuota con estado {estado_pago}")
            else:
                raise Exception(f"El código de pago '{cod_pago}' tiene estado : '{estado_pago}'")

    if not fila_encontrada:
        raise Exception(f"No se encontró el código de pago '{cod_pago}'")