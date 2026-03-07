#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,NoAlertPresentException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from Tiempo.fechas_horas import get_fecha_hoy
from Compañias.Positiva.metodos import mover_y_hacer_click_simple,escribir_lento,validar_pagina,validardeuda,leer_pdf
from LinuxDebian.Ventana.ventana import esperar_archivos_nuevos
from Chrome.google import tomar_capturar
#from Apis.Post.webhook import enviar_constancia_anular
#---- Import ---
import os
import re
import logging
import time
import random
import shutil

# --- Variables globales ---
ventana_menu_positiva = None

def solicitud_sctr(driver,wait,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,ba_codigo,ramo):
    
    sed = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='SED']")))
    sed.click()
    logging.info("🖱️ Clic en 'SED'")

    wait.until(lambda d: len(d.window_handles) > 1)

    driver.switch_to.window(driver.window_handles[-1])
    logging.info("⌛ Esperando la nueva ventana...")

    #Metodo para buscar y descargar constancias si es que falla la web

    transacciones = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Transacciones y Consultas")))
    transacciones.click()
    logging.info("🖱️ Clic en Transacciones y Consultas")

    time.sleep(3)

    try:
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, "//h3[contains(text(),'Lo sentimos, ha ocurrido un error inesperado')]")))
        raise Exception("Lo sentimos, ha ocurrido un error inesperado")
    except TimeoutException:
        pass

    poliza_input = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_txtNroPoliza")))
    poliza_input.clear()
    poliza_input.send_keys(ramo.poliza)
    logging.info(f"✅ Se Ingresó la póliza '{ramo.poliza}' en el campo")

    buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnBuscar")))
    buscar_btn.click()
    logging.info("🖱️ Clic en Buscar")

    try:
        wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado")))
        logging.info("⌛ Esperando la tabla con los campos...")
    except Exception as e:
        raise Exception(f"No se encontró la póliza {ramo.poliza} en la compañía")

    # Obtener la fecha de vigencia (formato: dd/mm/yyyy)
    fecha_vigencia_element =wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado_lblFechaVigencia_0")))
    fecha_vigencia_str = fecha_vigencia_element.text.strip()  # Ejemplo: "01/11/2025"
    
    # Convertir la fecha a objeto datetime
    fecha_vigencia_dmy = datetime.strptime(fecha_vigencia_str, "%d/%m/%Y")
    try:
        # Definir rango permitido
        rango_inicio = fecha_vigencia_dmy.date() - timedelta(days=2)  # 2 días antes
        rango_fin = fecha_vigencia_dmy.date() + timedelta(days=2)     # 2 días después

        logging.info(f"📅 Rango permitido para la {palabra_clave}: {rango_inicio} a {rango_fin}")

        logging.info(f"📅 Fecha Fin de la vigencia en la póliza: {fecha_vigencia_str}")
        #pos_fecha_dmy = get_pos_fecha_dmy()
        #logging.info(f"📅 Fecha de hoy: {pos_fecha_dmy}")

    except Exception as e:
        raise Exception(f"Error al parsear la fecha de vigencia,Motivo - {e}")

    if ramo.poliza == '0' : 
        raise Exception(f"Fuera del rango permitido, máximo son 2 días de diferencia para la {palabra_clave}")
    else:

        radio_seleccion = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_gvResultado_rbSeleccion_0")))
        radio_seleccion.click()
        logging.info("🖱️ Clic en el 'Radio Button' correspondiente")
        
        wait.until(EC.invisibility_of_element_located((By.ID, "ID_MODAL_PROCESS")))
        logging.info("⌛ Esperando que cargue...")

        if tipo_proceso == 'IN':  
            
            btn_incluir = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnIncluir")))
            btn_incluir.click()
            logging.info("🖱️ Clic en Incluir")

            try:
                WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "divTipoIncluir")))
                logging.info(f"⚠️ Apareció el modal con advertencia")
                
                if tipo_mes == 'MA':
                    btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnIncluirProformaSi")))
                else:
                    btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnIncluirProformaNo")))
                    
                nom_btn = 'Si' if tipo_mes == 'MA' else 'No'
                btn_aceptar.click()
                logging.info(f"🖱️ Clic en {nom_btn}")

            except TimeoutException:
                pass

            mensaje_deuda = validardeuda(driver,wait)

            if mensaje_deuda:
                raise Exception(f"{mensaje_deuda}")

            input_file = wait.until(EC.presence_of_element_located((By.ID, "fuPlanillaAjax")))
            logging.info("⌛ Esperando que cargue la nueva página donde se adjunta la Trama...")
      
            input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniFecVigInc")))
            input_fecha.clear()

            driver.execute_script("""
                arguments[0].value = arguments[1]; 
                arguments[0].dispatchEvent(new Event('change'));
                arguments[0].dispatchEvent(new Event('blur'));
            """, input_fecha, ramo.f_inicio)

            try:
                WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID, "divAlertas")))
                logging.info(f"⚠️ Apareció el modal con advertencia")
                mensaje_fecha = wait.until(EC.visibility_of_element_located((By.ID, "spMensaje"))).text
                raise Exception(mensaje_fecha)
            except TimeoutException:
                pass

            driver.find_element(By.TAG_NAME, "body").click()

            fecha_date = datetime.strptime(ramo.f_inicio,"%d/%m/%Y").date()

            if fecha_date < get_fecha_hoy().date():
                diferencia = (get_fecha_hoy().date() - fecha_date).days
                cant_dias = abs(diferencia)
                text_dia = "día" if cant_dias == 1 else "días"
                logging.info(f"📅 Vigencia de Inicio retroactiva ingresada: {ramo.f_inicio} con {cant_dias} {text_dia} de diferencia")
            else:
                logging.info(f"📅 Vigencia de Inicio ingresada: {ramo.f_inicio}")

            if ramo.poliza == '9231375':

                select_tipo_calculo = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_cboTipoCalculoEspecial")))
                select_tp = Select(select_tipo_calculo)
                select_tp.select_by_value("21")
                logging.info("✅ Se seleccionó la opción 'Por días' ")

        else:

            btn_incluir = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnRenovar")))
            btn_incluir.click()
            logging.info("🖱️ Clic en Renovar")

            mensaje_deuda = validardeuda(driver,wait)

            if mensaje_deuda:
                raise Exception(f"{mensaje_deuda}")

            wait.until(EC.visibility_of_element_located((By.ID, "divTipoRenovacion")))

            btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnRenovacionSi")))

            btn_aceptar.click()
            logging.info(f"🖱️ Clic en aceptar la Renovación")

            input_file = wait.until(EC.presence_of_element_located((By.ID, "fuPlanillaAjax")))
            logging.info("⌛ Esperando que cargue la nueva página donde se adjunta la Trama...")

            # Solo para este cliente: P & Q TELECOM EIRL en particular Salud: 9231375 Pensión: 30358377 
            # pq las otras polizas se ponen por defecto 1 mes
            if ramo.poliza == '9231375':

                fecha_str = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniVigPoliza"))).get_attribute("value")

                fecha = datetime.strptime(fecha_str, "%d/%m/%Y")

                fecha_siguiente_mes = fecha + relativedelta(months=1)

                resultado = fecha_siguiente_mes - timedelta(days=1)

                fecha_final = resultado.strftime("%d/%m/%Y")

                input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtFinVigPoliza")))

                input_fecha.clear()
                driver.execute_script("""
                    arguments[0].value = arguments[1]; 
                    arguments[0].dispatchEvent(new Event('change'));
                    arguments[0].dispatchEvent(new Event('blur'));
                """, input_fecha, fecha_final)

                # Clic fuera del input (por ejemplo, el body) para cerrar el calendario
                driver.find_element(By.TAG_NAME, "body").click()
                logging.info(f"📅 Vigencia de Póliza ingresada con JS Hasta : {fecha_final}")

        time.sleep(2)

        file_path = os.path.abspath(os.path.join(ruta_archivos_x_inclu,f"{ramo.poliza}.xlsx")) 

        if not os.path.exists(file_path):
            raise Exception (f"Archivo {ramo.poliza}.xlsx no encontrado")
        else:
            input_file.send_keys(file_path)
            logging.info(f"✅ Se subió el archivo '{ramo.poliza}.xlsx' para validar")

        wait.until(EC.invisibility_of_element_located((By.ID, "ID_MODAL_PROCESS")))
        logging.info("⌛ Esperando que cargue")

        try:
            validar_btn = wait.until( EC.visibility_of_element_located((By.ID, "btnValidarPlanillaExcel")))
            validar_btn.click()
            logging.info("🖱️ Clic en Validar Planilla")
        except Exception as e:
            raise Exception("No se puede Validar la Trama")

        try:
            WebDriverWait(driver,5).until(EC.alert_is_present())
            alerTrama = driver.switch_to.alert   
            logging.info("🚨 Texto de la alerta:", alerTrama.text)   
            alerTrama.accept()
            logging.info("✅ Alerta aceptada")
            raise Exception("Posiblemente un campo de la Trama no este correcta, verificarla porfavor")
        except TimeoutException:
            pass

        try:
            # Posible error para la hoja de la trama , debe ser Planilla no Trabajadores
            WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "divAlertaErrorValidacion")))
            logging.info(f"⚠️ Apareció el modal con errores")
            mensaje_error = wait.until(EC.visibility_of_element_located((By.ID, "spanAlertaErrorValidacion"))).text
            raise Exception(mensaje_error)
        except TimeoutException:

            try:
                WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "btnErroresPlanilla")))

                cantidad_errores = wait.until(EC.visibility_of_element_located((By.ID, "spnContadorError"))).text
                cantidad_texto = "error" if cantidad_errores == '1' else "errores"

                link = wait.until(EC.element_to_be_clickable((By.ID, "btnErroresPlanilla")))
                driver.execute_script("arguments[0].click();", link)

                wait.until(EC.visibility_of_element_located((By.ID, "divErrorPlanilla")))

                raise Exception(f"Se encontró {cantidad_errores} {cantidad_texto} en la planilla")

            except TimeoutException:
                pass

        alerta_detectada = False

        try:
            WebDriverWait(driver,5).until(EC.alert_is_present())
            alert = driver.switch_to.alert   
            logging.info("🚨 Texto de la alerta:", alert.text)   
            alert.accept()
            alerta_detectada = True
        except TimeoutException:
            logging.info("✅ No apareció ninguna alerta en el tiempo especificado.")
            
        if alerta_detectada:
            raise Exception("Hubo una alerta con la Trama")
        else:
                             
            procesar_btn = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnProcesar")))

            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", procesar_btn)

            procesar_btn.click()
            logging.info(f"🖱️ Clic en Procesar {palabra_clave}") 

            texto_mensaje = ""

            if tipo_proceso == 'IN':

                #aca carga un modal indicado si estas seguro de realziar la operacion de Inclusion

                btn_si = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnIncluirSi")))
                btn_si.click()
                logging.info("🖱️ Clic en Aceptar la Inclusión")

                time.sleep(3)

                mensaje_span = wait.until(EC.visibility_of_element_located((By.ID, "spanAlertaInclusion")))
                texto_mensaje = mensaje_span.text

            else:
                
                mensaje_span = wait.until(EC.visibility_of_element_located((By.ID, "spanAlertaRenovacion")))
                texto_mensaje = mensaje_span.text

                btn_si = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAceptarRenovacion")))
                btn_si.click()
                logging.info("🖱️ Clic en Aceptar la Renovación")

                time.sleep(3)
            
            logging.info(f"📩 Texto del mensaje: {texto_mensaje}")

            numero_doc = None

            if tipo_proceso == 'IN':

                match = re.search(r'Inclusión con nro\. (\d+)', texto_mensaje)
                if match:
                    numero_doc = match.group(1)
                    logging.info(f"Número de {palabra_clave} extraído: {numero_doc}")
                else:
                    logging.info(f"No se encontró el código de {palabra_clave}.")

            else:

                match = re.search(r'renovación con nro\. (\d+)', texto_mensaje)
                if match:
                    numero_doc = match.group(1)
                    logging.info(f"Número de {palabra_clave} extraído: {numero_doc}")
                else:
                    logging.info(f"No se encontró el código de {palabra_clave}.")

            prefijo = "INCL" if tipo_proceso == "IN" else "RENV"

            try:

                if tipo_proceso == 'IN':
                    btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAceptarInclusion")))
                    btn_aceptar.click()
                    time.sleep(3)
                else:
                    try:
                        btn_aceptar = WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAceptarRenovacion")))
                        btn_aceptar.click()
                    except TimeoutException:
                        logging.info(f"✅ No aparecio ningun otro boton.")

                wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-widget-overlay")))

                span_numero = wait.until(EC.visibility_of_element_located((By.XPATH, f"//span[contains(text(), '{prefijo}-{numero_doc}')]")))
                logging.info(f"✅ Span encontrado: {span_numero.text}")

                try:
                    # Esperar que aparezca la lupa (por cualquiera de los dos tipos) -- Esto puede fallar si solo es una póliza
                    if len(list_polizas) == 1 and ba_codigo == '1':
                        selector_xpath = (f"//img[@data-nropolizasalud='{ramo.poliza}']")
                    elif len(list_polizas) == 1 and ba_codigo == '2':
                        selector_xpath = (f"//img[@data-nropolizapension='{ramo.poliza}']")
                    else:
                        selector_xpath = (f"//img[@data-nropolizasalud='{list_polizas[0]}' or @data-nropolizapension='{list_polizas[1]}']")

                    lupa_element = wait.until(EC.element_to_be_clickable((By.XPATH, selector_xpath)))

                    driver.execute_script("arguments[0].scrollIntoView(true);", lupa_element)
                    lupa_element.click()

                    logging.info(f"🔎 Lupa encontrada y clickeada para póliza {ramo.poliza}")

                except TimeoutException:
                    raise Exception(f"No se encontró la lupa asociada a la póliza {ramo.poliza}")
                except Exception as e:
                    raise Exception(f"Error al hacer clic en la lupa: {e}")

            except TimeoutException:
                mensaje_span = f"No se encontró el span con número {prefijo}-{numero_doc}"
                raise Exception(f"{mensaje_span}")

            try:

                WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID, "btnAceptarError")))
                raise Exception (f"Se detecto una advertencia, el código para descagar el documento en la compañía es {prefijo}-{numero_doc}.")

            except TimeoutException:
                logging.info("✅ No apareció ninguna alerta.")

            wait.until(EC.visibility_of_element_located((By.ID, "divPanelPDFMaster")))
            time.sleep(3)
            logging.info("📄 Panel PDF de la constancia visible.")

            logging.info("----------------------------")

            boton_guardar = wait.until(EC.element_to_be_clickable((By.ID, "btnDescargarConstanciaM")))

            # Guardar archivos antes del clic
            archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

            driver.execute_script("arguments[0].click();", boton_guardar)
            logging.info("🖱 Clic con JS en el botón de descarga")

            archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=1)
            logging.info(f"✅ Archivo descargado exitosamente")

            if archivo_nuevo:
                ruta_original = archivo_nuevo[0]  # ya es ruta completa
                ruta_final = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
                os.rename(ruta_original, ruta_final)
                logging.info(f"🔄 Constancia renombrado a '{ramo.poliza}.pdf'")
            else:
                raise Exception("No se encontró constancia después de la descarga")

            if tipo_mes == 'MA':

                logging.info("----------------------------")

                driver.switch_to.frame("ifContenedorPDFMaster")#iframe     
                                
                boton_embebido = wait.until(EC.element_to_be_clickable((By.ID, "open-button")))

                # Guardar archivos antes del clic
                archivos_antes_2 = set(os.listdir(ruta_archivos_x_inclu))

                driver.execute_script("arguments[0].click();", boton_embebido)
                logging.info("🖱 Clic con JS en el botón de descarga")

                endoso = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes_2,".pdf",cantidad=1)
                logging.info(f"✅ Archivo descargado exitosamente")

                if archivo_nuevo:
                    ruta_original = endoso[0]  # ya es ruta completa
                    ruta_final = os.path.join(ruta_archivos_x_inclu, f"endoso_{ramo.poliza}.pdf")
                    os.rename(ruta_original, ruta_final)
                    logging.info(f"🔄 Endoso renombrado a 'endoso_{ramo.poliza}.pdf'")
                else:
                    raise Exception("No se descargo el endoso")

                driver.switch_to.default_content()
            
            time.sleep(1)

            if len(list_polizas) == 2:

                ruta_salud = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[0]}.pdf")
                ruta_pension = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[1]}.pdf")
                
                logging.info("----------------------------")

                if os.path.exists(ruta_salud):
                    shutil.copy2(ruta_salud,ruta_pension)
                    logging.info(f"📄 Copia creada para constancia de Pensión")
                else:
                    logging.error(f"❌ No existe el archivo base Constancia Salud")

                logging.info("----------------------------")

                if tipo_mes == 'MA':

                    ruta_endoso_salud = os.path.join(ruta_archivos_x_inclu,f"endoso_{list_polizas[0]}.pdf")
                    ruta_endoso_pension = os.path.join(ruta_archivos_x_inclu,f"endoso_{list_polizas[1]}.pdf")

                    if os.path.exists(ruta_endoso_salud):
                        shutil.copy2(ruta_endoso_salud,ruta_endoso_pension)
                        logging.info(f"📄 Copia creada para el endoso de Pensión")
                    else:
                        logging.error(f"❌ No existe el archivo base Endoso Salud")

                    logging.info("----------------------------")

            #btnPDFCancelarM,btnPDFEnviarM
            btn_cancelar_boton = wait.until(EC.element_to_be_clickable((By.ID, "btnPDFCancelarM")))
            btn_cancelar_boton.click()
            logging.info("✅ Cerrando panel de documentos.")

            time.sleep(2)

            logging.info(f"✅ {palabra_clave} en La Positiva realizada exitosamente")
       
def solicitud_vidaley_MV(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_proceso,actividad,ramo):
 
    ov = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OV']")))
    ov.click()
    logging.info("🖱️ Clic en 'OV'")

    time.sleep(2)

    # Espera a que se abra la nueva ventana (aumenta el tiempo de espera si es necesario)
    wait.until(lambda d: len(d.window_handles) > 1)

    # Cambiar a la nueva ventana
    driver.switch_to.window(driver.window_handles[-1])

    try:
        alert = WebDriverWait(driver,5).until(EC.alert_is_present())
        logging.info(f"⚠️ Alerta presente: {alert.text}")
        alert.accept()
        logging.info("✅ Alerta aceptada")
    except:
        logging.info("✅ No apareció ninguna alerta")
                
    logging.info ("--- Se ingresó a Oficina Virtual La Positiva 🌐---")
    action = ActionChains(driver)

    emision = wait.until(EC.visibility_of_element_located((By.ID, "stUIUserOV3_img")))
    action.move_to_element(emision).perform()
    logging.info ("🖱️ Mouse sobre Emisión")
 
    solicitudes = wait.until(EC.visibility_of_element_located((By.ID, "stUIUserOV5_txt")))
    action.move_to_element(solicitudes).perform()
    logging.info ("🖱️ Mouse sobre Solicitudes")

    formulario = wait.until(EC.element_to_be_clickable((By.ID, "stUIUserOV7_txt")))
    action.move_to_element(formulario).click().perform()
    logging.info ("🖱️ Clic en Formulario Genérico")

    logging.info ("--- Se ingresó a Formulario Genérico de Solicitud de Póliza 🌐---")

    # 4. Cambiar a la nueva pagina si es necesario, Si se abre en la misma pesta a, no hace falta cambiar de handle y Si se abre en otra pesta a, hay que cambiar el contexto:
    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])
        
    time.sleep(2)

    select_element_1 = wait.until(EC.presence_of_element_located((By.NAME, "cboTipo")))

    select_1 = Select(select_element_1)

    if tipo_proceso == 'IN':
        select_1.select_by_value("7")
        logging.info("✅ Se seleccionó la opción 'INCLUSIÓN' ")
    else:
        select_1.select_by_value("2")
        logging.info("✅ Se seleccionó la opción 'RENOVACIÓN' ")

    time.sleep(1)

    select_element_2 = wait.until(EC.presence_of_element_located((By.NAME, "sRamo")))
    select_2 = Select(select_element_2)
    select_2.select_by_value("29")
    logging.info("✅ Se seleccionó la opción 'VIDA LEY' ")

    time.sleep(2)

    boton = wait.until(EC.element_to_be_clickable((By.ID, "button1")))

    driver.execute_script("arguments[0].click();", boton)
    logging.info("✅ Se paso 'VIDA LEY' a la izquierda")

    time.sleep(2)

    iframe = wait.until(EC.presence_of_element_located((By.NAME, "frCliente")))
    driver.switch_to.frame(iframe)

    # Guardar las ventanas actuales
    handle_ventana_principal = driver.current_window_handle

    # Guardar handles antes de abrir la segunda ventana
    handles_antes_ventana2 = set(driver.window_handles)

    buscar_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='dtcliBuscar']//a")))
    driver.execute_script("arguments[0].click();", buscar_link)
    logging.info("🖱️ Clic en la lupa 🔍")

    # --- Paso 5: Manejar el alert (si aparece) ---
    try:
        alert = wait.until(EC.alert_is_present())
        logging.info(f"⚠️ Alerta presente: {alert.text}")
        alert.accept()
        logging.info("✅ Alerta aceptada")
    except:
        logging.info("✅ No apareció ninguna alerta")

    # Esperar y cambiar a la segunda ventana
    wait.until(lambda d: len(d.window_handles) > len(handles_antes_ventana2))
    nueva_ventana2 = (set(driver.window_handles) - handles_antes_ventana2).pop()
    driver.switch_to.window(nueva_ventana2)
    logging.info("--- Ventana 2---------🌐")

    time.sleep(3)

    # 💡 Lógica en segunda ventana
    tipo_documento = wait.until(EC.presence_of_element_located((By.NAME, "lscTipoBusqueda")))
    Select(tipo_documento).select_by_value("2" if len(str(ruc_empresa)) == 11 else "1")
    logging.info("✅ Se seleccionó 'RUC'" if len(str(ruc_empresa)) == 11 else "✅ Se eligió 'DNI'")

    input_element = driver.find_element(By.NAME, "tctTextoBusca")
    input_element.clear()
    input_element.send_keys(ruc_empresa)
    logging.info(f"✅ Se ingresó el RUC: {ruc_empresa}")

    tipo_persona = driver.find_element(By.NAME, "lscTipoPersona_T")
    tipo = Select(tipo_persona)

    prefijo = ruc_empresa[:2]

    if prefijo == "10":
        tipo.select_by_value("1")
        logging.info(f"✅ RUC de tipo Natural")
    else:
        tipo.select_by_value("2")
        logging.info(f"✅ RUC de tipo Jurídico")

    # Guardar handle actual (segunda ventana) ANTES de abrir tercera
    handle_segunda_ventana = driver.current_window_handle
    handles_antes_ventana3 = set(driver.window_handles)

    buscar_btn = driver.find_element(By.XPATH, "//a[contains(@href, 'buscarCliente')]")
    driver.execute_script("arguments[0].click();", buscar_btn)
    logging.info("🖱️ Se hizo clic en 'Buscar' 🔍")

    # Esperar y cambiar a la tercera ventana
    wait.until(lambda d: len(d.window_handles) > len(handles_antes_ventana3))
    handle_ventana3 = (set(driver.window_handles) - handles_antes_ventana3).pop()
    driver.switch_to.window(handle_ventana3)
    logging.info("--- Ventana 3---------🌐")

    try:
        iframe2 = wait.until(EC.presence_of_element_located((By.NAME, "frmClientes")))
        driver.switch_to.frame(iframe2)
        elemento = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, \"selecciona('\")]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento)
        time.sleep(3)
        driver.execute_script("arguments[0].click();", elemento)
        logging.info("🖱️ Clic en el Nombre del Cliente")
    except Exception as e:
        raise Exception(e)

    # --- Esperar que se cierre la ventana 3 y volver a la 2
    time.sleep(3)
    ventanas_actuales = set(driver.window_handles)

    if handle_segunda_ventana in ventanas_actuales:
        driver.switch_to.window(handle_segunda_ventana)
        logging.info("🔙 Volviendo a la Segunda Ventana")
        logging.info("--- Ventana 2 ---------🌐")
    else:
        logging.warning("⚠️ La segunda ventana ya no está disponible")
        if handle_ventana_principal in driver.window_handles:
           driver.switch_to.window(handle_ventana_principal)
           logging.info("✅ Se volvió a la ventana principal")

    aceptar_btn = wait.until(EC.element_to_be_clickable((By.NAME, "clienteAceptar")))
    driver.execute_script("arguments[0].scrollIntoView(true);", aceptar_btn)
    driver.execute_script("arguments[0].click();", aceptar_btn)
    logging.info("🖱️ Clic en botón 'Aceptar'")

    time.sleep(2)

    driver.switch_to.window(handle_ventana_principal)
    logging.info("🔙 Volviendo a la ventana principal")
    logging.info("--- Ventana 1 ---------🌐")

    iframe_constancia = wait.until(EC.presence_of_element_located((By.NAME, "ifrConstancia")))

    driver.switch_to.frame(iframe_constancia)

    email_input = wait.until(EC.presence_of_element_located((By.ID, "tctEmail")))
    email_input.clear()
    email_input.send_keys(ejecutivo_responsable)
    logging.info("📧 Correo del ejecutivo ingresado correctamente")

    input_archivo_xls = wait.until(EC.presence_of_element_located((By.ID, "fichero7")))
    ruta_trama_final_xls1 = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}_97.xls")
    input_archivo_xls.send_keys(ruta_trama_final_xls1)
    logging.info(f"📤 Se subió el archivo {ramo.poliza}_97.xls ")

    time.sleep(2)

    fecha_ini_input = wait.until(EC.presence_of_element_located((By.ID, "txtFechaInicio")))
    fecha_ini_input.clear()
    fecha_ini_input.send_keys(ramo.f_inicio)
    logging.info(f"📅 Fecha Inicio ingresada : {ramo.f_inicio}")

    fecha_fin_input = wait.until(EC.presence_of_element_located((By.ID, "txtFecFin")))
    fecha_fin_input.clear()
    fecha_fin_input.send_keys(ramo.f_fin)
    logging.info(f"📅 Fecha Fin ingresada : {ramo.f_fin}")

    sede_input = wait.until(EC.visibility_of_element_located((By.ID, "txtSede")))

    sede_input.clear()
    sede_input.send_keys(ramo.sede)
    logging.info(f"✅ Sede Ingresada: {ramo.sede}")

    time.sleep(1)

    select_tipo_sede = wait.until(EC.visibility_of_element_located((By.ID, "lscSede")))
    selectTS= Select(select_tipo_sede)
    actividad_upper = actividad.upper()

    if "MINA" in actividad_upper or "MINER" in actividad_upper:
        # Caso MINERA
        selectTS.select_by_value("2")
        logging.info("✅ Se seleccionó la opción 'MINERA'")
    else:
        # Caso NO MINERA
        selectTS.select_by_value("1")
        logging.info("✅ Se seleccionó la opción 'NO MINERA'")

    actividad_input = wait.until(EC.visibility_of_element_located((By.ID, "txtActividad")))
    actividad_input.clear()
    actividad_input.send_keys(actividad)
    logging.info(f"✅ Actividad registrada: {actividad}")

    # Salir del iframe y volver al contenido principal
    driver.switch_to.default_content()

    poliza_input = wait.until(EC.visibility_of_element_located((By.ID, "tctPoliza")))

    btn_registrar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#dtRegistrar a")))
    driver.execute_script("arguments[0].scrollIntoView(true);", btn_registrar)

    poliza_input.clear()
    poliza_input.send_keys(ramo.poliza)
    logging.info(f"✅ Número de póliza ingresado: {ramo.poliza}")

    # Scrollea hasta el botón por si está fuera de vista
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    time.sleep(2)

    forma_pago = wait.until(EC.presence_of_element_located((By.NAME, "lscPago")))
    form_pag = Select(forma_pago)
    form_pag.select_by_value("1")

    logging.info("✅ Forma de pago seleccionada")

    time.sleep(2)

    input_cert = wait.until(EC.presence_of_element_located((By.NAME, "tcnNumCert")))

    input_cert.clear()
    input_cert.send_keys("1")

    logging.info("✅ N° Certificados : 1 ")

    observacion = f"{palabra_clave} Vida Ley \n Vigencia: {ramo.f_inicio} al {ramo.f_fin} \n  N° de Póliza: {ramo.poliza} \n Modalidad: Mes Vencido"

    textarea = wait.until(EC.presence_of_element_located((By.ID, "textarea1")))
    textarea.clear()
    textarea.send_keys(observacion)

    logging.info(f"✅ Observación ingresada correctamente: {observacion}")

    time.sleep(2)

    input_archivo = wait.until(EC.presence_of_element_located((By.ID, "fichero4")))
    ruta_trama_final = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")
    input_archivo.send_keys(ruta_trama_final)
    logging.info(f"📤 Se subió el archivo: {ramo.poliza}.xlsx")

    # Guardar archivos antes del clic
    archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

    btn_registrar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#dtRegistrar a")))
    btn_registrar.click()
    logging.info("🖱️ Clic en 'Enviar' realizado correctamente.")

    # Primero verificar si aparece un alert
    try:
        WebDriverWait(driver,5).until(EC.alert_is_present())
        alert1 = driver.switch_to.alert
        logging.info(f"⚠️ Alerta : ¿{alert1.text}?")
        alert1.accept()
        logging.info("✅ Alerta aceptada")
    except TimeoutException:
        logging.info("✅ No apareció la primera alerta, revisamos si salió el iframe...")

        # Si no hay alert → entonces revisar el iframe
        try:
            iframe = wait.until(EC.presence_of_element_located((By.ID, "modalIframe")))
            driver.switch_to.frame(iframe)
            logging.info("⚠️ Se detectó el iframe de similitud de trámites.")

            driver.switch_to.default_content()

            try:

                tomar_capturar(driver,ruta_archivos_x_inclu,f"Similitud")

                cerrar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'cerrarVentanaModal')]")))

                cerrar_btn.click()
                logging.info("🖱️ Clic en botón 'Cerrar' dentro del iframe")

            except Exception:

                driver.execute_script("cerrarVentanaModal();")
                logging.info("🖱️ Clic en 'Cerrar' ejecutado con JS")

            # 👇 Aquí validas si aparece la alerta después del iframe
            try:
                WebDriverWait(driver,5).until(EC.alert_is_present())
                alert0 = driver.switch_to.alert
                logging.info(f"⚠️ Alerta : ¿{alert0.text}?")
                alert0.accept()
                logging.info("✅ Alerta aceptada después del iframe")
            except TimeoutException:
                logging.info("❌ No apareció alerta después del iframe")
            except NoAlertPresentException:
                logging.info("⚠️ Selenium detectó alerta, pero desapareció antes de leerla")
            finally:
                # Siempre regresar al DOM principal
                driver.switch_to.default_content()

        except TimeoutException:
            logging.info("✅ No apareció el iframe de similitud, se continúa con el flujo.")

    except NoAlertPresentException:
        logging.info("⚠️ Selenium detectó alerta, pero desapareció antes de leerla")
    
    time.sleep(10)
    archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=1)

    try:

        if archivo_nuevo:
            logging.info(f"✅ Archivo descargado exitosamente")
            ruta_original = archivo_nuevo[0]
            ruta_final = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
            os.rename(ruta_original, ruta_final)
            logging.info(f"🔄 Constancia renombrado a '{ramo.poliza}.pdf'")
        else:
            raise Exception("No se descargo ningun archivo")

    except Exception as e:
        raise Exception(f"{e}")

    try:

        wait.until(EC.alert_is_present())
        alert = driver.switch_to.alert
        mensaje = alert.text.strip()

        logging.info(f"⚠️ Alerta : {mensaje}")

        # ---- Evaluación del mensaje ----
        if mensaje.startswith("Corregir errores encontrados en la Trama."):
            alert.accept()
            logging.info("✅ Alerta aceptada")

            #--- capturar el mensaje del pdf que se descargo para enviarlo por error
            ruta_pdf_errores = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
            error_pdf = leer_pdf(ruta_pdf_errores)
            raise Exception(f"{error_pdf}")

        elif mensaje.startswith("Se ha registrado la solicitud correctamente."):
            alert.accept()
            logging.info("✅ Alerta aceptada")

            try:
                wait.until(EC.alert_is_present())
                alert_impresion = driver.switch_to.alert
                logging.info(f"⚠️ Alerta : ¿{alert_impresion.text}?")
                alert_impresion.dismiss()
                logging.info("✅ Alerta Cancelada")

            except TimeoutException:
                logging.info("❌ No apareció ninguna alerta")
            finally:
                tomar_capturar(driver,ruta_archivos_x_inclu,f"Tramite")

        else:
            logging.warning("⚠️ Alerta no contemplada, se acepta por defecto")
            alert.accept()

    except TimeoutException:
        logging.info("❌ No apareció ninguna alerta")


    # ruta_cons = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
    # tramite = obtener_tramite_pdf(ruta_cons)

    # if tramite is False:

    #     # Enviar documento a ejecutivo para anular
    #     enviar_constancia_anular(ejecutivo_responsable,ramo,ruta_cons)

    #     raise Exception("Se genero la constancia sin valores de la Trama , enviando correo para anular")
    # elif tramite is None:
    #     raise Exception("No se encontró número de trámite de la constancia")
    # else:
    #     logging.info(f"✅ Número de trámite encontrado: {tramite}")

    logging.info(f"✅ Constancia obtenida para la {palabra_clave} con numero de póliza '{ramo.poliza}'")

def solicitud_vidaley_MA(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_proceso,ramo):

    vidaley = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Vida Ley']")))
    vidaley.click()
    logging.info("🖱️ Clic en 'Vida Ley'")

    wait.until(lambda d: len(d.window_handles) > 1)

    driver.switch_to.window(driver.window_handles[-1])
    logging.info("🔙  Redireccionando en la nueva ventana")

    time.sleep(5)

    icon = wait.until( EC.element_to_be_clickable((By.CLASS_NAME, "icon-menuside-right-arrow")))

    actions = ActionChains(driver)
    actions.move_to_element(icon).click().perform()
    logging.info("🖱️ Clic en el Menú despegable de la izquierda")

    time.sleep(3)

    try:
        transacciones_span = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Transacciones']")))
        transacciones_span.click()
        logging.info("🖱️ Clic en Transacciones")
    except Exception as e:
        raise e

    time.sleep(3)

    try:
        input_element = wait.until(EC.element_to_be_clickable((By.ID, "b12-b1-Input_PolicesNumber")))
        input_element.clear()
        input_element.send_keys(ramo.poliza)
        logging.info(f"✅ Numero de Póliza ingresado: {ramo.poliza}")
    except Exception as e:
        raise e

    time.sleep(3)

    # input_ruc = wait.until(EC.element_to_be_clickable((By.ID, "b12-b1-Input_RUCContratante2")))
    # input_ruc.clear()
    # input_ruc.send_keys(ruc_empresa)
    # logging.info(f"✅ Numero de RUC ingresado: {ruc_empresa}")

    # time.sleep(3)

    # fecha_dt = datetime.strptime(ramo.f_inicio, "%d/%m/%Y")
    # primer_dia_mes = fecha_dt.replace(day=1)

    # fechai_vigencia = wait.until(EC.element_to_be_clickable((By.ID, "b12-b1-Input_RegistrationDateStart")))
    # driver.execute_script("""
    #     arguments[0].removeAttribute('readonly');
    #     arguments[0].value = arguments[1];
    #     arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
    #     arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    # """, fechai_vigencia, primer_dia_mes)

    # # Esperar a que el sistema autocomplete el segundo campo
    # time.sleep(2)

    # input_fin = driver.find_element(By.ID, "b12-b1-Input_RegistrationDateEnd")
    # fecha_fin_actual = input_fin.get_attribute("value")
    # logging.info(f"📅 Fecha de Inicio de la vigencia: {primer_dia_mes}")
    # logging.info(f"📅 Fecha Fin autocompletada: {fecha_fin_actual}")

    # driver.execute_script("document.body.click();")

    boton_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Buscar']]")))
    boton_buscar.click()
    logging.info(f"🖱️ Clic en 'Buscar'")

    try:
        tit = WebDriverWait(driver,8).until(EC.visibility_of_element_located((By.ID, "b12-b11-TextTitlevalue"))).text
        con = WebDriverWait(driver,8).until(EC.visibility_of_element_located((By.ID, "b12-b11-TextContentValue"))).text
        raise Exception(f"{tit} {con}")
    except TimeoutException:
        pass

    wait.until(EC.presence_of_element_located((By.ID, "b12-Widget_TransactionRecordList")))
    logging.info(f"⌛ Esperando la tabla con resultados")

    try:
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID, "b12-limpiargrilla")))
        logging.error(f"❌ Se detectó un modal que no carga los resultados.")
        raise Exception("No carga la tabla")
    except TimeoutException:
        logging.info("✅ No se detectó modal")

    try:
        fila = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#b12-Widget_TransactionRecordList tbody tr")))
    except TimeoutException:
        raise Exception(f"Poliza {ramo.poliza} no figura con Modalidad Mes Adelantado")

    checkbox = fila.find_element(By.CSS_SELECTOR, "input[type='checkbox']")

    actions = ActionChains(driver)
    actions.move_to_element(checkbox).click().perform()
    logging.info(f"✅ Se seleccionó la póliza: {ramo.poliza}")

    time.sleep(3)

    if tipo_proceso == 'IN':
        boton_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Inclusión']]")))
    else:
        boton_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Renovación']]")))

    boton_proceso.click()
    logging.info(f"🖱️ Clic en {palabra_clave} ")

    time.sleep(3)

    try:

        titulo = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID, "b12-b2-b8-TextTitlevalue"))).text
        contenido = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID, "b12-b2-b8-TextContentValue"))).text
        raise Exception(f"{titulo} {contenido}")

    except TimeoutException:
        logging.info("✅ No se detectó modal de aviso")

    logging.info(f"--- Se ingresó a la ventana de {palabra_clave} 🌐----")

    ruta_archivo = os.path.join(ruta_archivos_x_inclu,f"{ramo.poliza}.xlsx")

    if tipo_proceso == 'IN':

        fecha_vigencia_solicitud = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-EffectiveDateTyping")))
        driver.execute_script("""
            arguments[0].removeAttribute('readonly');
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, fecha_vigencia_solicitud,ramo.f_inicio)
        logging.info(f"📅 Fecha de Inicio de la vigencia para la {palabra_clave}: {ramo.f_inicio}")

    time.sleep(2)

    if not os.path.exists(ruta_archivo):
        raise Exception (f"Archivo {ramo.poliza}.xlsx no encontrado")

    if tipo_proceso == 'IN':
        input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='b13-b1-b11-b2-b1-DropArea']//input[@type='file']")))
        input_file.send_keys(ruta_archivo)
    else:
        input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='b13-b1-b1-b2-b1-DropArea']//input[@type='file']")))
        input_file.send_keys(ruta_archivo)

    logging.info(f"⌛ Trama {ramo.poliza}.xlsx subido para validar")

    time.sleep(5)

    try:
        msj = 'La planilla no está en un formato válido de excel'
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,f"//span[contains(text(), '{msj}')]")))
        raise Exception(msj)
    except TimeoutException:
        logging.info("✅ No se detectó error de formato en la planilla")

    boton_validar = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-btnValidate")))
    boton_validar.click()
    logging.info(f"🖱️ Clic en Validar")

    try:
        boton_aceptar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Aceptar']")))
        boton_aceptar.click()
        logging.info("✅ Clic en el botón Aceptar.")
    except TimeoutException:
        logging.info("✅ No se detectó advertencia de inicio de vigencia")

    try:
        msj = 'Lo sentimos'
        WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{msj}')]")))
        raise Exception(msj)
    except TimeoutException:
        #logging.info("✅ No se detectó errores en la Trama. Se continúa con el flujo normalmente.")
        pass
        
    try:
        wait.until(
            EC.any_of(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[contains(text(),'La planilla superó exitosamente')]")
                ),
                EC.presence_of_element_located(
                    (By.XPATH, "//span[contains(text(),'Encontramos algunos errores en la planilla')]")
                )
            )
        )

        # Verificamos cuál mensaje apareció
        if driver.find_elements(By.XPATH, "//span[contains(text(),'Encontramos algunos errores en la planilla')]"):
            raise Exception("Encontramos algunos errores en la planilla. Descarga y corrige las observaciones")

        logging.info("✅ Planilla validada exitosamente. Continuando con el flujo.")

    except TimeoutException:
        raise Exception("No se pudo determinar el resultado de la validación de la planilla.")

    # try:
    #     wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'La planilla superó exitosamente')]")))
    #     logging.info("✅ Planilla validada exitosamente. Continuando con el flujo.")
    # except TimeoutException:       
    #     raise Exception("La validación de la planilla falló o no se completó")
            
    #Nuevo Correo del cliente
    input_correo = wait.until(EC.element_to_be_clickable((By.ID, "b13-Input_ContractingEmail")))

    # Hace scroll hasta el botón
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'})", input_correo)

    input_correo.clear()
    input_correo.send_keys(ejecutivo_responsable)
    logging.info("✅ Editando correo del ejecutivo")

    # Esperar que el botón esté presente
    boton_calcular = wait.until(EC.element_to_be_clickable((By.ID, "b13-CalculatePremium")))
    driver.execute_script("arguments[0].scrollIntoView(true);", boton_calcular)
    boton_calcular.click()
    logging.info("🖱️ Se hizo clic en el botón 'Calcular'")

    if tipo_proceso == 'IN':   
        finalizar_btn = wait.until(EC.element_to_be_clickable((By.ID, "b13-InclusionValidate")))
    else:
        finalizar_btn = wait.until(EC.element_to_be_clickable((By.ID, "b13-RenewalRequest2")))

    finalizar_btn.click()
    logging.info(f"🖱️ Clic en 'Finalizar'")

    time.sleep(5)

    # link_descargar = wait.until(
    #     EC.element_to_be_clickable(
    #         (By.XPATH, "//a[.//span[normalize-space()='Descargar documentos']]")
    #     )
    # )
    # link_descargar.click()

    # span = wait.until(EC.presence_of_element_located((By.ID, "b13-b48-b1-TextContentValue")))
    # texto = span.text.strip()
    # logging.info(f"📄 Texto obtenido del popup: {texto}")

    # 1. Esperar que exista en el DOM
    wait.until(EC.presence_of_element_located((By.XPATH,"//span[normalize-space()='Descargar documentos']")))

    # 2. Esperar que sea visible
    wait.until(EC.visibility_of_element_located((By.XPATH,"//span[normalize-space()='Descargar documentos']")))

    # 3. Esperar que sea clickeable y clicar
    boton_descargar = wait.until(EC.element_to_be_clickable((By.XPATH,"//span[normalize-space()='Descargar documentos']/ancestor::a")))
    #driver.save_screenshot(os.path.join(ruta_archivos_x_inclu,f"solicitud_{get_timestamp()}.png"))
    tomar_capturar(driver,ruta_archivos_x_inclu,f"Solicitud")

    archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

    driver.execute_script("arguments[0].click();", boton_descargar)
    logging.info("🖱 Clic con JS en el botón de descarga")

    archivos_nuevos = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=2)
    logging.info("✅ Documentos descargados correctamente")

    if archivos_nuevos:
        # Ordenarlos por fecha de creación (más seguro)
        archivos_nuevos = sorted(
            archivos_nuevos,
            key=os.path.getctime
        )

        nombre_1 = os.path.join(
            ruta_archivos_x_inclu,
            f"{ramo.poliza}.pdf"
        )

        nombre_2 = os.path.join(
            ruta_archivos_x_inclu,
            f"endoso_{ramo.poliza}.pdf"
        )

        os.rename(archivos_nuevos[0], nombre_1)
        logging.info(f"✅ Constancia '{ramo.poliza}.pdf' renombrado correctamente")
        os.rename(archivos_nuevos[1], nombre_2)
        logging.info(f"✅ Endoso 'endoso_{ramo.poliza}.pdf' renombrado correctamente")
    else:
        logging.error("❌ No se detectaron los 2 archivos")

    # if descargar_documento(driver,boton_descargar,ramo.poliza,impresion=False,pestaña=False):
    #     logging.info(f"✅ Constancia {ramo.poliza}.pdf obtenida")
    # else:
    #     raise Exception ("No se logró descargar la Constancia")

    # time.sleep(5)
    # endoso_poliza = f"endoso_{ramo.poliza}"

    # try:

    #     logging.info("⌛ Esperando la segunda ventana descarga de Linux Debian")

    #     if not esperar_ventana("Save File"):
    #         raise Exception("No apareció la ventana de descarga")

    #     subprocess.run(["xdotool", "search", "--name", "Save File", "windowactivate", "windowfocus"])
    #     logging.info("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
    #     subprocess.run(["xdotool", "type","--delay", "100", endoso_poliza])
    #     logging.info("📄 Se cambio el nombre del documento")
    #     time.sleep(2)
    #     subprocess.run(["xdotool", "key", "Return"])
    #     logging.info("🖱️ Se presionó 'Enter' para confirmar la descarga")
    #     time.sleep(2)

    # except Exception as e:
    #     raise Exception(f"Error al descargar el endoso, Motivo -> {e}")

    # logging.info(f"✅ Endoso de {palabra_clave} obtenida, Documento : {endoso_poliza}.pdf")

    boton_aceptar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Aceptar')]")))
    boton_aceptar.click()
    logging.info("🖱️ Clic en 'Aceptar' final")

def login_la_positiva(driver,wait,list_polizas,ba_codigo,bab_codigo,tipo_mes,ruta_archivos_x_inclu,
                                   ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_proceso,actividad,ramo):
  
    global ventana_menu_positiva

    tipoError = ""
    detalleError = ""


    if ba_codigo == '3' and bab_codigo == '4':
        conVL,proVL,tipErVL,detErVL = solicitud_vidaley_x_tipo_Mes(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ejecutivo_responsable,
                                                                  palabra_clave,tipo_proceso,actividad,ramo,tipo_mes,tipoError,
                                                                  detalleError)

        return conVL,proVL,tipErVL,detErVL

    try:

        logging.info("----------------------------")
        driver.get('https://web.lapositiva.com.pe/sso_login_ui/')
        logging.info("⌛ Cargando la Web de Positiva")

        for intento in range(2):

            logging.info(f"🔄 Intento de login número {intento + 1}...")
            time.sleep(3)

            try:               

                user_field = wait.until(EC.presence_of_element_located((By.ID, "b5-Input_User")))
                user_field.clear()

                mover_y_hacer_click_simple(driver, user_field)
                time.sleep(random.uniform(0.97, 0.99))

                escribir_lento(user_field, ramo.usuario, min_delay=0.97, max_delay=0.99)
                logging.info("⌨️ Digitando el Username")

                time.sleep(1 + random.random() * 1.5)

                password_field = wait.until(EC.presence_of_element_located((By.ID, "b5-Input_PassWord")))
                password_field.clear()

                mover_y_hacer_click_simple(driver, password_field)
                time.sleep(random.uniform(0.97, 0.99))

                escribir_lento(password_field, ramo.clave, min_delay=0.97, max_delay=0.99)
                logging.info("⌨️ Digitando el Password")

                time.sleep(1 + random.random() * 1.5)

                login_button = wait.until(EC.element_to_be_clickable((By.ID, "b5-btnAction")))
                mover_y_hacer_click_simple(driver, login_button)
                logging.info("🖱️ Clic en Iniciar Sesión")

                time.sleep(3)

                try:

                    popup_text = WebDriverWait(driver,7).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Usuario o contraseña incorrectos')]")))

                    if popup_text:

                        logging.info("❌ Usuario o contraseña incorrectos")
 
                        if intento == 1:
                            raise Exception("Usuario o contraseña incorrectos")

                        aceptar_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//span[text()='Aceptar']]")))
                        aceptar_btn.click()
                        logging.info("🖱️ Clic en Aceptar")
                        time.sleep(2)
                        driver.refresh()
                        continue
                except TimeoutException:
                    logging.info("✅ Login exitoso")

                autogestion = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class,'menu-item')]//span[normalize-space()='Autogestión']/parent::div")))
                driver.execute_script("arguments[0].click();", autogestion)
                logging.info("🖱️ Clic en Autogestión")

                ventana_menu_positiva = driver.current_window_handle

                break

            except Exception as e:
                logging.info(f"❌ Error durante el intento {intento+1}: {e}")
                driver.refresh()

    except Exception as e:

        logging.error(f"❌ Error inesperado entrando a la url: {e}")
        return False,False,"Login Fallido", str(e)

    if bab_codigo in ['1', '2', '3']:

        try:
            solicitud_sctr(driver,wait,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,ba_codigo,ramo)
            return True, True if tipo_mes == 'MA' else False, tipoError, detalleError
        except Exception as e:
            resultado, asunto = validar_pagina(driver)
            tomar_capturar(driver, ruta_archivos_x_inclu, f"ERROR_SCTR_{tipo_mes}")
            detalle = f"{asunto} - {e}" if not resultado else str(e)
            logging.error(f"❌ Error en La Positiva (SCTR) - {tipo_mes}: {detalle}")
            return False, False, f"LAPO-SCTR-{tipo_mes}", detalle
        finally:
            driver.close()
            logging.info("✅ Cerrando la Pestaña SED Positiva-SCTR")
            driver.switch_to.window(driver.window_handles[0])
            logging.info("🔙 Retornando al menú principal tras cerrar SED")

    else: 
        conVL,proVL,tipErVL,detErVL = solicitud_vidaley_x_tipo_Mes(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ejecutivo_responsable,
                                                                  palabra_clave,tipo_proceso,actividad,ramo,tipo_mes,tipoError,detalleError)
        return conVL,proVL,tipErVL,detErVL

def solicitud_vidaley_x_tipo_Mes(driver, wait, ruta_archivos_x_inclu, ruc_empresa,ejecutivo_responsable, palabra_clave,
                                tipo_proceso,actividad, ramo, tipo_mes, tipoError, detalleError):

    # 🔹 Mapear función según tipo_mes
    funciones = {
        "MV": lambda: solicitud_vidaley_MV(
            driver, wait, ruta_archivos_x_inclu, ruc_empresa,
            ejecutivo_responsable, palabra_clave,
            tipo_proceso, actividad, ramo
        ),
        "MA": lambda: solicitud_vidaley_MA(
            driver, wait, ruta_archivos_x_inclu, ruc_empresa,
            ejecutivo_responsable, palabra_clave,
            tipo_proceso, ramo
        )
    }

    funcion = funciones.get(tipo_mes)

    if not funcion:
        return False, False, f"LAPO-VIDALEY-{tipo_mes}", "Tipo de mes inválido"

    try:
        funcion()

        # 🔹 Solo cambia este flag según tipo
        flag_extra = True if tipo_mes == "MA" else False

        return True, flag_extra, tipoError, detalleError

    except Exception as e:
        logging.error(f"❌ Error en La Positiva (Vida Ley) - {tipo_mes}: {e}")

        resultado, asunto = validar_pagina(driver)
        tomar_capturar(driver,ruta_archivos_x_inclu,f"ERROR_VIDALEY_{tipo_mes}")
        detalle = f"{asunto} - {e}" if not resultado else str(e)
        return False, False, f"LAPO-VIDALEY-{tipo_mes}", detalle

    finally:
        driver.close()
        logging.info("✅ Cerrando la pestaña Vida Ley")
        driver.switch_to.window(driver.window_handles[0])
        logging.info("🔙 Retornando al menú principal")