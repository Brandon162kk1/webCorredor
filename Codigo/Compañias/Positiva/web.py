#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,NoAlertPresentException,StaleElementReferenceException,ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from Tiempo.fechas_horas import get_fecha_hoy
from Compañias.Positiva.metodos import mover_y_hacer_click_simple,escribir_lento,validar_pagina,validardeuda,leer_pdf
from LinuxDebian.Ventana.ventana import esperar_archivos_nuevos
from Chrome.google import tomar_capturar
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

    error_locator = (By.XPATH, "//h3[contains(text(),'Lo sentimos, ha ocurrido un error inesperado')]")
    poliza_locator = (By.ID, "ContentPlaceHolder1_txtNroPoliza")

    resultado = wait.until(
        EC.any_of(
            EC.presence_of_element_located(error_locator),
            EC.presence_of_element_located(poliza_locator)
        )
    )

    if resultado.get_attribute("id") == "ContentPlaceHolder1_txtNroPoliza":
        poliza_input = resultado
        poliza_input.clear()
        poliza_input.send_keys(ramo.poliza)
        logging.info(f"✅ Se ingresó la póliza '{ramo.poliza}'")

    else:
        logging.error("❌ Error en la web detectado")
        raise Exception("Lo sentimos, ha ocurrido un error inesperado")

    # try:
    #     WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, "//h3[contains(text(),'Lo sentimos, ha ocurrido un error inesperado')]")))
    #     raise Exception("Lo sentimos, ha ocurrido un error inesperado")
    # except TimeoutException:
    #     pass

    # poliza_input = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_txtNroPoliza")))
    # poliza_input.clear()
    # poliza_input.send_keys(ramo.poliza)
    # logging.info(f"✅ Se Ingresó la póliza '{ramo.poliza}' en el campo")

    buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnBuscar")))
    buscar_btn.click()
    logging.info("🖱️ Clic en Buscar")

    try:
        wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado")))
        logging.info("⌛ Esperando la tabla con los campos...")
    except Exception as e:
        raise Exception(f"No se encontró la póliza {ramo.poliza} en la compañía")

    try:

        # Obtener la fecha de vigencia (formato: dd/mm/yyyy)
        fecha_emision_element = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado_lblFechaEmision_0")))
        fecha_emision_str = fecha_emision_element.text.strip()
        fecha_emision = datetime.strptime(fecha_emision_str, "%d/%m/%Y")
        # Calcular el primer día del mes siguiente
        if fecha_emision.month == 12:
            primer_dia_mes_siguiente = datetime(fecha_emision.year + 1, 1, 1)
        else:
            primer_dia_mes_siguiente = datetime(fecha_emision.year, fecha_emision.month + 1, 1)

        primer_dia_mes_siguiente_str = primer_dia_mes_siguiente.strftime("%d/%m/%Y")

        # Obtener la fecha de vigencia (formato: dd/mm/yyyy)
        fecha_vigencia_element = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado_lblFechaVigencia_0")))
        fecha_vigencia_str = fecha_vigencia_element.text.strip() 
        # Convertir la fecha a objeto datetime
        fecha_vigencia_dmy = datetime.strptime(fecha_vigencia_str, "%d/%m/%Y")
        fecha_vigencia_str_for = fecha_vigencia_dmy.strftime("%d/%m/%Y")

        logging.info(f"📅 Fecha Fin de la vigencia en la póliza : {fecha_vigencia_str}")

        if tipo_proceso == 'IN':

            # # Definir rango permitido solo para inclusiones
            # rango_inicio = fecha_vigencia_dmy.date() - timedelta(days=2)  # 2 días antes
            # rango_fin = fecha_vigencia_dmy.date() + timedelta(days=2)     # 2 días después
            logging.info(f"📅 Rango permitido para la {palabra_clave}: {primer_dia_mes_siguiente_str} a {fecha_vigencia_str_for}")

    except Exception as e:
        raise Exception(f"Error al parsear la fecha de vigencia,Motivo - {e}")
    
    if tipo_proceso == 'IN' and ramo.f_fin != fecha_vigencia_str: 
        raise Exception(f"Las fechas Fin de la vigencia en la Póliza {ramo.poliza} no coinciden")
    # elif tipo_proceso == 'RE' and ramo.f_fin == fecha_vigencia_str:
    #     raise Exception(f"La Póliza {ramo.poliza} ya fue renovada con el periodo: {ramo.f_inicio} al {fecha_vigencia_str}")
    else:

        # radio_seleccion = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_gvResultado_rbSeleccion_0")))
        # radio_seleccion.click()
        # logging.info("🖱️ Clic en el 'Radio Button' correspondiente")

        # loader = (By.ID, "ID_MODAL_PROCESS")
        # error_modal = (By.ID, "divAlertaErrorGeneral")
        # btn_incluir_locator = (By.ID, "ContentPlaceHolder1_btnIncluir") if tipo_proceso == 'IN' else (By.ID, "ContentPlaceHolder1_btnRenovar")

        # resultado = wait.until(
        #     EC.any_of(
        #         EC.invisibility_of_element_located(loader),
        #         EC.visibility_of_element_located(error_modal),
        #         EC.element_to_be_clickable(btn_incluir_locator)
        #     )
        # )

        # # Caso error
        # if driver.find_elements(*error_modal):
        #     logging.erro("Mucho demora 1")
        #     mensaje_error = driver.find_element(By.XPATH, "//div[@id='divAlertaErrorGeneral']//p/span[2]").text.strip()
        #     raise Exception(mensaje_error)
        
        # wait.until(EC.invisibility_of_element_located((By.ID, "ID_MODAL_PROCESS")))
        # logging.info("⌛ Esperando que cargue...")

        # try:
        #     modal_error_general = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "divAlertaErrorGeneral")))
        #     mensaje_error = modal_error_general.find_element(By.XPATH, ".//p/span[2]").text.strip()
        #     raise Exception(mensaje_error)
        # except TimeoutException:
        #     pass

        #------ ACA me quede -------

        # if tipo_proceso == 'IN':  
            
        #     btn_incluir = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnIncluir")))
        #     btn_incluir.click()
        #     logging.info("🖱️ Clic en Incluir")

        #     modal_incluir = (By.ID, "divTipoIncluir")
        #     siguiente_paso = (By.ID, "fuPlanillaAjax")  # o algo que aparezca si NO hay modal

        #     resultado = wait.until(
        #         EC.any_of(
        #             EC.visibility_of_element_located(modal_incluir),
        #             EC.presence_of_element_located(siguiente_paso)
        #         )
        #     )

        #     if resultado.get_attribute("id") == "divTipoIncluir":
        #         logging.info("⚠️ Apareció el modal con advertencia")

        #         btn_aceptar = wait.until(
        #             EC.element_to_be_clickable(
        #                 (By.ID, "ContentPlaceHolder1_btnIncluirProformaSi" if tipo_mes == 'MA'
        #                  else "ContentPlaceHolder1_btnIncluirProformaNo")
        #             )
        #         )
        #         btn_aceptar.click()
        #         nom_btn = 'Si' if tipo_mes == 'MA' else 'No'
        #         logging.info(f"🖱️ Clic en {nom_btn}")

        #     # try:
        #     #     WebDriverWait(driver,20).until(EC.visibility_of_element_located((By.ID, "divTipoIncluir")))
        #     #     logging.info(f"⚠️ Apareció el modal con advertencia")
                
        #     #     if tipo_mes == 'MA':
        #     #         btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnIncluirProformaSi")))
        #     #     else:
        #     #         btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnIncluirProformaNo")))
                    
        #     #     nom_btn = 'Si' if tipo_mes == 'MA' else 'No'
        #     #     btn_aceptar.click()
        #     #     logging.info(f"🖱️ Clic en {nom_btn}")

        #     # except TimeoutException:
        #     #     pass

        #     #-------------------------------
        #     # mensaje_deuda = validardeuda(driver,wait)

        #     # if mensaje_deuda:
        #     #     raise Exception(f"{mensaje_deuda}")

        #     #-------------------------------

        #     deuda_modal = (By.ID, "divAlertas")
        #     file_input = (By.ID, "fuPlanillaAjax")

        #     resultado2 = wait.until(
        #         EC.any_of(
        #             EC.visibility_of_element_located(deuda_modal),
        #             EC.presence_of_element_located(file_input)
        #         )
        #     )

        #     if resultado2.get_attribute("id") == "divAlertas":
        #         logging.warning("⚠️ Apareció el modal de advertencia")
        #         mensaje = driver.find_element(By.ID, "spMensaje").text.strip()
        #         raise Exception(mensaje)

        #     #input_file = wait.until(EC.presence_of_element_located((By.ID, "fuPlanillaAjax")))
        #     logging.info("⌛ Esperando que cargue la nueva página donde se adjunta la Trama...")
      
        #     input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniFecVigInc")))
        #     input_fecha.clear()

        #     driver.execute_script("""
        #         arguments[0].value = arguments[1]; 
        #         arguments[0].dispatchEvent(new Event('change'));
        #         arguments[0].dispatchEvent(new Event('blur'));
        #     """, input_fecha, ramo.f_inicio)

        #     alerta_fecha = (By.ID, "divAlertas")
        #     body_clickable = (By.TAG_NAME, "body")

        #     resultado3 = wait.until(
        #         EC.any_of(
        #             EC.visibility_of_element_located(alerta_fecha),
        #             EC.element_to_be_clickable(body_clickable)
        #         )
        #     )

        #     if resultado3.get_attribute("id") == "divAlertas":
        #         logging.info(f"⚠️ Apareció el modal con advertencia")
        #         mensaje = driver.find_element(By.ID, "spMensaje").text
        #         raise Exception(mensaje)

        #     # try:
        #     #     WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "divAlertas")))
        #     #     logging.info(f"⚠️ Apareció el modal con advertencia")
        #     #     mensaje_fecha = wait.until(EC.visibility_of_element_located((By.ID, "spMensaje"))).text
        #     #     raise Exception(mensaje_fecha)
        #     # except TimeoutException:
        #     #     pass

        #     driver.find_element(By.TAG_NAME, "body").click()

        #     fecha_date = datetime.strptime(ramo.f_inicio,"%d/%m/%Y").date()

        #     if fecha_date < get_fecha_hoy().date():
        #         diferencia = (get_fecha_hoy().date() - fecha_date).days
        #         cant_dias = abs(diferencia)
        #         text_dia = "día" if cant_dias == 1 else "días"
        #         logging.info(f"📅 Vigencia de Inicio retroactiva ingresada: {ramo.f_inicio} con {cant_dias} {text_dia} de diferencia")
        #     else:
        #         logging.info(f"📅 Vigencia de Inicio ingresada: {ramo.f_inicio}")

        #     if ramo.poliza == '9231375':

        #         select_tipo_calculo = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_cboTipoCalculoEspecial")))
        #         select_tp = Select(select_tipo_calculo)
        #         select_tp.select_by_value("21")
        #         logging.info("✅ Se seleccionó la opción 'Por días' ")

        # else:

        #     btn_incluir = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnRenovar")))
        #     btn_incluir.click()
        #     logging.info("🖱️ Clic en Renovar")

        #     mensaje_deuda = validardeuda(driver,wait)

        #     if mensaje_deuda:
        #         raise Exception(f"{mensaje_deuda}")

        #     wait.until(EC.visibility_of_element_located((By.ID, "divTipoRenovacion")))

        #     btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnRenovacionSi")))

        #     btn_aceptar.click()
        #     logging.info(f"🖱️ Clic en aceptar la Renovación")

        #     input_file = wait.until(EC.presence_of_element_located((By.ID, "fuPlanillaAjax")))
        #     logging.info("⌛ Esperando que cargue la nueva página donde se adjunta la Trama...")

        #     # Solo para este cliente: P & Q TELECOM EIRL en particular Salud: 9231375 Pensión: 30358377 
        #     # pq las otras polizas se ponen por defecto 1 mes
        #     if ramo.poliza == '9231375':

        #         fecha_str = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniVigPoliza"))).get_attribute("value")

        #         fecha = datetime.strptime(fecha_str, "%d/%m/%Y")

        #         fecha_siguiente_mes = fecha + relativedelta(months=1)

        #         resultado = fecha_siguiente_mes - timedelta(days=1)

        #         fecha_final = resultado.strftime("%d/%m/%Y")

        #         input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtFinVigPoliza")))

        #         input_fecha.clear()
        #         driver.execute_script("""
        #             arguments[0].value = arguments[1]; 
        #             arguments[0].dispatchEvent(new Event('change'));
        #             arguments[0].dispatchEvent(new Event('blur'));
        #         """, input_fecha, fecha_final)

        #         # Clic fuera del input (por ejemplo, el body) para cerrar el calendario
        #         driver.find_element(By.TAG_NAME, "body").click()
        #         logging.info(f"📅 Vigencia de Póliza ingresada con JS Hasta : {fecha_final}")
        #--------------

        # =========================
        # RADIO BUTTON
        # =========================
        radio_seleccion = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_gvResultado_rbSeleccion_0")))
        radio_seleccion.click()
        logging.info("🖱️ Clic en el 'Radio Button' correspondiente")

        # 🔥 ESPERAR que el DOM cambie
        wait.until(EC.staleness_of(radio_seleccion))

        wait.until(EC.invisibility_of_element_located((By.ID, "ID_MODAL_PROCESS")))

        # =========================
        # LOADER vs ERROR vs BOTÓN
        # =========================
        error_modal = (By.ID, "divAlertaErrorGeneral")

        btn_locator = (
            (By.ID, "ContentPlaceHolder1_btnIncluir")
            if tipo_proceso == 'IN'
            else (By.ID, "ContentPlaceHolder1_btnRenovar")
        )

        resultadobtn = wait.until(
            EC.any_of(
                EC.visibility_of_element_located(error_modal),
                EC.element_to_be_clickable(btn_locator)
            )
        )

        if resultadobtn.get_attribute("id") == "divAlertaErrorGeneral":
        #if driver.find_elements(By.XPATH, "//div[@id='divAlertaErrorGeneral' and not(contains(@style,'display: none'))]"):
            logging.warning("⚠️ Apareció el modal con advertencia")
            #mensaje_error = driver.find_element(By.XPATH, "//div[@id='divAlertaErrorGeneral']//p/span[2]").text.strip()
            #mensaje_error = resultado.find_element(By.XPATH, ".//p/span[2]").text.strip()
            mensaje_error = driver.find_element(
                By.XPATH, "//div[@id='divAlertaErrorGeneral']//p/span[2]"
            ).text.strip()
            raise Exception(mensaje_error)

        # =========================
        # INCLUSIÓN
        # =========================
        if tipo_proceso == 'IN':

            btn_incluir = wait.until(EC.element_to_be_clickable(btn_locator))
            btn_incluir.click()
            logging.info("🖱️ Clic en Incluir")

            wait.until(EC.staleness_of(btn_incluir))

            # =========================
            # MODAL OPCIONAL vs SIGUIENTE PASO
            # =========================
            modal_incluir = (By.ID, "divTipoIncluir")
            file_input = (By.ID, "fuPlanillaAjax")

            resultado2 = wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(modal_incluir),
                    EC.presence_of_element_located(file_input)
                )
            )

            # 🔎 si aparece modal
            if resultado2.get_attribute("id") == "divTipoIncluir":
                logging.warning("⚠️ Apareció el modal con advertencia")

                btn_aceptar = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.ID,
                            "ContentPlaceHolder1_btnIncluirProformaSi"
                            if tipo_mes == 'MA'
                            else "ContentPlaceHolder1_btnIncluirProformaNo"
                        )
                    )
                )

                nom_btn = 'Si' if tipo_mes == 'MA' else 'No'
                btn_aceptar.click()
                logging.info(f"🖱️ Clic en {nom_btn}")

            # =========================
            # DEUDA vs CONTINUAR
            # =========================
            deuda_modal = (By.ID, "divAlertas")

            resultado3 = wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(deuda_modal),
                    EC.presence_of_element_located(file_input)
                )
            )

            if resultado3.get_attribute("id") == "divAlertas":
                logging.warning("⚠️ Apareció el modal con advertencia")
                mensaje = driver.find_element(By.ID, "spMensaje").text.strip()
                raise Exception(mensaje)

            # =========================
            # INPUT FILE
            # =========================
            input_file = wait.until(EC.presence_of_element_located(file_input))
            logging.info("⌛ Página de carga lista (Trama)")

            # =========================
            # FECHA
            # =========================
            input_fecha = wait.until(
                EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniFecVigInc"))
            )
            input_fecha.clear()

            driver.execute_script("""
                arguments[0].value = arguments[1]; 
                arguments[0].dispatchEvent(new Event('change'));
                arguments[0].dispatchEvent(new Event('blur'));
            """, input_fecha, ramo.f_inicio)

            # =========================
            # ALERTA FECHA vs CONTINUAR
            # =========================
            alerta_fecha = (By.ID, "divAlertas")

            resultado4 = wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(alerta_fecha),
                    EC.element_to_be_clickable((By.TAG_NAME, "body")) #puede cambiar
                )
            )

            if resultado4.get_attribute("id") == "divAlertas":
                logging.warning("⚠️ Apareció el modal con advertencia")
                mensaje = driver.find_element(By.ID, "spMensaje").text
                raise Exception(mensaje)

            driver.find_element(By.TAG_NAME, "body").click()

            # =========================
            # LOG FECHA
            # =========================
            fecha_date = datetime.strptime(ramo.f_inicio, "%d/%m/%Y").date()

            if fecha_date < get_fecha_hoy().date():
                diferencia = (get_fecha_hoy().date() - fecha_date).days
                cant_dias = abs(diferencia)
                text_dia = "día" if cant_dias == 1 else "días"
                logging.info(
                    f"📅 Vigencia retroactiva: {ramo.f_inicio} ({cant_dias} {text_dia})"
                )
            else:
                logging.info(f"📅 Vigencia ingresada: {ramo.f_inicio}")

            # =========================
            # CASO ESPECIAL
            # =========================
            if ramo.poliza == '9231375':

                select_tipo_calculo = wait.until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_cboTipoCalculoEspecial"))
                )
                Select(select_tipo_calculo).select_by_value("21")

                logging.info("✅ Se seleccionó 'Por días'")

        # =========================
        # RENOVACIÓN
        # =========================
        else:

            btn_renovar = wait.until(EC.element_to_be_clickable(btn_locator))
            btn_renovar.click()
            logging.info("🖱️ Clic en Renovar")

            # =========================
            # DEUDA vs MODAL RENOVACIÓN
            # =========================
            deuda_modal = (By.ID, "divAlertas")
            modal_renovacion = (By.ID, "divTipoRenovacion")

            resultado5 = wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(deuda_modal),
                    EC.visibility_of_element_located(modal_renovacion)
                )
            )

            if resultado5.get_attribute("id") == "divAlertas":
                mensaje = driver.find_element(By.ID, "spMensaje").text.strip()
                raise Exception(mensaje)

            # =========================
            # ACEPTAR RENOVACIÓN
            # =========================
            btn_aceptar = wait.until(
                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnRenovacionSi"))
            )
            btn_aceptar.click()
            logging.info("🖱️ Clic en aceptar la Renovación")

            # =========================
            # SIGUIENTE PANTALLA
            # =========================
            input_file = wait.until(
                EC.presence_of_element_located((By.ID, "fuPlanillaAjax"))
            )
            logging.info("⌛ Página de carga lista (Trama)")

            # =========================
            # CASO ESPECIAL
            # =========================
            if ramo.poliza == '9231375':

                fecha_str = wait.until(
                    EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniVigPoliza"))
                ).get_attribute("value")

                fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
                fecha_final = (fecha + relativedelta(months=1) - timedelta(days=1)).strftime("%d/%m/%Y")

                input_fecha = wait.until(
                    EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtFinVigPoliza"))
                )

                input_fecha.clear()
                driver.execute_script("""
                    arguments[0].value = arguments[1]; 
                    arguments[0].dispatchEvent(new Event('change'));
                    arguments[0].dispatchEvent(new Event('blur'));
                """, input_fecha, fecha_final)

                driver.find_element(By.TAG_NAME, "body").click()

                logging.info(f"📅 Vigencia Hasta: {fecha_final}")

        #--------------
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
            validar_btn = wait.until(EC.visibility_of_element_located((By.ID, "btnValidarPlanillaExcel")))
            validar_btn.click()
            logging.info("🖱️ Clic en Validar Planilla")
        except Exception as e:
            raise Exception("No se puede Validar la Trama")

        try:
            WebDriverWait(driver,5).until(EC.alert_is_present())
            alerTrama = driver.switch_to.alert   
            logging.info("⚠️ Modal de Alerta")   
            alerTrama.accept()
            logging.info("✅ Alerta aceptada")
            raise Exception("Posiblemente el tipo de Trabajador no tiene un formato correcto")
        except TimeoutException:
            pass

        #---------------------------------------------------------------------------------------------------------------------------------------------------------
        error_validacion = (By.ID, "divAlertaErrorValidacion")
        btn_errores = (By.ID, "btnErroresPlanilla")
        progress_bar = (By.ID, "progressbar")
        
        resultado = wait.until(
            EC.any_of(
                EC.visibility_of_element_located(error_validacion),
                EC.visibility_of_element_located(btn_errores),
                EC.visibility_of_element_located(progress_bar)
            )
        )

        # 🔎 ERROR DE VALIDACIÓN (modal rojo)
        modal = driver.find_elements(*error_validacion)

        if modal and modal[0].is_displayed():
            logging.warning("⚠️ Apareció el modal con errores")

            mensaje_error = driver.find_element(By.ID, "spanAlertaErrorValidacion").text
            raise Exception(mensaje_error)

        # 🔎 ERRORES EN PLANILLA
        btn = driver.find_elements(*btn_errores)

        if btn and btn[0].is_displayed():

            cantidad_errores = driver.find_element(By.ID, "spnContadorError").text
            cantidad_texto = "error" if cantidad_errores == '1' else "errores"

            driver.execute_script("arguments[0].click();", btn[0])

            wait.until(EC.visibility_of_element_located((By.ID, "divErrorPlanilla")))

            filas = wait.until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, "//table[@id='gridErrorPlanilla']//tr[contains(@class,'jqgrow')]")
                )
            )

            errores = []

            for fila in filas:
                error = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sError']").text
                nombres = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sNombres']").get_attribute("title")
                paterno = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sPaterno']").get_attribute("title")
                materno = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sMaterno']").get_attribute("title")
                nrodoc = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sNroDoc']").get_attribute("title")

                errores.append(f"{error} - {nombres} {paterno} {materno} | {nrodoc}")

            detalle_errores = "\n".join(errores)

            raise Exception(f"Se encontró {cantidad_errores} {cantidad_texto} en la planilla:\n{detalle_errores}")

        # 👇 si aparece progress, esperar a 100%
        if driver.find_elements(*progress_bar):
            wait.until(lambda d: "100%" in d.find_element(By.ID, "progressbar").text)
            logging.info("✅ Proceso completado al 100%")

        # btn_calcularPrima = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnCalcular")))
        # btn_calcularPrima.click()
        # logging.info("🖱️ Clic en Calcular Prima")


        #---------------------------------------------------------------------------------------------------------------------------------------------------------
        # try:
        #     # Posible error para la hoja de la trama , debe ser Planilla no Trabajadores
        #     WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "divAlertaErrorValidacion")))
        #     logging.info(f"⚠️ Apareció el modal con errores")
        #     mensaje_error = wait.until(EC.visibility_of_element_located((By.ID, "spanAlertaErrorValidacion"))).text
        #     raise Exception(mensaje_error)
        # except TimeoutException:

        #     try:
        #         WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "btnErroresPlanilla")))

        #         cantidad_errores = wait.until(EC.visibility_of_element_located((By.ID, "spnContadorError"))).text
        #         cantidad_texto = "error" if cantidad_errores == '1' else "errores"

        #         link = wait.until(EC.element_to_be_clickable((By.ID, "btnErroresPlanilla")))
        #         driver.execute_script("arguments[0].click();", link)

        #         wait.until(EC.visibility_of_element_located((By.ID, "divErrorPlanilla")))

        #         filas = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='gridErrorPlanilla']//tr[contains(@class,'jqgrow')]")))

        #         errores = []

        #         for fila in filas:

        #             error = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sError']").text

        #             nombres = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sNombres']").get_attribute("title")

        #             paterno = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sPaterno']").get_attribute("title")

        #             materno = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sMaterno']").get_attribute("title")

        #             nrodoc = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sNroDoc']").get_attribute("title")

        #             errores.append(f"{error} - {nombres} {paterno} {materno} | {nrodoc}")

        #         detalle_errores = "\n".join(errores)

        #         raise Exception(f"Se encontró {cantidad_errores} {cantidad_texto} en la planilla:\n{detalle_errores}")

        #     except TimeoutException:
        #         pass
        #---------------------------------------------------------------------------------------------------------------------------------------------------------
        # #ESTO FUNCIONA
        # alerta_detectada = False

        # try:
        #     WebDriverWait(driver,5).until(EC.alert_is_present())
        #     alert = driver.switch_to.alert   
        #     logging.info(f"⚠️ Alerta: {alert.text}")   
        #     alert.accept()
        #     alerta_detectada = True
        # except TimeoutException:
        #     pass
            
        # if alerta_detectada:
        #     raise Exception("Hubo una alerta con la Trama")
        # else:
                             
        #     procesar_btn = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnProcesar")))

        #     driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", procesar_btn)

        #     procesar_btn.click()
        #     logging.info(f"🖱️ Clic en Procesar {palabra_clave}") 

        #     texto_mensaje = ""

        #     if tipo_proceso == 'IN':

        #         #aca carga un modal indicado si estas seguro de realziar la operacion de Inclusion

        #         btn_si = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnIncluirSi")))
        #         btn_si.click()
        #         logging.info("🖱️ Clic en Aceptar la Inclusión")

        #         time.sleep(3)

        #         mensaje_span = wait.until(EC.visibility_of_element_located((By.ID, "spanAlertaInclusion")))
        #         texto_mensaje = mensaje_span.text

        #     else:
                
        #         mensaje_span = wait.until(EC.visibility_of_element_located((By.ID, "spanAlertaRenovacion")))
        #         texto_mensaje = mensaje_span.text

        #         btn_si = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAceptarRenovacion")))
        #         btn_si.click()
        #         logging.info("🖱️ Clic en Aceptar la Renovación")

        #         time.sleep(3)
            
        #     logging.info(f"📩 Texto del mensaje: {texto_mensaje}")

        #     numero_doc = None

        #     if tipo_proceso == 'IN':

        #         match = re.search(r'Inclusión con nro\. (\d+)', texto_mensaje)
        #         if match:
        #             numero_doc = match.group(1)
        #             logging.info(f"✅ Número de {palabra_clave} extraído: {numero_doc}")
        #         else:
        #             logging.info(f"No se encontró el código de {palabra_clave}.")

        #     else:

        #         match = re.search(r'renovación con nro\. (\d+)', texto_mensaje)
        #         if match:
        #             numero_doc = match.group(1)
        #             logging.info(f"✅ Número de {palabra_clave} extraído: {numero_doc}")
        #         else:
        #             logging.info(f"No se encontró el código de {palabra_clave}.")

        #     prefijo = "INCL" if tipo_proceso == "IN" else "RENV"

        #     if not numero_doc:
        #         raise Exception(f"No se pudo obtener el número de {palabra_clave} desde el mensaje: {texto_mensaje}")

        #     codigo_documento = f"{prefijo}-{numero_doc}"

        #     try:

        #         if tipo_proceso == 'IN':
        #             btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAceptarInclusion")))
        #             btn_aceptar.click()
        #             logging.info("🖱️ Clic en Aceptar")
        #             time.sleep(3)
        #         else:
        #             try:
        #                 btn_aceptar = WebDriverWait(driver,7).until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAceptarRenovacion")))
        #                 btn_aceptar.click()
        #                 logging.info("🖱️ Clic en Aceptar")
        #             except TimeoutException:
        #                 pass

        #         wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-widget-overlay")))

        #         try:
        #             span_numero = wait.until(EC.visibility_of_element_located((By.XPATH, f"//span[contains(text(),'{codigo_documento}')]")))
        #             logging.info(f"✅ Span encontrado: {span_numero.text}")
        #         except TimeoutException:
        #             raise Exception(f"No se encontró el span con número {codigo_documento}")

        #         try:
        #             # Esperar que aparezca la lupa (por cualquiera de los dos tipos) -- Esto puede fallar si solo es una póliza
        #             if len(list_polizas) == 1 and ba_codigo == '1':
        #                 selector_xpath = (f"//img[@data-nropolizasalud='{ramo.poliza}']")
        #             elif len(list_polizas) == 1 and ba_codigo == '2':
        #                 selector_xpath = (f"//img[@data-nropolizapension='{ramo.poliza}']")
        #             else:
        #                 selector_xpath = (f"//img[@data-nropolizasalud='{list_polizas[0]}' or @data-nropolizapension='{list_polizas[1]}']")

        #             lupa_element = wait.until(EC.element_to_be_clickable((By.XPATH, selector_xpath)))

        #             driver.execute_script("arguments[0].scrollIntoView(true);", lupa_element)
        #             lupa_element.click()
        #             logging.info(f"🖱️ Clic en la lupa {codigo_documento} para póliza {ramo.poliza}")

        #         except TimeoutException:
        #             raise Exception(f"No se encontró la lupa {codigo_documento} asociada a la póliza {ramo.poliza}")
        #         except Exception as e:
        #             raise Exception(f"Error al hacer clic en la lupa -> {e}")

        #         try:
        #             WebDriverWait(driver,7).until(EC.element_to_be_clickable((By.ID, "btnAceptarError")))
        #             raise Exception (f"Se detecto una advertencia, el código para descagar el documento en la compañía es {codigo_documento}.")
        #         except TimeoutException:
        #             pass

        #         wait.until(EC.visibility_of_element_located((By.ID, "divPanelPDFMaster")))
        #         time.sleep(3)
        #         logging.info("📄 Panel PDF de la constancia visible.")

        #         logging.info("----------------------------")

        #         boton_guardar = wait.until(EC.element_to_be_clickable((By.ID, "btnDescargarConstanciaM")))

        #         # Guardar archivos antes del clic
        #         archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

        #         driver.execute_script("arguments[0].click();", boton_guardar)
        #         logging.info("🖱 Clic con JS en el botón de descarga")

        #         archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=1)
        #         logging.info(f"✅ Archivo descargado exitosamente")

        #         if archivo_nuevo:
        #             ruta_original = archivo_nuevo[0]  # ya es ruta completa
        #             ruta_final = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
        #             os.rename(ruta_original, ruta_final)
        #             logging.info(f"🔄 Constancia renombrado a '{ramo.poliza}.pdf'")
        #         else:
        #             raise Exception("No se encontró constancia después de la descarga")

        #         if tipo_mes == 'MA':

        #             logging.info("----------------------------")

        #             driver.switch_to.frame("ifContenedorPDFMaster")#iframe     
                                
        #             boton_embebido = wait.until(EC.element_to_be_clickable((By.ID, "open-button")))

        #             # Guardar archivos antes del clic
        #             archivos_antes_2 = set(os.listdir(ruta_archivos_x_inclu))

        #             driver.execute_script("arguments[0].click();", boton_embebido)
        #             logging.info("🖱 Clic con JS en el botón de descarga")

        #             endoso = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes_2,".pdf",cantidad=1)
        #             logging.info(f"✅ Archivo descargado exitosamente")

        #             if archivo_nuevo:
        #                 ruta_original = endoso[0]  # ya es ruta completa
        #                 ruta_final = os.path.join(ruta_archivos_x_inclu, f"endoso_{ramo.poliza}.pdf")
        #                 os.rename(ruta_original, ruta_final)
        #                 logging.info(f"🔄 Endoso renombrado a 'endoso_{ramo.poliza}.pdf'")
        #             else:
        #                 raise Exception("No se descargo el endoso")

        #             driver.switch_to.default_content()
            
        #         time.sleep(1)

        #         if len(list_polizas) == 2:

        #             ruta_salud = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[0]}.pdf")
        #             ruta_pension = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[1]}.pdf")
                
        #             logging.info("----------------------------")

        #             if os.path.exists(ruta_salud):
        #                 shutil.copy2(ruta_salud,ruta_pension)
        #                 logging.info(f"📄 Copia creada para constancia de Pensión")
        #             else:
        #                 logging.error(f"❌ No existe el archivo base Constancia Salud")

        #             logging.info("----------------------------")

        #             if tipo_mes == 'MA':

        #                 ruta_endoso_salud = os.path.join(ruta_archivos_x_inclu,f"endoso_{list_polizas[0]}.pdf")
        #                 ruta_endoso_pension = os.path.join(ruta_archivos_x_inclu,f"endoso_{list_polizas[1]}.pdf")

        #                 if os.path.exists(ruta_endoso_salud):
        #                     shutil.copy2(ruta_endoso_salud,ruta_endoso_pension)
        #                     logging.info(f"📄 Copia creada para el endoso de Pensión")
        #                 else:
        #                     logging.error(f"❌ No existe el archivo base Endoso Salud")

        #                 logging.info("----------------------------")

        #         #btnPDFCancelarM,btnPDFEnviarM
        #         btn_cancelar_boton = wait.until(EC.element_to_be_clickable((By.ID, "btnPDFCancelarM")))
        #         btn_cancelar_boton.click()
        #         logging.info("✅ Cerrando panel de documentos.")

        #         time.sleep(2)

        #         logging.info(f"✅ {palabra_clave} en La Positiva realizada exitosamente")

        #     except Exception as e:
        #         logging.error(f"Detalles del error: {str(e)}")
        #         raise Exception(
        #             f"No reprocesar la {palabra_clave}, ya que se generó la solicitud con el código -> {codigo_documento}.\n"
        #             f"Buscar y descargar manualmente en la compañía."
        #         )
        
        #---------------------------------------------------------------------------------------------------------------------------------------------------------    
        # PROBANDO MENOS TIEMPO
        # btn_procesar_locator = (By.ID, "ContentPlaceHolder1_btnProcesar")

        # def boton_habilitado(driver):
        #     try:
        #         btn = driver.find_element(*btn_procesar_locator)
        #         return btn if btn.is_displayed() and btn.is_enabled() and "ui-state-disabled" not in btn.get_attribute("class") else False
        #     except:
        #         return False

        # resultadohb = wait.until(
        #     EC.any_of(
        #         EC.alert_is_present(),
        #         boton_habilitado
        #     )
        # )

        # # 🔎 ALERTA
        # try:
        #     alert = driver.switch_to.alert
        #     texto = alert.text
        #     alert.accept()
        #     raise Exception(f"Hubo una alerta con la Trama | {texto}")

        # except NoAlertPresentException:

        #     for _ in range(3):
        #         try:
        #             driver.execute_script("arguments[0].scrollIntoView({block:'center'});", resultadohb)
        #             time.sleep(0.5)

        #             resultadohb.click()
        #             logging.info(f"🖱️ Clic en Procesar {palabra_clave}")
        #             break

        #         except (StaleElementReferenceException, ElementClickInterceptedException):
        #             time.sleep(1)
        #             resultadohb = wait.until(boton_habilitado)

        #     else:
        #         driver.execute_script("arguments[0].click();", resultadohb)
        #         logging.info("🖱️ Clic forzado con JS")

        #------ CALCULAR Y PROCESAR ----------
        btn_procesar_locator = (By.ID, "ContentPlaceHolder1_btnProcesar")
        btn_calcular_locator = (By.ID, "ContentPlaceHolder1_btnCalcular")

        def boton_procesar_habilitado(driver):
            try:
                btn = driver.find_element(*btn_procesar_locator)
                return btn if (
                    btn.is_displayed()
                    and btn.is_enabled()
                    and "ui-state-disabled" not in btn.get_attribute("class")
                ) else False
            except:
                return False

        def boton_calcular_habilitado(driver):
            try:
                btn = driver.find_element(*btn_calcular_locator)

                if (
                    btn.is_displayed()
                    and btn.is_enabled()
                    and btn.get_attribute("disabled") is None
                    and "ui-state-disabled" not in btn.get_attribute("class")
                ):
                    return btn

                return False
            except:
                return False

        # 🔍 Intentar Calcular (máx 2 veces por seguridad)
        for intento in range(2):

            btn_calcular = boton_calcular_habilitado(driver)

            if not btn_calcular:
                break  # 👉 ya no se puede calcular, ir a procesar

            try:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});",
                    btn_calcular
                )
                time.sleep(0.5)

                btn_calcular.click()
                logging.info(f"🖱️ Clic en Calcular ({intento+1})")

                # 🔥 Esperar cambio REAL
                wait.until(
                    EC.any_of(
                        EC.alert_is_present(),
                        boton_procesar_habilitado
                    )
                )

            except (StaleElementReferenceException, ElementClickInterceptedException):
                logging.warning("⚠️ Error al hacer clic en Calcular, reintentando...")
                time.sleep(1)
                continue


        # 🟢 PROCESAR (sí o sí)
        btn_procesar = wait.until(boton_procesar_habilitado)

        for _ in range(3):
            try:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});",
                    btn_procesar
                )
                time.sleep(0.5)

                btn_procesar.click()
                logging.info(f"🖱️ Clic en Procesar {palabra_clave}")
                break

            except (StaleElementReferenceException, ElementClickInterceptedException):
                time.sleep(1)
                btn_procesar = wait.until(boton_procesar_habilitado)

        else:
            driver.execute_script("arguments[0].click();", btn_procesar)
            logging.info("🖱️ Clic forzado en Procesar (JS)")

        # MENSAJE + BOTONES
        texto_mensaje = ""

        if tipo_proceso == 'IN':

            btn_si = (By.ID, "ContentPlaceHolder1_btnIncluirSi")
            mensaje = (By.ID, "spanAlertaInclusion")

            wait.until(
                EC.any_of(
                    EC.element_to_be_clickable(btn_si),
                    EC.visibility_of_element_located(mensaje)
                )
            )

            if driver.find_elements(*btn_si):
                driver.find_element(*btn_si).click()
                logging.info("🖱️ Clic en Aceptar la Inclusión")

            texto_mensaje = wait.until(
                EC.visibility_of_element_located(mensaje)
            ).text

        else:

            mensaje = (By.ID, "spanAlertaRenovacion")
            btn_si = (By.ID, "ContentPlaceHolder1_btnAceptarRenovacion")

            wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(mensaje),
                    EC.element_to_be_clickable(btn_si)
                )
            )

            texto_mensaje = wait.until(
                EC.visibility_of_element_located(mensaje)
            ).text

            driver.find_element(*btn_si).click()
            logging.info("🖱️ Clic en Aceptar la Renovación")

        logging.info(f"📩 Texto del mensaje: {texto_mensaje}")

        # EXTRAER NÚMERO DE LA SOLICITUD
        numero_doc = None

        if tipo_proceso == 'IN':
            match = re.search(r'Inclusión con nro\. (\d+)', texto_mensaje)
        else:
            match = re.search(r'renovación con nro\. (\d+)', texto_mensaje)

        if match:
            numero_doc = match.group(1)
            logging.info(f"✅ Número de {palabra_clave} extraído: {numero_doc}")
        else:
            raise Exception(f"No se encontró el código en el mensaje: {texto_mensaje}")

        prefijo = "INCL" if tipo_proceso == "IN" else "RENV"
        codigo_documento = f"{prefijo}-{numero_doc}"

        # BOTÓN ACEPTAR FINAL SOLO PARA INCLUSIÓN
        btn_inc = (By.ID, "ContentPlaceHolder1_btnAceptarInclusion")
        #btn_ren = (By.ID, "ContentPlaceHolder1_btnAceptarRenovacion")

        if tipo_proceso == 'IN':

            btn_aceptar = wait.until(EC.element_to_be_clickable(btn_inc))
            btn_aceptar.click()
            logging.info("🖱️ Clic en Aceptar")

        # ESPERAR QUE CIERRE MODAL
        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-widget-overlay")))

        # VALIDAR CÓDIGO
        span_numero = wait.until(EC.visibility_of_element_located((By.XPATH, f"//span[contains(text(),'{codigo_documento}')]")))
        logging.info(f"✅ Span encontrado: {span_numero.text}")

        # LUPA vs ERROR
        if len(list_polizas) == 1 and ba_codigo == '1':
            selector_xpath = f"//img[@data-nropolizasalud='{ramo.poliza}']"
        elif len(list_polizas) == 1 and ba_codigo == '2':
            selector_xpath = f"//img[@data-nropolizapension='{ramo.poliza}']"
        else:
            selector_xpath = f"//img[@data-nropolizasalud='{list_polizas[0]}' or @data-nropolizapension='{list_polizas[1]}']"

        lupa = (By.XPATH, selector_xpath)
        error_btn = (By.ID, "btnAceptarError")

        resultadol = wait.until(
            EC.any_of(
                EC.element_to_be_clickable(lupa),
                EC.element_to_be_clickable(error_btn)
            )
        )

        if resultadol.get_attribute("id") == "btnAceptarError":
            raise Exception(f"Advertencia detectada. Código: {codigo_documento}")

        # 🟢 Caso lupa
        for _ in range(3):
            try:
                resultadol.click()
                logging.info(f"🖱️ Clic en la lupa {codigo_documento}")
                break
            except StaleElementReferenceException:
                resultado = wait.until(EC.element_to_be_clickable(lupa))

        # PANEL PDF
        wait.until(EC.visibility_of_element_located((By.ID, "divPanelPDFMaster")))
        logging.info("📄 Panel PDF visible")

        # DESCARGA CONSTANCIA
        boton_guardar = wait.until(EC.element_to_be_clickable((By.ID, "btnDescargarConstanciaM")))

        archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

        driver.execute_script("arguments[0].click();", boton_guardar)
        logging.info(f"🖱️ Clic en Descargar Constancia")

        archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu, archivos_antes, ".pdf", cantidad=1)

        if archivo_nuevo:
            ruta_final = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
            os.rename(archivo_nuevo[0], ruta_final)
            logging.info(f"🔄 Constancia renombrada")
        else:
            raise Exception("No se descargó constancia")

        # ENDOSO (MA)
        if tipo_mes == 'MA':

            driver.switch_to.frame("ifContenedorPDFMaster")

            boton_embebido = wait.until(EC.element_to_be_clickable((By.ID, "open-button")))

            archivos_antes_2 = set(os.listdir(ruta_archivos_x_inclu))

            driver.execute_script("arguments[0].click();", boton_embebido)
            logging.info(f"🖱️ Clic en Descargar Endoso")

            endoso = esperar_archivos_nuevos(ruta_archivos_x_inclu, archivos_antes_2, ".pdf", cantidad=1)

            if endoso:
                ruta_final = os.path.join(ruta_archivos_x_inclu, f"endoso_{ramo.poliza}.pdf")
                os.rename(endoso[0], ruta_final)
                logging.info("🔄 Endoso renombrado")
            else:
                raise Exception("No se descargó el endoso")

            driver.switch_to.default_content()

        # COPIAS
        if len(list_polizas) == 2:

            ruta_salud = os.path.join(ruta_archivos_x_inclu, f"{list_polizas[0]}.pdf")
            ruta_pension = os.path.join(ruta_archivos_x_inclu, f"{list_polizas[1]}.pdf")

            if os.path.exists(ruta_salud):
                shutil.copy2(ruta_salud, ruta_pension)

            if tipo_mes == 'MA':
                ruta_endoso_salud = os.path.join(ruta_archivos_x_inclu, f"endoso_{list_polizas[0]}.pdf")
                ruta_endoso_pension = os.path.join(ruta_archivos_x_inclu, f"endoso_{list_polizas[1]}.pdf")

                if os.path.exists(ruta_endoso_salud):
                    shutil.copy2(ruta_endoso_salud, ruta_endoso_pension)

        # CERRAR PANEL
        wait.until(EC.element_to_be_clickable((By.ID, "btnPDFCancelarM"))).click()

        logging.info(f"✅ {palabra_clave} realizada exitosamente")

def solicitud_vidaley_MV(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_mes,tipo_proceso,actividad,ramo):
 
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
        #logging.info("✅ No apareció ninguna alerta")
        pass

    resultado,asunto = validar_pagina(driver)

    if not resultado:
        raise Exception (f"{asunto}")
                
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

    # # Primero verificar si aparece un alert
    # try:
    #     WebDriverWait(driver,5).until(EC.alert_is_present())
    #     alert_v2 = driver.switch_to.alert
    #     msj_alert =  alert_v2.text
    #     logging.info(f"⚠️ Alerta : {msj_alert}")
    #     alert_v2.accept()
    #     logging.info("✅ Alerta aceptada")
    #     raise Exception (msj_alert)
    # except TimeoutException:
    #     pass

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
        texto_alerta = alert1.text.strip()

        # Validar texto
        if "Esta seguro de registrar la Solicitud, luego de ello no podrá ser modificada." not in texto_alerta:
            logging.info(f"⚠️ Alerta : {texto_alerta}")
            alert1.accept()
            logging.info("✅ Alerta aceptada")
            raise Exception(texto_alerta)

        logging.info(f"⚠️ Alerta : ¿{texto_alerta}?")
        alert1.accept()
        logging.info("✅ Alerta aceptada")
    except TimeoutException:
        #logging.info("✅ No apareció la primera alerta, revisamos si salió el iframe...")

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

def solicitud_vidaley_MA(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_mes,tipo_proceso,ramo):

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
        raise str(e)

    time.sleep(3)

    try:
        input_element = wait.until(EC.element_to_be_clickable((By.ID, "b12-b1-Input_PolicesNumber")))
        input_element.clear()
        input_element.send_keys(ramo.poliza) # 1005537087 ramo.poliza
        logging.info(f"✅ Numero de Póliza ingresado: {ramo.poliza}")
    except Exception as e:
        raise str(e)

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
    logging.info("🖱️ Clic en 'Buscar'")


    # 👇 Esperar cualquiera de los posibles resultados
    resultado1 = wait.until(
        EC.any_of(
            EC.visibility_of_element_located((By.ID, "b12-b11-TextTitlevalue")),   # error
            EC.presence_of_element_located((By.ID, "b12-limpiargrilla")),          # modal sin resultados
            EC.presence_of_element_located((By.ID, "b12-Widget_TransactionRecordList"))  # tabla
        )
    )

    # 👇 Evaluar qué apareció
    #html = driver.page_source

    # 🔴 Caso 1: mensaje de error
    if resultado1.get_attribute("id") == "b12-b11-TextTitlevalue":
    #if "b12-b11-TextTitlevalue" in html:
        tit = driver.find_element(By.ID, "b12-b11-TextTitlevalue").text
        con = driver.find_element(By.ID, "b12-b11-TextContentValue").text
        raise Exception(f"{tit} {con}")


    # 🔴 Caso 2: modal sin resultados
    if resultado1.get_attribute("id") == "b12-limpiargrilla":
    #elif "b12-limpiargrilla" in html:
        logging.error("❌ Se detectó un modal que no carga los resultados")
        raise Exception(f"No hay resultados para la póliza {ramo.poliza}")

    #----------------------------------------------------------------------------------------
    # logging.info("✅ Tabla cargada correctamente")

    # try:
    #     # fila = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"#b12-Widget_TransactionRecordList tbody tr")))
    #     filas = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,"#b12-Widget_TransactionRecordList tbody tr")))

    #     fila_objetivo = None

    #     for fila in filas:
    #         try:
    #             estado = fila.find_element(By.XPATH, ".//td[@data-header='Estado']//span").text.strip()
        
    #             if estado.lower() == "vigente":
    #                 fila_objetivo = fila
    #                 break

    #         except Exception:
    #             continue

    #     if not fila_objetivo:
    #         raise Exception(f"No se encontró fila Vigente para la póliza {ramo.poliza}")                                         # CIA          WEBHOOK

    #     inicio_vigencia = fila_objetivo.find_element(By.XPATH, ".//td[@data-header='Inicio de Vigencia']//span").text.strip() # 01/03/2026 # 01/04/2026

    #     fin_vigencia = fila_objetivo.find_element(By.XPATH, ".//td[@data-header='Fin de Vigencia']//span").text.strip()       # 01/04/2026 # 01/05/2026

    #     estado = fila_objetivo.find_element(By.XPATH, ".//td[@data-header='Estado']//span").text.strip()

    #     indice_fila = filas.index(fila_objetivo) + 1

    # except TimeoutException:
    #     raise Exception(f"Poliza {ramo.poliza} no figura con Modalidad -> {tipo_mes}")
    #----------------------------------------------------------------------------------------
    logging.info("✅ Tabla cargada correctamente")

    try:
        # 🔍 Buscar SOLO filas con estado "Vigente"
        filas_vigentes = wait.until(
            EC.presence_of_all_elements_located((
                By.XPATH,
                "//table[@id='b12-Widget_TransactionRecordList']//tbody/tr[.//td[@data-header='Estado']//span[normalize-space()='Vigente']]"
            ))
        )

        if not filas_vigentes:
            raise Exception(f"No se encontró fila Vigente para la póliza {ramo.poliza}")

        fila_valida = None

        # 🔁 Validar cada fila vigente
        for fila in filas_vigentes:
            try:
                inicio_vigencia = fila.find_element(
                    By.XPATH, ".//td[@data-header='Inicio de Vigencia']//span"
                ).text.strip()

                fin_vigencia = fila.find_element(
                    By.XPATH, ".//td[@data-header='Fin de Vigencia']//span"
                ).text.strip()

                # 🔍 Validaciones según tipo de proceso
                if tipo_proceso == 'IN':
                    if ramo.f_fin == fin_vigencia:
                        fila_valida = fila
                        break

                elif tipo_proceso == 'RE':
                    if ramo.f_inicio == fin_vigencia:
                        fila_valida = fila
                        break

            except StaleElementReferenceException:
                continue  # 🔁 reintenta con la siguiente fila

        # ❌ No encontró ninguna válida
        if not fila_valida:
            raise Exception(
                f"No se encontró una fila vigente válida para la póliza {ramo.poliza}\n"
                f"Solicitud: {ramo.f_inicio} al {ramo.f_fin}\n" 
                f"Compañía: {inicio_vigencia} al {fin_vigencia}"
            )

        # (opcional) índice de fila
        indice_fila = filas_vigentes.index(fila_valida) + 1

        # ✅ Log de éxito
        #logging.info(f"✅ Fila {indice_fila} vigente válida encontrada")

        checkbox = fila_valida.find_element(By.CSS_SELECTOR, "input[type='checkbox']")

        ActionChains(driver).move_to_element(checkbox).click().perform()

        logging.info(f"✅ Fila seleccionada: {indice_fila} | Póliza: {ramo.poliza} | Inicio: {inicio_vigencia} | Fin: {fin_vigencia}")

    except TimeoutException:
        raise Exception(f"Póliza {ramo.poliza} no figura con Modalidad -> {tipo_mes}")
    #----------------------------------------------------------------------------------------

    # checkbox = fila.find_element(By.CSS_SELECTOR, "input[type='checkbox']")

    # ActionChains(driver).move_to_element(checkbox).click().perform()

    # logging.info(f"✅ Fila seleccionada: {indice_fila} | Póliza: {ramo.poliza} | Inicio: {inicio_vigencia} | Fin: {fin_vigencia} | Estado: {estado}")

    # #---------------------------------------------------------------------------------------------
    # boton_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Buscar']]")))
    # boton_buscar.click()
    # logging.info(f"🖱️ Clic en 'Buscar'")

    # try:
    #     tit = WebDriverWait(driver,8).until(EC.visibility_of_element_located((By.ID, "b12-b11-TextTitlevalue"))).text
    #     con = WebDriverWait(driver,8).until(EC.visibility_of_element_located((By.ID, "b12-b11-TextContentValue"))).text
    #     raise Exception(f"{tit} {con}")
    # except TimeoutException:
    #     pass

    # wait.until(EC.presence_of_element_located((By.ID, "b12-Widget_TransactionRecordList")))
    # logging.info(f"⌛ Esperando la tabla con resultados")

    # try:
    #     WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID, "b12-limpiargrilla")))
    #     logging.error(f"❌ Se detectó un modal que no carga los resultados")
    #     raise Exception(f"No hay resultados para la póliza {ramo.poliza}")
    # except TimeoutException:
    #     logging.info("✅ No se detectó modal")

    # try:
    #     fila = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#b12-Widget_TransactionRecordList tbody tr")))
    # except TimeoutException:
    #     raise Exception(f"Poliza {ramo.poliza} no figura con Modalidad Mes Adelantado")

    # checkbox = fila.find_element(By.CSS_SELECTOR, "input[type='checkbox']")

    # actions = ActionChains(driver)
    # actions.move_to_element(checkbox).click().perform()
    # logging.info(f"✅ Se seleccionó la póliza: {ramo.poliza}")
    # #---------------------------------------------------------------------------------------------
    time.sleep(3)

    # =========================
    # CLICK EN PROCESO
    # =========================
    if tipo_proceso == 'IN':
        boton_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Inclusión']]")))
    else:
        boton_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Renovación']]")))

    boton_proceso.click()
    logging.info(f"🖱️ Clic en {palabra_clave}")

    # =========================
    # ESPERAR (ERROR O VENTANA)
    # =========================
    resultado2 = wait.until(
        EC.any_of(
            EC.visibility_of_element_located((By.ID, "b12-b2-b8-TextTitlevalue")),  # error
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))     # ventana lista
        )
    )

    if resultado2.get_attribute("id") == "b12-b2-b8-TextTitlevalue":
        titulo = resultado2.text
        contenido = driver.find_element(By.ID, "b12-b2-b8-TextContentValue").text
        raise Exception(f"{titulo} \n{contenido}")

    logging.info(f"🌐 Ventana de {palabra_clave} cargada correctamente")

    # =========================
    # SET FECHA (INCLUSIÓN)
    # =========================
    if tipo_proceso == 'IN':

        fecha_vigencia = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-EffectiveDateTyping")))
        driver.execute_script("""
            arguments[0].removeAttribute('readonly');
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, fecha_vigencia, ramo.f_inicio)
        logging.info(f"📅 Fecha Inicio: {ramo.f_inicio}")

    # =========================
    # SUBIR ARCHIVO
    # =========================
    ruta_archivo = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")

    if not os.path.exists(ruta_archivo):
        raise Exception(f"Archivo {ramo.poliza}.xlsx no encontrado")

    if tipo_proceso == 'IN':
        input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='b13-b1-b11-b2-b1-DropArea']//input[@type='file']")))
    else:
        input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='b13-b1-b1-b2-b1-DropArea']//input[@type='file']")))

    input_file.send_keys(ruta_archivo)
    logging.info(f"📂 Trama {ramo.poliza}.xlsx subida")

    # =========================
    # VALIDAR FORMATO (ERROR O BOTÓN VALIDAR)
    # =========================
    resultado3 = wait.until(
        EC.any_of(
            EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'La planilla no está en un formato válido de excel')]")),
            EC.element_to_be_clickable((By.ID, "b13-b1-btnValidate"))
        )
    )

    if "formato válido" in resultado3.text:
        raise Exception("La planilla no está en un formato válido de excel")

    logging.info("✅ Archivo válido")

    # CLICK VALIDAR
    boton_validar = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-btnValidate")))
    boton_validar.click()
    logging.info("🖱️ Clic en Validar")

    # RESULTADO FINAL (TODO EN UNO)
    resultado4 = wait.until(
        EC.any_of(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Aceptar']")),
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Lo sentimos')]")),
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Encontramos algunos errores en la planilla')]")),
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'La planilla superó exitosamente')]"))
        )
    )

    texto = resultado4.text.lower()

    # ⚠️ Caso: Advertencia (botón Aceptar)
    if resultado4.tag_name == "button":
        resultado4.click()
        logging.info("✅ Se aceptó advertencia de vigencia")

        # 👇 luego de aceptar, esperar resultado final
        resultado5 = wait.until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Lo sentimos')]")),
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Encontramos algunos errores en la planilla')]")),
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'La planilla superó exitosamente')]"))
            )
        )
        texto = resultado5.text.lower()

    # ❌ Error general
    if "lo sentimos" in texto:
        tit = driver.find_element(By.ID, "b13-b1-b5-TextTitlevalue").text
        con = driver.find_element(By.ID, "b13-b1-b5-TextContentValue").text
        raise Exception(f"{tit} \n{con}")

    # ❌ Error en planilla
    elif "encontramos algunos errores" in texto:
        raise Exception("Encontramos algunos errores en la planilla")

    # ✅ Éxito
    elif "superó exitosamente" in texto:
        logging.info("✅ Planilla validada correctamente")

    # #---------------------------------------------------------------------------------------------------------------------------

    # if tipo_proceso == 'IN':
    #     boton_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Inclusión']]")))
    # else:
    #     boton_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Renovación']]")))

    # boton_proceso.click()
    # logging.info(f"🖱️ Clic en {palabra_clave} ")

    # time.sleep(3)

    # try:

    #     titulo = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "b12-b2-b8-TextTitlevalue"))).text
    #     contenido = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "b12-b2-b8-TextContentValue"))).text
    #     raise Exception(f"{titulo} {contenido}")

    # except TimeoutException:
    #     pass

    # logging.info(f"--- Se ingresó a la ventana de {palabra_clave} 🌐----")

    # ruta_archivo = os.path.join(ruta_archivos_x_inclu,f"{ramo.poliza}.xlsx")

    # if tipo_proceso == 'IN':

    #     fecha_vigencia_solicitud = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-EffectiveDateTyping")))
    #     driver.execute_script("""
    #         arguments[0].removeAttribute('readonly');
    #         arguments[0].value = arguments[1];
    #         arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
    #         arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    #     """, fecha_vigencia_solicitud,ramo.f_inicio)
    #     logging.info(f"📅 Fecha de Inicio de la vigencia para la {palabra_clave}: {ramo.f_inicio}")

    # time.sleep(2)

    # if not os.path.exists(ruta_archivo):
    #     raise Exception (f"Archivo {ramo.poliza}.xlsx no encontrado")

    # if tipo_proceso == 'IN':
    #     input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='b13-b1-b11-b2-b1-DropArea']//input[@type='file']")))
    #     input_file.send_keys(ruta_archivo)
    # else:
    #     input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='b13-b1-b1-b2-b1-DropArea']//input[@type='file']")))
    #     input_file.send_keys(ruta_archivo)

    # logging.info(f"⌛ Trama {ramo.poliza}.xlsx subido para validar")

    # time.sleep(5)

    # try:
    #     msj = 'La planilla no está en un formato válido de excel'
    #     WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,f"//span[contains(text(), '{msj}')]")))
    #     raise Exception(msj)
    # except TimeoutException:
    #     logging.info("✅ No se detectó error de formato en la planilla")

    # boton_validar = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-btnValidate")))
    # boton_validar.click()
    # logging.info(f"🖱️ Clic en Validar")

    # try:
    #     boton_aceptar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Aceptar']")))
    #     boton_aceptar.click()
    #     logging.info("✅ Clic en el botón Aceptar.")
    # except TimeoutException:
    #     logging.info("✅ No se detectó advertencia de inicio de vigencia")
    #     #pass

    # try:
    #     msj = 'Lo sentimos'
    #     WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{msj}')]")))
    #     raise Exception(msj)
    # except TimeoutException:
    #     logging.info("✅ No se detectó errores en la Trama. Se continúa con el flujo normalmente.")
    #     #pass
        
    # try:
    #     wait.until(
    #         EC.any_of(
    #             EC.presence_of_element_located(
    #                 (By.XPATH, "//span[contains(text(),'La planilla superó exitosamente')]")
    #             ),
    #             EC.presence_of_element_located(
    #                 (By.XPATH, "//span[contains(text(),'Encontramos algunos errores en la planilla')]")
    #             )
    #         )
    #     )

    #     # Verificamos cuál mensaje apareció
    #     if driver.find_elements(By.XPATH, "//span[contains(text(),'Encontramos algunos errores en la planilla')]"):
    #         raise Exception("Encontramos algunos errores en la planilla. Descarga y corrige las observaciones")

    #     logging.info("✅ Planilla validada exitosamente. Continuando con el flujo.")

    # except TimeoutException:
    #     raise Exception("No se pudo determinar el resultado de la validación de la planilla.")

    # #---------------------------------------------------------------------------------------------------------------------------

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
    logging.info(f"⌨️ Ingresando correo '{ejecutivo_responsable}'")

    #--------------------------
    while True:
        resultado6 = wait.until(
            EC.any_of(
                EC.element_to_be_clickable((By.ID, "b13-CalculatePremium")),
                EC.element_to_be_clickable((By.ID, "b13-InclusionValidate")),
                EC.element_to_be_clickable((By.ID, "b13-RenewalRequest2"))
            )
        )

        boton_id = resultado6.get_attribute("id")

        # 🔵 Caso 1: Calcular
        if boton_id == "b13-CalculatePremium":
            driver.execute_script("arguments[0].scrollIntoView(true);", resultado6)
            resultado6.click()
            logging.info("🖱️ Clic en 'Calcular'")
            time.sleep(2)  # opcional: pequeño respiro
            continue  # 🔁 volver a esperar el siguiente paso

        # 🟢 Caso 2: Finalizar (cualquiera de los dos)
        elif boton_id in ["b13-InclusionValidate", "b13-RenewalRequest2"]:
            resultado6.click()
            logging.info(f"🖱️ Clic en 'Finalizar' ({boton_id})")
            break  # ✅ terminamos flujo
    #--------------------------

    # #if tipo_proceso == 'IN':
    #     # Esperar que el botón esté presente
    # try:
    #     boton_calcular = wait.until(EC.element_to_be_clickable((By.ID, "b13-CalculatePremium")))
    #     driver.execute_script("arguments[0].scrollIntoView(true);", boton_calcular)
    #     boton_calcular.click()
    #     logging.info("🖱️ Clic en 'Calcular'")
    # except TimeoutException:
    #     pass

    # if tipo_proceso == 'IN':   
    #     finalizar_btn = wait.until(EC.element_to_be_clickable((By.ID, "b13-InclusionValidate")))
    # else:
    #     finalizar_btn = wait.until(EC.element_to_be_clickable((By.ID, "b13-RenewalRequest2")))

    # finalizar_btn.click()
    # logging.info(f"🖱️ Clic en 'Finalizar'")

    # time.sleep(5)

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

    cantidad = 2 if tipo_mes == 'MA' else 1 

    archivos_nuevos = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad)

    if archivos_nuevos and len(archivos_nuevos) >= cantidad:
        logging.info(f"✅ {cantidad} documento(s) descargado(s) correctamente")

        archivos_nuevos = sorted(
            archivos_nuevos,
            key=os.path.getctime
        )

        # Siempre existe el primero
        nombre_1 = os.path.join(
            ruta_archivos_x_inclu,
            f"{ramo.poliza}.pdf"
        )

        os.rename(archivos_nuevos[0], nombre_1)
        logging.info(f"✅ Constancia '{ramo.poliza}.pdf' renombrado correctamente")

        # Solo si hay 2 archivos
        if cantidad == 2:
            nombre_2 = os.path.join(
                ruta_archivos_x_inclu,
                f"endoso_{ramo.poliza}.pdf"
            )

            os.rename(archivos_nuevos[1], nombre_2)
            logging.info(f"✅ Endoso 'endoso_{ramo.poliza}.pdf' renombrado correctamente")

    else:
        logging.error(f"❌ No se detectaron los {cantidad} archivo(s)")

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
     
        user_field = wait.until(EC.presence_of_element_located((By.ID, "b5-Input_User")))
        user_field.clear()

        mover_y_hacer_click_simple(driver, user_field)
        time.sleep(random.uniform(0.97, 0.99))

        escribir_lento(user_field, ramo.usuario, min_delay=0.97, max_delay=0.99)
        logging.info(f"⌨️ Digitando el Username")

        time.sleep(1 + random.random() * 1.5)

        password_field = wait.until(EC.presence_of_element_located((By.ID, "b5-Input_PassWord")))
        password_field.clear()

        mover_y_hacer_click_simple(driver, password_field)
        time.sleep(random.uniform(0.97, 0.99))

        escribir_lento(password_field, ramo.clave, min_delay=0.97, max_delay=0.99)
        logging.info(f"⌨️ Digitando el Password")

        time.sleep(1 + random.random() * 1.5)

        login_button = wait.until(EC.element_to_be_clickable((By.ID, "b5-btnAction")))
        mover_y_hacer_click_simple(driver, login_button)
        logging.info("🖱️ Clic en Iniciar Sesión")

        # blocked_account = (By.XPATH, "//span[contains(text(),'Cuenta bloqueada')]")
        # error_login = (By.XPATH, "//span[contains(text(),'Usuario o contraseña incorrectos')]")
        # change_password = (By.ID, "b5-b22-ChangePassword")
        # temp_blocked = (By.XPATH, "//span[contains(text(),'Cuenta inhabilitada temporalmente')]")
        autogestion_locator = (By.XPATH, "//div[contains(@class,'menu-item')]//span[normalize-space()='Autogestión']/parent::div")

        try:

            autogestion = wait.until(EC.element_to_be_clickable(autogestion_locator))
            logging.info("✅ Login exitoso")
            driver.execute_script("arguments[0].click();", autogestion)
            logging.info("🖱️ Clic en Autogestión")

            ventana_menu_positiva = driver.current_window_handle

        except TimeoutException:
            raise Exception("Problemas en el Inicio de Sesión, comunícate con el ejecutivo responsable")

    except Exception as e:
        logging.error(f"❌ Error entrando a la Positiva: {e}")
        tomar_capturar(driver, ruta_archivos_x_inclu, f"LOGIN_FALLIDO")
        return False,False,"Login Fallido", str(e)

    if bab_codigo in ['1', '2', '3']:

        try:
            solicitud_sctr(driver,wait,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,ba_codigo,ramo)
            return True, True if tipo_mes == 'MA' else False, tipoError, detalleError
        except Exception as e:
            resultado, asunto = validar_pagina(driver)
            tomar_capturar(driver, ruta_archivos_x_inclu, f"ERROR_SCTR_{tipo_mes}")
            detalle = f"{asunto}, intentar entre 5 a 10 minutos de nuevo" if not resultado else str(e)
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
        # "MV": lambda: solicitud_vidaley_MV(
        #     driver, wait, ruta_archivos_x_inclu, ruc_empresa,
        #     ejecutivo_responsable, palabra_clave,tipo_mes,
        #     tipo_proceso, actividad, ramo
        # ),
        "MV": lambda: solicitud_vidaley_MA(
            driver, wait, ruta_archivos_x_inclu, ruc_empresa,
            ejecutivo_responsable, palabra_clave,tipo_mes,
            tipo_proceso, ramo
        ),
        "MA": lambda: solicitud_vidaley_MA(
            driver, wait, ruta_archivos_x_inclu, ruc_empresa,
            ejecutivo_responsable, palabra_clave,tipo_mes,
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
        detalle = f"{asunto}, intentar entre 5 a 10 minutos de nuevo" if not resultado else str(e)
        return False, False, f"LAPO-VIDALEY-{tipo_mes}", detalle
    finally:
        driver.close()
        logging.info("✅ Cerrando la pestaña Vida Ley")
        driver.switch_to.window(driver.window_handles[0])
        logging.info("🔙 Retornando al menú principal")
