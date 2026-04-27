#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,NoAlertPresentException,StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from Tiempo.fechas_horas import get_fecha_hoy
from Compañias.Positiva.metodos import mover_y_hacer_click_simple,escribir_lento,validar_pagina,leer_pdf
from LinuxDebian.Ventana.ventana import esperar_archivos_nuevos
from Chrome.google import tomar_capturar
#---- Import ---
import os
import re
import logging
import time
import random
import shutil
import pandas as pd

# --- Variables globales ---
ventana_menu_positiva = None
login_exitoso = False

def solicitud_sctr(driver,wait,list_polizas,ruta_archivos_x_inclu,tipo_mes,palabra_clave,tipo_proceso,ba_codigo,ramo):
    
    sed = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='SED']")))
    sed.click()
    logging.info("🖱️ Clic en 'SED'")

    wait.until(lambda d: len(d.window_handles) > 1)

    driver.switch_to.window(driver.window_handles[-1])
    logging.info("⌛ Esperando la nueva ventana...")

    transacciones = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Transacciones y Consultas")))
    transacciones.click()
    logging.info("🖱️ Clic en Transacciones y Consultas")

    #descargar_documento_por_codigo(driver,wait,"RENV-140755","140755",ramo,ruta_archivos_x_inclu)

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

    buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnBuscar")))
    buscar_btn.click()
    logging.info("🖱️ Clic en Buscar")

    try:
        wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado")))
        logging.info("⌛ Esperando la tabla con los campos...")
    except Exception as e:
        raise Exception(f"No se encontró la póliza {ramo.poliza} en la compañía")

    tipo_facturacion = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado_lblTipoFacturacion_0"))).text

    logging.info(f"📌 Tipo de Facturación: {tipo_facturacion}")

    if ramo.facturacion and tipo_facturacion != "Facturación Multiple":
            raise Exception(f"El tipo de facturación de la póliza {ramo.poliza} es {tipo_facturacion} y no coincide con el de la compañia")

    #fecha_emision_element = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado_lblFechaEmision_0")))
    #fecha_emision_str = fecha_emision_element.text.strip()
    #fecha_emision = datetime.strptime(fecha_emision_str, "%d/%m/%Y")

    # if fecha_emision.month == 12:
    #     primer_dia_mes_siguiente = datetime(fecha_emision.year + 1, 1, 1)
    # else:
    #     primer_dia_mes_siguiente = datetime(fecha_emision.year, fecha_emision.month + 1, 1)

    #primer_dia_mes_siguiente_str = primer_dia_mes_siguiente.strftime("%d/%m/%Y")

    # En formato dateTime para hacer validaciones
    fecha_ramo_fin = datetime.strptime(ramo.f_fin, "%d/%m/%Y")

    fecha_vigencia_element = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_gvResultado_lblFechaVigencia_0")))
    fecha_vigencia_str = fecha_vigencia_element.text.strip() 
    fecha_vigencia_dmy = datetime.strptime(fecha_vigencia_str, "%d/%m/%Y")
    #fecha_vigencia_str_for = fecha_vigencia_dmy.strftime("%d/%m/%Y")

    logging.info(f"📅 Fecha Fin de la vigencia en la póliza : {fecha_vigencia_str}")

    # if tipo_proceso == 'IN':

    #     # # Definir rango permitido solo para inclusiones
    #     # rango_inicio = fecha_vigencia_dmy.date() - timedelta(days=2)  # 2 días antes
    #     # rango_fin = fecha_vigencia_dmy.date() + timedelta(days=2)     # 2 días después
    #     logging.info(f"📅 Rango permitido para la {palabra_clave}: {primer_dia_mes_siguiente_str} a {fecha_vigencia_str_for}")

    #except Exception as e:
        #raise Exception(f"Error al parsear la fecha de vigencia,Motivo - {e}")
    
    if tipo_proceso == 'IN':
        #Como es texto no hay problemas (valido para  ==  y !=)
        #if ramo.f_fin != fecha_vigencia_str:
        if fecha_ramo_fin != fecha_vigencia_dmy:
            raise Exception(f"Las fechas Fin de la vigencia en la Póliza {ramo.poliza} no coinciden")
    else: 

        if fecha_ramo_fin == fecha_vigencia_dmy:
            raise Exception(f"La Póliza {ramo.poliza} ya fue renovada hasta el {fecha_vigencia_str}.")

        elif fecha_ramo_fin < fecha_vigencia_dmy:
            raise Exception(
                f"La Póliza {ramo.poliza} tiene vigencia desfasada (atrasada).\n"
                f"Web: {fecha_vigencia_str} | Ramo: {ramo.f_fin}"
            )

    radio_seleccion = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_gvResultado_rbSeleccion_0")))
    radio_seleccion.click()
    logging.info("🖱️ Clic en el 'Radio Button' correspondiente")

    wait.until(EC.staleness_of(radio_seleccion))

    wait.until(EC.invisibility_of_element_located((By.ID, "ID_MODAL_PROCESS")))

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
        logging.warning("⚠️ Apareció el modal con advertencia")
        mensaje_error = driver.find_element(By.XPATH, "//div[@id='divAlertaErrorGeneral']//p/span[2]").text.strip()
        raise Exception(mensaje_error)

    if tipo_proceso == 'IN':

        btn_incluir = wait.until(EC.element_to_be_clickable(btn_locator))
        btn_incluir.click()
        logging.info("🖱️ Clic en Incluir")

        wait.until(EC.staleness_of(btn_incluir))

        modal_incluir = (By.ID, "divTipoIncluir")
        file_input = (By.ID, "fuPlanillaAjax")
        deuda_modal = (By.ID, "divAlertas")

        resultado2 = wait.until(
            EC.any_of(
                EC.visibility_of_element_located(modal_incluir),
                EC.presence_of_element_located(file_input),
                EC.visibility_of_element_located(deuda_modal),
            )
        )

        if resultado2.get_attribute("id") == "divAlertas":
            logging.warning("⚠️ Apareció el modal con advertencia")
            mensaje = driver.find_element(By.ID, "spMensaje").text.strip()
            raise Exception(mensaje)

        if resultado2.get_attribute("id") == "divTipoIncluir":
            logging.warning("⚠️ Apareció el modal con advertencia")

            btn_aceptar = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.ID,
                        "ContentPlaceHolder1_btnIncluirProformaSi"
                        if not ramo.facturacion #tipo_mes == 'MA'
                        else "ContentPlaceHolder1_btnIncluirProformaNo"
                    )
                )
            )

            nom_btn = 'Si' if tipo_mes == 'MA' else 'No'
            btn_aceptar.click()
            logging.info(f"🖱️ Clic en {nom_btn}")

        input_file = wait.until(EC.presence_of_element_located(file_input))
        logging.info("⌛ Página de carga lista (Trama)")

        input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniFecVigInc")))
        input_fecha.clear()

        driver.execute_script("""
            arguments[0].value = arguments[1]; 
            arguments[0].dispatchEvent(new Event('change'));
            arguments[0].dispatchEvent(new Event('blur'));
        """, input_fecha, ramo.f_inicio)

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

        if ramo.poliza == '9231375':

            select_tipo_calculo = wait.until(EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_cboTipoCalculoEspecial")))
            Select(select_tipo_calculo).select_by_value("21")
            logging.info("✅ Se seleccionó 'Por días'")

    else:

        btn_renovar = wait.until(EC.element_to_be_clickable(btn_locator))
        btn_renovar.click()
        logging.info("🖱️ Clic en Renovar")

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

        btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnRenovacionSi")))
        btn_aceptar.click()
        logging.info("🖱️ Clic en aceptar la Renovación")

        input_file = wait.until(EC.presence_of_element_located((By.ID, "fuPlanillaAjax")))
        logging.info("⌛ Página de carga lista (Trama)")

        # Caso Especial para estas pólizas
        if ramo.poliza in ['9231375', '30358377']:

            fecha_str = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniVigPoliza"))).get_attribute("value")

            fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
            fecha_final = (fecha + relativedelta(months=1) - timedelta(days=1)).strftime("%d/%m/%Y")

            input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtFinVigPoliza")))
            input_fecha.clear()
            driver.execute_script("""
                arguments[0].value = arguments[1]; 
                arguments[0].dispatchEvent(new Event('change'));
                arguments[0].dispatchEvent(new Event('blur'));
            """, input_fecha, fecha_final)

            driver.find_element(By.TAG_NAME, "body").click()

            logging.info(f"📅 Vigencia Hasta: {fecha_final}")

        else:

            fecha_inicio_web = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtIniVigPoliza"))).get_attribute("value")

            fecha_fin_web = wait.until(EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_txtFinVigPoliza"))).get_attribute("value")
            fecha_fin_web_dt = datetime.strptime(fecha_fin_web, "%d/%m/%Y")

            if fecha_ramo_fin != fecha_fin_web_dt :
                raise Exception(
                    f"Intentas renovar la póliza {ramo.poliza} del {fecha_inicio_web} al {fecha_fin_web}.\n"
                    f"Comunícate con el ejecutivo."
                )

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
        raise Exception("Posiblemente el tipo de Trabajador no tiene un formato correcto en la Trama")
    except TimeoutException:
        pass

    error_validacion = (By.ID, "divAlertaErrorValidacion")
    btn_errores = (By.ID, "btnErroresPlanilla")

    def visible(driver, locator):
        elems = driver.find_elements(*locator)
        return elems and elems[0].is_displayed()

    def proceso_finalizado(driver):

        if visible(driver, error_validacion):
            return "modal"

        if visible(driver, btn_errores):
            return "errores"

        elems = driver.find_elements(By.ID, "progressbar")
        if elems and "100%" in elems[0].text:
            return "ok"

        return False

    estado = wait.until(lambda d: proceso_finalizado(d))

    if estado == "modal":
        mensaje_error = driver.find_element(By.ID, "spanAlertaErrorValidacion").text
        raise Exception(mensaje_error)

    elif estado == "errores":
        cantidad_errores = driver.find_element(By.ID, "spnContadorError").text
        cantidad_texto = "error" if cantidad_errores == '1' else "errores"

        btn = driver.find_element(*btn_errores)
        driver.execute_script("arguments[0].click();", btn)

        wait.until(EC.visibility_of_element_located((By.ID, "divErrorPlanilla")))

        filas = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//table[@id='gridErrorPlanilla']//tr[contains(@class,'jqgrow')]")
        ))

        errores = []
        for fila in filas:
            error = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sError']").text
            nombres = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sNombres']").get_attribute("title")
            paterno = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sPaterno']").get_attribute("title")
            materno = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sMaterno']").get_attribute("title")
            nrodoc = fila.find_element(By.XPATH, ".//td[@aria-describedby='gridErrorPlanilla_sNroDoc']").get_attribute("title")

            errores.append(f"{error} - {nombres} {paterno} {materno} | {nrodoc}")

        detalle_errores = "\n".join(errores)
        raise Exception(f"Se encontró {cantidad_errores} {cantidad_texto} en la planilla\n{detalle_errores}")

    elif estado == "ok":
        logging.info("✅ Proceso completado al 100%")
        wait.until(EC.invisibility_of_element_located((By.ID, "ID_MODAL_PROCESS")))

    #----------------------------------------------------------------

    btn_calcular_locator = (By.ID, "ContentPlaceHolder1_btnCalcular")
    btn_procesar_locator = (By.ID, "ContentPlaceHolder1_btnProcesar")

    def esperar_boton_habilitado(locator):

        return wait.until(lambda d: (
            (btn := d.find_element(*locator)) and
            btn.is_displayed()
            and btn.is_enabled()
            and "ui-state-disabled" not in (btn.get_attribute("class") or "")
        ) and btn)

    def click_boton_seguro(locator):

        for _ in range(3):
            try:

                btn = esperar_boton_habilitado(locator)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});",btn)
                btn.click()
                return True

            except StaleElementReferenceException:
                time.sleep(0.5)

        btn = esperar_boton_habilitado(locator)
        driver.execute_script("arguments[0].click();", btn)
        return True

    click_boton_seguro(btn_calcular_locator)
    logging.info("🖱️ Clic en Calcular")

    click_boton_seguro(btn_procesar_locator)
    logging.info("🖱️ Clic en Procesar")
    
    texto_mensaje = ""

    if tipo_proceso == 'IN':

        btn_si = (By.ID, "ContentPlaceHolder1_btnIncluirSi")
        mensaje = (By.ID, "spanAlertaInclusion")

        resultado7 = wait.until(
            EC.any_of(
                EC.element_to_be_clickable(btn_si),
                EC.visibility_of_element_located(mensaje)
            )
        )

        elemento_id = resultado7.get_attribute("id")

        if elemento_id == "ContentPlaceHolder1_btnIncluirSi":
            resultado7.click()
            logging.info("🖱️ Clic en Aceptar la Inclusión")

        texto_mensaje = wait.until(EC.visibility_of_element_located(mensaje)).text

    else:

        mensaje = (By.ID, "spanAlertaRenovacion")
        btn_si = (By.ID, "ContentPlaceHolder1_btnAceptarRenovacion")

        wait.until(
            EC.any_of(
                EC.visibility_of_element_located(mensaje),
                EC.element_to_be_clickable(btn_si)
            )
        )

        texto_mensaje = wait.until(EC.visibility_of_element_located(mensaje)).text
        driver.find_element(*btn_si).click()
        logging.info("🖱️ Clic en Aceptar la Renovación")

    logging.info(f"📩 Texto del mensaje: {texto_mensaje}")

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

    btn_inc = (By.ID, "ContentPlaceHolder1_btnAceptarInclusion")
    #btn_ren = (By.ID, "ContentPlaceHolder1_btnAceptarRenovacion")

    if tipo_proceso == 'IN':
        btn_aceptar = wait.until(EC.element_to_be_clickable(btn_inc))
        btn_aceptar.click()
        logging.info("🖱️ Clic en Aceptar")

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "ui-widget-overlay")))

    span_numero = wait.until(EC.visibility_of_element_located((By.XPATH,f"//span[contains(text(),'{codigo_documento}')]")))
    logging.info(f"✅ Span encontrado: {span_numero.text}")

    # if len(list_polizas) == 1 and ba_codigo == '1':
    #     selector_xpath = f"//img[@data-nropolizasalud='{ramo.poliza}']"
    # elif len(list_polizas) == 1 and ba_codigo == '2':
    #     selector_xpath = f"//img[@data-nropolizapension='{ramo.poliza}']"
    # else:
    #     #selector_xpath = f"//img[@data-nropolizasalud='{list_polizas[0]}' or @data-nropolizapension='{list_polizas[1]}']"
    #     selector_xpath = f"//img[@data-nropolizasalud='{list_polizas[0]}' and @data-nropolizapension='{list_polizas[1]}']"

    #span = wait.until(EC.visibility_of_element_located(( By.XPATH, f"//span[contains(text(),'{codigo_documento}')]")))

    bloque = span_numero.find_element(By.XPATH, "./ancestor::div[1]")

    lupa = bloque.find_element(By.XPATH,f".//img[contains(@data-nropolizasalud,'{ramo.poliza}') or contains(@data-nropolizapension,'{ramo.poliza}')]")

    #lupa = span_numero.find_element(By.XPATH,f".//ancestor::tr//img[contains(@data-nropolizasalud,'{list_polizas[0]}') or contains(@data-nropolizapension,'{list_polizas[0]}')]")
    #lupa = (By.XPATH, selector_xpath)
    error_btn = (By.ID, "btnAceptarError")

    resultadol = wait.until(
        EC.any_of(
            EC.element_to_be_clickable(lupa),
            EC.element_to_be_clickable(error_btn)
        )
    )

    if resultadol.get_attribute("id") == "btnAceptarError":
        raise Exception(f"Advertencia detectada. Código: {codigo_documento}")

    for _ in range(3):
        try:
            resultadol.click()
            logging.info(f"🖱️ Clic en la lupa {codigo_documento}")
            break
        except StaleElementReferenceException:
            resultado = wait.until(EC.element_to_be_clickable(lupa))

    wait.until(EC.visibility_of_element_located((By.ID, "divPanelPDFMaster")))
    logging.info("📄 Panel PDF visible")

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
        raise Exception(f"No se descargó constancia, buscar en la compania con el código '{numero_doc}'")

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

    if len(list_polizas) == 2:

        ruta_salud = os.path.join(ruta_archivos_x_inclu, f"{list_polizas[0]}.pdf")
        ruta_pension = os.path.join(ruta_archivos_x_inclu, f"{list_polizas[1]}.pdf")

        if os.path.exists(ruta_pension):
            shutil.copy2(ruta_pension, ruta_salud)

        if tipo_mes == 'MA':
            ruta_endoso_salud = os.path.join(ruta_archivos_x_inclu, f"endoso_{list_polizas[0]}.pdf")
            ruta_endoso_pension = os.path.join(ruta_archivos_x_inclu, f"endoso_{list_polizas[1]}.pdf")

            if os.path.exists(ruta_endoso_pension):
                shutil.copy2(ruta_endoso_pension, ruta_endoso_salud)

    wait.until(EC.element_to_be_clickable((By.ID, "btnPDFCancelarM"))).click()

    logging.info(f"✅ {palabra_clave} realizada exitosamente")

def solicitud_vidaley_ov(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_mes,tipo_proceso,actividad,ramo):
 
    ov = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='OV']")))
    ov.click()
    logging.info("🖱️ Clic en 'OV'")

    time.sleep(2)

    wait.until(lambda d: len(d.window_handles) > 1)

    driver.switch_to.window(driver.window_handles[-1])

    try:
        alert = WebDriverWait(driver,5).until(EC.alert_is_present())
        logging.info(f"⚠️ Alerta presente: {alert.text}")
        alert.accept()
        logging.info("✅ Alerta aceptada")
    except:
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

    handle_ventana_principal = driver.current_window_handle

    handles_antes_ventana2 = set(driver.window_handles)

    buscar_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='dtcliBuscar']//a")))
    driver.execute_script("arguments[0].click();", buscar_link)
    logging.info("🖱️ Clic en la lupa 🔍")

    try:
        alert = wait.until(EC.alert_is_present())
        logging.info(f"⚠️ Alerta presente: {alert.text}")
        alert.accept()
        logging.info("✅ Alerta aceptada")
    except:
        logging.info("✅ No apareció ninguna alerta")

    wait.until(lambda d: len(d.window_handles) > len(handles_antes_ventana2))
    nueva_ventana2 = (set(driver.window_handles) - handles_antes_ventana2).pop()
    driver.switch_to.window(nueva_ventana2)
    logging.info("--- Ventana 2---------🌐")

    time.sleep(3)

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

    handle_segunda_ventana = driver.current_window_handle
    handles_antes_ventana3 = set(driver.window_handles)

    buscar_btn = driver.find_element(By.XPATH, "//a[contains(@href, 'buscarCliente')]")
    driver.execute_script("arguments[0].click();", buscar_btn)
    logging.info("🖱️ Se hizo clic en 'Buscar' 🔍")

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

    if not os.path.exists(ruta_trama_final_xls1):
        raise Exception (f"Archivo {ramo.poliza}.xlsx no encontrado")
    else:
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
        selectTS.select_by_value("2")
        logging.info("✅ Se seleccionó la opción 'MINERA'")
    else:
        selectTS.select_by_value("1")
        logging.info("✅ Se seleccionó la opción 'NO MINERA'")

    actividad_input = wait.until(EC.visibility_of_element_located((By.ID, "txtActividad")))
    actividad_input.clear()
    actividad_input.send_keys(actividad)
    logging.info(f"✅ Actividad registrada: {actividad}")

    driver.switch_to.default_content()

    poliza_input = wait.until(EC.visibility_of_element_located((By.ID, "tctPoliza")))

    btn_registrar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#dtRegistrar a")))
    driver.execute_script("arguments[0].scrollIntoView(true);", btn_registrar)

    poliza_input.clear()
    poliza_input.send_keys(ramo.poliza)
    logging.info(f"✅ Número de póliza ingresado: {ramo.poliza}")

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

    #observacion = f"{palabra_clave} Vida Ley \n Vigencia: {ramo.f_inicio} al {ramo.f_fin} \n  N° de Póliza: {ramo.poliza} \n Modalidad: {'Mes Vencido' if tipo_mes == 'MV' else 'Mes Adelantado'}"
    observacion = f"{palabra_clave} Vida Ley \n Vigencia: {ramo.f_inicio} al {ramo.f_fin} \n  N° de Póliza: {ramo.poliza} \n Modalidad: {'Facturación Multiple' if ramo.facturacion else 'Facturación Simple'}"

    textarea = wait.until(EC.presence_of_element_located((By.ID, "textarea1")))
    textarea.clear()
    textarea.send_keys(observacion)

    logging.info(f"✅ Observación ingresada correctamente: {observacion}")

    time.sleep(2)

    input_archivo = wait.until(EC.presence_of_element_located((By.ID, "fichero4")))
    ruta_trama_final = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")

    if not os.path.exists(ruta_trama_final):
        raise Exception (f"Archivo {ramo.poliza}.xlsx no encontrado")
    else:
        input_archivo.send_keys(ruta_trama_final)
        logging.info(f"📤 Se subió el archivo: {ramo.poliza}.xlsx")

    archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

    btn_registrar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#dtRegistrar a")))
    btn_registrar.click()
    logging.info("🖱️ Clic en 'Enviar' realizado correctamente.")

    try:

        WebDriverWait(driver,5).until(EC.alert_is_present())
        alert1 = driver.switch_to.alert
        texto_alerta = alert1.text.strip()

        if "Esta seguro de registrar la Solicitud, luego de ello no podrá ser modificada." not in texto_alerta:
            logging.info(f"⚠️ Alerta : {texto_alerta}")
            alert1.accept()
            logging.info("✅ Alerta aceptada")
            raise Exception(texto_alerta)

        logging.info(f"⚠️ Alerta : ¿{texto_alerta}?")
        alert1.accept()
        logging.info("✅ Alerta aceptada")

    except TimeoutException:

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

        if mensaje.startswith("Corregir errores encontrados en la Trama."):
            alert.accept()
            logging.info("✅ Alerta aceptada")

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

    logging.info(f"✅ Constancia obtenida para la {palabra_clave} con numero de póliza '{ramo.poliza}'")

def solicitud_vidaley_vl(driver,wait,ruta_archivos_x_inclu,ejecutivo_responsable,palabra_clave,tipo_mes,tipo_proceso,ramo):

    vidaley = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Vida Ley']")))
    vidaley.click()
    logging.info("🖱️ Clic en 'Vida Ley'")

    wait.until(lambda d: len(d.window_handles) > 1)

    driver.switch_to.window(driver.window_handles[-1])
    logging.info("🔙 Redireccionando en la nueva ventana")

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
        time.sleep(3)
    except Exception as e:
        raise str(e)
    finally:
        time.sleep(3)

    try:
        input_element = wait.until(EC.element_to_be_clickable((By.ID, "b12-b1-Input_PolicesNumber")))
        input_element.clear()
        input_element.send_keys(ramo.poliza)
        logging.info(f"✅ Numero de Póliza ingresado: {ramo.poliza}")
    except Exception as e:
        raise str(e)
    finally:
        time.sleep(3)

    # input_ruc = wait.until(EC.element_to_be_clickable((By.ID, "b12-b1-Input_RUCContratante2")))
    # input_ruc.clear()
    # input_ruc.send_keys(ruc_empresa)
    # logging.info(f"✅ Numero de RUC ingresado: {ruc_empresa}")

    # fecha_dt = datetime.strptime(ramo.f_inicio, "%d/%m/%Y")
    # primer_dia_mes = fecha_dt.replace(day=1)

    # fechai_vigencia = wait.until(EC.element_to_be_clickable((By.ID, "b12-b1-Input_RegistrationDateStart")))
    # driver.execute_script("""
    #     arguments[0].removeAttribute('readonly');
    #     arguments[0].value = arguments[1];
    #     arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
    #     arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    # """, fechai_vigencia, primer_dia_mes)

    # input_fin = driver.find_element(By.ID, "b12-b1-Input_RegistrationDateEnd")
    # fecha_fin_actual = input_fin.get_attribute("value")
    # logging.info(f"📅 Fecha de Inicio de la vigencia: {primer_dia_mes}")
    # logging.info(f"📅 Fecha Fin autocompletada: {fecha_fin_actual}")

    # driver.execute_script("document.body.click();")

    boton_buscar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Buscar']]")))
    boton_buscar.click()
    logging.info("🖱️ Clic en 'Buscar'")

    resultado1 = wait.until(
        EC.any_of(
            EC.visibility_of_element_located((By.ID, "b12-b11-TextTitlevalue")),   # error
            EC.presence_of_element_located((By.ID, "b12-limpiargrilla")),          # modal sin resultados
            EC.presence_of_element_located((By.ID, "b12-Widget_TransactionRecordList"))  # tabla
        )
    )

    if resultado1.get_attribute("id") == "b12-b11-TextTitlevalue":
        tit = driver.find_element(By.ID, "b12-b11-TextTitlevalue").text
        con = driver.find_element(By.ID, "b12-b11-TextContentValue").text
        raise Exception(f"{tit} {con}")

    if resultado1.get_attribute("id") == "b12-limpiargrilla":
        logging.error("❌ Se detectó un modal que no carga los resultados")
        raise Exception(f"No hay resultados para la póliza {ramo.poliza}")

    try:

        filas_vigentes = wait.until(
            EC.presence_of_all_elements_located((
                By.XPATH,
                "//table[@id='b12-Widget_TransactionRecordList']//tbody/tr[.//td[@data-header='Estado']//span[normalize-space()='Vigente']]"
            ))
        )

        logging.info("✅ Tabla cargada correctamente")

        if not filas_vigentes:
            raise Exception(f"No se encontró fila Vigente para la póliza {ramo.poliza}")

        fila_valida = None

        for fila in filas_vigentes:
            try:

                inicio_vigencia = fila.find_element(By.XPATH, ".//td[@data-header='Inicio de Vigencia']//span").text.strip()
                fin_vigencia = fila.find_element(By.XPATH, ".//td[@data-header='Fin de Vigencia']//span").text.strip()

                # - conversion a date time para comparar fechas sin importar el formato o posibles espacios
                fecha_ramo_fin = datetime.strptime(ramo.f_fin, "%d/%m/%Y")
                fecha_ramo_inicio = datetime.strptime(ramo.f_inicio, "%d/%m/%Y")
                fecha_fin_vigencia = datetime.strptime(fin_vigencia, "%d/%m/%Y")

                if tipo_proceso == 'IN':
                    #if ramo.f_fin == fin_vigencia:
                    if fecha_ramo_fin == fecha_fin_vigencia:
                        fila_valida = fila
                        break

                elif tipo_proceso == 'RE':
                    #if ramo.f_inicio == fin_vigencia:
                    if fecha_ramo_inicio == fecha_fin_vigencia:
                        fila_valida = fila
                        break

            except StaleElementReferenceException:
                continue 

        if not fila_valida:
            raise Exception(
                f"No se encontró una fila vigente válida para la póliza {ramo.poliza}\n"
                f"Solicitud: {ramo.f_inicio} al {ramo.f_fin}\n" 
                f"Compañía: {inicio_vigencia} al {fin_vigencia}"
            )

        indice_fila = filas_vigentes.index(fila_valida) + 1
        checkbox = fila_valida.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
        ActionChains(driver).move_to_element(checkbox).click().perform()
        logging.info(f"✅ Fila seleccionada: {indice_fila} | Póliza: {ramo.poliza} | Inicio: {inicio_vigencia} | Fin: {fin_vigencia}")

    except TimeoutException:
        raise Exception(f"Póliza {ramo.poliza} figura con estado 'No vigente'")

    time.sleep(3)

    if tipo_proceso == 'IN':
        boton_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Inclusión']]")))
    else:
        boton_proceso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//div[text()='Renovación']]")))

    boton_proceso.click()
    logging.info(f"🖱️ Clic en {palabra_clave}")

    time.sleep(3)

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

    if tipo_proceso == 'IN':

        fecha_vigencia = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-EffectiveDateTyping")))
        driver.execute_script("""
            arguments[0].removeAttribute('readonly');
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, fecha_vigencia, ramo.f_inicio)
        logging.info(f"📅 Fecha Inicio: {ramo.f_inicio}")

        driver.find_element(By.TAG_NAME, "body").click()

    else:

        for _ in range(3):
            try:

                input_fecha_incio = wait.until(EC.visibility_of_element_located((By.ID, "b13-EffectiveDateTyping")))
                input_fecha_fin = wait.until(EC.visibility_of_element_located((By.ID, "b13-EndOfValidityTyping")))

                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'})", input_fecha_fin)

                fecha_inicio= input_fecha_incio.get_attribute("value")
                fecha_fin = input_fecha_fin.get_attribute("value")

                break
            except StaleElementReferenceException:
                continue

        fecha_fin_ob = datetime.strptime(fecha_fin, "%d/%m/%Y")
        fecha_ramo_fin = datetime.strptime(ramo.f_fin, "%d/%m/%Y")

        if fecha_ramo_fin != fecha_fin_ob :
            raise Exception(
                f"Intentas renovar la póliza {ramo.poliza} del {fecha_inicio} al {fecha_fin}.\n"
                f"Comunícate con el ejecutivo."
            )

    time.sleep(5)
    ruta_archivo = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")

    if not os.path.exists(ruta_archivo):
        raise Exception(f"Archivo {ramo.poliza}.xlsx no encontrado")

    if tipo_proceso == 'IN':
        input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='b13-b1-b11-b2-b1-DropArea']//input[@type='file']")))
    else:
        input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='b13-b1-b1-b2-b1-DropArea']//input[@type='file']")))

    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'})", input_file)
    input_file.send_keys(ruta_archivo)
    logging.info(f"📂 Trama {ramo.poliza}.xlsx subida")

    resultado3 = wait.until(
        EC.any_of(
            EC.presence_of_element_located((By.XPATH,"//span[contains(text(),'La planilla no está en un formato válido de excel')]")),
            EC.element_to_be_clickable((By.ID, "b13-b1-btnValidate"))
        )
    )

    if "formato válido" in resultado3.text:
        raise Exception("La planilla no está en un formato válido de excel")

    logging.info("✅ Archivo válido")

    boton_validar = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-btnValidate")))
    boton_validar.click()
    logging.info("🖱️ Clic en Validar")

    resultado4 = wait.until(
        EC.any_of(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Aceptar']")),
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Lo sentimos')]")),
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Encontramos algunos errores en la planilla')]")),
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'La planilla superó exitosamente')]"))
        )
    )

    texto = resultado4.text.lower()
    texto = resultado4.text.lower()

    if resultado4.tag_name == "button":
        resultado4.click()
        logging.info("✅ Se aceptó advertencia de vigencia")

        resultado5 = wait.until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Lo sentimos')]")),
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Encontramos algunos errores en la planilla')]")),
                EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'La planilla superó exitosamente')]"))
            )
        )
        texto = resultado5.text.lower()

    if "lo sentimos" in texto:
        tit = driver.find_element(By.ID, "b13-b1-b5-TextTitlevalue").text
        con = driver.find_element(By.ID, "b13-b1-b5-TextContentValue").text
        raise Exception(f"{tit} \n{con}")

    elif "encontramos algunos errores" in texto:

        btn_des_obs = wait.until(EC.element_to_be_clickable((By.ID, "b13-b1-b11-b3-btnObserv")))

        archivos_antes_des_obs = set(os.listdir(ruta_archivos_x_inclu))

        driver.execute_script("arguments[0].click();", btn_des_obs)
        logging.info("🖱 Clic con JS en el botón 'Descargar Observaciones'")

        archivo_obs = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes_des_obs,".xlsx",cantidad=1)

        if archivo_obs:

            ruta_des_obs = archivo_obs[0]
            ruta_final_des_obs = os.path.join(ruta_archivos_x_inclu, f"observaciones_{ramo.poliza}.xlsx")
            os.rename(ruta_des_obs, ruta_final_des_obs)
            logging.info(f"📂 Archivo con Observaciones renombrado")

            df = pd.read_excel(ruta_final_des_obs, engine="openpyxl")

            mensajes = []

            df_obs = df[df["Observaciones"].notna()]

            for _, row in df_obs.iterrows():
                dni = row["NroDoc"]
                observacion = row["Observaciones"]
                mensajes.append(f"{dni}: {observacion}")

            mensaje_final = "\n".join(mensajes)

            raise Exception(f"Encontramos algunos errores en la planilla \n{mensaje_final}")
        else:   
            raise Exception("No se puedo descargar el archivo con las observaciones")

    elif "superó exitosamente" in texto:
        logging.info("✅ Planilla validada correctamente")

    input_correo = wait.until(EC.element_to_be_clickable((By.ID, "b13-Input_ContractingEmail")))
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'})", input_correo)
    input_correo.clear()
    input_correo.send_keys(ejecutivo_responsable)
    logging.info(f"⌨️ Ingresando correo '{ejecutivo_responsable}'")

    while True:
        resultado6 = wait.until(
            EC.any_of(
                EC.element_to_be_clickable((By.ID, "b13-CalculatePremium")),
                EC.element_to_be_clickable((By.ID, "b13-InclusionValidate")),
                EC.element_to_be_clickable((By.ID, "b13-RenewalRequest2"))
            )
        )

        boton_id = resultado6.get_attribute("id")

        if boton_id == "b13-CalculatePremium":
            driver.execute_script("arguments[0].scrollIntoView(true);", resultado6)
            resultado6.click()
            logging.info("🖱️ Clic en 'Calcular'")
            time.sleep(2) 
            continue  

        elif boton_id in ["b13-InclusionValidate", "b13-RenewalRequest2"]:
            resultado6.click()
            logging.info(f"🖱️ Clic en 'Finalizar' ({boton_id})")
            break

    time.sleep(5)

    wait.until(EC.presence_of_element_located((By.XPATH,"//span[normalize-space()='Descargar documentos']")))

    wait.until(EC.visibility_of_element_located((By.XPATH,"//span[normalize-space()='Descargar documentos']")))

    boton_descargar = wait.until(EC.element_to_be_clickable((By.XPATH,"//span[normalize-space()='Descargar documentos']/ancestor::a")))
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

        nombre_1 = os.path.join(
            ruta_archivos_x_inclu,
            f"{ramo.poliza}.pdf"
        )

        os.rename(archivos_nuevos[0], nombre_1)
        logging.info(f"✅ Constancia '{ramo.poliza}.pdf' renombrado correctamente")

        if cantidad == 2:
            nombre_2 = os.path.join(
                ruta_archivos_x_inclu,
                f"endoso_{ramo.poliza}.pdf"
            )

            os.rename(archivos_nuevos[1], nombre_2)
            logging.info(f"✅ Endoso 'endoso_{ramo.poliza}.pdf' renombrado correctamente")

    else:
        logging.error(f"❌ No se detectaron los {cantidad} archivo(s)")

    boton_aceptar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Aceptar')]")))
    boton_aceptar.click()
    logging.info("🖱️ Clic en 'Aceptar' final")

def realizar_solicitud_positiva(driver,wait,list_polizas,ba_codigo,bab_codigo,tipo_mes,ruta_archivos_x_inclu,
                                   ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_proceso,actividad,ramo):
  
    global ventana_menu_positiva
    global login_exitoso

    if not login_exitoso:

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
                login_exitoso = True
                logging.info("✅ Login exitoso")
                driver.execute_script("arguments[0].click();", autogestion)
                logging.info("🖱️ Clic en Autogestión")
                ventana_menu_positiva = driver.current_window_handle
            except TimeoutException:
                raise Exception("Problemas en el Inicio de Sesión, comunícate con el ejecutivo responsable para reprocesarlo")

        except Exception as e:
            logging.error(f"❌ Error entrando a la Positiva: {e}")
            tomar_capturar(driver, ruta_archivos_x_inclu,f"ERROR_{'SCTR' if bab_codigo != '4' else 'VIDALEY'}_LOGIN_FALLIDO")
            driver.refresh()
            wait.until(EC.presence_of_element_located((By.ID, "b5-Input_User")))
            return False,False,"Login Fallido", str(e)

    if bab_codigo in ['1', '2', '3']:
        return ejecutar_con_manejo(driver,wait,ruta_archivos_x_inclu,"SCTR",tipo_mes,lambda: solicitud_sctr(driver, wait, list_polizas, ruta_archivos_x_inclu,tipo_mes, palabra_clave, tipo_proceso, ba_codigo, ramo))
    else:
        return solicitud_vidaley_x_tipo_Mes(driver, wait, ruta_archivos_x_inclu, ruc_empresa,ejecutivo_responsable, palabra_clave, tipo_proceso,actividad, ramo, tipo_mes)

def solicitud_vidaley_x_tipo_Mes(driver, wait, ruta_archivos_x_inclu, ruc_empresa,ejecutivo_responsable, palabra_clave,
                                tipo_proceso,actividad, ramo, tipo_mes):

    tipo_vl = detectar_tipo_mes(ruta_archivos_x_inclu,ramo)

    if not tipo_vl:
        tomar_capturar(driver,ruta_archivos_x_inclu,f"ERROR_VIDALEY_{tipo_mes}")
        logging.error(f"❌ Error en La Positiva (Vida Ley) - {tipo_mes}: No se pudo determinar el tipo (OV/VL)")
        return False, False, f"LAPO-VIDALEY-{tipo_mes}", "No se pudo determinar el tipo (OV/VL)"

    logging.info(f"📂 Tipo detectado: {tipo_vl}")

    funciones = {
        "OV": lambda: solicitud_vidaley_ov(
            driver, wait, ruta_archivos_x_inclu, ruc_empresa,
            ejecutivo_responsable, palabra_clave,tipo_mes,
            tipo_proceso, actividad, ramo
        ),
        "VL": lambda: solicitud_vidaley_vl(
            driver, wait, ruta_archivos_x_inclu,
            ejecutivo_responsable, palabra_clave,tipo_mes,
            tipo_proceso, ramo
        )
    }

    funcion = funciones.get(tipo_vl)

    if not funcion:
        tomar_capturar(driver,ruta_archivos_x_inclu,f"ERROR_VIDALEY_{tipo_mes}")
        logging.error(f"❌ Error en La Positiva (Vida Ley) - {tipo_mes}: Tipo de Vida Ley inválido")
        return False, False, f"LAPO-VIDALEY-{tipo_mes}", "Tipo de Vida Ley inválido"

    return ejecutar_con_manejo(driver,wait,ruta_archivos_x_inclu,"VIDALEY",tipo_mes,funcion)

def ejecutar_con_manejo(driver,wait,ruta_archivos_x_inclu,tipo,tipo_mes,funcion):

    try:
        funcion()
        flag_extra = True if tipo_mes == "MA" else False
        return True, flag_extra, "", ""
    except Exception as e:
        resultado, asunto = validar_pagina(driver)
        tomar_capturar(driver,ruta_archivos_x_inclu,f"ERROR_{tipo}_{tipo_mes}")
        detalle = (f"{asunto}, intentar entre 5 a 10 minutos de nuevo"if not resultado else str(e))
        logging.error(f"❌ Error en La Positiva ({tipo}) - {tipo_mes}: {detalle}")
        return False, False, f"LAPO-{tipo}-{tipo_mes}", detalle
    finally:
        driver.close()
        logging.info(f"✅ Cerrando la pestaña {tipo}")
        driver.switch_to.window(driver.window_handles[0])
        logging.info("🔙 Retornando al menú principal")

def detectar_tipo_mes(ruta_archivos,ramo):

    archivos = os.listdir(ruta_archivos)

    poliza = str(ramo.poliza)

    archivos_poliza = [f for f in archivos if f.startswith(poliza)]

    archivos_xlsx = [f for f in archivos_poliza if f.endswith(".xlsx")]
    archivos_xls = [f for f in archivos_poliza if f.endswith(".xls")]

    logging.info(f"📂 Archivos '.xlsx' encontrados: {len(archivos_xlsx)} y '.xls' encontrados: {len(archivos_xls)}")

    if len(archivos_xlsx) == 1 and len(archivos_xls) == 0:
        return "VL"
    elif len(archivos_xls) == 1 and len(archivos_xlsx) == 1:
        return "OV"
    else:
        return None
