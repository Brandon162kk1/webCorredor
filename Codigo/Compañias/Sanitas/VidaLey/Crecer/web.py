#   --- Froms ----
from dataclasses import asdict
from turtle import clear
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import (
    TimeoutException,
    NoAlertPresentException,
    StaleElementReferenceException,
    WebDriverException,
    UnexpectedAlertPresentException
)
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from Tiempo.fechas_horas import get_timestamp,get_fecha_menos_x_dias
from LinuxDebian.Ventana.ventana import desbloquear_interaccion,bloquear_interaccion,esperar_archivos_nuevos
from LinuxDebian.OtrosMetodos.metodos import subir_trama
from Compañias.Sanitas.metodos import imagen_a_pdf
from Chrome.google import tomar_capturar
from Apis.Put.web_corredor import enviar_aviso_captcha
from Apis.Post.webhook import enviar_msj_wsp
#   --- Imports ----
import time
import os
import logging
import zipfile
import re

url_n8n_base = os.getenv("url_host_prod")
puerto = os.getenv("puerto")
entorno = os.getenv("entorno","false").strip().lower() == "true"

url_crecer = f"{url_n8n_base}:{puerto}"

def login_completado(driver):

    if driver.find_elements(By.XPATH, "//a[contains(normalize-space(),'Cerrar sesión')]"):
        return True

    if "login" not in driver.current_url.lower():
        return True

    if not driver.find_elements(By.ID, "suser"):
        return True

    return False

def login_crecer(driver,wait,tipo_proceso,ruta_archivos_x_inclu,ejecutivo_responsable,
                    palabra_clave,ruc_empresa,tipo_mes,ramo):

    if tipo_proceso in ('NE','IN','RE') and entorno:
        msj = f"Agente no disponible para realizar {palabra_clave.lower()} en esta compañía"
        return False,False,f"CREC-VL-{tipo_mes}",msj

    constancia = False
    factura = False
    error = False
    tipoError = ""
    detalleError = ""

    try:

        driver.get(os.getenv("login_url_crecer_vida_ley"))
        logging.info("--- Ingresando a Crecer Seguros🌐 ---")

        user_input = wait.until(EC.presence_of_element_located((By.ID, "suser")))
        user_input.clear()
        user_input.send_keys(ramo.usuario)
        logging.info("⌨️ Digitando el Username")
        
        pass_input = wait.until(EC.presence_of_element_located((By.ID, "spassword")))
        pass_input.clear()
        pass_input.send_keys(ramo.clave)
        logging.info(f"⌨️ Digitando el Password")

        if entorno:

            # Envio de aviso de captcha al API para que el usuario resuelva el captcha manualmente
            enviar_aviso_captcha(ramo)

            mensaje = f"Ingresar al siguiente enlace {url_crecer} para resolver el captcha manualmente de la {palabra_clave}."
            enviar_msj_wsp(ramo,"notificacion",mensaje)

            desbloquear_interaccion()

        # Tiempo estimado para resolver captcha 5 minutos
        wait_humano = WebDriverWait(driver, 300, poll_frequency=1)

        resultado = wait_humano.until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//a[contains(normalize-space(),'Cerrar sesión')]")),
                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Debe crear una contraseña nueva.')]")),
                EC.visibility_of_element_located((By.CLASS_NAME, "modal-content"))
            )
        )

        clase = resultado.get_attribute("class") or ""
        texto = resultado.text or ""

        if "Debe crear una contraseña nueva" in texto:

            logging.warning("⚠️ Alerta de contraseña temporal detectada")
            raise Exception("El usuario tiene contraseña temporal, es necesario cambiarla antes de continuar  o comunícate con tu administrador")

        elif "modal-content" in clase:

            mensaje = resultado.find_element(By.CLASS_NAME, "security-body").text
            logging.info(f"📩 Mensaje del modal: {mensaje}")
            boton_aceptar = resultado.find_element(By.XPATH, ".//button[contains(text(),'Aceptar')]")
            boton_aceptar.click()
            logging.info("🖱️ Clic en 'Aceptar'")
            wait_humano.until(EC.presence_of_element_located((By.XPATH, "//a[contains(normalize-space(),'Cerrar sesión')]")))
            logging.info("✅ Login exitoso")

        else:
            logging.info("✅ Login exitoso")

        if entorno :
            bloquear_interaccion()

        if tipo_proceso == 'IN':
            inclusion_crecer(driver,wait,ruta_archivos_x_inclu,ramo)
        elif tipo_proceso == 'RE':
            renovacion_crecer(driver,wait,ruta_archivos_x_inclu,ejecutivo_responsable,ramo)
        else:
            buscar_cuota_y_descargar_constancia(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ramo)

    except Exception as e :
        error = True
        detalleError = str(e)
    except WebDriverException as e :
        error = True
        logging.exception(f"⚠️ Error técnico de Selenium | {e}")
        detalleError = "Problemas Técnicos del Agente"
    finally:

        if error:
            try:
                tomar_capturar(driver,ruta_archivos_x_inclu,f"ERROR_VL_{tipo_mes}")
            except Exception:
                logging.exception("⚠️ No se pudo tomar screenshot")

            logging.error(f"❌ Error en Crecer (VL) - {tipo_mes}: {detalleError}")
            return constancia,factura,f"CREC-VL-{tipo_mes}",detalleError
        else:
            return constancia,factura,tipoError,detalleError

def inclusion_crecer(driver,wait,ruta_archivos_x_inclu,ramo):

    span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Gestión de Endoso']")))
    logging.info("🔍 Ubicándose en 'Gestión de Endoso'")
    time.sleep(3)

    actions = ActionChains(driver)
    actions.move_to_element(span_element).click().perform()

    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Generar Endoso Inclusión / Exclusión")))
    link.click()
    logging.info("🖱️ Clic en 'Generar Endoso Inclusión / Exclusión'")

    time.sleep(3)

    input_fullname = wait.until(EC.element_to_be_clickable((By.ID, "spolizanumber")))
    input_fullname.clear()
    input_fullname.send_keys(ramo.poliza)
    logging.info(f"✅ Se ingresó la póliza: {ramo.poliza}")
        
    select_ramos = wait.until(EC.element_to_be_clickable((By.ID, "Ramos")))
    select = Select(select_ramos)
    select.select_by_value("2")
    logging.info("🖱️ Clic en la opción Vida Ley")

    select_accion = wait.until(EC.element_to_be_clickable((By.ID, "Action")))
    select2= Select(select_accion)
    select2.select_by_value("1")
    logging.info("🖱️ Clic en la opción Inclusión")

    #time.sleep(2)

    boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Buscar Póliza')]")))
    boton.click()
    logging.info("🖱️ Clic en 'Buscar Póliza")

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "overlay")))

    # validar que salga modal para mostrar error


    #Esperar que cargue la tabla con mas de 0 filas y esten visibles
    wait.until(EC.visibility_of_element_located((By.XPATH,"//table[contains(@class, 'table-striped') and contains(@class, 'table-bordered')]//tbody/tr")))

    try:
        campo_fecha = wait.until(EC.element_to_be_clickable((By.ID, "dFechaInicio1")))
        # Hacer clic para abrir el calendario
        actions = ActionChains(driver)
        actions.move_to_element(campo_fecha).click().perform()

        fecha_inicio = ramo.f_inicio
        fecha_actual = fecha_inicio.strftime('%d/%m/%Y')
        campo_fecha.send_keys(fecha_actual)
        logging.info("✅ Se ingresó la fecha actual")
    except WebDriverException as e:
        logging.error(f"❌ Ocurrió un error tecnico al ingresar la fecha: {e}")
    except Exception as e:
        logging.error(f"❌ Ocurrió un error al ingresar la fecha: {e}")
    finally:
        pass

    #time.sleep(3)

    input_file = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Cargar excel')]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", input_file)

    ruta_trama_xlsx_sanitas_vl = f"{ruta_archivos_x_inclu}/{ramo.poliza}.xlsx"

    if not os.path.exists(ruta_trama_xlsx_sanitas_vl):
        raise FileNotFoundError(f"El Archivo {ramo.poliza}.xlsx no existe")

    if subir_trama(input_file,ruta_trama_xlsx_sanitas_vl):               
        time.sleep(3)
        logging.info(f"✅ Trama {ramo.poliza}.xlsx adjuntada")
    else:
        raise Exception(f"No se pudo subir la trama {ramo.poliza}.xlsx")

    try:
        # Modal de Endoso Retroactivo
        boton_ok = WebDriverWait(driver,8).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'OK')]")))
        # boton_ok.click()
        # logging.error(f"🖱️Clic en Ok")
        raise  Exception("Mensaje de Endoso Retroactivo")
    except :
        pass

    # Encuentra el botón por su texto
    boton_endoso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Generar Endoso')]")))
    driver.execute_script("arguments[0].scrollIntoView(true);", boton_endoso)
    boton_endoso.click()
    logging.info("🖱️ Clic en 'Generar Endoso'")

    try:

        modal_advertencia = (By.ID, "swal2-content")

        resultado = wait.until(
            EC.any_of(
                EC.visibility_of_element_located(modal_advertencia),
                EC.staleness_of(boton_ok)
            )
        )

        if isinstance(resultado, WebElement):
            # Modal de confirmacion de generación de endoso
            texto_modal = resultado.text
            logging.warning(f"⚠️ Mensaje de advertencia: {texto_modal}")
            raise Exception (texto_modal)
        else:
            pass

        wait.until(EC.staleness_of(boton_ok))
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        input("Esperar")

        btn_pagar_endoso = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Pagar Endoso')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_pagar_endoso)
        btn_pagar_endoso.click()
        logging.info("🖱️ Clic en 'Pagar Endoso'")

        # Enviar por wsp al ejecutivo responsable para decidir que metodo de pago usar
        #enviar_msj_wsp(ramo,"notificacion",None,palabra_clave,nombre_cliente)

        input("Esperar")

        pago_efectivo_img = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[contains(@src,'pago_efectivo')]")))
        pago_efectivo_img.click()
        logging.info("🖱️ Clic en 'Pago Efectivo'")

        contenido = wait.until(EC.visibility_of_element_located((By.ID, "swal2-content")))
        texto_cip = contenido.text
        logging.info(f"📩 Mensaje CIP: {texto_cip}")

        btn_ok = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
        driver.execute_script("arguments[0].click();", btn_ok)
        logging.info("🖱️ Clic en 'Ok'")

        wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-popup")))

        # Esperar que el iframe esté disponible y cambiar contexto
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@src,'pagoefectivo.pe')]")))
        logging.info("✅ Ya estoy dentro del iframe")

        # Esperar que aparezca el texto indicador
        wait.until(EC.visibility_of_element_located((By.XPATH, "//p[contains(text(),'Código de pago')]")))

        # Ahora obtener el número (solo números largos)
        cip_elemento = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[normalize-space()[string-length()>=8 and number(.)=number(.)]]")))

        codigo_cip = cip_elemento.text
        logging.info(f"💳 CIP obtenido: {codigo_cip}")

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",cip_elemento)
        logging.info("📌 Scroll hasta el codigo para tomar captura")

        cip_elemento.screenshot(os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.png"))

        imagen_a_pdf(os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.png"), os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.pdf"))

        driver.save_screenshot(os.path.join(ruta_archivos_x_inclu,"cod_cip.png"))

        imagen_a_pdf(os.path.join(ruta_archivos_x_inclu,"cod_cip.png"),os.path.join(ruta_archivos_x_inclu,f"endoso_{ramo.poliza}.pdf"))

        # Cuando termines
        driver.switch_to.default_content()
        logging.info("🔙 Saliendo del iframe")

    except Exception as e:
        logging.error(f"❌ Error durante el proceso de inclusión: {e}")
    except WebDriverException as e :
        logging.exception(f"❌ Ocurrió un error tecnico: {e}")
    finally:
        logging.info("📌 Aca nos quedamos para validar")
        input("Esperar")

def renovacion_crecer(driver,wait,ruta_archivos_x_inclu,correo,ramo):
    
    span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Gestión de Cotización']")))
    logging.info("🔍 Ubicándose en 'Gestión de Cotización'")

    actions = ActionChains(driver)
    actions.move_to_element(span_element).click().perform()
    logging.info("🖱️ Clic en 'Gestión de Cotización'")

    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Renovación Vida Ley")))
    link.click()
    logging.info("🖱️ Clic en 'Renovación Vida Ley'")
  
    input_fullname = wait.until(EC.element_to_be_clickable((By.ID, "spolizanumber")))
    input_fullname.clear()
    input_fullname.send_keys(ramo.poliza)
    logging.info(f"✅ Se ingresó la póliza: {ramo.poliza}")

    boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Busca Póliza')]")))
    boton.click()
    logging.info("🖱️ Clic en 'Busca Póliza")

    resultado = wait.until(
        EC.any_of(
            EC.visibility_of_element_located((By.ID, "swal2-content")),
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Cargar plantilla')]")
            )
        )
    )

    if resultado.get_attribute("id") == "swal2-content":
        texto_alerta_poliza = resultado.text
        raise Exception(texto_alerta_poliza)

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "overlay")))

    boton_subir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Cargar plantilla')]")))

    ruta_trama_crecer_vl_xlsx = f"{ruta_archivos_x_inclu}/{ramo.poliza}.xlsx"

    if not os.path.exists(ruta_trama_crecer_vl_xlsx):
        raise FileNotFoundError(f"La ruta {ruta_trama_crecer_vl_xlsx} no existe")

    if subir_trama(boton_subir,ruta_trama_crecer_vl_xlsx):               
        time.sleep(3)
        logging.info(f"✅ Trama {ramo.poliza}.xlsx adjuntada")
    else:
        raise Exception(f"No se pudo subir la trama {ramo.poliza}.xlsx")

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "overlay")))

    resultado1 = wait.until(
        EC.any_of(
            EC.visibility_of_element_located((By.ID, "swal2-content")),
            EC.element_to_be_clickable(
                (By.XPATH, '//input[@value="Validar"]')
            )
        )
    )

    if resultado1.get_attribute("id") == "swal2-content":
        texto_error_poliza = resultado1.text
        tomar_capturar(driver,ruta_archivos_x_inclu,f"errorTrama_{ramo.poliza}")
        raise Exception(texto_error_poliza)

    boton_validar = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@value="Validar"]')))
    driver.execute_script("arguments[0].scrollIntoView(true);", boton_validar)
    boton_validar.click()
    logging.info("🖱️ Clic en 'Validar'")

    # Validar que salga modal de error de Renovacióno el boton ok presente

    # <div class="swal2-container swal2-center swal2-shown" style="overflow-y: auto;"><div aria-labelledby="swal2-title" aria-describedby="swal2-content" class="swal2-popup swal2-modal swal2-show" tabindex="-1" role="dialog" aria-live="assertive" aria-modal="true" style="display: flex;"><div class="swal2-header"><ul class="swal2-progress-steps" style="display: none;"></ul><div class="swal2-icon swal2-error" style="display: none;"><span class="swal2-x-mark"><span class="swal2-x-mark-line-left"></span><span class="swal2-x-mark-line-right"></span></span></div><div class="swal2-icon swal2-question" style="display: none;"></div><div class="swal2-icon swal2-warning" style="display: none;"></div><div class="swal2-icon swal2-info swal2-animate-info-icon" style="display: flex;"></div><div class="swal2-icon swal2-success" style="display: none;"><div class="swal2-success-circular-line-left" style="background-color: rgb(255, 255, 255);"></div><span class="swal2-success-line-tip"></span> <span class="swal2-success-line-long"></span><div class="swal2-success-ring"></div> <div class="swal2-success-fix" style="background-color: rgb(255, 255, 255);"></div><div class="swal2-success-circular-line-right" style="background-color: rgb(255, 255, 255);"></div></div><img class="swal2-image" style="display: none;"><h2 class="swal2-title" id="swal2-title" style="display: none;"></h2><button type="button" class="swal2-close" aria-label="Close this dialog" style="display: none;">×</button></div><div class="swal2-content"><div id="swal2-content" style="display: block;">Por regla, no se puede generar un endoso retroactivo de más de 45 días.</div><input class="swal2-input" style="display: none;"><input type="file" class="swal2-file" style="display: none;"><div class="swal2-range" style="display: none;"><input type="range"><output></output></div><select class="swal2-select" style="display: none;"></select><div class="swal2-radio" style="display: none;"></div><label for="swal2-checkbox" class="swal2-checkbox" style="display: none;"><input type="checkbox"><span class="swal2-label"></span></label><textarea class="swal2-textarea" style="display: none;"></textarea><div class="swal2-validation-message" id="swal2-validation-message"></div></div><div class="swal2-actions"><button type="button" class="swal2-confirm swal2-styled" aria-label="" style="display: inline-block; border-left-color: rgb(48, 133, 214); border-right-color: rgb(48, 133, 214);">OK</button><button type="button" class="swal2-cancel swal2-styled" aria-label="" style="display: none;">Cancel</button></div><div class="swal2-footer" style="display: none;"></div></div></div>

    resultado2 = wait.until(
        EC.any_of(
            EC.visibility_of_element_located((By.ID, "swal2-content")),
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class,'swal2-confirm') and normalize-space()='OK']")
            )
        )
    )

    if resultado2.get_attribute("id") == "swal2-content":
        texto_error_modal = resultado2.text
        raise Exception(texto_error_modal)

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

        btn_aceptar = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='Aceptar']")))
        btn_aceptar.click()
        logging.info("🖱️ Clic en 'Aceptar'")
  
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

    except WebDriverException as e:
        logging.error(f"❌ Ocurrió un error tecnico: {e}")
    except Exception as e:
        logging.error(f"❌ Ocurrió un error: {e}")
    finally:
        logging.info("📌 Aca nos quedamos para validar")
        input("Esperar")

    try:

        email_input = WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.ID, "semailpolicy")))
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

    # Dejar interaccion humana habilitada para que el ejecutivo decida si es codigo de pago , por banco ,etc
    try:
        # Esperar que el iframe esté disponible y cambiar contexto
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@src,'pagoefectivo.pe')]")))
        logging.info("✅ Ya estoy dentro del iframe")

        # Esperar que aparezca el texto indicador
        wait.until(EC.visibility_of_element_located((By.XPATH, "//p[contains(text(),'Código de pago')]")))

        # Ahora obtener el número (solo números largos)
        cip_elemento = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[normalize-space()[string-length()>=8 and number(.)=number(.)]]")))

        codigo_cip = cip_elemento.text
        logging.info(f"✅ CIP obtenido: {codigo_cip}")

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",cip_elemento)
        logging.info("✅ Scroll hasta el codigo para tomar captura")

        cip_elemento.screenshot(os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.png"))
        logging.info("✅ Captura solo del cip")

        imagen_a_pdf(os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.png"), os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.pdf"))
        logging.info("✅ CIP en pdf")

        tomar_capturar(driver, ruta_archivos_x_inclu, "codigo_de_pago")
        logging.info("✅ Captura toda la pantalla")

        imagen_a_pdf(os.path.join(ruta_archivos_x_inclu,"codigo_de_pago"),os.path.join(ruta_archivos_x_inclu,f"endoso_{ramo.poliza}.pdf"))
        logging.info("✅ Captura en pdf")
    except WebDriverException as e:
        logging.error(f"❌ Ocurrió un error tecnico: {e}")
    except Exception as e:
        logging.error(f"❌ Ocurrió un error: {e}")
    finally:
        logging.info("📌 Aca nos quedamos para validar")
        input("Esperar")

        #Clic en Pago Efectivo me falta
        #------------------------------------

        # contenido = wait.until(EC.visibility_of_element_located((By.ID, "swal2-content")))
        # texto_cip = contenido.text
        # logging.info(f"Mensaje CIP: {texto_cip}")

        # btn_ok = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
        # driver.execute_script("arguments[0].click();", btn_ok)
        # logging.info("🖱️ Clic en 'Ok'")

        # wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "swal2-popup")))

        # # 1️⃣ Esperar que el iframe esté disponible y cambiar contexto
        # wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@src,'pagoefectivo.pe')]")))
        # logging.info("✅ Ya estoy dentro del iframe")

        # # Esperar que aparezca el texto indicador
        # wait.until(EC.visibility_of_element_located((By.XPATH, "//p[contains(text(),'Código de pago')]")))

        # # Ahora obtener el número (solo números largos)
        # cip_elemento = wait.until(EC.visibility_of_element_located((By.XPATH, "//span[normalize-space()[string-length()>=8 and number(.)=number(.)]]")))

        # codigo_cip = cip_elemento.text
        # logging.info(f"✅ CIP obtenido: {codigo_cip}")

        # driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",cip_elemento)
        # logging.info("Scroll hasta el codigo para tomar captura")

        # cip_elemento.screenshot(os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.png"))

        # imagen_a_pdf(os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.png"), os.path.join(ruta_archivos_x_inclu,"solo_cod_cip.pdf"))

        # tomar_capturar(driver, ruta_archivos_x_inclu, "codigo_de_pago")

        # imagen_a_pdf(os.path.join(ruta_archivos_x_inclu,"codigo_de_pago"),os.path.join(ruta_archivos_x_inclu,f"endoso_{ramo.poliza}.pdf"))

        # # Cuando termines
        # driver.switch_to.default_content()
        # logging.info("Saliendo del iframe")

def buscar_cuota_y_descargar_constancia(driver,wait,ruta_archivos_x_inclu,ruc_empresa,ramo,cod_pago='346593644'):

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
