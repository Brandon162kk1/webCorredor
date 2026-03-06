#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from LinuxDebian.Ventana.ventana import bloquear_interaccion,desbloquear_interaccion,esperar_archivos_nuevos
from Compañias.Positiva.metodos import validar_pagina
from Chrome.google import tomar_capturar
#---- Import ---
import os
import logging
import time
import shutil

def cambiar_a_nueva_ventana(wait, driver, handles_antes, timeout=10):
    wait.until(lambda d: len(set(d.window_handles) - handles_antes) == 1)
    nueva = (set(driver.window_handles) - handles_antes).pop()
    driver.switch_to.window(nueva)
    return nueva

def volver_a_ventana(wait, driver, handle_objetivo, total_handles_esperados, timeout=10):
    wait.until(lambda d: len(d.window_handles) == total_handles_esperados)
    driver.switch_to.window(handle_objetivo)

def reentrar_frames(wait, driver):
    driver.switch_to.default_content()

    WebDriverWait(driver, 15).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    WebDriverWait(driver, 15).until(
        lambda d: len(d.find_elements(By.TAG_NAME, "frame")) > 0
    )

    wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Principal")))
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Contenido")))

def seleccionar_radio(id_radio,opcion,wait,driver):
    radio = wait.until(EC.element_to_be_clickable((By.ID, id_radio)))
    driver.execute_script("arguments[0].click();", radio)

    # esperar postback → el radio vuelva a existir y esté marcado
    wait.until(lambda d: d.find_element(By.ID, id_radio).is_selected())

    logging.info(f"✅ Radio {opcion} seleccionado")

def detectar_recaptcha(driver):

    try:
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.XPATH, "//iframe[contains(@src,'recaptcha/api2/bframe')]")
            )
        )
        return True
    except TimeoutException:
        return False

def cerrar_popup(driver, ventana_principal, wait):
    try:
        wait.until(lambda d: len(d.window_handles) > 1)

        for ventana in driver.window_handles:
            if ventana != ventana_principal:
                driver.switch_to.window(ventana)
                driver.close()

        driver.switch_to.window(ventana_principal)

    except TimeoutException:
        pass

def realizar_solicitud_pacifico(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,
                          palabra_clave,ruc_empresa,ejecutivo_responsable,bab_codigo,ramo):
    
    tipoError = ""
    detalleError = ""
    constancia = False
    proforma = False
  
    try:

        driver.get("https://www0.pacificoseguros.com/loginPacifico/login.aspx") 
        logging.info("⌛ Cargando la Web de Pacifico en Linea")
           
        resultado,asunto = validar_pagina(driver)

        if not resultado:
            raise Exception (f"{asunto}")

        user_input = wait.until(EC.element_to_be_clickable((By.ID, "txtNumero")))
        user_input.clear()
        user_input.send_keys(ramo.usuario)
        logging.info("⌨️ Digitando el Usuario")
    
        pass_input = wait.until(EC.element_to_be_clickable((By.ID, "txtClave")))
        pass_input.clear()
        pass_input.send_keys(ramo.clave)
        logging.info("⌨️ Digitando el Password")

        buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "imbIngresar2")))
        driver.execute_script("arguments[0].click();", buscar_btn)
        logging.info(f"🖱️ Clic en 'Ingresar'")

        try:
            mensaje = WebDriverWait(driver,7).until(EC.visibility_of_element_located((By.ID, "lblMensaje")))
            texto = mensaje.text.strip()
            if "Contraseña incorrecta" in texto:
                raise Exception(texto)

        except TimeoutException:
            pass

        if detectar_recaptcha(driver):

            logging.warning("🚫 Flujo detenido: reCAPTCHA activo")
            desbloquear_interaccion()

            try:
                wait_humano = WebDriverWait(driver, 600)  # 10 minutos máx
                wait_humano.until(EC.presence_of_element_located((By.ID, "Label1")))
                logging.info("✅ reCAPTCHA resuelto, Label1 detectado")

            except TimeoutException:
                logging.error("⏰ Timeout: reCAPTCHA no fue resuelto a tiempo")
                raise Exception("🧩 Captcha no resuelto")

            finally:
                bloquear_interaccion()

        #id_ramo = "imgbtn_EPS" if ba_codigo != '4' else "imgbtn_VIDA"
        seguro_vida = wait.until(EC.element_to_be_clickable((By.ID, "imgbtn_EPS")))
        driver.execute_script("arguments[0].click();", seguro_vida)

        logging.info("✅ Login exitoso detectado ")
        logging.info("🖱️ Clic en 'Seguros de Salud'")

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Menu")))

        # Expandir carpeta
        folder = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='spn_5']//img")))
        driver.execute_script("arguments[0].click();", folder)
        time.sleep(0.5)

        link = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='spn_5']//a[contains(.,'Renovación')]")))
        driver.execute_script("arguments[0].click();", link)
        logging.info(f"🖱️ Clic en {palabra_clave}")

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Principal")))

        #logging.info("🔀 Cambiado al frame 'Principal'")

        ruc_input = wait.until(EC.presence_of_element_located((By.ID, "txtNumDoc")))
        ruc_input.clear()
        ruc_input.send_keys(ruc_empresa)
        logging.info(f"✅ Se ingresó el RUC '{ruc_empresa}'")

        buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnBusqueda")))
        buscar_btn.click()
        logging.info("🖱️ Clic en Buscar")

        buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "dgClientes_hlLink_0")))
        driver.execute_script("arguments[0].click();", buscar_btn)
        logging.info(f"🖱️ Clic en el cliente")

        try:
            WebDriverWait(driver,5).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            logging.info(f"⚠️ Texto de la alerta: {alert.text}")
            alert.accept()
            logging.info("✅ Alerta Aceptada")

            driver.switch_to.default_content()

        except TimeoutException:
            pass

        wait.until(EC.presence_of_element_located((By.NAME, "Menu")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Menu")))

        folder = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='spn_5']//img")))

        if "folder.gif" in folder.get_attribute("src").lower():
            driver.execute_script("arguments[0].click();", folder)
            time.sleep(0.5)

        link = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='spn_5']//a[contains(.,'Renovación')]")))

        driver.execute_script("arguments[0].click();", link)
        logging.info(f"🖱️ Clic en {palabra_clave} nuevamente")

        #input("Esperar")
        driver.switch_to.default_content()

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Principal")))
        #logging.info("🔀 En frame Principal (Renovación cargada)")

        #frames = driver.find_elements(By.TAG_NAME, "frame")
        #logging.info(f"🧩 Frames detectados: {[f.get_attribute('name') for f in frames]}")

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Tab")))
        #logging.info("🔀 Cambiado al frame 'Tab'")

        id_li = "rptTab_hlkTab_0" if bab_codigo != '4' else "rptTab_hlkTab_1"

        li_tab = wait.until(EC.element_to_be_clickable((By.ID, id_li)))
        driver.execute_script("arguments[0].click();", li_tab)

        logging.info(f"🖱️ Clic en pestaña {'SCTR' if bab_codigo != '4' else 'Vida Ley'}")
    
        driver.switch_to.default_content()

        #frames2 = driver.find_elements(By.TAG_NAME, "frame")
        #logging.info(f"🧩 Frames detectados: {[f.get_attribute('name') for f in frames2]}")

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Principal")))
        #logging.info("🔀 Cambiado al frame 'Principal'")

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Contenido")))
        #logging.info("🔀 Entré a frame Contenido")

        #--------
        ventana_principal = driver.current_window_handle
        handles_iniciales = set(driver.window_handles)

        buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "imgBuscarCliente")))
        driver.execute_script("arguments[0].click();", buscar_btn)
        logging.info("🖱️ Clic en la Lupa del cliente")

        ventana_2 = cambiar_a_nueva_ventana(wait, driver, handles_iniciales)
        logging.info("--- Ventana 2---------🌐")

        id_ruc_input_2 = "txtDocumento" if bab_codigo != '4' else "txtNroDocumento"

        ruc_input2 = wait.until(EC.presence_of_element_located((By.ID, id_ruc_input_2)))
        ruc_input2.clear()
        ruc_input2.send_keys(ruc_empresa)
        logging.info(f"⌨️ Digitando RUC -> {ruc_empresa}")

        buscar_btn2 = wait.until(EC.element_to_be_clickable((By.ID, "btnBuscar" if bab_codigo != '4' else "btnBuscarCliente")))
        buscar_btn2.click()
        logging.info("🖱️ Clic en Buscar")

        cliente = wait.until(EC.element_to_be_clickable((By.XPATH,"(//a[contains(@onclick,'fr_DevolverCodigo')]|//td[contains(@onclick,'seleccionarCliente')])[1]")))
        cliente.click()
        logging.info("🖱️ Cliente seleccionado")

        volver_a_ventana(wait, driver, ventana_principal, len(handles_iniciales))
        logging.info("🔙 Regresando a la ventana principal")

        try:
            WebDriverWait(driver,20).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            texto_alerta = alert.text   # 👈 guardar antes
            alert.accept()
            logging.info(f"⚠️ Texto de la alerta: {texto_alerta}")
            logging.info("✅ Alerta Aceptada")
            raise Exception(texto_alerta)
        except TimeoutException:
            pass

        # 🔑 REENTRAR A LOS FRAMES
        driver.switch_to.default_content() #Sale de TODOS los frames / iframes
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Principal")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Contenido")))

        if bab_codigo == '4':
            handles_antes = set(driver.window_handles)

            lupa_poliza = wait.until(EC.element_to_be_clickable((By.ID, "imgBuscarPoliza")))
            driver.execute_script("arguments[0].click();", lupa_poliza)
            logging.info("🖱️ Clic en la Lupa de la Póliza")

            ventana_3 = cambiar_a_nueva_ventana(wait, driver, handles_antes)
            logging.info("--- Ventana 3---------🌐")

            poliza = wait.until(EC.element_to_be_clickable((By.ID, "grvPoliza_ctl03_NroPoliza")))
            poliza.click()
            logging.info("🖱️ Póliza seleccionada")

            volver_a_ventana(wait, driver, ventana_principal, len(handles_iniciales))
            logging.info("🔙 Regresando a la ventana principal")

            reentrar_frames(wait, driver)

        if bab_codigo == '2':  # Pensión
            seleccionar_radio("rblProducto_1","Pension",wait,driver)

        elif bab_codigo == '3':  # Ambos
            seleccionar_radio("rblProducto_2","Ambos",wait,driver)

        if tipo_proceso == 'IN':
            id_mov = "rdbTipoMov_1" if bab_codigo != '4' else "rblTipoMovimiento_1" # SCTR - VL
            seleccionar_radio(id_mov,"Inclusión",wait,driver)
        else:
            id_mov = "rdbTipoMov_0" if bab_codigo != '4' else "rblTipoMovimiento_0" # SCTR - VL
            seleccionar_radio(id_mov,"Renovación",wait,driver)

            try:
                WebDriverWait(driver,10).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                texto_alerta = alert.text   # 👈 guardar antes

                if texto_alerta.startswith("VREN02: No se puede renovar este periodo."):
                    alert.accept()
                    logging.info(f"⚠️ Texto de la alerta: {texto_alerta}")
                    logging.info("✅ Alerta Aceptada")
                    raise Exception(texto_alerta)

            except TimeoutException:
                pass

        if bab_codigo != '4':
            ruta_archivo = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}_97.xls")
            msj_tra = f"✅ Trama {ramo.poliza}_97.xls subida"
        else:
            ruta_archivo = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")
            msj_tra = f"✅ Trama {ramo.poliza}.xlsx subida"

        id_input_trama = "filArchivo" if bab_codigo != '4' else "fupExaminarPlantilla"
        input_file = wait.until(EC.presence_of_element_located((By.ID, id_input_trama)))
        input_file.send_keys(ruta_archivo)
        logging.info(msj_tra)

        btnCargar= wait.until(EC.element_to_be_clickable((By.ID, "btnCargar")))
        driver.execute_script("arguments[0].click();", btnCargar)
        logging.info("🖱️ Clic en Cargar la trama")

        time.sleep(5)
        try:

            wait.until(EC.alert_is_present())
            alert_trama = driver.switch_to.alert
            texto_alerta_trama = alert_trama.text.strip()

            if "Ingreso de planilla de trabajadores, satisfactorio" in texto_alerta_trama:
                alert_trama.accept()
                logging.info("✅ Alerta Aceptada")
            else:
                raise Exception(texto_alerta_trama)

        except TimeoutException:
            pass

        id_btn_procesar = "btnGrabar" if bab_codigo != '4' else "btnProcesar"
        btn_procesar = wait.until(EC.element_to_be_clickable((By.ID, id_btn_procesar)))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_procesar)
        logging.info("✅ Scroll hasta btn Procesar")
        btn_procesar.click()
        logging.info("🖱️ Clic en Procesar")

        try:
            wait.until(EC.alert_is_present())
            alert = driver.switch_to.alert
            logging.info(f"⚠️ Texto de la alerta: {alert.text}")
            alert.accept()
            logging.info("✅ Alerta Aceptada")
        except TimeoutException:
            pass

        ventana_principal = driver.current_window_handle

        # Flujo para descargar constancia
        try:

            id_btn_verConstancia = "btnGenerarPdf" if bab_codigo != '4' else "btnVerConstancia"
            btn_verConstancia = wait.until(EC.element_to_be_clickable((By.ID, id_btn_verConstancia)))

            # Guardar archivos antes del clic
            archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

            driver.execute_script("arguments[0].click();", btn_verConstancia)
            logging.info("🖱 Clic con JS en el botón 'Ver Constancia'")

            archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=1)

            cerrar_popup(driver, ventana_principal, wait)

            #La web te lo descarga con el mismo nombre de la poliza , ya no es necesario renombrar
            if archivo_nuevo:

                #SCTR no descarga con el nombre que es , vida ley si // arreglar
                if bab_codigo != '4':

                    logging.info(f"✅ Archivo descargado exitosamente")
                    ruta_original = archivo_nuevo[0]
                    ruta_final = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
                    os.rename(ruta_original, ruta_final)
                    logging.info(f"🔄 Constancia renombrada a '{ramo.poliza}.pdf'")
                else:
                    logging.info(f"✅ Constancia '{ramo.poliza}.pdf' descargado exitosamente")

                constancia = True
               
            else:
                raise Exception("No se encontró archivo nuevo después de descargar")

        except TimeoutException:
            raise Exception(f"No esta habilitado la opcion para descargar la constancia")

        # Flujo para descargar Proformas - Liquidacion y Facturacion
        if tipo_mes == 'MA':

            if bab_codigo != '4':
                
                try:

                    btn_proforma_salud = wait.until(EC.element_to_be_clickable((By.ID, "btnGenerarPdfFactura")))
                    archivos_antes_proforma_Salud = set(os.listdir(ruta_archivos_x_inclu))
                    driver.execute_script("arguments[0].click();", btn_proforma_salud)
                    logging.info("🖱 Clic con JS en el botón 'Ver Documento Salud'")

                    proforma_salud = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes_proforma_Salud,".pdf",cantidad=1)

                    cerrar_popup(driver, ventana_principal, wait)

                    if proforma_salud:

                        logging.info(f"✅ Archivo descargado exitosamente")
                        ruta_proforma_salud = proforma_salud[0]
                        ruta_nueva_proforma = os.path.join(ruta_archivos_x_inclu, f"endoso_{ramo.poliza}.pdf")
                        os.rename(ruta_proforma_salud, ruta_nueva_proforma)
                        proforma = True
                        logging.info(f"🔄 Facturación renombrado a 'endoso_{ramo.poliza}.pdf'")
                    else:
                        logging.error("❌ No se encontró archivo nuevo después de la descarga")

                except TimeoutException:
                    logging.error(f"El boton Ver Documento Salud no esta clickeable aún")

                try:

                    btn_proforma_pension= wait.until(EC.element_to_be_clickable((By.ID, "btnVerDocPension")))
                    archivos_antes_proforma_Pension = set(os.listdir(ruta_archivos_x_inclu))

                    driver.execute_script("arguments[0].click();", btn_proforma_pension)
                    logging.info("🖱 Clic con JS en el botón 'Ver Documento Pension'")

                    proforma_pension = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes_proforma_Pension,".pdf",cantidad=1)

                    cerrar_popup(driver, ventana_principal, wait)

                    if proforma_pension:

                        logging.info(f"✅ Archivo descargado exitosamente")
                        ruta_proforma_pension = proforma_pension[0]

                        nom_p_2 = f"{list_polizas[1]}" if len(list_polizas) == 2 else f"{ramo.poliza}"

                        ruta_nueva_proformaP = os.path.join(ruta_archivos_x_inclu, f"endoso_{nom_p_2}.pdf")
                        os.rename(ruta_proforma_pension, ruta_nueva_proformaP)
                        proforma = True
                        logging.info(f"🔄 Liquidación renombrado a 'endoso_{nom_p_2}.pdf'")
                    else:
                        logging.error("❌ No se encontró archivo nuevo después de la descarga")

                except TimeoutException:
                    logging.error(f"El boton Ver Documento Pension no esta clickeable aún")

            else:
                # En vida ley mes adelantado se descarga la constancia
                try:
                    btn_verLiqui = wait.until(EC.element_to_be_clickable((By.ID, "btnVerLiquidaciones")))

                    # Guardar archivos antes del clic
                    archivos_antes2 = set(os.listdir(ruta_archivos_x_inclu))

                    driver.execute_script("arguments[0].click();", btn_verLiqui)
                    logging.info("🖱 Clic con JS en el botón 'Ver Liquidacion'")

                    archivo_nuevo2 = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes2,".pdf",cantidad=1)

                    cerrar_popup(driver, ventana_principal, wait)

                    if archivo_nuevo2:

                        logging.info(f"✅ Archivo descargado exitosamente")
                        ruta_original = archivo_nuevo2[0]
                        ruta_final = os.path.join(ruta_archivos_x_inclu, f"endoso_{ramo.poliza}.pdf")
                        os.rename(ruta_original, ruta_final)
                        proforma = True
                        logging.info(f"🔄 Endoso renombrado a 'endoso_{ramo.poliza}.pdf'")
                    else:
                        logging.error("❌ No se encontró archivo nuevo después de la descarga")

                    proforma = True

                except TimeoutException:
                    logging.error(f"El boton Ver Liquidación no esta clickeable aún")

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

        #return True,True if tipo_mes == 'MA' else False,tipoError,detalleError
        logging.info(f"✅ {palabra_clave} en Pacifico realizada exitosamente.")

    except Exception as e:
        ramo_s = "VIDALEY" if bab_codigo == '4' else "SCTR"
        tomar_capturar(driver, ruta_archivos_x_inclu, f"ERROR_{ramo_s}_{tipo_mes}")
        logging.error(f"❌ Error en Pacifico {ramo_s} - {tipo_mes}: {e}")
        tipoError = f"PACI-{ramo_s}-{tipo_mes}"
        detalleError = str(e)
    
    return constancia,proforma,tipoError,detalleError
