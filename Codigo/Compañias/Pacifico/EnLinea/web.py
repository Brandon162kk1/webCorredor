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

def seleccionar_radio(id_radio,opcion,wait,driver):
    radio = wait.until(EC.element_to_be_clickable((By.ID, id_radio)))
    driver.execute_script("arguments[0].click();", radio)

    # esperar postback → el radio vuelva a existir y esté marcado
    wait.until(lambda d: d.find_element(By.ID, id_radio).is_selected())

    logging.info(f"✅ Radio {opcion} seleccionado")

def realizar_solicitud_pacifico(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,
                          palabra_clave,ruc_empresa,ejecutivo_responsable,bab_codigo,ramo):
    
    numero_poliza = list_polizas[0]
    tipoError = ""
    detalleError = ""
  
    try:

        driver.get("https://www0.pacificoseguros.com/loginPacifico/login.aspx") 
        logging.info("⌛ Cargando la Web de Pacifico en Linea")
           
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

        # Click real
        link = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='spn_5']//a[contains(.,'Renovación')]")))
        driver.execute_script("arguments[0].click();", link)
        logging.info("🖱️ Clic en 'Renovación / Inlclusión'")

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Principal")))

        logging.info("🔀 Cambiado al frame 'Principal'")

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

            # MUY IMPORTANTE
            driver.switch_to.default_content()

        except TimeoutException:
            logging.info("✅ No apareció ninguna alerta en el tiempo especificado")

        wait.until(EC.presence_of_element_located((By.NAME, "Menu")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Menu")))

        folder = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='spn_5']//img")))

        if "folder.gif" in folder.get_attribute("src").lower():
            driver.execute_script("arguments[0].click();", folder)
            time.sleep(0.5)

        link = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@id='spn_5']//a[contains(.,'Renovación')]")))

        driver.execute_script("arguments[0].click();", link)
        logging.info("🖱️ Clic en 'Renovación / Inclusión' nuevamente")

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

        logging.info(f"🖱️ Clic REAL en pestaña {'SCTR' if bab_codigo != '4' else 'Vida Ley'}")
    
        driver.switch_to.default_content()

        #frames2 = driver.find_elements(By.TAG_NAME, "frame")
        #logging.info(f"🧩 Frames detectados: {[f.get_attribute('name') for f in frames2]}")

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Principal")))
        #logging.info("🔀 Cambiado al frame 'Principal'")

        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Contenido")))
        #logging.info("🔀 Entré a frame Contenido")

        # Guardar handles antes de abrir la segunda ventana
        ventana_principal = driver.current_window_handle
        handles_ventana_principal = set(driver.window_handles)

        buscar_btn = wait.until(EC.element_to_be_clickable((By.ID, "imgBuscarCliente")))
        driver.execute_script("arguments[0].click();", buscar_btn)
        logging.info("🖱️ Clic en la Lupa del cliente")

        # Esperar y cambiar a la segunda ventana
        wait.until(lambda d: len(d.window_handles) > len(handles_ventana_principal))
        nueva_ventana2 = (set(driver.window_handles) - handles_ventana_principal).pop()
        driver.switch_to.window(nueva_ventana2)
        logging.info("--- Ventana 2---------🌐")

        time.sleep(3)

        id_ruc_input_2 = "txtDocumento" if bab_codigo != '4' else "txtNroDocumento"

        ruc_input2 = wait.until(EC.presence_of_element_located((By.ID, id_ruc_input_2))) 
        ruc_input2.clear()
        ruc_input2.send_keys(ruc_empresa)
        logging.info(f"✅ Se ingresó el RUC '{ruc_empresa}'")

        id_buscar_btn_2 = "btnBuscar" if bab_codigo != '4' else "btnBuscarCliente"

        buscar_btn2 = wait.until(EC.element_to_be_clickable((By.ID, id_buscar_btn_2)))
        buscar_btn2.click()
        logging.info("🖱️ Clic en Buscar")
  
        id_cod_cliente = "fr_DevolverCodigo" if bab_codigo != '4' else "seleccionarCliente"

        cliente = wait.until(EC.element_to_be_clickable((By.XPATH, f"""(//a[contains(@onclick,'{id_cod_cliente}')]|//td[contains(@onclick,'{id_cod_cliente}')])[1]""")))

        cliente.click()
        logging.info("🖱️ Clic en primer cliente encontrado")

        driver.switch_to.window(ventana_principal) #Cambia de VENTANA o PESTAÑA del navegador
        logging.info("🔙 Regresando a la ventana principal")

        try:
            WebDriverWait(driver,10).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            texto_alerta = alert.text   # 👈 guardar antes
            alert.accept()
            logging.info(f"⚠️ Texto de la alerta: {texto_alerta}")
            logging.info("✅ Alerta Aceptada")
            raise Exception(texto_alerta)
        except TimeoutException:
            logging.info("✅ No apareció ninguna alerta en el tiempo especificado")

        if bab_codigo == '4':

            buscar_lupa_poliza = wait.until(EC.element_to_be_clickable((By.ID, "imgBuscarPoliza")))
            driver.execute_script("arguments[0].click();", buscar_lupa_poliza)
            logging.info("🖱️ Clic en la Lupa de la Póliza")

            # Esperar y cambiar a la tecera ventana
            wait.until(lambda d: len(d.window_handles) > len(handles_ventana_principal))
            nueva_ventana3 = (set(driver.window_handles) - handles_ventana_principal).pop()
            driver.switch_to.window(nueva_ventana3)
            logging.info("--- Ventana 3---------🌐")

            cliente_vl = wait.until(EC.element_to_be_clickable((By.ID, "grvPoliza_ctl03_NroPoliza")))
            cliente_vl.click()
            logging.info("🖱️ Clic en la primer poliza encontrada")

            driver.switch_to.window(ventana_principal) #Cambia de VENTANA o PESTAÑA del navegador
            logging.info("🔙 Regresando a la ventana principal")

        # 🔑 REENTRAR A LOS FRAMES
        driver.switch_to.default_content() #Sale de TODOS los frames / iframes
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Principal")))
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Contenido")))
        #----------------------------

        if bab_codigo == '2':  # Pensión
            seleccionar_radio("rblProducto_1","Pension",wait,driver)

        elif bab_codigo == '3':  # Ambos
            seleccionar_radio("rblProducto_2","Ambos",wait,driver)

        if tipo_proceso == 'IN':
            seleccionar_radio("rdbTipoMov_1","Inclusión",wait,driver)

        ruta_archivo = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}_97.xls")

        input_file = wait.until(EC.presence_of_element_located((By.ID, "filArchivo")))
        input_file.send_keys(ruta_archivo)
        logging.info(f"✅ Trama {ramo.poliza}_97.xls subida")

        btnCargar= wait.until(EC.element_to_be_clickable((By.ID, "btnCargar")))
        driver.execute_script("arguments[0].click();", btnCargar)
        logging.info("🖱️ Clic en Cargar la trama")

        try:
            WebDriverWait(driver,10).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            texto_alerta = alert.text   # 👈 guardar antes

            if texto_alerta.startswith("Ingreso de planilla de trabajadores, satisfactorio."):
                alert.accept()
                logging.info(f"⚠️ Texto de la alerta: {texto_alerta}")
                logging.info("✅ Alerta Aceptada")
            else:
                raise Exception(f"{texto_alerta}")

        except TimeoutException:
            logging.info("✅ No apareció ninguna alerta en el tiempo especificado")

        btn_procesar = wait.until(EC.element_to_be_clickable((By.ID, "btnGrabar")))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_procesar)
        logging.info("✅ Scroll hasta btn Procesar")
        #btn_procesar.click()

        # try:
        #     WebDriverWait(driver,10).until(EC.alert_is_present())
        #     alert = driver.switch_to.alert
        #     logging.info(f"⚠️ Texto de la alerta: {alert.text}")
        #     alert.accept()
        #     logging.info("✅ Alerta Aceptada")
        # except TimeoutException:
        #     raise Exception(f"No apareció ninguna alerta de confirmación")

        #clic en ver constancia

        #if es MA ver liquidacion

    except Exception as e:
        logging.error(f"Error: {e}")
    finally:
        input("Esperar")
