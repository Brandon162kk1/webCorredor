#   --- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

import time
import os
import subprocess
import logging

#----- Variables de Entorno -------
urlRimacPW = os.getenv("urlRimacPW")
usernameRimac = os.getenv("usernameRimac")
passwordRimac = os.getenv("passwordRimac")

def subir_trama(driver,ruta_archivos_x_inclu,boton,ruta_trama):
    
    try:

        if not os.path.exists(ruta_trama):
            raise FileNotFoundError(f"La ruta {ruta_trama} no existe.")

        boton.click()
        logging.info("✅ Se hizo clic con JS en el botón 'Subir'.")

        logging.info("⌛ Esperando la ventana descarga de Linux Debian...")
        time.sleep(3)

        posibles_nombres = ["Open", "Open Files", "Abrir", "Abrir archivo"]
        ventana_id = None

        for nombre in posibles_nombres:
            resultado = subprocess.run(
                ["xdotool", "search", "--name", nombre],
                capture_output=True, text=True
            )
            if resultado.returncode == 0 and resultado.stdout.strip():
                ventana_id = resultado.stdout.strip().split("\n")[0]
                break

        if not ventana_id:
            raise Exception("❌ No se encontró ninguna ventana de selección de archivos.")

        subprocess.run(["xdotool", "windowactivate", "--sync", ventana_id])
        time.sleep(2)
        subprocess.run(["xdotool", "key", "ctrl+l"])
        time.sleep(0.5)
        subprocess.run(["xdotool", "type", "--delay", "200", ruta_trama])
        logging.info("📄 Se escribió la ruta del documento.")
        time.sleep(2)
        subprocess.run(["xdotool", "key", "Return"])
        logging.info("🖱️ Clic en Open")
        return True

    except Exception as ex:
        logging.error(f"❌ Error durante el flujo de subida: {ex}")
        driver.save_screenshot(os.path.join(ruta_archivos_x_inclu, "error_click_subir_archivo.png"))
        return False

def realizar_solicitud_PortalWeb(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,ejecutivo_responsable,
                                 ba_codigo,nombre_cliente,ruc_empresa,ramo):

    driver.get("https://www.rimac.com.pe/PortalWebEmpresa/login.do") 
    logging.info("🔐 Iniciando sesión en RIMAC Portal Web...")
 
    user_input = wait.until(EC.presence_of_element_located((By.ID, "CODUSUARIO")))
    user_input.clear()
    user_input.send_keys(ramo.usuario)
    logging.info("⌨️ Usuario ha sido Digitando.")
 
    pass_input = wait.until(EC.presence_of_element_located((By.ID, "CLAVE")))
    pass_input.clear()
    pass_input.send_keys(ramo.clave)
    logging.info("⌨️ Password ha sido Digitado")
 
    ingresar_btn = wait.until(EC.element_to_be_clickable((By.ID, "btningresar")))
    driver.execute_script("arguments[0].click();", ingresar_btn)
    logging.info("🖱️ Clic en 'Ingresar'.")

    #-------------------------------
    codigo_path = "/codigo_rimac_SAS/codigo.txt"

    print("⏳ Esperando código...")
    while not os.path.exists(codigo_path):
        #print("⏳ Esperando que se cree codigo.txt...")
        time.sleep(2)

    with open(codigo_path, "r") as f:
        codigo = f.read().strip()

    print(f"✅ Código recibido desde volumen: {codigo}")

    token_input = wait.until(EC.element_to_be_clickable((By.ID, "TOKEN")))
    token_input.clear()
    token_input.send_keys(codigo)
    logging.info(f"⌨️ Código {codigo} digitado correctamente en TOKEN")

    # --- Eliminar el archivo después de usarlo ---
    try:
        os.remove(codigo_path)
        #print("🧹 Archivo codigo.txt eliminado del volumen.")
    except FileNotFoundError:
        print("⚠️ No se encontró codigo.txt al intentar eliminarlo (ya fue borrado).")
    except Exception as e:
        print(f"❌ Error al eliminar codigo.txt: {e}")

    #---------------------------

    ingresar_btn2 = wait.until(EC.element_to_be_clickable((By.ID, "btningresar")))
    driver.execute_script("arguments[0].click();", ingresar_btn2)
    logging.info("🖱️ Clic en 'Ingresar' Luego del TOKEN.")
 
    time.sleep(4)
 
    try:
 
        link_seguros = wait.until(EC.element_to_be_clickable((By.ID, "panelSeguros")))
        link_seguros.click()
        logging.info("🖱️ Clic en 'Panel Seguros'")
        link_seguros2 = wait.until(EC.element_to_be_clickable((By.ID, "panelSaludd"))) #Este es el sub panel de seguros
        link_seguros2.click()
        logging.info("🖱️ Clic en 'Sub Panel Seguros'")
 
        link_riegosLab = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Riesgos Laborales")))
        link_riegosLab.click()
        logging.info("🖱️ Clic en 'Riesgos Laborales'")
 
        link_sctr = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "SCTR Pensión y SCTR Salud")))
        link_sctr.click()
        logging.info("🖱️ Clic en 'SCTR Pensión y SCTR Salud'")
 
        input_doc = wait.until(EC.element_to_be_clickable((By.ID, "nroDocumento")))
        input_doc.click()
        logging.info("🖱️ Clic en el input 'Número de Documento'")
 
        input_doc.clear()
        input_doc.send_keys(ruc_empresa)
        logging.info(f"⌨️ Numero de RUC ingresado: {ruc_empresa}")
 
        btn_buscar = wait.until(EC.element_to_be_clickable((By.ID, "btn-bsq-search")))
        btn_buscar.click()
        logging.info("🖱️ Clic en el Botón 'Buscar'")
 
        time.sleep(3)
 
        tabla = wait.until(EC.presence_of_element_located((By.ID, "table-rpta-bsq-broker")))
 
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")
 
        encontrado = False
 
        time.sleep(2.1)
 
        for fila in filas:
            columnas = fila.find_elements(By.TAG_NAME, "td")
            nro_doc = columnas[1].text.strip()
            if nro_doc == ruc_empresa:
                logging.info(f"✅ Documento encontrado: {nro_doc}")
                fila.click()
                logging.info(f"🖱️ Clic en la fila con el RUC {ruc_empresa}")
                btn_select = wait.until(EC.element_to_be_clickable((By.ID, "btn-bsq-select")))
                btn_select.click()
                logging.info("🖱️ Clic en el Botón 'Seleccionar'.")
                encontrado = True
                break
        time.sleep(2)
        if not encontrado:
            logging.info("❌ Documento NO encontrado en la tabla.")
 
        ms2 = "Inclusiones" if tipo_proceso == 'IN' else "Renovación"
        
        link_inclusiones = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, f"{ms2}")))
        link_inclusiones.click()
        logging.info(f"🖱️ Clic en '{ms2}' ")
 
        tabla = wait.until(EC.presence_of_element_located((By.ID, "table_polizas")))
 
        logging.info("⏳ Esperando a que las filas estén cargadas")
        filas = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#table_polizas tbody tr")))
        logging.info(f"✅ Filas cargadas: {len(filas)}")
 
        encontrado = False
 
        for fila in filas:
            columnas = fila.find_elements(By.TAG_NAME, "td")
            nro_poliza = columnas[0].text.strip()
 
            logging.info(f"Leyendo póliza: {repr(nro_poliza)}")
 
            if nro_poliza in list_polizas:
                logging.info(f"✅ Póliza encontrada: {nro_poliza}")
 
                # Asegurar que la fila sea clickeable
                wait.until(EC.element_to_be_clickable(fila)).click()
                logging.info("🖱️ Clic en la fila")
 
                encontrado = True
 
        if not encontrado:
            raise Exception("Ninguna póliza de la lista fue encontrada en la tabla.")
        else:

            #Scroll hacia abajo
            driver.execute_script("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});")
     
            input_fecha = wait.until(EC.presence_of_element_located((By.ID, "fecinicio2")))
 
            driver.execute_script("""
                const input = arguments[0];
                const today = new Date();
 
                const year = today.getFullYear();
                const month = String(today.getMonth() + 1).padStart(2, '0');
 
                const fecha = `01/${month}/${year}`;
 
                input.removeAttribute('readonly');
                input.value = fecha;
 
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
            """, input_fecha)
 
            logging.info("📅 Fecha seteada: día 1 del mes actual")

            continuar = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Continuar")))
            continuar.click()
            logging.info("🖱️ Clic en 'Continuar'")
 
        time.sleep(4)
 
        wait.until(EC.element_to_be_clickable((By.ID, "radio_trama"))).click()
        logging.info("✅ Opción 1 Seleccionada ")
 
        boton = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='file']")))
        ruta_trama_xls_rimacPW = f"{ruta_archivos_x_inclu}/{list_polizas[0]}.xls"
 
        if subir_trama(driver,ruta_archivos_x_inclu,boton,ruta_trama_xls_rimacPW):
            logging.info ("✅ Trama subida")
        else:
            raise Exception("Error al subir la trama")
 
        time.sleep(2)
 
        subir_archivo = wait.until(EC.element_to_be_clickable((By.ID, "subirArchivo")))
        subir_archivo.click()
        logging.info("🖱️ Clic en 'Subir Archivo'")
 
        wait.until(EC.visibility_of_element_located((By.ID, "miDiv")))
 
        while True:

            time.sleep(1.2)
 
            estado = driver.find_element(By.ID, "Estado_validacion").get_attribute("value").strip()
            mensaje = driver.find_element(By.ID, "Mensaje_validacion").get_attribute("value").strip()
 
            #logging.info(f"🔄 Estado actual: {estado} | Mensaje: {mensaje}")
 
            if estado == "Error":
                logging.info(f"❌ {estado}: {mensaje}")
 
                # texto_error = f"Estado: {estado}\nMensaje: {mensaje}\n"
 
                # with open("errores_trama.txt", "a", encoding="utf-8") as f:
                #     f.write(texto_error + "\n")
 
                # ruta_imagen = os.path.join(ruta_archivos_x_inclu, f"{get_timestamp()}.png")
                # driver.save_screenshot(ruta_imagen)
                # logging.info(f"📸 Captura guardada en: {ruta_imagen}")
 
                #logging.info("❌ Proceso detenido por ERROR.")
                #return False
                raise Exception(f"{mensaje}")
 
            if estado == "Procesado":
                logging.info(f"✅ {estado}")
 
                driver.execute_script("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});")
                time.sleep(1.2)
 
                boton_cargar = wait.until(EC.element_to_be_clickable((By.ID, "buttonCargar")))
                boton_cargar.click()
                logging.info("🖱️ Clic en 'Cargar' ejecutado")
 
                break
                
            #logging.info("⏳ Estado aún no final (ni Error ni Procesado)")
 
        time.sleep(4)
 
        driver.execute_script("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});")
 
        time.sleep(1.2)
 
        label = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='aceptoafil'], label[style*='padding-left']")))
        label.click()
        logging.info("☑️ Checkbox marcado mediante label")
 
        time.sleep(1.2)
 
        wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(@onclick,'continuarPaso3')]")))
        driver.execute_script("continuarPaso3();")
        logging.info("🔥 Click Continuar.")

        #Falta clic en confirmar la solicitud
 
    except Exception as e:
        logging.info(f" ❌ Error en el flujo de selección de póliza en Portal Web: {e}")
        raise Exception(e)
    finally:
        input("Esperar")