#   --- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import ActionChains
from Tiempo.fechas_horas import get_timestamp,get_fecha_hoy,get_dia
#   --- Imports ----
import time
import os
import subprocess
import logging
from datetime import datetime
import unicodedata
import re
import json
import shutil

def esperar_ventana(nombre, timeout=30):
    inicio = time.time()
    while time.time() - inicio < timeout:
        result = subprocess.run(
            ["xdotool", "search", "--name", nombre],
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL
        )
        if result.stdout.strip():
            return True
        time.sleep(0.5)
    return False

# def normalizar(texto):
#     return unicodedata.normalize('NFKD', texto.lower()).encode('ascii', 'ignore').decode('utf-8')

def normalizar(texto):

    if not texto:
        return ""

    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = re.sub(r"\s+", " ", texto)  # colapsar espacios
    return texto.lower().strip()

def subir_trama(boton,ruta_trama):
    
    try:

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
        logging.error(f"Ventana de Dialogo: {ex}")
        return False

def click_descarga_documento(driver,boton_descarga,nombre_documento,impresion):
    
    try:

        driver.execute_script("arguments[0].click();", boton_descarga)
        logging.info("🖱 Se hizo clic con JS en el botón de descarga.")

        time.sleep(3)

        if impresion:
            time.sleep(3)
            subprocess.run(["xdotool", "key", "Return"])
            logging.info(f"🖱️ Clic en 'Save'.")
            time.sleep(3)

        time.sleep(3)

        logging.info("⌛ Esperando la ventana descarga de Linux Debian...")
        time.sleep(2)
        subprocess.run(["xdotool", "search", "--name", "Save File", "windowactivate", "windowfocus"])
        logging.info("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
        subprocess.run(['xdotool', 'type',"--delay", "200", nombre_documento])
        logging.info("📄 Se cambio el nombre del documento")
        time.sleep(2)
        subprocess.run(["xdotool", "key", "Return"])
        logging.info(f"🖱️ Clic en 'Save' para confirmar")

        if impresion:
            try:
                time.sleep(3)
                aceptar_btn = WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID, "CloseModal")))
                aceptar_btn.click()
                logging.info(f"🖱️ Clic en 'Aceptar' para cerrar el Modal.")
            except TimeoutException:
                logging.info(f"✅ No apareció nada despues de descargar.")
                pass

        return True

    except Exception as ex:
        logging.info("❌ Error durante el flujo de descarga:", ex)
        return False

def detectar_constancia_en_tabla(driver):

    """Detecta qué tipo de constancia aparece realmente en la tabla con prioridad."""
    filas = driver.find_elements(By.XPATH, "//table[@id='createdDocumentsTable']//tr")

    encontrada_salud = False
    encontrada_pension = False

    for fila in filas:
        try:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if len(celdas) < 2:
                continue

            texto = normalizar(celdas[1].text)

            # 🔥 PRIORIDAD ALTA
            if "constancia conjunta" in texto:
                return "CONJUNTA"

            # guardar pero NO devolver aún
            if "constancia salud" in texto:
                encontrada_salud = True

            if "constancia pension" in texto or "constancia pensión" in texto:
                encontrada_pension = True

        except:
            continue

    # Si NO hubo conjunta → revisar prioridades siguientes
    if encontrada_salud:
        return "SALUD"

    if encontrada_pension:
        return "PENSION"

    return "No hay constancia, solo Proformas"

def realizar_solicitud_sanitas(driver,wait,list_url_san,list_polizas,tipo_mes,ruta_archivos_x_inclu,
                               tipo_proceso,palabra_clave,ramo):
    
    numero_poliza = list_polizas[0]
    encontrado = False
    constancia = False
    constancia_conjunta = False
    proforma = False
    tipoError = ""
    detalleError = ""

    try:

        for i, url in enumerate(list_url_san):

            if i == 0:
                compania = "Crecer" 
                logging.info(f"---------Sanitas {compania}-----------")
            else:
                compania = "Protecta" 
                logging.info(f"-------Sanitas {compania}----------")
                        
            driver.get(url)
            logging.info(f"⌛ Cargando la Web de Sanitas {compania}")

            username_input = wait.until(EC.presence_of_element_located((By.ID, 'Login')))
            username_input.clear()
            username_input.send_keys(ramo.usuario)
            logging.info("⌨️ Digitando el Usuario")

            password_input = wait.until(EC.presence_of_element_located((By.ID, 'Password')))
            password_input.clear()
            password_input.send_keys(ramo.clave)
            logging.info("⌨️ Digitando la Contraseña")
    
            submit_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@type='submit']")))
            submit_button.click()
            logging.info("🖱️ Clic en Iniciar Sesión")
    
            wait.until(EC.url_changes(url))

            afiliaciones_link = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'Afiliaciones')]")))
            afiliaciones_link.click()
            logging.info("🖱️ Clic en Afiliaciones")

            texto_link = "Inclusión" if tipo_proceso == "IN" else "Renovación"

            link = wait.until(EC.presence_of_element_located((By.XPATH, f"//a[contains(text(),'{texto_link}')]")))
            link.click()

            logging.info(f"🖱️ Clic en {palabra_clave}")

            search_contract_button = wait.until(EC.presence_of_element_located((By.ID, 'btnSearchContract')))
            search_contract_button.click()
            logging.info("🖱️ Clic en Buscar")

            wait.until(EC.visibility_of_element_located((By.ID, 'searchContractPopup')))
            logging.info("⌛ Esperando el Modal")

            contract_number_input = wait.until(EC.presence_of_element_located((By.ID, 'ContractNumber')))
            contract_number_input.send_keys(numero_poliza)
            logging.info("✅ Ingresando Póliza")

            search_button = wait.until(EC.presence_of_element_located((By.ID, 'btnSearchContracts')))
            search_button.click()
            logging.info("🖱️ Clic en Buscar")

            try:               
                filas = WebDriverWait(driver,5).until(EC.presence_of_all_elements_located((By.XPATH, f"//td[contains(text(),'{numero_poliza}')]/parent::tr")))
                fila_correcta = filas[0]
                fila_correcta.click()
                logging.info(f"✅ Se seleccionó la fila para la póliza {numero_poliza}")
                encontrado = True
            except:
                logging.error(f"❌ No se encontraron resultados para el número de contrato {numero_poliza}")
                continue

            #tipoError = "Solo existe proformas"
            #detalleError = "Primero pagar para buscar constancia"
            #break
            time.sleep(2)
            
            accept_button = wait.until(EC.presence_of_element_located((By.ID, 'btnSelectContract')))
            accept_button.click()
            logging.info("🖱️ Clic en Aceptar")

            try:
                modal_poliza = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID, "MessageBox")))                  
                span_txt_modal = modal_poliza.find_element(By.ID, "message").text
                logging.warning(f"⚠️ Apareció el modal con advertencia: {span_txt_modal}.")
                raise Exception(f"{span_txt_modal}")
            except TimeoutException:
                logging.info("✅ No hay advertencias, el flujo continúa.")

            client_branch_office_select = wait.until(EC.presence_of_element_located((By.ID, 'ClientBranchOfficeId')))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", client_branch_office_select)

            time.sleep(2)
    
            select = Select(client_branch_office_select)

            for option in select.options:                                       #----- for 1
                texto_opcion = option.text.strip()
                if normalizar(ramo.sede) == normalizar(texto_opcion):
                    select.select_by_visible_text(texto_opcion)
                    logging.info(f"✅ Opción seleccionada: {texto_opcion}")
                    break                                                       #----- for 1
            else:
                raise Exception("No se encontró ninguna Sede que coincida")

            time.sleep(2)

            if tipo_proceso == 'IN':

                input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "InsureFrom")))
                valor_fecha_web = input_fecha.get_attribute("value")
                fecha_web = datetime.strptime(valor_fecha_web, "%d/%m/%Y")
                dia_web = fecha_web.day
                logging.info(f"📅 Fecha de la web de Sanitas {compania}: {valor_fecha_web} | Día: {dia_web}")

                # Por defecto se pone en el primer dia del mes
                if datetime.strptime(ramo.f_inicio, "%d/%m/%Y").day != dia_web:

                    input_fecha = wait.until(EC.visibility_of_element_located((By.ID, "InsureFrom")))
                    input_fecha.clear()
                    input_fecha.send_keys(ramo.f_inicio)

                    driver.find_element(By.TAG_NAME,"body").click()

                    fecha_date = datetime.strptime(ramo.f_inicio,"%d/%m/%Y").date()

                    if fecha_date < get_fecha_hoy().date():
                        diferencia = (get_fecha_hoy().date() - fecha_date).days
                        logging.info(f"📅 Vigencia de Inicio retroactiva ingresada: {ramo.f_inicio} con {diferencia} {'dia' if diferencia == 1 else 'dias'} de diferencia.")
                    else:
                        logging.info(f"📅 Vigencia de Inicio ingresada: {ramo.f_inicio}")

                    try:
                        # Modal con Advertencias de la fecha
                        modal_fecha = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID,"MessageBox")))
                        span_txt_fecha = modal_fecha.find_element(By.ID, "message").text
                        logging.warning(f"⚠️ Apareció el modal con advertencia: {span_txt_fecha}")

                        if span_txt_fecha == "No es posible realizar movimientos retroactivos superiores a 2 días.":
                            raise Exception(f"{span_txt_fecha}")

                        btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "CloseModal")))
                        btn_aceptar.click()
                        logging.info("🖱️ Se hizo clic en 'Aceptar'.")

                    except TimeoutException:
                        logging.info("✅ No se detectó el modal de advertencias en la fecha. Continuando flujo.")

                else:
                    logging.info("📅 La fecha de vigencia de inicio no requiere ser modificada.")

            if tipo_proceso == 'RE':
                select_element = wait.until(EC.presence_of_element_located((By.ID, 'OfficeProductionId')))
                select = Select(select_element)
                select.select_by_value("2")
                logging.info("🖱️ Se hizo clic en 'Lima'.")

            cargar_planilla_button = wait.until(EC.presence_of_element_located((By.ID, "btnCargarPlanilla")))
            cargar_planilla_button.click()
            logging.info("🖱️ Clic en Cargar Planilla")

            wait.until(EC.visibility_of_element_located((By.ID, "uploadAffiliationsFileUpload")))
            logging.info("⌛ Esperando que aparezca Modal...")

            buscar_span = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'btn-file') and contains(., 'Buscar')]")))

            ruta_trama_xls_sanitas = f"{ruta_archivos_x_inclu}/{numero_poliza}_97.xls"

            if not os.path.exists(ruta_trama_xls_sanitas):
                raise FileNotFoundError(f"La ruta {ruta_trama_xls_sanitas} no existe.")

            if subir_trama(buscar_span,ruta_trama_xls_sanitas):               
                time.sleep(3)
                logging.info(f"✅ Trama {numero_poliza}.xls adjuntada")
            else:
                raise Exception(f" Problemas al subir la trama")

            btn_subir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Subir Archivo')]")))
            btn_subir.click()
            logging.info("🖱️ Clic en 'Subir Archivo' para validar la trama en la compañía")

            time.sleep(3)

            try:
                # Modal con Errores en la Trama
                modal_error_trama = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID, "MessageBoxWithScroll")))             
                span_txtArea_modal = modal_error_trama.find_element(By.ID, "message").text
                logging.warning(f"⚠️ Apareció el modal con errores de la Trama")
                raise Exception(f"{span_txtArea_modal}")
            except TimeoutException:
                logging.info("✅ No se detectó el modal de errores. Continuando flujo.")

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            logging.info("🖱️ Scroll hasta abajo de la pagina")

            confirmar_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnConfirm')))
            
            confirmar_button.click()
            logging.info("🖱️ Clic en Confirmar.")

            try:
                modal_advertencia = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "MessageBox")))
                spana_dv_txt_modal = modal_advertencia.find_element(By.ID, "message").text
                logging.warning(f"⚠️ Apareció el modal con advertencia: {spana_dv_txt_modal}")
                raise Exception(f"{spana_dv_txt_modal}")
            except TimeoutException:
                logging.info("✅ No se detectó el modal de advertencias. Continuando flujo.")

            try:
                # Modal de Retroactividad para autorizar no siniestros
                modal_advertencia = WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.ID, "ConfirmationMessageBox")))
                logging.warning(f"⚠️ Apareció el modal con advertencia de retroactividad.")
                boton_si = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn-primary') and contains(text(), 'Sí')]")))
                boton_si.click()
                logging.info("🖱️ Se hizo clic en el botón 'Sí'.")
            except TimeoutException:
                logging.info("✅ No se detectó el modal de Retroactividad. Continuando flujo.")
    
            wait.until(EC.visibility_of_element_located((By.ID, 'createdDocumentsPopUp')))
            logging.info("⌛ Cargando Modal de los documentos...")

            modal_body = driver.find_element(By.CLASS_NAME, 'dataTables_scrollBody')
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", modal_body)

            ruta_imagen = os.path.join(ruta_archivos_x_inclu,f"{get_timestamp()}_verDocumentos.png")
            driver.save_screenshot(ruta_imagen)

            tipo_constancia = detectar_constancia_en_tabla(driver)
            documentos_a_descargar = []

            logging.info(f"🔎 Tipo de constancia detectada en la tabla: {tipo_constancia}")

            constancia_conjunta = False

            if tipo_constancia == "CONJUNTA":
                documentos_a_descargar.append({
                    "nombreDoc": "Constancia conjunta",
                    "alias": list_polizas[0],
                    "tittle": "Descargar PDF",
                    "impresion": False
                })
                constancia_conjunta = True
                constancia = True

            elif tipo_constancia == "SALUD":
                documentos_a_descargar.append({
                    "nombreDoc": "Constancia Salud - Sanitas Perú S.A.",
                    "alias": list_polizas[0],
                    "tittle": "Descargar PDF",
                    "impresion": False
                })
                constancia = True

            elif tipo_constancia == "PENSION":
                documentos_a_descargar.append({
                    "nombreDoc": "Constancia Pensión - Crecer S.A.",
                    "alias": list_polizas[0],
                    "tittle": "Descargar PDF",
                    "impresion": False
                })
                constancia = True

            else:
                logging.warning("⚠️ No se encontró ninguna constancia para descargar")
                constancia = False
                tipoError = "Solo existe Proformas"
                detalleError = "Primero pagar para obtener constancias"
                #raise Exception("No se encontró ninguna constancia en la tabla 'Documentos'.")

            if tipo_mes != "MV":

                if len(list_polizas) == 2:

                    proforma = True

                    documentos_a_descargar.append({
                        "nombreDoc": "Proforma - Sanitas Perú S.A.",
                        "alias": f"endoso_{list_polizas[0]}", #proforma_
                        "tittle": "Imprimir",
                        "impresion": True
                    })

                    documentos_a_descargar.append({
                        "nombreDoc":
                            "Aviso de cobranza - Protecta S.A." if i == 1
                            else "Aviso de cobranza - Crecer Seguros S.A.",
                        "alias": f"endoso_{list_polizas[1]}", #proforma
                        "tittle": "Imprimir",
                        "impresion": True
                    })

                elif tipo_constancia == "SALUD":
                    documentos_a_descargar.append({
                        "nombreDoc": "Proforma - Sanitas Perú S.A.",
                        "alias": f"endoso_{list_polizas[0]}", #proforma
                        "tittle": "Imprimir",
                        "impresion": True
                    })

            logging.info(f"📂 Documentos a descargar: {documentos_a_descargar}")

            for doc in documentos_a_descargar:                                                          ## --- for 2

                nombre_documento = doc["nombreDoc"]
                alias_archivo = doc["alias"]
                title_documento = doc["tittle"]
                impresion = doc["impresion"]

                filas = driver.find_elements(By.XPATH, "//table[@id='createdDocumentsTable']//tr")
                fila_objetivo = None

                for i, fila in enumerate(filas, 1):                        # ----------- for 3

                    try:
                        celdas = fila.find_elements(By.XPATH, "./td")
                        if len(celdas) < 2:
                            continue

                        texto_celda = celdas[1].text.strip()
                        logging.info(f"-----------")
                        logging.info(f"🔍 Fila {i}: '{texto_celda}'")

                        if normalizar(nombre_documento) in normalizar(texto_celda):
                            fila_objetivo = fila
                            break                                           # salimos de este for3  porque ya encontramos la fila
                    except Exception as e:
                        logging.warning(f"⚠️ Fila {i} error: {str(e)}")

                if not fila_objetivo:
                    logging.warning(f"❌ No se encontró la fila para: {nombre_documento}")
                    continue

                try:

                    descargar_link = fila_objetivo.find_element(By.XPATH, f".//a[@title='{title_documento}']")

                    driver.execute_script("arguments[0].scrollIntoView(true);", descargar_link)
                        
                    if click_descarga_documento(driver,descargar_link,alias_archivo,impresion):
                        time.sleep(3)
                        ruta_doc_descargado = f"{ruta_archivos_x_inclu}/{alias_archivo}.pdf"
                        if os.path.exists(ruta_doc_descargado):
                            logging.info(f"✅ {alias_archivo}.pdf descargado exitosamente")
                        else:
                            logging.warning(f"⚠️ No se encontró el archivo descargado en {ruta_archivos_x_inclu}")
                    else:
                        logging.info(f"❌ No se logró descargar {alias_archivo}, Verifica manualmente.")

                except Exception as e:
                    logging.error(f"❌ Error al intentar descargar: {nombre_documento} ,Detalle: {e}")
                    continue

            # if not constancia_conjunta:
            #     logging.info("No existe constancia conjunta")
            
            time.sleep(3)

            if encontrado:

                # Si hay constancia conjunta, se guardo con nombre de salud y hay que renombrarla para pension
                if constancia_conjunta: #constancia_conjunta
                    logging.info("🔄 Renombrando para Pensión")

                    ruta_salud = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[0]}.pdf")

                    # Nuevo nombre para Pensión
                    ruta_pension = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[1]}.pdf")

                    if os.path.exists(ruta_salud):
                        shutil.copy2(ruta_salud,ruta_pension)
                        logging.info(f"📄 Copia creada: {ruta_pension}")
                    else:
                        logging.error(f"❌ No existe el archivo base: {ruta_salud}")

                break

        # Cierre del navegador o continúa según resultado
        if not encontrado:
            raise Exception(f"No se encontró la póliza '{numero_poliza}' en ninguna de las compañías")
        else:
            logging.info("-----------")
            logging.info(f"✅ {palabra_clave} en SANITAS realizada exitosamente.")

        time.sleep(3) 

    except Exception as e:
        logging.error(f"❌ Error en Sanitas (SCTR) - {tipo_mes}: {e}")
        tipoError = f"SANI-SCTR-{tipo_mes}"
        detalleError = e
    finally:
        #logging.info(f"1: {tipoError} , 2 : {detalleError}")
        #return True,True,tipoError,detalleError
        return constancia,proforma,tipoError,detalleError
            
def procesar_solicitud_san_crecer_vl(driver,wait,numero_poliza,tipo_proceso,ruta_archivos_x_inclu,palabra_clave):
 
    driver.get("https://app.crecerseguros.pe/Plataforma/manage/login")
    logging.info("---Ingresando a Crecer Seguros🌐---")

    user_input = wait.until(EC.presence_of_element_located((By.ID, "suser")))
    user_input.clear()
    user_input.send_keys()
    logging.info("✅ Digitando el Username")
        
    pass_input = wait.until(EC.presence_of_element_located((By.ID, "spassword")))
    pass_input.clear()
    pass_input.send_keys()
    logging.info("✅ Digitando el Password")

    #Importante: Primero cambiar al iframe del reCAPTCHA
    iframe = driver.find_element(By.XPATH, "//iframe[contains(@title, 'reCAPTCHA')]")
    driver.switch_to.frame(iframe)

    # Ahora puedes ubicar el checkbox y hacer clic
    checkbox = driver.find_element(By.ID, "recaptcha-anchor")

    # Algunas veces es mejor mover el mouse encima antes de hacer clic
    actions = ActionChains(driver)
    actions.move_to_element(checkbox).click().perform()
    logging.info("🖱️ Clic en el Captcha")

    # Luego debes volver al contexto principal si necesitas seguir navegando
    driver.switch_to.default_content()
    
    #input("Resuelve el CAPTCHA manualmente y presiona ENTER para continuar...")
    time.sleep(3)

    ingresar_btn = wait.until(EC.element_to_be_clickable(
    (By.XPATH, "//button[contains(text(),'Ingresar')]")
    ))
    ingresar_btn.click()
    logging.info("🖱️ Clic en 'Ingresar' para Crecer Seguros.")
    time.sleep(8)

    if tipo_proceso == 'IN' :

        try:
            inclusion_crecer_vly(driver,wait,numero_poliza,ruta_archivos_x_inclu)
        except Exception as e:
            logging.error(f"❌ Error en Crecer (VL) - MA: {e}")

    else:

        try:
            renovacion_crecer_vly(driver,wait,numero_poliza,ruta_archivos_x_inclu)
        except Exception as e :
            logging.error(f"❌ Error en Crecer (VL) - MA: {e}")
     
def inclusion_crecer_vly(driver,wait,numero_poliza,ruta_archivos_x_inclu):

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
        input_fullname.send_keys(numero_poliza)
        logging.info(f"✅ Se ingresó la póliza: {numero_poliza}")
        
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
        wait.until(EC.visibility_of_element_located((
            By.XPATH,
            "//table[contains(@class, 'table-striped') and contains(@class, 'table-bordered')]//tbody/tr"
        )))

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

        ruta_trama_xls_sanitas_vl = f"{ruta_archivos_x_inclu}/{numero_poliza}.xls"
        input_file.send_keys(ruta_trama_xls_sanitas_vl)
        logging.info(f"✅ Trama y/o Excel subido para la póliza: '{numero_poliza}'")

        try:
            # Modal de Endoso Retroactivo
            boton_ok = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'OK')]")))
            driver.screenshot(os.path.join(ruta_archivos_x_inclu, f"error_subidaTramaCrecer_{numero_poliza}.png"))
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
        driver.save_screenshot(os.path.join(ruta_archivos_x_inclu, f"codigoPagoCrecer_{numero_poliza}.png"))

def renovacion_crecer_vly(driver,wait,numero_poliza,ruta_archivos_x_inclu):
    
    span_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Gestión de Cotización']")))
    logging.info("🔍 Ubicandose en 'Gestión de Cotización'")
    time.sleep(3)

    actions = ActionChains(driver)
    actions.move_to_element(span_element).click().perform()
    logging.info("🖱️ Clic en 'Gestión de Cotización'")

    # Esperar que el enlace esté disponible
    link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Renovación Vida Ley")))
    # Clic en el enlace
    link.click()
    logging.info("🖱️ Clic en 'Renovación Vida Ley'")

    time.sleep(3)
        
    input_fullname = wait.until(EC.visibility_of_element_located((By.ID, "spolizanumber")))
    input_fullname.clear()
    input_fullname.send_keys(numero_poliza)
    logging.info(f"✅ Se ingresó la póliza: {numero_poliza}")

    boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Busca Póliza')]")))
    boton.click()
    logging.info("🖱️ Clic en 'Busca Póliza")

    try:
        div_poliza= WebDriverWait(driver,8).until(EC.visibility_of_element_located((By.ID, "swal2-content")))
        texto_alerta_poliza = div_poliza.text
        logging.warning(f"⚠️ Mensaje de Advertencia encontrado: {texto_alerta_poliza}")
        driver.screenshot(os.path.join(ruta_archivos_x_inclu, f"{numero_poliza}_noExiste.png"))
        raise Exception(texto_alerta_poliza)    
    except TimeoutException:
        pass

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "overlay")))

    boton_subir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Cargar plantilla')]")))

    ruta_trama_crecer_vl_xlsx = f"{ruta_archivos_x_inclu}/{numero_poliza}.xlsx"

    intentos = 0
    max_intentos = 2

    while intentos < max_intentos:
        intentos += 1
        try:
            subido_correctamente = subir_trama(boton_subir,ruta_trama_crecer_vl_xlsx)
            if subido_correctamente:
                logging.info(f"✅ Trama subido para la póliza: '{numero_poliza}'")
                break
            else:
                logging.warning(f"⚠️ Intento {intentos}: No se logró subir la Trama. Reintentando...")
                time.sleep(2)
        except Exception as e:
            logging.error(f"❌ Error al intentar subir archivo: {e}")
            time.sleep(2)

    if intentos == max_intentos:
        msj = "No se logró subir la Trama después de varios intentos."
        logging.error(f"❌ {msj}")
        raise Exception(msj)

    wait.until(EC.invisibility_of_element_located((By.CLASS_NAME, "overlay")))

    try:
        # Modal de error al subir la trama
        div_alerta = WebDriverWait(driver,8).until(EC.visibility_of_element_located((By.ID, "swal2-content")))
        texto_alerta = div_alerta.text
        logging.warning(f"⚠️ Mensaje de Error encontrado: {texto_alerta}")
        driver.screenshot(os.path.join(ruta_archivos_x_inclu, f"errorTrama_{numero_poliza}.png"))
        raise Exception(f"❌ Error al subir la Trama en Crecer Vida Ley: {texto_alerta}.")
    except TimeoutException:
        pass

    boton_validar = wait.until(EC.visibility_of_element_located((By.XPATH, '//input[@value="Validar"]')))
    driver.execute_script("arguments[0].scrollIntoView(true);", boton_validar)
    boton_validar.click()
        
    time.sleep(5)

def procesar_solicitud_san_protecta_vl(driver,wait,poliza3,ruc_empresa,tipo_proceso,ruta_archivos_x_inclu):

    try:

        # ------------------ Inicio del Flujo de Automatización ------------------
        driver.get("https://plataformadigital.protectasecurity.pe/ecommerce/extranet/login")
        logging.info("---Ingresando a Protecta Seguros🌐---")

        # 2. Rellenar el formulario de login en Protecta
        user_input = wait.until(EC.presence_of_element_located((By.ID, "username")))
        user_input.clear()
        user_input.send_keys()
        logging.info("✅ Digitando el Username")
        
        pass_input = wait.until(EC.presence_of_element_located((By.ID, "password")))
        pass_input.clear()
        pass_input.send_keys()
        logging.info("✅ Digitando el Password")

        time.sleep(3)

        ingresar_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(text(),'Ingresar')]")
        ))
        ingresar_btn.click()
        logging.info("🖱️ Clic en 'Ingresar' en Protecta.")

        time.sleep(8)

        # Busca el <span> con el texto 'VIDA LEY' y haz clic
        span_element = driver.find_element(By.XPATH, "//span[text()='VIDA LEY']")
        span_element.click()

        # Espera a que el elemento sea clickeable y haz clic
        polizas_element = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Pólizas']"))
        )
        polizas_element.click()

        consulta_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[@title='Consulta de transacciones']"))
        )
        consulta_link.click()

        time.sleep(3)

        # Encuentra el elemento <select>
        dropdown = Select(driver.find_element(By.XPATH, "//select[@formcontrolname='typeTransac']"))

        if tipo_proceso == 'IN':
            dropdown.select_by_value("2")
        else:
            dropdown.select_by_value("4")

        time.sleep(2)

        select_element1 = wait.until(
            EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='branch']"))
        )

        dropdown = Select(select_element1)

        # Selecciona la opción por default
        dropdown.select_by_value("73") # -- > VIDA LEY

        time.sleep(2)

        select_element2 = wait.until(
            EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='product']"))
        )

        dropdown = Select(select_element2)

        # Selecciona la opción con value="1" Seguro Vida Ley Trabajadores
        dropdown.select_by_value("1") 

        time.sleep(2)

        input_element = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//input[@formcontrolname='policy']"))
        )

        input_element.clear()
        input_element.send_keys(poliza3) 

        time.sleep(2)

        select_element1 = wait.until(
            EC.presence_of_element_located((By.XPATH, "//select[@formcontrolname='contractorDocumentType']"))
        )

        dropdown = Select(select_element1)

        # Selecciona la opción con value="1" RUC
        dropdown.select_by_value("1")

        time.sleep(2)

        input_num_contra = wait.until(
            EC.visibility_of_element_located((By.XPATH, "//input[@formcontrolname='contractorDocumentNumber']"))
        )

        input_num_contra.clear()
        input_num_contra.send_keys(ruc_empresa)

        time.sleep(2)

        buscar_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Buscar']]"))
        )

        buscar_button.click()

        time.sleep(5)

        id_dinamico_poliza = f"policy{poliza3}"

        # Seleccionar la póliza haciendo clic en el radio button correspondiente
        radio_seleccion = wait.until(
            EC.element_to_be_clickable((By.ID,id_dinamico_poliza))
        )
        radio_seleccion.click()
        
        time.sleep(3)
        
        if tipo_proceso == 'IN':
            label_tipo_proceso = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Incluir')]"))
            )
        else:
            label_tipo_proceso  = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Renovar')]"))
            )

        # Clic en el label (esto activa el input radio asociado)
        label_tipo_proceso.click()
        
        # Esperar a que cargue la nueva página de adjuntar el archivo
        time.sleep(8)

        file_input = wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )

        # Hacer scroll hasta el input
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", file_input)

        ruta_trama_protecta = f"{ruta_archivos_x_inclu}/{poliza3}.xls"

        file_input.send_keys(ruta_trama_protecta)
        logging.info(f" ✔ Trama Macro de Protecta de la póliza: {poliza3} subida para validar' ")

        time.sleep(5)

        # Click en Validar
        boton_validar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Validar']]")))
        boton_validar.click()
        logging.info(f" ✔ Click en Validar Trama ")

        # Aparece un Modal para confirmar
        boton_ok = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "swal2-confirm")))
        boton_ok.click()
        logging.info(f" ✔ Click en Aceptar la Validación")

        # Paso 1: Ubicar el botón por el texto del <span>
        boton_procesar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[span[text()='PROCESAR']]")))

        # Paso 2: Hacer scroll hasta el botón
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", boton_procesar)

        # Paso 3: Esperar que sea clickeable y hacer clic
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='PROCESAR']]"))).click()

        #Hacer pruebas con inclusiones reales para ver como se descarga

        return True
    except:
        return False
