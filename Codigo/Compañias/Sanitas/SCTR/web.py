#   --- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from Tiempo.fechas_horas import get_timestamp,get_fecha_hoy
from LinuxDebian.Ventana.ventana import esperar_archivos_nuevos
from selenium.webdriver import ActionChains
from LinuxDebian.OtrosMetodos.metodos import esperar_ventana
from Compañias.Sanitas.metodos import es_error_502
from Chrome.google import tomar_capturar
#   --- Imports ----
import time
import os
import logging
from datetime import datetime
import unicodedata
import re
import shutil
import subprocess

def pruebas_impresion(driver,wait,ramo):
    
    try:
        texto_link = 'Contratos'
        link = wait.until(EC.presence_of_element_located((By.XPATH, f"//a[contains(text(),'{texto_link}')]")))
        link.click()

        logging.info(f"🖱️ Clic en {texto_link}")

        contract_number_input = wait.until(EC.presence_of_element_located((By.ID, 'ContractNumber')))
        contract_number_input.send_keys(ramo.poliza)
        logging.info("⌨️ Digitando Póliza")

        search_button = wait.until(EC.presence_of_element_located((By.ID, 'btnSearchContracts')))
        search_button.click()
        logging.info("🖱️ Clic en Filtrar")

        try:
            wait.until(
                lambda d: len(
                    d.find_elements(
                        By.XPATH,
                        "//table[@id='searchContractTable']/tbody/tr[td]"
                    )
                ) > 0
            )
            logging.info("✅ La tabla tiene al menos 1 fila válida.")
        except TimeoutException:
            logging.info("⏰ No se cargaron las filas en el tiempo esperado.")

        rows = driver.find_elements(
            By.XPATH,
            "//table[@id='searchContractTable']/tbody/tr[td]"
        )
 
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            logging.info(f"Columnas encontradas: {len(cells)}")

            if len(cells) >= 4:
                movimiento = cells[3].text.strip()
                logging.info(f"El movimiento es : {movimiento}")

                if movimiento == 'Inclusión':

                    celda_accion = cells[-1]
                    icono = celda_accion.find_element(By.XPATH, ".//a[contains(@class, 'dropdown-toggle')]")
                    driver.execute_script("arguments[0].scrollIntoView(true);", icono)
                    ActionChains(driver).move_to_element(icono).click().perform()
                    logging.info("🖱️ Icono desplegable cliceado.")
                    time.sleep(2)  # espera a que aparezca la ventana

                    # 2. Forzar visualización del menú
                    menu = celda_accion.find_element(By.XPATH, ".//ul[contains(@class, 'dropdown-menu')]")
                    driver.execute_script("arguments[0].style.display = 'block';", menu)
                    logging.info("✅ Menú desplegable forzado a visible.")
                    time.sleep(1)

                    #---Haciendo dinamico el encontrar el ID
                    # Encuentra TODOS los enlaces <a> en el menú desplegable
                    links = menu.find_elements(By.TAG_NAME, "a")
                    boton_imprimir = None

                    for link in links:
                        texto = link.text.strip()
                        titulo = link.get_attribute("title")
                        id_link = link.get_attribute("id")

                        logging.info(f"Probando link: texto='{texto}', title='{titulo}', id='{id_link}'")

                        if "Imprimir con prefactura" in texto or "Imprimir con prefactura" in (titulo or ""):
                            boton_imprimir = link
                            break

                    if boton_imprimir is None:
                        logging.error("❌ No se encontró el botón 'Imprimir con prefactura'")
                        raise Exception("No se encontró Imprimir con prefactura")

                    wait.until(EC.visibility_of(boton_imprimir))
                    driver.execute_script("arguments[0].click();", boton_imprimir)

                    logging.info("✅ Se hizo clic en 'Imprimir con prefactura'")

                    time.sleep(3)
                    subprocess.run(["xdotool", "key", "Return"])
                    logging.info(f"🖱️ Clic en 'Save'.")
                    time.sleep(3)

                    logging.info("⌛ Esperando la ventana 'Open File' de Linux Debian")

                    ventana_id = esperar_ventana("Save File")
                    if not ventana_id:
                        raise Exception("No apareció la ventana 'Open File'")

                    ids = subprocess.check_output(
                        ["xdotool", "search", "--onlyvisible", "--name", "."]
                    ).decode().split()

                    for win in ids:
                        name = subprocess.check_output(
                            ["xdotool", "getwindowname", win]
                        ).decode().strip()
                        logging.info(str(win) + " -> " + name)

                    logging.info(f"Id Ventana: {ventana_id}")
                    # Activar la ventana específica
                    subprocess.run(["xdotool", "windowactivate", "--sync", ventana_id])
                    logging.info("💡 Se hizo FOCO en la nueva ventana de dialogo de Linux Debian")
                    time.sleep(0.5)

                    subprocess.run(["xdotool", "type","--delay", "100", "pancho"])
                    logging.info("📄 Se escribió el nombre del archivo")
                    time.sleep(2)
                    subprocess.run(["xdotool", "key", "--clearmodifiers", "Alt+o"])
                    time.sleep(2)
                    subprocess.run(["xdotool", "key", "Return"])
                    logging.info("🖱️ Se presionó Enter para confirmar la descarga.")

                    break

    except Exception as e :
        logging.error(f"El error es : {e}")
    finally:
        input("Esperar")

def normalizar(texto):

    if not texto:
        return ""

    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = re.sub(r"\s+", " ", texto)  # colapsar espacios
    return texto.lower().strip()

def detectar_constancia_en_tabla(wait):

    #filas = driver.find_elements(By.XPATH, "//table[@id='createdDocumentsTable']//tr")
    tabla = wait.until(EC.presence_of_element_located((By.ID, "createdDocumentsTable")))
    filas = tabla.find_elements(By.TAG_NAME, "tr")

    #filas = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table[@id='createdDocumentsTable']//tr")))

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
                               tipo_proceso,palabra_clave,ba_codigo,ramo):

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
                logging.info(f"--------- Sanitas {compania} -----------")
            else:
                compania = "Protecta" 
                logging.info(f"------- Sanitas {compania} ----------")
                        
            driver.get(url)
            logging.info(f"⌛ Cargando la Web de Sanitas {compania}")

            if es_error_502(driver):
                raise("Pagina Fuera de Servicio")

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

            #pruebas_impresion(driver,wait,ramo)

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
            contract_number_input.send_keys(ramo.poliza)
            logging.info("⌨️ Digitando Póliza")

            search_button = wait.until(EC.presence_of_element_located((By.ID, 'btnSearchContracts')))
            search_button.click()
            logging.info("🖱️ Clic en Buscar")

            try:               
                filas = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH, f"//td[contains(text(),'{ramo.poliza}')]/parent::tr")))
                fila_correcta = filas[0]
                fila_correcta.click()
                logging.info(f"✅ Se seleccionó la fila para la póliza {ramo.poliza}")
                encontrado = True
            except:
                logging.error(f"❌ No se encontraron resultados para el número de contrato {ramo.poliza}")
                continue

            time.sleep(2)
            
            accept_button = wait.until(EC.presence_of_element_located((By.ID, 'btnSelectContract')))
            accept_button.click()
            logging.info("🖱️ Clic en Aceptar")

            try:
                modal_poliza = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "MessageBox")))                  
                span_txt_modal = modal_poliza.find_element(By.ID, "message").text
                logging.warning(f"⚠️ Apareció el modal con advertencia: {span_txt_modal}.")
                raise Exception(f"{span_txt_modal}")
            except TimeoutException:
                pass

            client_branch_office_select = wait.until(EC.presence_of_element_located((By.ID, 'ClientBranchOfficeId')))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", client_branch_office_select)

            time.sleep(2)
    
            select = Select(client_branch_office_select)

            for option in select.options:                                       #----- for 1
                texto_opcion = option.text.strip()
                if normalizar(ramo.sede) == normalizar(texto_opcion):
                    select.select_by_visible_text(texto_opcion)
                    logging.info(f"🖱️ Opción seleccionada: {texto_opcion}")
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
                        logging.info(f"📅 Vigencia de Inicio retroactiva ingresada: {ramo.f_inicio} con {diferencia} {'día' if diferencia == 1 else 'días'} de diferencia")
                    else:
                        logging.info(f"📅 Vigencia de Inicio ingresada: {ramo.f_inicio}")

                    try:
                        # Modal con Advertencias de la fecha
                        modal_fecha = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID,"MessageBox")))
                        span_txt_fecha = modal_fecha.find_element(By.ID, "message").text
                        logging.warning(f"⚠️ Apareció el modal con advertencia: {span_txt_fecha}")

                        if span_txt_fecha == "No es posible realizar movimientos retroactivos superiores a 2 días.":
                            raise Exception(f"{span_txt_fecha}")

                        btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "CloseModal")))
                        btn_aceptar.click()
                        logging.info("🖱️ Clic en 'Aceptar'")

                    except TimeoutException:
                        pass

                else:
                    logging.info("📅 La fecha de vigencia de inicio no requiere ser modificada")

            if tipo_proceso == 'RE':
                select_element = wait.until(EC.presence_of_element_located((By.ID, 'OfficeProductionId')))
                select = Select(select_element)
                select.select_by_value("2")
                logging.info("🖱️ Clic en 'Lima'")

            cargar_planilla_button = wait.until(EC.presence_of_element_located((By.ID, "btnCargarPlanilla")))
            cargar_planilla_button.click()
            logging.info("🖱️ Clic en Cargar Planilla")

            wait.until(EC.visibility_of_element_located((By.ID, "uploadAffiliationsFileUpload")))
            logging.info("⌛ Esperando que aparezca Modal")

            buscar_span = wait.until(EC.presence_of_element_located((By.NAME, "file")))
            #buscar_span = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'btn-file') and contains(., 'Buscar')]")))

            ruta_trama_xls_sanitas = f"{ruta_archivos_x_inclu}/{ramo.poliza}_97.xls"

            if not os.path.exists(ruta_trama_xls_sanitas):
                raise FileNotFoundError(f"La ruta {ruta_trama_xls_sanitas} no existe")

            buscar_span.send_keys(ruta_trama_xls_sanitas)
            logging.info(f"✅ Trama {ramo.poliza}_97.xls adjuntada")

            # if subir_trama(buscar_span,ruta_trama_xls_sanitas):               
            #     time.sleep(3)
            #     logging.info(f"✅ Trama {ramo.poliza}.xls adjuntada")
            # else:
            #     raise Exception(f"No se pudo subir la trama {ramo.poliza}.xls")

            btn_subir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Subir Archivo')]")))
            btn_subir.click()
            logging.info("🖱️ Clic en 'Subir Archivo' para validar la trama en la compañía")

            time.sleep(3)

            try:
                # Modal con Errores en la Trama
                modal_error_trama = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "MessageBoxWithScroll")))             
                span_txtArea_modal = modal_error_trama.find_element(By.ID, "message").text
                logging.warning(f"⚠️ Apareció el modal con errores de la Trama")
                raise Exception(f"{span_txtArea_modal}")
            except TimeoutException:
                pass

            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            logging.info("🖱️ Scroll hasta abajo de la página")

            confirmar_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnConfirm')))
            
            confirmar_button.click()
            logging.info("🖱️ Clic en Confirmar")

            try:
                modal_advertencia = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "MessageBox")))
                spana_dv_txt_modal = modal_advertencia.find_element(By.ID, "message").text
                logging.warning(f"⚠️ Apareció el modal con advertencia")
                raise Exception(f"{spana_dv_txt_modal}")
            except TimeoutException:
                pass

            try:
                # Modal de Retroactividad para autorizar no siniestros
                modal_advertencia = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.ID, "ConfirmationMessageBox")))
                logging.warning(f"⚠️ Apareció el modal con advertencia de retroactividad")
                btn_si = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, \"data-dialog-response', 'YES'\")]")))
                btn_si.click()
                logging.info("🖱️ Se hizo clic en el botón 'Sí'.")
            except TimeoutException:
                pass
    
            wait.until(EC.visibility_of_element_located((By.ID, 'createdDocumentsPopUp')))
            logging.info("⌛ Cargando Modal de los documentos")

            modal_body = driver.find_element(By.CLASS_NAME, 'dataTables_scrollBody')
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", modal_body)

            ruta_imagen = os.path.join(ruta_archivos_x_inclu,f"{get_timestamp()}_verDocumentos.png")
            driver.save_screenshot(ruta_imagen)

            tipo_constancia = detectar_constancia_en_tabla(wait)
            documentos_a_descargar = []

            logging.info(f"🔎 Tipo de constancia detectada en la tabla: {tipo_constancia}")

            constancia_conjunta = False

            if tipo_constancia == "CONJUNTA":
                documentos_a_descargar.append({
                    "nombreDoc": "Constancia conjunta",
                    "alias": ramo.poliza,
                    "tittle": "Descargar PDF",
                    "impresion": False
                })
                constancia_conjunta = True
                constancia = True

            elif tipo_constancia == "SALUD":
                documentos_a_descargar.append({
                    "nombreDoc": "Constancia Salud - Sanitas Perú S.A.",
                    "alias": ramo.poliza,
                    "tittle": "Descargar PDF",
                    "impresion": False
                })
                constancia = True

            elif tipo_constancia == "PENSION":
                documentos_a_descargar.append({
                    "nombreDoc": "Constancia Pensión - Crecer S.A.",
                    "alias": ramo.poliza,
                    "tittle": "Descargar PDF",
                    "impresion": False
                })
                constancia = True

            else:
                logging.warning("⚠️ No se encontró ninguna constancia para descargar")
                constancia = False
                tipoError = "Solo existe Proformas"
                detalleError = "Primero pagar para obtener constancias"

            if tipo_mes != "MV":

                proforma = True

                if len(list_polizas) == 2:

                    #proforma = True

                    documentos_a_descargar.append({
                        "nombreDoc": "Proforma - Sanitas Perú S.A.",
                        "alias": f"endoso_{list_polizas[0]}", 
                        "tittle": "Imprimir",
                        "impresion": True
                    })

                    documentos_a_descargar.append({
                        "nombreDoc":
                            "Aviso de cobranza - Protecta S.A." if i == 1
                            else "Aviso de cobranza - Crecer Seguros S.A.",
                        "alias": f"endoso_{list_polizas[1]}", 
                        "tittle": "Imprimir",
                        "impresion": True
                    })

                elif ba_codigo == '1': # Solo Salud
                    documentos_a_descargar.append({
                        "nombreDoc": "Proforma - Sanitas Perú S.A.",
                        "alias": f"endoso_{ramo.poliza}",
                        "tittle": "Imprimir",
                        "impresion": True
                    })
                elif ba_codigo == '2': # Solo Pensión
                    documentos_a_descargar.append({
                        "nombreDoc":
                            "Aviso de cobranza - Protecta S.A." if i == 1
                            else "Aviso de cobranza - Crecer Seguros S.A.",
                        "alias": f"endoso_{ramo.poliza}",
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

                    # Guardar archivos antes del clic
                    archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

                    driver.execute_script("arguments[0].click();", descargar_link)
                    logging.info("🖱 Clic con JS en el botón de descarga")

                    archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=1)
                    logging.info(f"✅ Documento '{nombre_documento}' descargado exitosamente")

                    if archivo_nuevo:
                        ruta_original = archivo_nuevo[0]
                        ruta_final = os.path.join(ruta_archivos_x_inclu, f"{alias_archivo}.pdf")
                        os.rename(ruta_original, ruta_final)
                        logging.info(f"🔄 {nombre_documento} renombrado a '{alias_archivo}.pdf'")
                    else:
                        logging.error("❌ No se encontró archivo nuevo después de descargar")

                    if impresion:
                        try:
                            time.sleep(3)
                            aceptar_btn = WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.ID, "CloseModal")))
                            aceptar_btn.click()
                            logging.info(f"🖱️ Clic en 'Aceptar' para cerrar el Modal")
                        except TimeoutException:
                            pass

                    # #if click_descarga_documento(driver,descargar_link,alias_archivo,impresion):
                    # if descargar_documento(driver,descargar_link,alias_archivo,impresion,pestaña=False):
                    #     time.sleep(3)
                    #     ruta_doc_descargado = f"{ruta_archivos_x_inclu}/{alias_archivo}.pdf"
                    #     if os.path.exists(ruta_doc_descargado):
                    #         logging.info(f"✅ {alias_archivo}.pdf descargado exitosamente")
                    #     else:
                    #         logging.warning(f"⚠️ No se encontró el archivo descargado en {ruta_archivos_x_inclu}")
                    # else:
                    #     logging.info(f"❌ No se logró descargar {alias_archivo}, Verifica manualmente.")

                except Exception as e:
                    logging.error(f"❌ Error al intentar descargar: {nombre_documento} ,Detalle: {e}")
                    continue
                
            time.sleep(3)

            if encontrado:

                # Si hay constancia conjunta, se guardo con nombre de salud y hay que renombrarla para pension
                if constancia_conjunta:

                    ruta_salud = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[0]}.pdf")

                    # Nuevo nombre para Pensión
                    ruta_pension = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[1]}.pdf")

                    if os.path.exists(ruta_salud):
                        shutil.copy2(ruta_salud,ruta_pension)
                        logging.info(f"📄 Copia creada como '{list_polizas[1]}.pdf'")
                    else:
                        logging.error(f"❌ No existe el archivo base: {ruta_salud}")
                else:
                    logging.info("No hay constancia conjunta, o salud o pension, pero verificar cual falta para volver a descargar")

                break

        # Cierre del navegador o continúa según resultado
        if not encontrado:
            raise Exception(f"No se encontró la póliza '{ramo.poliza}' en ninguna de las compañías")
        else:
            logging.info("-----------")
            logging.info(f"✅ {palabra_clave} en SANITAS realizada exitosamente.")

        time.sleep(3) 

    except Exception as e:

        logging.error(f"❌ Error en Sanitas (SCTR) - {tipo_mes}: {e}")
        tipoError = f"SANI-SCTR-{tipo_mes}"
        detalleError = str(e)
        tomar_capturar(driver,ruta_archivos_x_inclu,f"ERROR_SCTR_{tipo_mes}")
    finally:
        return constancia,proforma,tipoError,detalleError