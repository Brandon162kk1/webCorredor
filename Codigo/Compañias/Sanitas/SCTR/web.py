#   --- Froms ----
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException,StaleElementReferenceException, WebDriverException
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
        link = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(),'{texto_link}')]")))
        link.click()

        logging.info(f"🖱️ Clic en {texto_link}")

        contract_number_input = wait.until(EC.element_to_be_clickable((By.ID, 'ContractNumber')))
        contract_number_input.send_keys(ramo.poliza)
        logging.info("⌨️ Digitando Póliza")

        search_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnSearchContracts')))
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
    texto = texto.replace("-", " ")
    texto = re.sub(r"\s+", " ", texto)  
    return texto.lower().strip()

def detectar_constancia_en_tabla(wait):

    tabla = wait.until(EC.presence_of_element_located((By.ID, "createdDocumentsTable")))
    filas = tabla.find_elements(By.TAG_NAME, "tr")

    encontrada_salud = False
    encontrada_pension = False

    for fila in filas:
        try:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if len(celdas) < 2:
                continue

            texto = normalizar(celdas[1].text)

            if "constancia conjunta" in texto:
                return "CONJUNTA"

            if "constancia salud" in texto:
                encontrada_salud = True

            if "constancia pension" in texto:
                encontrada_pension = True

        except:
            continue

    if encontrada_salud:
        return "SALUD"

    if encontrada_pension:
        return "PENSION"

    return "No hay constancia, solo Proformas"

def login_sanitas(driver,wait,list_url_san,list_polizas,tipo_mes,ruta_archivos_x_inclu,
                  tipo_proceso,palabra_clave,ba_codigo,ramo):

    encontrado = False
    constancia = False
    constancia_conjunta = False
    proforma = False
    tipoError = ""
    companiaTipoError = ""
    detalleError = ""
    error = False

    try:

        for i, url in enumerate(list_url_san):

            if i == 0:
                compania = "Crecer"
                logging.info(f"--------- Sanitas {compania} -----------")
            else:
                compania = "Protecta" 
                logging.info(f"------- Sanitas {compania} ----------")

            companiaTipoError = compania

            driver.get(url)
            logging.info(f"⌛ Cargando la Web de Sanitas {compania}")

            if es_error_502(driver):
                raise Exception("Pagina Fuera de Servicio")

            username_input = wait.until(EC.presence_of_element_located((By.ID, 'Login')))
            username_input.clear()
            username_input.send_keys(ramo.usuario)
            logging.info("⌨️ Digitando el Usuario")

            password_input = wait.until(EC.presence_of_element_located((By.ID, 'Password')))
            password_input.clear()
            password_input.send_keys(ramo.clave)
            logging.info("⌨️ Digitando la Contraseña")
    
            submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']")))
            submit_button.click()
            logging.info("🖱️ Clic en Iniciar Sesión")

            mensaje_error = (By.XPATH,"//div[contains(@class,'alert-error')]//li")
            afiliaciones = (By.XPATH, "//a[contains(text(),'Afiliaciones')]")

            resultado_login = wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(mensaje_error),
                    EC.element_to_be_clickable(afiliaciones)
                )
            )

            if resultado_login.tag_name == "li":
                mensaje = resultado_login.text.strip()
                raise Exception(mensaje)
            else:
                resultado_login.click()
                logging.info("🖱️ Clic en Afiliaciones")

            # pruebas_impresion(driver,wait,ramo)

            textos_proceso = {
                "IN": "Inclusión",
                "RE": "Renovación",
                "NE": "Neteo"
            }

            try:
                texto_link = textos_proceso[tipo_proceso]
            except KeyError:
                raise ValueError(f"Tipo de proceso no válido: {tipo_proceso}")

            link = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(),'{texto_link}')]")))
            link.click()
            logging.info(f"🖱️ Clic en {palabra_clave}")

            search_contract_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnSearchContract')))
            search_contract_button.click()
            logging.info("🖱️ Clic en Buscar")

            logging.info("⌛ Esperando el Modal")
            wait.until(EC.visibility_of_element_located((By.ID, 'searchContractPopup')))
            logging.info("✅ Modal cargado")

            contract_number_input = wait.until(EC.presence_of_element_located((By.ID, 'ContractNumber')))
            contract_number_input.send_keys(ramo.poliza)
            logging.info("⌨️ Digitando Póliza")

            search_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnSearchContracts')))
            search_button.click()
            logging.info("🖱️ Clic en Buscar")

            try:

                fila_correcta = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, f"//td[contains(text(),'{ramo.poliza}')]/parent::tr")))
                fila_correcta.click()
                logging.info(f"✅ Se seleccionó la fila para la póliza {ramo.poliza}")
                encontrado = True

            except:
                logging.warning(f"⚠️ No se encontraron la póliza '{ramo.poliza}'")
                tomar_capturar(driver,ruta_archivos_x_inclu,f"n_poliza")
                continue

            time.sleep(2)

            accept_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnSelectContract')))
            accept_button.click()
            logging.info("🖱️ Clic en Aceptar")

            if tipo_proceso == 'NE':

                sede_impresion_id = (By.ID, "PrintBranchOfficeId")

                def sede_69_disponible(driver):
                    try:
                        elemento = driver.find_element(*sede_impresion_id)
                        select = Select(elemento)

                        return any(
                            option.get_attribute("value") == "69"
                            for option in select.options
                        )

                    except StaleElementReferenceException:
                        return False

                wait.until(sede_69_disponible)

                select_element = wait.until(EC.element_to_be_clickable(sede_impresion_id))

                Select(select_element).select_by_value("69")

                logging.info("🖱️ Sede de impresión seleccionada: Lima - San Isidro")

                datos_movimiento_icono = (By.XPATH,"//a[contains(@href, '#collapseFive')]//span[contains(@class, 'collapse-icon')]")

                icono = wait.until(EC.element_to_be_clickable(datos_movimiento_icono))

                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});",icono)

                icono.click()

            #----------------------------------------------------------------------------------

            modal_advertencia = (By.ID, "MessageBox")
            client = (By.ID, "ClientBranchOfficeId")

            resultado_poliza = wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(modal_advertencia),
                    EC.element_to_be_clickable(client)
                )
            )

            if resultado_poliza is None:
                raise TimeoutException("No apareció el modal ni la sede del cliente")

            if resultado_poliza.get_attribute("id") == "MessageBox":
                logging.warning(f"⚠️ Apareció el modal con advertencia")
                span_txt_modal = resultado_poliza.find_element(By.ID, "message").text
                raise Exception(f"{span_txt_modal}")
            else:
                client_branch_office_select = wait.until(EC.element_to_be_clickable((client)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", client_branch_office_select)

            time.sleep(2)

            sede_buscada = normalizar(ramo.sede)
            sedes_disponibles = []

            def sede_disponible(driver):
                global sedes_disponibles
                try:
                    select_element = driver.find_element(*client)

                    # Activar el combo para que se muestren/carguen las opciones
                    if not select_element.is_displayed():
                        return False

                    if not select_element.is_enabled():
                        return False

                    select = Select(select_element)

                    opciones = select.options

                    sedes_disponibles = [
                        option.text.strip()
                        for option in opciones
                        if option.get_attribute("value")
                        and option.get_attribute("value") != "-1"
                        and option.text.strip()
                    ]

                    for option in opciones:
                        texto = option.text.strip()

                        if normalizar(texto) == sede_buscada:
                            return select_element

                    return False

                except StaleElementReferenceException:
                    return False

            try:

                select_element = wait.until(sede_disponible)
                select = Select(select_element)

                opcion_encontrada = next(
                    option
                    for option in select.options
                    if normalizar(option.text.strip()) == sede_buscada
                )

                select.select_by_value(opcion_encontrada.get_attribute("value"))

                logging.info(f"🖱️ Sede seleccionada: {opcion_encontrada.text.strip()}")

            except TimeoutException:
                raise Exception(f"No se encontró la sede solicitada: '{ramo.sede}'")

            # for option in select.options:                                       
            #     texto_opcion = option.text.strip()
            #     if normalizar(ramo.sede) == normalizar(texto_opcion):
            #         select.select_by_visible_text(texto_opcion)
            #         logging.info(f"🖱️ Opción seleccionada: {texto_opcion}")
            #         break                                                       
            # else:
            #     raise Exception("No se encontró ninguna Sede que coincida")

            time.sleep(2)

            if tipo_proceso == 'IN':

                input_fecha = (By.ID, "InsureFrom")

                resultado_0 = wait.until(
                        EC.any_of(
                            EC.visibility_of_element_located(modal_advertencia),
                            EC.visibility_of_element_located(input_fecha)
                        )
                    )

                if resultado_0.get_attribute("id") == "MessageBox":
                    span_txt_fecha = resultado_0.find_element(By.ID, "message").text
                    logging.warning(f"⚠️ Apareció el modal con advertencia: {span_txt_fecha}")

                    if span_txt_fecha == "No es posible realizar movimientos retroactivos superiores a 2 días.":
                        raise Exception(f"{span_txt_fecha}")

                    btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "CloseModal")))
                    btn_aceptar.click()
                    logging.info("🖱️ Clic en 'Aceptar'")

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

                    id_planilla = (By.ID, "btnCargarPlanilla")

                    resultado_vigencia = wait.until(
                        EC.any_of(
                            EC.visibility_of_element_located(modal_advertencia),
                            EC.invisibility_of_element_located(id_planilla)
                        )
                    )

                    if resultado_vigencia.get_attribute("id") == "MessageBox":
                        span_txt_fecha = resultado_vigencia.find_element(By.ID, "message").text
                        logging.warning(f"⚠️ Apareció el modal con advertencia: {span_txt_fecha}")

                        if span_txt_fecha == "No es posible realizar movimientos retroactivos superiores a 2 días.":
                            raise Exception(f"{span_txt_fecha}")

                        btn_aceptar = wait.until(EC.element_to_be_clickable((By.ID, "CloseModal")))
                        btn_aceptar.click()
                        logging.info("🖱️ Clic en 'Aceptar'")
 
                    else:
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

            buscar_span = wait.until(EC.presence_of_element_located((By.NAME, "file")))

            ruta_trama_xls_sanitas = f"{ruta_archivos_x_inclu}/{ramo.poliza}_97.xls"

            if not os.path.exists(ruta_trama_xls_sanitas):
                raise FileNotFoundError(f"La ruta {ruta_trama_xls_sanitas} no existe")

            buscar_span.send_keys(ruta_trama_xls_sanitas)
            logging.info(f"✅ Trama {ramo.poliza}_97.xls adjuntada")

            btn_subir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Subir Archivo')]")))
            btn_subir.click()
            logging.info("🖱️ Clic en 'Subir Archivo'")

            modal_error_trama = (By.ID, "MessageBoxWithScroll")

            resultado_validar_trama = wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(modal_error_trama),
                    EC.visibility_of_element_located(modal_advertencia),
                    # Tabla de inclusiones con al menos un registro
                    lambda d: (
                        len(d.find_elements(By.CSS_SELECTOR, "#inclusionsTable tbody tr")) > 0
                        and
                        not d.find_elements(By.CSS_SELECTOR, "#inclusionsTable td.dataTables_empty")
                    ),
                    # Tabla de renovaciones con registros reales
                    lambda d: (
                        len(d.find_elements(
                            By.CSS_SELECTOR, "#renovationsTable tbody tr"
                        )) > 0
                        and not d.find_elements(
                            By.CSS_SELECTOR, "#renovationsTable td.dataTables_empty"
                        )
                    ),
                    # Al menos una tabla de neteos con registros reales
                    lambda d: any(
                        len(filas := d.find_elements(
                            By.CSS_SELECTOR, f"#{tabla} tbody tr"
                        )) > 0
                        and not d.find_elements(
                            By.CSS_SELECTOR, f"#{tabla} td.dataTables_empty"
                        )
                        for tabla in (
                            "neteosTotalsTableAI",
                            "neteosTotalsTableAE",
                            "neteosTotalsTableBI",
                            "neteosTotalsTableBE"
                        )
                    )
                )
            )

            if resultado_validar_trama is True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                logging.info("🖱️ Scroll hasta abajo de la página")
            else:

                modal_id = resultado_validar_trama.get_attribute("id")

                if modal_id == "MessageBox":
                    logging.warning("⚠️ Apareció el modal con advertencia")
                    span_txt_modal = resultado_validar_trama.find_element(By.ID, "message").text
                    raise Exception(span_txt_modal)

                elif modal_id == "MessageBoxWithScroll":
                    logging.warning("⚠️ Apareció el modal con errores de la Trama")
                    span_txtArea_modal = resultado_validar_trama.find_element(By.ID, "message").text
                    raise Exception(span_txtArea_modal)

            confirmar_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnConfirm')))
            confirmar_button.click()
            logging.info("🖱️ Clic en Confirmar")

            modal_retroactividad = (By.ID, "ConfirmationMessageBox")
            modal_documentos = (By.ID, 'createdDocumentsPopUp')

            resultado_confirmar = wait.until(
                EC.any_of(
                    EC.visibility_of_element_located(modal_advertencia),
                    EC.visibility_of_element_located(modal_retroactividad),
                    EC.visibility_of_element_located(modal_documentos)
                )
            )

            elemento_id = resultado_confirmar.get_attribute("id")

            if elemento_id == "MessageBox":
                span_txt_modal = resultado_confirmar.find_element(By.ID, "message").text
                logging.warning(f"⚠️ Apareció el modal con advertencia")
                raise Exception(f"{span_txt_modal}")

            elif elemento_id == "ConfirmationMessageBox":
                logging.warning(f"⚠️ Apareció el modal con advertencia de retroactividad")
                btn_si = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, \"data-dialog-response', 'YES'\")]")))
                btn_si.click()
                logging.info("🖱️ Clic en el botón 'Sí'")
                wait.until(EC.visibility_of_element_located(modal_documentos))
                logging.info("✅ Modal de documentos cargados")

            else:
                logging.info("✅ Modal de los documentos cargados")

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
                    #"alias": ramo.poliza,
                    "alias": list_polizas[0],
                    "tittle": "Descargar PDF",
                    "impresion": False
                })
                constancia = True

            elif tipo_constancia == "PENSION":
                alias = ramo.poliza if ba_codigo == 2 else list_polizas[1]
                documentos_a_descargar.append({
                    "nombreDoc": "Constancia Pensión - Crecer S.A.",
                    #"alias": ramo.poliza,
                    "alias": alias,
                    "tittle": "Descargar PDF",
                    "impresion": False
                })
                constancia = True

            else:
                logging.warning("⚠️ No se encontró ninguna constancia para descargar")
                constancia = False
                tipoError = "Solo existe Proformas"
                detalleError = "Primero pagar para obtener constancias"

            if tipo_mes != "MV" and tipo_proceso != 'NE':

                proforma = True

                if len(list_polizas) == 2:

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

                elif ba_codigo == '1':
                    documentos_a_descargar.append({
                        "nombreDoc": "Proforma - Sanitas Perú S.A.",
                        "alias": f"endoso_{ramo.poliza}",
                        "tittle": "Imprimir",
                        "impresion": True
                    })
                elif ba_codigo == '2':
                    documentos_a_descargar.append({
                        "nombreDoc":
                            "Aviso de cobranza - Protecta S.A." if i == 1
                            else "Aviso de cobranza - Crecer Seguros S.A.",
                        "alias": f"endoso_{ramo.poliza}",
                        "tittle": "Imprimir",
                        "impresion": True
                    })

            logging.info(f"📂 Documentos a descargar: {documentos_a_descargar}")

            for doc in documentos_a_descargar:                                                          

                nombre_documento = doc["nombreDoc"]
                alias_archivo = doc["alias"]
                title_documento = doc["tittle"]
                impresion = doc["impresion"]

                filas = driver.find_elements(By.XPATH, "//table[@id='createdDocumentsTable']//tr")
                fila_objetivo = None

                for i, fila in enumerate(filas, 1):                       

                    try:
                        celdas = fila.find_elements(By.XPATH, "./td")
                        if len(celdas) < 2:
                            continue

                        texto_celda = celdas[1].text.strip()
                        logging.info(f"-----------")
                        logging.info(f"🔍 Fila {i}: '{texto_celda}'")

                        if normalizar(nombre_documento) in normalizar(texto_celda):
                            fila_objetivo = fila
                            break                                           
                    except Exception as e:
                        logging.warning(f"⚠️ Fila {i} error: {str(e)}")

                if not fila_objetivo:
                    logging.warning(f"❌ No se encontró la fila para: {nombre_documento}")
                    continue

                try:

                    descargar_link = fila_objetivo.find_element(By.XPATH, f".//a[@title='{title_documento}']")

                    driver.execute_script("arguments[0].scrollIntoView(true);", descargar_link)

                    archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

                    driver.execute_script("arguments[0].click();", descargar_link)
                    logging.info("🖱 Clic con JS en el botón de descarga")

                    archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=1)

                    if archivo_nuevo:
                        logging.info(f"✅ Documento '{nombre_documento}' descargado exitosamente")
                        ruta_original = archivo_nuevo[0]
                        ruta_final = os.path.join(ruta_archivos_x_inclu, f"{alias_archivo}.pdf")
                        os.rename(ruta_original, ruta_final)
                        logging.info(f"🔄 {nombre_documento} renombrado a '{alias_archivo}.pdf'")
                    else:
                        raise Exception(f"No se descargo ningun archivo")

                    if impresion:
                        try:
                            aceptar_btn = WebDriverWait(driver,8).until(EC.element_to_be_clickable((By.ID, "CloseModal")))
                            aceptar_btn.click()
                            logging.info(f"🖱️ Clic en 'Aceptar' para cerrar el Modal")
                        except TimeoutException:
                            pass

                except Exception as e:
                    logging.error(f"❌ Error al intentar descargar: {nombre_documento} | Detalle: {e}")
                    continue
                
            time.sleep(3)

            if encontrado:

                # Si hay constancia conjunta, se guardo con nombre de salud/pensión y hay que renombrarla para pension/salud
                if constancia_conjunta:

                    ruta_salud = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[0]}.pdf")
                    ruta_pension = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[1]}.pdf")

                    if os.path.exists(ruta_pension):
                        shutil.copy2(ruta_pension,ruta_salud)
                        logging.info(f"📄 Copia creada como '{list_polizas[0]}.pdf'")

                else:
                    logging.warning(f"⚠️ No hay constancia conjunta, solo {tipo_constancia.lower()}, ya que la otra póliza no paso la primeta neta en la emisión")

                break

        # Cierre del navegador o continúa según resultado
        if not encontrado:
            raise Exception(f"No se encontró la póliza '{ramo.poliza}' en ninguna de las compañías")
        else:
            logging.info("-----------------")
            logging.info(f"✅ {palabra_clave} en SANITAS realizada exitosamente")

    except WebDriverException as e:
        logging.exception(f"⚠️ Error técnico de Selenium: {str(e)}")
        detalleError = "Problemas Técnicos del Agente"
        error = True
    except Exception as e:
        detalleError = str(e)
        error = True
    finally:

        if error:

            try:
                tomar_capturar(driver,ruta_archivos_x_inclu,f"ERROR_SCTR_{tipo_mes}")
            except Exception:
                logging.exception("⚠️ No se pudo tomar screenshot")

            logging.error(f"❌ Error en Sanitas {companiaTipoError} (SCTR) - {tipo_mes}: {detalleError}")
            tipoError = f"{companiaTipoError.upper()[:4]}-SCTR-{tipo_mes}"
            return constancia,proforma,tipoError,detalleError
        else:
            #return True,True,tipoError,detalleError
            return constancia,proforma,tipoError,detalleError
