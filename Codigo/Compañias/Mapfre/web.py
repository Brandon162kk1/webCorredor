#--- Froms ---
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from Tiempo.fechas_horas import tipo_vigencia
from LinuxDebian.ventana import esperar_archivos_nuevos
#---- Import ---
import os
import logging
import time
import shutil
import pandas as pd

def solicitud_sctr_vl(driver,wait,palabra_clave,tipo_proceso,ruta_archivos_x_inclu,ba_codigo,bab_codigo,list_polizas,
                                     ejecutivo_responsable,nombre_cliente,tipo_mes,ramo):


    ramo_opcion = "Vida Ley" if bab_codigo == '4' else "SCTR General"
    opcion = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[a/div[normalize-space()='{ramo_opcion}']]/a")))
    driver.execute_script("arguments[0].click();", opcion)
    logging.info(f"🖱️ Clic en '{ramo_opcion}'")

    try:

        label = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-label[normalize-space()='Nro. Póliza']")))
        driver.execute_script("arguments[0].click();", label)
        time.sleep(2)
        logging.info("🖱️ Clic en el # de Póliza")

        input_poliza = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='nNumPolizaFilter']")))
        input_poliza.clear()
        input_poliza.send_keys(ramo.poliza)
        logging.info(f"⌨️ Digitando {ramo.poliza}")

    except Exception as e:
        logging.info(f"❌ No se pudo escribir en el input: {e}")

    boton_filtrar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Filtrar']")))
    driver.execute_script("arguments[0].click();", boton_filtrar)
    logging.info("🖱️ Clic en botón 'Filtrar'")           

    containers = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.g-myd-result")))
    logging.info(f"✅ Cantidad de contenedores encontrados: {len(containers)}")

    encontrado = False

    for cont in containers:
        
        elementos = cont.find_elements(By.XPATH,f".//div[contains(@class,'item-dato') and normalize-space()='3322 - BIRLIK CORREDORES DE SEGUROS SAC']")

        if elementos:
            logging.info("✅ Agente Birlik encontrado dentro de un contenedor")

            boton = cont.find_element(By.XPATH,".//span[normalize-space()='Seleccionar']/parent::a")
            driver.execute_script("arguments[0].click();", boton)
            logging.info("🖱️ Clic en 'Seleccionar'")
            encontrado = True
            break

    if not encontrado:
        raise Exception ("No se encontro el agente 'BIRLIK CORREDORES DE SEGUROS SAC'")

    span_expand = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class,'mat-expansion-indicator')]")))
    driver.execute_script("arguments[0].click();", span_expand)
    logging.info("🖱️ Clic en 'Detalle'")

    time.sleep(2)

    driver.execute_script("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});")

    time.sleep(2)

    logging.info("-------------------------")

    bloques = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//div[contains(@class,'h-myd-bg--gray4')]")))
    logging.info(f"🔍 Bloques encontrados: {len(bloques)}")

    # bloques = driver.find_elements(By.XPATH, "//div[contains(@class,'h-myd-bg--gray4')]")
    # logging.info(f"🔍 Bloques encontrados: {len(bloques)}")

    bloque_correcto = None

    # --- BUSCAR EL BLOQUE QUE TIENE LA FECHA CORRECTA ---
    for indice, bloque in enumerate(bloques):
        try:
            # Primero intentamos buscar formato SCTR
            elementos_p = bloque.find_elements(By.XPATH, ".//p[contains(text(),'Desde')]")

            if elementos_p:
                texto_vigencia = elementos_p[0].text
                ramo_text = "SCTR"
            else:
                # Si no existe <p>, buscamos formato VL
                elementos_div = bloque.find_elements(By.XPATH, ".//div[contains(text(),'Desde')]")
            
                if elementos_div:
                    texto_vigencia = elementos_div[0].text
                    ramo_text = "VL"
                else:
                    #logging.error(f"No se encontró vigencia en bloque {indice}")
                    continue

            #logging.info(f"Bloque {indice} ({ramo}) → {texto_vigencia}")

            if ramo.f_fin in texto_vigencia:
                logging.info(f"✅ Bloque correcto encontrado ({ramo_text})")
                bloque_correcto = bloque
                
                parte_desde, parte_hasta = texto_vigencia.split(" - ")

                fecha_inicio_web = parte_desde.replace("Desde:", "").strip()
                fecha_fin_web = parte_hasta.replace("Hasta:", "").strip()

                modalidad = tipo_vigencia(fecha_inicio_web,fecha_fin_web)
                logging.info(f"📅 Fecha Inicio: {fecha_inicio_web} -- Fecha Fin : {fecha_fin_web} , Modalidad : {modalidad}")
                break

        except Exception as e:
            logging.error(f"Error en bloque {indice}: {e}")
            continue

    if bloque_correcto is None:
        raise Exception("No se encontró un bloque con la Fechas de Vigencia indicada")
    else:

        logging.info("⌛ Analizando dentro del bloque correcto")

        # Dentro del bloque: obtener los UL de SALUD y PENSIÓN
        listas = bloque_correcto.find_elements(By.XPATH, ".//ul[contains(@class,'g-row-fz12')]")

        for ul in listas:
            # Leer el label (SALUD o PENSIÓN)
            try:
                label = " ".join(ul.find_element(By.XPATH, ".//b").text.split()).upper()
                logging.info(f"✅ Tipo detectado: {label}")
            except:
                continue

            # Obtener checkbox
            checkbox = ul.find_element(By.XPATH, ".//input[@type='checkbox']")

            # ------ LÓGICA PARA MARCAR SEGÚN ba_codigo ------
            marcar = True #False           
            if marcar:
                driver.execute_script("arguments[0].click();", checkbox)
                logging.info(f"✅ Checkbox marcado para: {label}")
                break

    time.sleep(5)

    # Buscar el contenedor completo según el texto de la póliza
    div_objetivo = wait.until(EC.element_to_be_clickable((By.XPATH,f"//div[contains(@class,'item-dato')][normalize-space()='{ramo.poliza}']/ancestor::div[contains(@class,'cnt-item')]")))
    driver.execute_script("arguments[0].click();", div_objetivo)
    logging.info(f"🖱️ Clic en el 'div' de la póliza: {ramo.poliza}")

    text = "Incluir" if tipo_proceso == "IN" else "Declarar"

    #Hacer Click a Declarar
    boton_declarar = wait.until(EC.element_to_be_clickable((By.XPATH, f"//a[normalize-space()='{text}']")))
    driver.execute_script("arguments[0].click();", boton_declarar)
    logging.info(f"🖱️ Clic en botón '{text}'")

    file_path = os.path.abspath(os.path.join(ruta_archivos_x_inclu,f"{ramo.poliza}_97.xls"))
    file_path_leer = os.path.abspath(os.path.join(ruta_archivos_x_inclu,f"{ramo.poliza}.xlsx"))

    if bab_codigo != '4':

        bloques2 = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.row.gnSecToggle.ng-star-inserted")))
        logging.info("-------------------------")
        logging.info(f"🔍 Bloques encontrados: {len(bloques2)}")

        df = pd.read_excel(file_path_leer, engine="openpyxl", dtype={"Sueldo": float})

        for bloqueX in bloques2:

            label_trab = bloqueX.find_element(By.XPATH, ".//label[contains(., 'Trabajadores')]")

            driver.execute_script("arguments[0].click();", label_trab)
            numero_de_trabajadores = df.shape[0]

            input_trab = bloqueX.find_element(By.XPATH, ".//label[contains(., 'Trabajadores')]/ancestor::mat-form-field//input")
            input_trab.clear()
            input_trab.send_keys(numero_de_trabajadores)
            logging.info(f"✅ Se ingresó {numero_de_trabajadores} trabajadores")

            label_monto = bloqueX.find_element(By.XPATH, ".//label[contains(., 'Monto')]")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", label_monto)
            driver.execute_script("arguments[0].click();", label_monto)

            multiplicadores = {
                "Mensual": 1,
                "Bimestral": 2,
                "Trimestral": 3,
                "Semestral": 6,
                "Anual": 12
            }

            # Obtener multiplicador según modalidad
            multiplicador = multiplicadores.get(modalidad)  # default = 1
            # Calcular el total sumando sueldo*multiplicador
            total_sueldo = (df["Sueldo"] * multiplicador).sum()

            input_monto = bloqueX.find_element(By.XPATH, ".//label[contains(., 'Monto')]/ancestor::mat-form-field//input")
            input_monto.clear()
            input_monto.send_keys(total_sueldo)
            logging.info(f"💰 Monto ingresado: S/.{total_sueldo}")

    #Scroll para ir hasta abajo
    driver.execute_script("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});")

    input_file = wait.until(EC.presence_of_element_located((By.ID, "iImportarPlanilla")))
    input_file.send_keys(file_path)
    logging.info(f"📤 Trama {ramo.poliza}.xls cargada")
    time.sleep(2)

    # driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    
    if bab_codigo == '4':
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Procesar')]")))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", btn)
        # Forzar click compatible con Angular
        driver.execute_script("""
        var element = arguments[0];
        ['mouseover', 'mousedown', 'mouseup', 'click'].forEach(eventType => {
            var event = new MouseEvent(eventType, { bubbles: true, cancelable: true, view: window });
            element.dispatchEvent(event);
        });
        """, btn)
        logging.info(f"🖱️ Clic en 'Procesar'")

        try:
            errores = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//li[@class='col-10 cnt-item item-dato g-text-uppercase g-text-left-xs']")))
            lista_errores = [e.text.strip() for e in errores]
            raise Exception(lista_errores)
        except TimeoutException:
            pass

    boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'g-button') and normalize-space()='Siguiente']")))
    # Forzar click compatible con Angular
    driver.execute_script("""
    var element = arguments[0];
    ['mouseover', 'mousedown', 'mouseup', 'click'].forEach(eventType => {
        var event = new MouseEvent(eventType, { bubbles: true, cancelable: true, view: window });
        element.dispatchEvent(event);
    });
    """, boton)
 
    logging.info(f"🖱️ Clic en 'Siguiente'")
  
    mensaje_error = None
 
    try:
        # Espera que aparezac el modal dentro de 10 segundos
        WebDriverWait(driver,7).until( EC.visibility_of_element_located((By.CSS_SELECTOR, ".swal2-popup.swal2-modal")))
 
        # Si lo encuentra, obtener mensaje
        mensaje_error = driver.find_element(By.CSS_SELECTOR, "#swal2-html-container").text
        logging.info(f"❌ ERROR DETECTADO: {mensaje_error}")
 
        # Guardar error en archivo
        with open("errores_trama.txt", "a", encoding="utf-8") as f:
            f.write(mensaje_error + "\n")
 
        raise Exception(mensaje_error)
 
    except TimeoutException:

        try:

            modal = WebDriverWait(driver,7).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "mat-dialog-container")))
            logging.info("⚠️ Modal de observaciones detectado.")
            nota_element = modal.find_element(By.XPATH, ".//*[contains(text(), 'NOTA:')]")
            mensaje_error = nota_element.text.strip()

            with open("errores_trama.txt", "a", encoding="utf-8") as f:
                f.write(mensaje_error + "\n\n")
       
            raise Exception(mensaje_error)

        except TimeoutException:
            logging.info("✅ No apareció modal de observaciones. Continuando...")
           
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    logging.info("✅ Scroll hasta abajo")
    time.sleep(2)

    boton = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'g-button') and normalize-space()='Generar']")))
 
    #Forzar click compatible con Angular
    driver.execute_script("""
    var element = arguments[0];
    ['mouseover', 'mousedown', 'mouseup', 'click'].forEach(eventType => {
        var event = new MouseEvent(eventType, { bubbles: true, cancelable: true, view: window });
        element.dispatchEvent(event);
    });
    """, boton)
 
    logging.info("🖱️ Clic al botón Generar")

    boton_generar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'GENERAR')]")))
    boton_generar.click()
    logging.info("🖱️ Clic al Modal 'GENERAR'")

    boton_ok = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(),'OK')]")))
    boton_ok.click()
    logging.info("🖱️ Clic en 'Ok'")

    boton_enviar = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'Enviar documentos')]")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", boton_enviar)

    if tipo_mes == 'MA':

        if bab_codigo == '3':
            alias = {
            "Constancia": f"{list_polizas[0]}",
            "Recibo Pensión": f"endoso_{list_polizas[0]}",
            "Recibo Salud": f"endoso_{list_polizas[1]}"
            }
        else:
            alias = {
                "Constancia": f"{ramo.poliza}",
                "Recibo Vida Ley": f"endoso_{ramo.poliza}"
            }
    else:
        if bab_codigo == '3':
            alias = {
                "Constancia": f"{list_polizas[0]}"
            }
        else:
            alias = {
                "Constancia": f"{ramo.poliza}"
            }

    #---For para descargar---
    try:

        # Esperar que la lista de documentos esté presente
        ul_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.col-xs-12 > ul")))

        # Solo li visibles (los que realmente se muestran en la web)
        document_items = ul_element.find_elements(By.CSS_SELECTOR, "li.h-myd--show")

        # Recorrer cada fila con enumeración para tener un contador
        for idx, item in enumerate(document_items):
            try:
                logging.info("---------------------------------------")
                # Capturar tipo de documento y número de documento
                tipo_doc = item.find_element(By.CSS_SELECTOR, "div.info > p").text.strip()
                logging.info(f"Documento: {tipo_doc}")
                # Verificar si tipo_doc está en multiplicadores y definir nombre
                if tipo_doc in alias:
                    nombre_guardar = f"{alias[tipo_doc]}"
                else:
                    #nombre_guardar = tipo_doc  # Si no está, usar el nombre original
                    continue
      
                # Buscar el botón Descargar dentro de la fila
                descargar_btn = item.find_element(By.XPATH, ".//a[contains(text(), 'Descargar')]")
            
                # Guardar archivos antes del clic
                archivos_antes = set(os.listdir(ruta_archivos_x_inclu))

                driver.execute_script("arguments[0].click();", descargar_btn)
                logging.info("🖱 Clic con JS en el botón de descarga")

                archivo_nuevo = esperar_archivos_nuevos(ruta_archivos_x_inclu,archivos_antes,".pdf",cantidad=1)
                logging.info(f"✅ Archivo '{tipo_doc}' descargado exitosamente")

                if archivo_nuevo:
                    ruta_original = os.path.join(ruta_archivos_x_inclu, archivo_nuevo)
                    ruta_final = os.path.join(ruta_archivos_x_inclu, f"{nombre_guardar}.pdf")
                    os.rename(ruta_original, ruta_final)
                    logging.info(f"🔄 Archivo renombrado a '{nombre_guardar}.pdf'")
                else:
                    raise Exception("No se encontró archivo después de la descarga")

                # # Intentar descargar usando tu función personalizada
                # if descargar_documento(driver, descargar_btn, nombre_guardar, impresion=False,pestaña=True):
                #     time.sleep(2)  # espera para no saturar la web
                #     logging.info(f"✅ Descargado: Fila {idx}, Tipo: {tipo_doc}")

            except Exception as e:
                logging.error(f"❌ Error en fila {idx} ({tipo_doc if 'tipo_doc' in locals() else 'Desconocido'}): {e}")

    except Exception as e:
        logging.error(f"❌ Error general descargando en el for: {e}")
    #------------------------

    if len(list_polizas) == 2:
        logging.info("----------------------------")
        ruta_salud = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[0]}.pdf")
        ruta_pension = os.path.join(ruta_archivos_x_inclu,f"{list_polizas[1]}.pdf")

        if os.path.exists(ruta_salud):
            shutil.copy2(ruta_salud,ruta_pension)
            logging.info(f"📄 Copia creada para constancia de Pensión")
        else:
            logging.error(f"❌ No existe el archivo base Constancia Salud")

        logging.info("----------------------------")


    boton_enviar = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'Enviar documentos')]")))
    driver.execute_script("arguments[0].click();", boton_enviar)
    logging.info("🖱️ Clic en 'Enviar Documentos'.")

    try:

        input_para = wait.until(EC.presence_of_element_located((By.XPATH,"//mat-label[normalize-space()='Para']/ancestor::mat-form-field//input")))
        driver.execute_script("arguments[0].focus();", input_para)
        time.sleep(1)
        input_para.clear()
        input_para.send_keys(ejecutivo_responsable)
        logging.info(f"📧 Correo ingresado: {ejecutivo_responsable}")

        input_asunto = wait.until(EC.presence_of_element_located((By.XPATH,"//mat-label[normalize-space()='Asunto']/ancestor::mat-form-field//input")))
        driver.execute_script("arguments[0].focus();", input_asunto)
        input_asunto.clear()
        texto_final = f"{palabra_clave}-{ramo.poliza}-{nombre_cliente}"
        input_asunto.send_keys(texto_final)

        logging.info(f"📝 Asunto ingresado: {texto_final}")

        boton_enviar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(text(),'Enviar')]")))
        boton_enviar.click()
        logging.info("🖱️ Clic en el botón 'Enviar'.")
 
    except Exception as e:
        logging.info (f"❌ Error enviando documento: {e}")
        raise Exception(f"{e}")

    if ba_codigo == '3':
        
        rutas = {
            "salud": os.path.join(ruta_archivos_x_inclu, f"endoso_{list_polizas[0]}.pdf"),
            "pension": os.path.join(ruta_archivos_x_inclu, f"endoso_{list_polizas[1]}.pdf"),
        }

        faltantes = [nombre for nombre, ruta in rutas.items() if not os.path.exists(ruta)]

        if faltantes:
            raise Exception(f"No se encontraron los archivos {' y '.join(faltantes)}")

def realizar_solicitud_mapfre(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,
                              ejecutivo_responsable,ba_codigo,bab_codigo,nombre_cliente,ramo):

    global ventana_solicitud_mapfre
    tipoError = ""
    detalleError = ""
    constancia = False
    proforma = False

    if ba_codigo == '3' and bab_codigo == '4':

        try:
            solicitud_sctr_vl(driver,wait,palabra_clave,tipo_proceso,ruta_archivos_x_inclu,ba_codigo,bab_codigo,
                              list_polizas,ejecutivo_responsable,nombre_cliente,tipo_mes,ramo)
            return True,True,tipoError,detalleError
        except Exception as e:
            logging.error(f"❌ Error en Mapfre (VL) - {tipo_mes}: {e}")
            return constancia,proforma,f"MAPF-VL-{tipo_mes}",f"{e}"

    try:

        driver.get("https://oim.mapfre.com.pe/")
        logging.info("⌛ Cargando la Web de Mapfre")

        user_input = wait.until(EC.presence_of_element_located((By.ID, "mat-input-1")))
        user_input.clear()
        user_input.send_keys(ramo.usuario)
        logging.info("⌨️ Digitando el Username")
        
        pass_input = wait.until(EC.presence_of_element_located((By.ID, "mat-input-0")))
        pass_input.clear()
        pass_input.send_keys(ramo.clave)
        logging.info("⌨️ Digitando el Password")

        ingresar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Ingresar')]")))
        ingresar_btn.click()
        logging.info("🖱️ Clic en 'Ingresar' en Mapfre")

        elemento = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".card-modality__item--left")))
        elemento.click()
        logging.info("🖱️ Clic en enviar por Correo Electronico")

        codigo_mapfre_path = "/codigo_mapfre/codigo.txt"

        while not os.path.exists(codigo_mapfre_path):
            time.sleep(2)

        with open(codigo_mapfre_path, "r") as f:
            codigo = f.read().strip()

        logging.info(f"✅ Código recibido desde volumen: {codigo}")

        inputs = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input.g-input-codes__code")))

        if len(inputs) == len(codigo):
            for i, inp in enumerate(inputs):
                inp.clear()
                inp.send_keys(codigo[i])
        else:
            raise Exception("Los inputs no coinciden con la longitud del código")

        time.sleep(1)

        try:
            comprobar_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Comprobar')]")))
            comprobar_btn.click()
            logging.info("🖱️ Clic en 'Comprobar'")
        except Exception as e:
            raise Exception(f"No se pudo hacer clic en Comprobar -> {e}")

        # --- Eliminar el archivo después de usarlo ---
        try:
            os.remove(codigo_mapfre_path)
            logging.info("🧹 Archivo codigo.txt eliminado desde volumen")
        except FileNotFoundError:
            logging.warning("⚠️ No se encontró codigo.txt al intentar eliminarlo (ya fue borrado)")
        except Exception as e:
            logging.error(f"❌ Error al eliminar codigo.txt: {e}")

        try:
        
            # Esperar a que aparezca el texto del modal
            mensaje_elemento = WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.c-modal p.txt")))
            mensaje = mensaje_elemento.text.strip()
            print(f"⚠️ Modal detectado: {mensaje}")

            boton = wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//span[contains(text(),'Cerrar')]]")))

            driver.execute_script("arguments[0].click();", boton)
            print("✅ Modal cerrado correctamente")

        except TimeoutException:
            pass

        try:
            boton_ok = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Ok')]]")))
            boton_ok.click()
            logging.info("🖱️ Clic en 'Ok'")
        except TimeoutException:
            pass
        
        consulta_gestion = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='CONSTANCIAS SCTR Y VL']")))
        consulta_gestion.click()
        logging.info("🖱️ Clic en 'CONSTANCIAS SCTR Y VL'")

        ventana_solicitud_mapfre = driver.current_window_handle
    
    except Exception as e:
        return constancia,proforma,"Página Web", "Hubo problemas al iniciar sessión en la compañía"
    
    try:
        solicitud_sctr_vl(driver,wait,palabra_clave,tipo_proceso,ruta_archivos_x_inclu,ba_codigo,bab_codigo,list_polizas,
                          ejecutivo_responsable,nombre_cliente,tipo_mes,ramo)
        return True,True,tipoError,detalleError
    except Exception as e:

        ramo_s = "SCTR" if all(c == '3' for c in (ba_codigo, bab_codigo)) else "VL"
        logging.error(f"❌ Error en Mapfre {ramo_s} - {tipo_mes}: {e}")
        try:
            ok_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
            ok_btn.click()
            logging.info("🖱️ Clic en 'Confirmar'")
        except :
            try:
                btn_cancelar = wait.until(EC.element_to_be_clickable((By.XPATH,"//a[contains(@class,'g-button') and contains(normalize-space(.), 'Cancelar')]")))
                btn_cancelar.click()
                logging.info("🖱️ Clic en 'Cancelar'")
            except:
                pass

        return constancia,proforma,f"MAPF-{ramo_s}-{tipo_mes}",f"{e}"
    finally:
        time.sleep(3)
        link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[span[text()='Constancias SCTR y VL']]")))  
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", link)
        driver.execute_script("arguments[0].click();", link)
        logging.info("🔙 Regresando al menú de Constancias")

