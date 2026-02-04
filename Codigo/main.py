# -*- coding: utf-8 -*-
# -- Froms ---
#from jinja2.utils import
from numpy import iinfo
from Compañias.Mapfre.web import realizar_solicitud_mapfre
from Compañias.Pacifico.web import realizar_solicitud_pacifico
from Compañias.Positiva.web import realizar_solicitud_la_positiva
from Compañias.Sanitas.web import realizar_solicitud_sanitas,procesar_solicitud_san_protecta_vl,procesar_solicitud_san_crecer_vl
from Compañias.Rimac.VidaLey.web_sas import realizar_solicitud_SAS
from Compañias.Rimac.SCTR.web_PortalWeb import realizar_solicitud_PortalWeb
#from Compañias.Rimac.SCTR.web_Corredores import realizar_solicitud_corredores
from Plantillas.Crecer.generarplantilla import generarConstanciaInCrecer,generarConstanciaReCrecer
from Apis.metodos_put import enviar_documentos,enviar_error_movimiento,enviar_puerto,enviar_estaca
from LinuxDebian.rutas import armar_ruta_archivos
from Chrome.google import abrirDriver
from selenium.webdriver.support.ui import WebDriverWait
from LinuxDebian.ventana import esperar_ventana
# -- Imports --
import os,json
import logging
import sys
import time
import io
import subprocess
from pprint import pprint
from pprint import pformat

# Forzar la salida en UTF-8 para evitar UnicodeEncodeError
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
# Variables de entorno
puerto = os.getenv("puerto")

def to_bool(value):
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip().lower() in ("true", "1", "yes", "y")
    if isinstance(value, (int, float)):
        return value == 1
    return False

#--- Json que llegan desde Webhook como variable de Entorno -------
data = json.loads(os.getenv("DATA"))

class ContextoRPA:

    def __init__(self, data: dict):
        self.cliente = data["cliente"]
        self.correo = data["correo"]
        self.ruc = data["ruc"]
        self.giro = data["giro"]
        self.proceso = data["proceso"]
        self.tipo_mes = data["tipoMes"]

        self.salud = Salud(data)
        self.pension = Pension(data)
        self.vida = VidaLey(data)

    def __str__(self):
        data = {
            "cliente": self.cliente,
            "correo": self.correo,
            "ruc": self.ruc,
            "giro": self.giro,
            "proceso": self.proceso,
            "tipo_mes": self.tipo_mes,
            "salud": self.salud.to_dict(ocultar_clave=True),
            "pension": self.pension.to_dict(ocultar_clave=True),
            "vida": self.vida.to_dict(ocultar_clave=True),
        }
        return pformat(data)

class RamoBase:
    def __init__(self, activo: bool,proforma: bool, poliza: str):
        self.activo = activo
        self.proforma = proforma
        self.poliza = str(poliza or "0")

    def tiene_poliza(self) -> bool:
        return self.poliza != "0"

    def debe_procesarse(self) -> bool:
        return not self.activo and self.tiene_poliza()

    def to_dict(self, ocultar_clave=False):
        data = self.__dict__.copy()
        if ocultar_clave and "clave" in data:
            data["clave"] = "****"
        return data

class Salud(RamoBase):
    def __init__(self, data: dict):
        super().__init__(
            activo=to_bool(data.get("estacaSalud")),
            proforma=to_bool(data.get("proformaSalud")),
            poliza=data.get("poliza_salud")
        )

        self.compania = data["companiaSalud"]
        self.usuario = data["usuarioSalud"]
        self.clave = data["contraseniaSalud"]
        self.celular = data["celularSalud"]
        self.id_poliza = data["id_poliza_salud"]
        self.f_inicio = data["fInicioVigSalud"]
        self.f_fin = data["fFinVigSalud"]
        self.trama_97 = data["tramaSalud97"]
        self.trama = data["tramaSalud"]
        self.sede = data["sedeSalud"]

    def __str__(self):
        return (
            f"Salud ("
            f"Póliza:{self.poliza}, "
            f"Proforma: {self.proforma}, "
            f"Sede: {self.sede}, "
            f"Vigencia: {self.f_inicio} al {self.f_fin}"
            f")"
        )

class Pension(RamoBase):
    def __init__(self, data: dict):
        super().__init__(
            activo=to_bool(data.get("estacaPension")),
            proforma=to_bool(data.get("proformaSalud")),
            poliza=data.get("poliza_pension")
        )

        self.compania = data["companiaPension"]
        self.usuario = data["usuarioPension"]
        self.clave = data["contraseniaPension"]
        self.celular = data["celularPension"]
        self.id_poliza = data["id_poliza_pension"]
        self.f_inicio = data["fInicioVigPension"]
        self.f_fin = data["fFinVigPension"]
        self.trama_97 = data["tramaPension97"]
        self.trama = data["tramaPension"]
        self.sede = data["sedePension"]

    def __str__(self):
        return (
            f"Pensión ("
            f"Póliza: {self.poliza}, "
            f"Proforma: {self.proforma}, "
            f"Sede: {self.sede}, "
            f"Vigencia: {self.f_inicio} al {self.f_fin}"
            f")"
        )

class VidaLey(RamoBase):
    def __init__(self, data: dict):
        super().__init__(
            activo=to_bool(data.get("estacaVida")),
            proforma=to_bool(data.get("proformaSalud")),
            poliza=data.get("poliza_vida")
        )

        self.compania = data["companiaVida"]
        self.usuario = data["usuarioVida"]
        self.clave = data["contraseniaVida"]
        self.celular = data["celularVidaLey"]
        self.id_poliza = data["id_poliza_vida"]
        self.f_inicio = data["fInicioVigVida"]
        self.f_fin = data["fFinVigVida"]
        self.trama_97 = data["tramaVida97"]
        self.trama = data["tramaVida"]
        self.sede = data["sedeVida"]

    def __str__(self):
        return (
            f"VidaLey ("
            f"Póliza: {self.poliza}, "
            f"Proforma: {self.proforma}, "
            f"Sede: {self.sede}, "
            f"Vigencia: {self.f_inicio} al {self.f_fin}"
            f")"
        )

ctx = ContextoRPA(data)
         
def derivar_compania_sctr(driver,wait,list_polizas,compania_BA,ba_codigo,tipo_mes,ruta_archivos_x_inclu,
                          ruc_empresa,ejecutivo_responsable,tipo_proceso,palabra_clave,actividad,nombre_cliente,
                          ramo):

    # def ejecutar_rimac_corredores():
    #     logging.info("✅ Compañía: Rímac - Corredores")
    #     return realizar_solicitud_corredores(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,
    #                                          ejecutivo_responsable,ba_codigo,nombre_cliente,ruc_empresa,ramo)

    def ejecutar_rimac_portal_web():
        logging.info("✅ Compañía: Rímac - Portal Web")
        return realizar_solicitud_PortalWeb(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,
                                                                    palabra_clave,ejecutivo_responsable,ba_codigo,nombre_cliente,ruc_empresa,ramo)

    def ejecutar_mapfre():
        logging.info("✅ Compañía: Mapfre")
        return realizar_solicitud_mapfre(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,
                                         ejecutivo_responsable,ba_codigo,ba_codigo,nombre_cliente,ramo)

    def ejecutar_sanitas():

        urls = [os.getenv("login_url_sanitas_crecer"),os.getenv("login_url_sanitas_protecta")]
        logging.info(f"✅ Compañía: Sanitas")
        return realizar_solicitud_sanitas(driver,wait,urls,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,ramo)

    def ejecutar_lapositiva():
        logging.info("✅ Compañía: La Positiva")
        return realizar_solicitud_la_positiva(driver,wait,list_polizas,ba_codigo,ba_codigo,tipo_mes,ruta_archivos_x_inclu,
                                              ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_proceso,actividad,ramo)

    def ejecutar_pacifico():
        logging.info("✅ Compañía: Pacifico")
        return realizar_solicitud_pacifico(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,
                                           tipo_proceso,palabra_clave,ejecutivo_responsable,ba_codigo,ramo)

    dispatch = {
        'SANI': ejecutar_sanitas,
        'CREC': ejecutar_sanitas,
        'PROT': ejecutar_sanitas,
        'MAPF': ejecutar_mapfre,
        'LAPO': ejecutar_lapositiva,
        'PACI': ejecutar_pacifico,
        'RIMA': ejecutar_rimac_portal_web
    }

    if compania_BA not in dispatch:
        logging.error(f"❌ Compañía no reconocida: '{compania_BA}'")
        return False,False,"Compañía no reconocida","Revisar compañía"

    resultado = dispatch[compania_BA]()
    constancia,proforma,tipoError,detalleErrror = resultado
   
    return constancia,proforma,tipoError,detalleErrror

def derivar_compania_vidaley(driver,wait,list_polizas,compania_BB,ba_codigo,bb_codigo,tipo_mes, 
                             ruta_archivos_x_inclu,ruc_empresa,ejecutivo_responsable,palabra_clave, 
                             tipo_proceso,nombre_cliente,actividad,ramo):

    tipoError = ""
    detalleError = ""

    # -------------------------------
    # 🔹 Funciones internas por compañía
    # -------------------------------
    # PortalWeb: Facturas, Incluciones-VidaLey(MA), Polizas y SOAT.
    # SAS: MV VidaLey.
    # WebCorredores: SCTR (MV-MA) y Ver Facturas.
 
    def ejecutar_rimac_SAS():
        logging.info("✅ Compañía: Rímac - SAS")
        return realizar_solicitud_SAS(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,ejecutivo_responsable,
                                      ba_codigo, bb_codigo,nombre_cliente,ruc_empresa,ramo)

    def ejecutar_rimac_portal_web():
        logging.info("✅ Compañía: Rímac - Portal Web")
        return realizar_solicitud_PortalWeb(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,
                                            ejecutivo_responsable,ba_codigo,nombre_cliente,ruc_empresa,ramo)

    def manejar_error_crecer(e,tipo_mes):
        logging.error(f"❌ Error en Crecer Vida Ley: {e}")
        return False,False, f"{compania_BB}-VL-{tipo_mes}", e

    def ejecutar_mapfre():
        logging.info("✅ Compañía: Mapfre")
        return realizar_solicitud_mapfre(driver,wait,list_polizas,tipo_mes,ruta_archivos_x_inclu,tipo_proceso,palabra_clave,
                                         ejecutivo_responsable,ba_codigo,bb_codigo,nombre_cliente,ramo)

    def ejecutar_crecer():
        logging.info("✅ Compañía: Crecer")

        try:
            if tipo_mes == 'MV':
                logging.info("Tipo: Mes Vencido")

                if tipo_proceso == 'IN':
                    generarConstanciaInCrecer(ruta_archivos_x_inclu,palabra_clave,tipo_proceso,nombre_cliente,ramo)
                else:
                    generarConstanciaReCrecer(ruta_archivos_x_inclu,palabra_clave,tipo_proceso,nombre_cliente,ruc_empresa,ramo)
                time.sleep(2)
                return True,True, tipoError, detalleError

            else:
                logging.info("Tipo: Mes Adelantado")
                procesar_solicitud_san_crecer_vl(driver,wait,tipo_proceso,ruta_archivos_x_inclu,palabra_clave)
                return True,True, tipoError, detalleError

        except Exception as e:
            return manejar_error_crecer(e,tipo_mes)

    def ejecutar_protecta():
        logging.info("✅ Compañía: Protecta")

        if tipo_mes == 'MV':
            logging.info("Tipo: Mes Vencido (manual por compañía)")
            time.sleep(5)
            return True,True, tipoError, detalleError
        else:
            logging.info("Tipo: Mes Adelantado")
            procesar_solicitud_san_protecta_vl(driver, wait, ruc_empresa, tipo_proceso, ruta_archivos_x_inclu,ramo)
            return True,True, tipoError, detalleError

    def ejecutar_lapositiva():
        logging.info("✅ Compañía: La Positiva")
        return realizar_solicitud_la_positiva(driver,wait,list_polizas,ba_codigo,bb_codigo,tipo_mes,ruta_archivos_x_inclu,
                                                                      ruc_empresa,ejecutivo_responsable,palabra_clave,tipo_proceso,actividad,ramo)

    dispatch = {
        'CREC': ejecutar_crecer,
        'PROT': ejecutar_protecta,
        'MAPF': ejecutar_mapfre,
        'LAPO': ejecutar_lapositiva,
        'RIMA': ejecutar_rimac_SAS if tipo_mes == 'MV' else ejecutar_rimac_portal_web
    }

    if compania_BB not in dispatch:
        logging.error(f"❌ Compañía no reconocida: '{compania_BB}'")
        return False,False,"Compañía no reconocida","Revisar compañía"

    resultado = dispatch[compania_BB]()
    constancia,proforma,tipoError,detalleErrror = resultado
   
    return constancia,proforma,tipoError,detalleErrror

def enviar_puerto_por_ramos(RAMOS, puerto):
    for r in RAMOS:

        ctx_ramo = r["ctx"]

        # ⛔ No existe movimiento
        if not ctx_ramo.id_poliza:
            continue

        # ⛔ Ya está activo
        if ctx_ramo.activo:
            continue

        logging.info(
            f"⌛ Enviando puerto {puerto} al Id -> {ctx_ramo.id_poliza} de '{r['nombre']}'"
        )

        if not enviar_puerto(ctx_ramo.id_poliza, puerto):
            raise Exception(f"No se pudo enviar puerto {puerto}")

        time.sleep(1)

        return True   # ✅ SE ENVIÓ UNA VEZ

    logging.info("ℹ️ No hubo ramos disponibles para enviar puerto")
    return False

def click_descarga_documento():
    
    try:

        if not esperar_ventana("Save File"):
            raise Exception("No apareció la ventana de descarga")

        time.sleep(1)
        subprocess.run(["xdotool", "search", "--name", "Save File", "windowactivate", "windowfocus"])
        subprocess.run(["xdotool", "key", "Return"])
        return True

    except Exception as ex:
        logging.error(f"❌ Error durante el flujo de descarga: {ex}")
        return False

def main():
 
    ok_sctr = False
    ok_vl = False
    ok_proforma_sctr = False
    ok_proforma_vl = False
    tipoErrorSCTR = ""
    detalleErrorSCTR = ""
    tipoErrorVL = ""
    detalleErrorVL = ""

    match ctx.tipo_mes.upper():
        case 'ADELANTADO':
            tipo_mes = 'MA'
        case _:
            tipo_mes = 'MV'

    # Mapa de compañías
    dispatch = {
        'SANITAS': 'SANI',
        'CRECER': 'CREC',
        'PROTECTA': 'PROT',
        'MAPFRE': 'MAPF',
        'LA POSITIVA': 'LAPO',
        'PACIFICO': 'PACI',
        'RIMAC': 'RIMA',
        'UNIVERSAL': 'NING'
    }

    if not ctx.salud.debe_procesarse() and not ctx.pension.debe_procesarse():
            ba_codigo = '0'
    elif ctx.salud.debe_procesarse() and ctx.pension.debe_procesarse():
        ba_codigo = '3'
    elif ctx.salud.debe_procesarse() and not ctx.pension.debe_procesarse():
        ba_codigo = '1'
    else:
        ba_codigo = '2'

    def get_compania_abrev(nombre):
        return dispatch.get(nombre, '')

    if ba_codigo in ('1','3'):
        compania_BA = get_compania_abrev(ctx.salud.compania)
    elif ba_codigo == '2':
        compania_BA = get_compania_abrev(ctx.pension.compania)
    else:
        compania_BA = 'NING'

    if not ctx.vida.debe_procesarse() :
        bb_codigo = '5'
        compania_BB = 'NING'
    else:
        bb_codigo = '4'
        compania_BB = get_compania_abrev(ctx.vida.compania)

    try:

        if ba_codigo == '0' and bb_codigo == '5':
            raise Exception("No hay pólizas para procesar")

        #--- Cases para tipos de proceso y mes ---
        match ctx.proceso:
            case 'INCLUSION':
                tipo_proc = 'IN'
                palabra_clave = 'Inclusion'
                tiempo = 60
            case _:
                tipo_proc = 'RE'
                palabra_clave = 'Renovacion'
                tiempo = 120

        ruta_archivos_x_inclu = armar_ruta_archivos(tipo_proc,ba_codigo,bb_codigo,compania_BA,compania_BB,ctx.salud.poliza,ctx.pension.poliza,ctx.vida.poliza)

        #--- Iniciando el Driver de Chrome ---
        driver = abrirDriver(ruta_archivos_x_inclu)
        wait = WebDriverWait(driver,tiempo)

        PRE_RAMOS = [
            {
                "nombre": "SALUD",
                "ctx": ctx.salud
            },
            {
                "nombre": "PENSION",
                "ctx": ctx.pension
            },
            {
                "nombre": "VIDALEY",
                "ctx": ctx.vida
            }
        ]

        logging.info("-----------------------------")
        #-- Para descargar las tramas ---
        for r in PRE_RAMOS:

            ctx_ramo = r["ctx"]

            if ctx_ramo.activo:
                continue

            # Trama principal
            if ctx_ramo.trama:
                logging.info(f"⌛ Descargando trama {r['nombre']}")

                driver.get(ctx_ramo.trama)

                if click_descarga_documento():
                    logging.info(f"✅ Trama {r['nombre']} descargada")
                else:

                    raise Exception(f"No se pudo descargar la trama {r['nombre']}")

                time.sleep(1)

            # Trama 97 (si aplica)
            if ctx_ramo.trama_97:
                logging.info(f"⌛ Descargando trama 97 {r['nombre']}")

                driver.get(ctx_ramo.trama_97)

                if click_descarga_documento():
                    logging.info(f"✅ Trama 97 {r['nombre']} descargada")
                else:
                    raise Exception(f"No se pudo descargar la trama {r['nombre']}")

                time.sleep(1)
        
        logging.info("-----------------------------")
        # --- Enviando puerto a Birlik (solo una vez) ---
        if any(r["ctx"].id_poliza and not r["ctx"].activo for r in PRE_RAMOS):
            enviar_puerto_por_ramos(PRE_RAMOS, puerto)

        if tipo_proc and ba_codigo and bb_codigo and compania_BA and compania_BB and ctx.salud.poliza and ctx.pension.poliza and ctx.vida.poliza and tipo_mes:
            
            polizas_sctr = []
            contexto_sctr = None

            if ctx.salud.debe_procesarse():

                logging.info("-----------------------------")
                polizas_sctr.append(ctx.salud.poliza)
                logging.info(f"✅ {ctx.salud}")

            if ctx.pension.debe_procesarse():

                logging.info("-----------------------------")
                polizas_sctr.append(ctx.pension.poliza)
                logging.info(f"✅ {ctx.pension}")

            if polizas_sctr:
                logging.info(f"⌛ Procesando {palabra_clave} en SCTR ({' y '.join(polizas_sctr)})")

                contexto_sctr = ctx.salud if ctx.salud.debe_procesarse() else ctx.pension

                ok_sctr,ok_proforma_sctr, tipoErrorSCTR, detalleErrorSCTR = derivar_compania_sctr(
                    driver, wait, polizas_sctr, compania_BA, ba_codigo, tipo_mes,
                    ruta_archivos_x_inclu, ctx.ruc, ctx.correo, tipo_proc,
                    palabra_clave, ctx.giro, ctx.cliente, contexto_sctr
                )

                #logging.info(f"TipoErrorSCTR : {tipoErrorSCTR} y DetalleErrorSCTR {detalleErrorSCTR}")
           
            if ctx.vida.debe_procesarse():
                logging.info(f"⌛ Procesando {palabra_clave} en VIDA LEY - póliza {ctx.vida.poliza}")
                logging.info(f"✅ {ctx.vida}")

                ok_vl,ok_proforma_vl, tipoErrorVL, detalleErrorVL = derivar_compania_vidaley(
                    driver,wait,[ctx.vida.poliza],compania_BB,ba_codigo,bb_codigo,tipo_mes,
                    ruta_archivos_x_inclu,ctx.ruc, ctx.correo, palabra_clave,tipo_proc,ctx.cliente,
                    ctx.giro,ctx.vida)
                
        else:
            raise Exception(f" Dato no válido -> {tipo_proc}-{ba_codigo}-{bb_codigo}-{compania_BA}-{compania_BB}-{ctx.salud.poliza}-{ctx.pension.poliza}-{ctx.vida.poliza}-{tipo_mes}")
         
    except Exception as e:
        logging.error(f"⚠️ Conclusión: {e}")
        tipoErrorSCTR = "Fallas en la Trama con Azure"
        detalleErrorSCTR = "Trama no descargada"
        tipoErrorVL = "Fallas en la Trama con Azure"
        detalleErrorVL = "Trama no descargada"
    finally:

        RAMOS = [
            {
                "codigo": 0,
                "nombre": "SALUD",
                "ctx": ctx.salud,
                "ok": ok_sctr,
                "proforma": ok_proforma_sctr
            },
            {
                "codigo": 1,
                "nombre": "PENSION",
                "ctx": ctx.pension,
                "ok": ok_sctr,
                "proforma": ok_proforma_sctr
            },
            {
                "codigo": 2,
                "nombre": "VIDALEY",
                "ctx": ctx.vida,
                "ok": ok_vl,
                "proforma": ok_proforma_vl
            }
        ]

        def obtener_error(r):
            if r["nombre"] == "VIDALEY":
                return tipoErrorVL, detalleErrorVL
            return tipoErrorSCTR, detalleErrorSCTR

        #------------- Enviando Estacas,Errores y Oocumentos---------------------
        for r in RAMOS:

            ctx_ramo = r["ctx"]
            nombre = r["nombre"]
            ok = r["ok"]
            proforma = r["proforma"]
            id_mov = ctx_ramo.id_poliza

            if not id_mov:
                continue  # ⛔ No existe movimiento

            logging.info("-----------------------------")
            # ---------------- ESTACA ----------------

            if tipo_mes == 'MV' and ok and not proforma:
                logging.info(f"⌛ Actualizando registros de Constancia → '{ok}' para el Id → {id_mov} de '{nombre}'")

                enviar_estaca(id_mov, nombre,ok,proforma)
                time.sleep(1)
            elif tipo_mes == 'MA' and ok and proforma:

                logging.info(f"⌛ Actualizando registros de Constancia → '{ok}' y Proforma → '{proforma}' para el Id → {id_mov} de '{nombre}'")

                enviar_estaca(id_mov, nombre, ok,proforma)
                time.sleep(1)
            elif tipo_mes == 'MA' and not ok and proforma:
            #--------------------------------
            #if ok or proforma:
                logging.info(f"⌛ Actualizando registros de Proforma → '{proforma}' para el Id → {id_mov} de '{nombre}'")

                enviar_estaca(id_mov, nombre, ok,proforma)
                time.sleep(1)

            # ---------------- ERRORES ----------------
            #else:
            tipo_error, detalle_error = obtener_error(r)

            if tipo_error and detalle_error:
                logging.info(f"⌛ Enviando errores para '{nombre}' con Id → {id_mov}")

                enviar_error_movimiento(
                    id_mov,
                    nombre,
                    tipo_error,
                    detalle_error
                )
                time.sleep(1)

            #continue  # ⛔ Si hubo error, no enviar documentos

            # ---------------- DOCUMENTOS ----------------
            if ok:
                constancia = f"{ctx_ramo.poliza}.pdf"
                ruta_constancia = os.path.join(ruta_archivos_x_inclu, constancia)

                logging.info(f"⌛ Enviando Constancia de '{nombre}' al Id → {id_mov}")

                enviar_documentos(id_mov, ruta_constancia, nombre, "Constancia")
                time.sleep(1)

            if proforma: #tipo_mes == 'MA':
                endoso = f"endoso_{ctx_ramo.poliza}.pdf"
                ruta_endoso = os.path.join(ruta_archivos_x_inclu, endoso)

                logging.info(f"⌛ Enviando Endoso de '{nombre}' al Id →{id_mov}")

                enviar_documentos(id_mov, ruta_endoso, nombre, "Endoso")
                time.sleep(1)
        #------------------------------------------
        if ba_codigo != '0' or bb_codigo != '5':
            driver.quit()

if __name__ == "__main__":
    main()
