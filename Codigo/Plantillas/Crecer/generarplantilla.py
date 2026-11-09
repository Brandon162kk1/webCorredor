from jinja2 import Environment, FileSystemLoader
from Tiempo.fechas_horas import get_fecha_hoy
from weasyprint import HTML # type: ignore
#import pdfkit
import logging
import time
import os
import pandas as pd
import subprocess
from docxtpl import DocxTemplate

# Apagar completamente logs de weasyprint y fonttools
logging.getLogger("weasyprint").handlers = []
logging.getLogger("weasyprint").propagate = False
logging.getLogger("weasyprint").setLevel(logging.CRITICAL)

logging.getLogger("fontTools").handlers = []
logging.getLogger("fontTools").propagate = False
logging.getLogger("fontTools").setLevel(logging.CRITICAL)

# Opcional: bajar nivel global
logging.basicConfig(level=logging.CRITICAL)

def fecha_lima_formateada():
    try:
        meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        hoy = get_fecha_hoy()
        return f"Lima, {hoy.day} de {meses[hoy.month - 1]} del {hoy.year}"
    except:
        return get_fecha_hoy().strftime("Lima, %d de %B del %Y").capitalize()

def obtener_rango_vigencia(ramo):

    return f"{ramo.f_inicio} al {ramo.f_fin}"

def leer_asegurados_excel(ruta_excel, incluir_nombres=True):

    df = pd.read_excel(ruta_excel, engine="openpyxl", dtype={"N° Documento": str})

    df.columns = df.columns.str.strip()
    df = df.fillna("")

    columnas = ["Nombre 1", "Nombre 2", "Ap. Paterno", "Ap. Materno", "N° Documento"]
    df = df[columnas].copy()

    if incluir_nombres:
        df["Nombres"] = df.apply(
            lambda r: r["Nombre 1"] if pd.isna(r["Nombre 2"]) or str(r["Nombre 2"]).strip() == ""
            else f"{r['Nombre 1']} {r['Nombre 2']}",
            axis=1
        )
    else:
        df["Nombre 2"] = df["Nombre 2"].fillna("")

    df.insert(0, "N°", range(1, len(df) + 1))
    return df.to_dict(orient="records")

def generarConstanciaInReCrecer(ruta_archivos_x_inclu,palabra_clave,nombre_cliente,ruc,ramo,tipo_proceso):

    tipoError = ""
    detalleError = ""
    constancia = True

    try:
        rango = obtener_rango_vigencia(ramo)

        datos_tabla = leer_asegurados_excel(os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx"),False if tipo_proceso == 'RE' else True)

        ruta_plantilla_base = "/app/Codigo/Plantillas/Crecer"

        if tipo_proceso == 'RE':

            ruta_plantilla = "/app/Codigo/Plantillas/Crecer/Renovaciones"
            encabezado = (
                "Lima: (01) 417 4400<br>"
                "Provincias: (0801) 17440<br>"
                "gestionalcliente@crecerseguros.pe<br>"
                "Av. Ricardo Rivera Navarrete 501 oficina 1101-A, San Isidro - Lima"
            )
            titulo = "CONSTANCIA - SEGURO VIDA LEY TRABAJADORES"
            tramite = f"{palabra_clave.upper()}"
            descripcion = """Por medio de la presente constancia indicamos la relación de Asegurados que integran la póliza vida ley en la vigencia del"""

        else:

            ruta_plantilla = "/app/Codigo/Plantillas/Crecer/Inclusiones"
            encabezado = (
                "Crecer Seguros S.A. Compañía de Seguros – RUC: 20600098633<br>"
                "Av. Ricardo Rivera Navarrete 501 oficina 1101-A, San Isidro, Lima – Perú<br>"
                "T: Lima (01) 4174400 / Provincia (0801) 17440<br>"
                "gestionalcliente@crecerseguros.pe"
            )
            titulo = "ENDOSO DE INCLUSIÓN DE ASEGURADOS"
            tramite = f"{palabra_clave.upper().capitalize()} de Asegurados"
            descripcion = """Por medio del presente endoso se deja constancia que, a solicitud del Contratante, se procede a incluir al (los) siguientes Asegurados a partir de"""

        contexto = {
            "poliza": ramo.poliza,
            "contratante": nombre_cliente,
            "tramite": tramite,
            "rango_fech": rango,
            "fecha_texto": fecha_lima_formateada(),
            "datos_tabla": datos_tabla,
            "titulo_html": titulo,
            "encabezado_texto": encabezado,
            "descripcion_html": descripcion,

            # Rutas sin file://
            "logo": os.path.join(ruta_plantilla_base, "crecer_re_logo.jpg"),
            "firma_ger_ope": os.path.join(ruta_plantilla_base, "gerente_operaciones.jpg"),
            "firma_ger_com": os.path.join(ruta_plantilla_base, "gerente_comercial.jpg")
        }

        if tipo_proceso == 'RE':
            contexto["ruc"] = ruc
            contexto["subtitulo_html"] = "CODIGO SBS N° VI1787300005"
            contexto["sede"] = ramo.sede
            contexto["img_footer"] = os.path.join(ruta_plantilla_base, "img_footer.jpg")
            contexto["redes_footer"] = os.path.join(ruta_plantilla_base, "redesSocialesCrecer.jpg")
        else:
            contexto["producto"] = "Vida Ley Trabajadores"

        # Cargar plantilla
        env = Environment(loader=FileSystemLoader(ruta_plantilla))

        if tipo_proceso == 'RE':
            template = env.get_template("crecerReNew.html")
        else:
            template = env.get_template("crecerInNew.html")

        html_renderizado = template.render(contexto)

        salida_pdf = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")

        # Generar PDF
        HTML(string=html_renderizado,base_url=ruta_plantilla_base).write_pdf(salida_pdf)
        logging.info("✅ PDF generado correctamente")

        time.sleep(2)

        ruta_vl = os.path.join(ruta_archivos_x_inclu,f"{ramo.poliza}.pdf")

        if not os.path.exists(ruta_vl):
            raise Exception(f"No se genero la constancia manual")

    except Exception as e:
        detalleError = str(e)
        constancia = False
        logging.error(f"❌ Error en Crecer (VL) - MV: {detalleError}")
        tipoError = f"CREC-VL-MV"
    finally:
        return constancia,False,tipoError,detalleError

def generarConstanciaWordCrecer(ruta_archivos_x_inclu, palabra_clave, nombre_cliente, ruc, ramo, tipo_proceso):
    """
    Genera la constancia de Crecer Seguros utilizando la plantilla Word (.docx)
    y convirtiéndola a PDF con LibreOffice en modo headless.
    """
    tipoError = ""
    detalleError = ""
    constancia = True

    try:
        f_ini = getattr(ramo, 'f_inicio', '')
        f_fin = getattr(ramo, 'f_fin', '')
        sede = getattr(ramo, 'sede', '')
        
        # Leer asegurados del Excel
        ruta_excel = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")
        df_raw = pd.read_excel(ruta_excel, engine="openpyxl", dtype={"N° Documento": str})
        df_raw.columns = df_raw.columns.str.strip()
        df_raw = df_raw.fillna("")

        datos_tabla = []
        for _, row in df_raw.iterrows():
            datos_tabla.append({
                "paterno": str(row.get("Ap. Paterno", "")).strip(),
                "materno": str(row.get("Ap. Materno", "")).strip(),
                "nombre1": str(row.get("Nombre 1", "")).strip(),
                "nombre2": str(row.get("Nombre 2", "")).strip(),
                "documento": str(row.get("N° Documento", "")).strip()
            })

        # Rutas de plantilla y salida
        ruta_plantilla_base = "/app/Codigo/Plantillas/Crecer"
        if not os.path.exists(ruta_plantilla_base):
            ruta_plantilla_base = os.path.dirname(os.path.abspath(__file__))

        ruta_plantilla_docx = os.path.join(ruta_plantilla_base, "plantilla_crecer_base.docx")
        if not os.path.exists(ruta_plantilla_docx):
            raise Exception(f"No se encontró la plantilla Word en: {ruta_plantilla_docx}")

        tramite = f"{palabra_clave.upper().capitalize()}"

        contexto = {
            "fecha_texto": fecha_lima_formateada().replace("Lima, ", "").strip(),
            "poliza": str(ramo.poliza),
            "contratante": str(nombre_cliente),
            "ruc": str(ruc) if ruc else "",
            "tramite": tramite,
            "sede": str(sede),
            "f_inicio": str(f_ini),
            "f_fin": str(f_fin),
            "datos_tabla": datos_tabla
        }

        # 1. Renderizar plantilla Word
        doc = DocxTemplate(ruta_plantilla_docx)
        doc.render(contexto)

        ruta_docx_temporal = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.docx")
        doc.save(ruta_docx_temporal)
        logging.info(f"✅ Word temporal generado: {ruta_docx_temporal}")

        # 2. Convertir a PDF mediante LibreOffice headless
        cmd = [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", ruta_archivos_x_inclu,
            ruta_docx_temporal
        ]
        
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=60)
        
        if proc.returncode != 0:
            logging.error(f"Error en LibreOffice: {proc.stderr}")
            raise Exception(f"Error convirtiendo Word a PDF con LibreOffice: {proc.stderr}")

        salida_pdf = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")

        time.sleep(2)

        if not os.path.exists(salida_pdf):
            raise Exception(f"No se generó el PDF de la constancia en {salida_pdf}")

        # Eliminar docx temporal si se generó correctamente
        # try:
        #     if os.path.exists(ruta_docx_temporal):
        #         os.remove(ruta_docx_temporal)
        # except Exception:
        #     pass

        logging.info("✅ PDF generado correctamente desde Word con LibreOffice")

    except Exception as e:
        detalleError = str(e)
        constancia = False
        logging.error(f"❌ Error en Crecer Word (VL) - MV: {detalleError}")
        tipoError = f"CREC-VL-MV"
    finally:
        return constancia, False, tipoError, detalleError
