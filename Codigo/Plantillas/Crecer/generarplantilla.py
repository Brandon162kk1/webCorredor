import os
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit
import logging
from Tiempo.fechas_horas import get_fecha_hoy

# Apagar completamente logs de weasyprint y fonttools
logging.getLogger("weasyprint").handlers = []
logging.getLogger("weasyprint").propagate = False
logging.getLogger("weasyprint").setLevel(logging.CRITICAL)

logging.getLogger("fontTools").handlers = []
logging.getLogger("fontTools").propagate = False
logging.getLogger("fontTools").setLevel(logging.CRITICAL)

# Opcional: bajar nivel global
logging.basicConfig(level=logging.CRITICAL)

from weasyprint import HTML # type: ignore

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

def renderizar_pdf1(ruta_plantilla,template_html,contexto,salida_pdf,header_logo,encabezado,footer_vars=None):

    env = Environment(loader=FileSystemLoader(ruta_plantilla))

    # HTML principal
    template = env.get_template(template_html)
    ruta_html = os.path.join(ruta_plantilla, "temp_constancia.html")
    with open(ruta_html, "w", encoding="utf-8") as f:
        f.write(template.render(**contexto))
    
    # Header
    header = env.get_template("header.html").render(
        logo="file://" + header_logo,
        encabezado_logo=encabezado
    )
    ruta_header = os.path.join(ruta_plantilla, "temp_header.html")
    with open(ruta_header, "w", encoding="utf-8") as f:
        f.write(header)

    # Footer
    footer = env.get_template("footer.html").render(**(footer_vars or {}))
    ruta_footer = os.path.join(ruta_plantilla, "temp_footer.html")
    with open(ruta_footer, "w", encoding="utf-8") as f:
        f.write(footer)

    options = {
        "enable-local-file-access": None,
        "header-html": os.path.abspath(ruta_header),
        "footer-html": os.path.abspath(ruta_footer),
        "margin-top": "60mm",
        "margin-bottom": "25mm",
        "margin-left": "15mm",
        "margin-right": "15mm",
        "header-spacing": "5",
        "footer-spacing": "5",
        "encoding": "UTF-8",
    }

    pdfkit.from_file(ruta_html, salida_pdf, options=options)
    logging.info(f"✅ Documento PDF generado")

def renderizar_pdf(ruta_plantilla, template_html, contexto, salida_pdf):

    env = Environment(loader=FileSystemLoader(ruta_plantilla))

    # Renderizar HTML principal
    template = env.get_template(template_html)
    html_renderizado = template.render(**contexto)

    # Generar PDF
    HTML(
        string=html_renderizado,
        base_url=ruta_plantilla  # 🔥 MUY IMPORTANTE
    ).write_pdf(salida_pdf)

    print("✅ PDF generado correctamente")

def generarConstanciaInCrecer1(ruta_archivos_x_inclu,palabra_clave,nombre_cliente,ramo):

    rango = obtener_rango_vigencia(ramo)
    datos_tabla = leer_asegurados_excel(os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx"))

    encabezado = (
    "Crecer Seguros S.A. Compañía de Seguros – RUC: 20600098633<br>"
    "Av. Jorge Basadre 310, piso 2, San Isidro, Lima – Perú<br>"
    "T: Lima (01) 4174400 / Provincia (0801) 17440<br>"
    "gestionalcliente@crecerseguros.pe"
    )

    ruta_plantilla = "/app/Codigo/Plantillas/Crecer/Inclusiones"

    contexto = {
        "poliza": ramo.poliza,
        "producto": "VIDA LEY",
        "contratante": nombre_cliente,
        "tramite": f"{palabra_clave.upper()} DE ASEGURADOS",
        "rango_fech": rango,
        "fecha_texto": fecha_lima_formateada(),
        "datos_tabla": datos_tabla,
        "titulo_html": f"ENDOSO DE {palabra_clave.upper()} DE ASEGURADOS",
        "subtitulo_html": "ASEGURADOS INCLUIDOS:",
        "encabezado_logo": encabezado,
        "descripcion_html": f"""Por medio del presente endoso se deja constancia que, a solicitud del Contratante, 
        se procede a incluir al (los) siguientes Asegurados a partir del""",
        "firma_ger_op": "file://" + os.path.join(ruta_plantilla, "ger_ope_vl_crecer.jpg"),
        "firma_vic": "file://" + os.path.join(ruta_plantilla, "vice_com_vl_crecer.jpg")
    }

    salida_pdf = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")

    renderizar_pdf(
        ruta_plantilla,
        "crecerIn.html",
        contexto,
        salida_pdf,
        os.path.join(ruta_plantilla, "logo_crecer.jpg"),
        contexto["encabezado_logo"]
    )

def generarConstanciaInCrecer(ruta_archivos_x_inclu, palabra_clave, nombre_cliente, ramo):

    rango = obtener_rango_vigencia(ramo)
    datos_tabla = leer_asegurados_excel(
        os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")
    )

    encabezado = (
        "Crecer Seguros S.A. Compañía de Seguros – RUC: 20600098633<br>"
        "Av. Jorge Basadre 310, piso 2, San Isidro, Lima – Perú<br>"
        "T: Lima (01) 4174400 / Provincia (0801) 17440<br>"
        "gestionalcliente@crecerseguros.pe"
    )

    ruta_plantilla = "/app/Codigo/Plantillas/Crecer/Inclusiones"

    contexto = {
        "poliza": ramo.poliza,
        "producto": "VIDA LEY",
        "contratante": nombre_cliente,
        "tramite": f"{palabra_clave.upper()} DE ASEGURADOS",
        "rango_fech": rango,
        "fecha_texto": fecha_lima_formateada(),
        "datos_tabla": datos_tabla,
        "encabezado_logo": encabezado,
        "titulo_html": f"ENDOSO DE {palabra_clave.upper()} DE ASEGURADOS",
        "subtitulo_html": "ASEGURADOS INCLUIDOS:",
        "descripcion_html": f"""Por medio del presente endoso se deja constancia que, 
        a solicitud del Contratante, se procede a incluir al (los) siguientes 
        Asegurados a partir del""",
        # 👇 Ahora SIN file://
        "firma_ger_op": os.path.join(ruta_plantilla, "ger_ope_vl_crecer.jpg"),
        "firma_vic": os.path.join(ruta_plantilla, "vice_com_vl_crecer.jpg"),
        "logo": os.path.join(ruta_plantilla, "logo_crecer.jpg")
    }

    # 🔹 Cargar plantilla Jinja
    env = Environment(loader=FileSystemLoader(ruta_plantilla))
    template = env.get_template("crecerInNew.html")
    html_renderizado = template.render(contexto)

    salida_pdf = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")

    # 🔹 Generar PDF
    HTML(string=html_renderizado, base_url=ruta_plantilla).write_pdf(salida_pdf)
    logging.info("✅ PDF generado correctamente")

def generarConstanciaReCrecer1(ruta_archivos_x_inclu,palabra_clave,nombre_cliente,ruc,ramo):
    
    rango = obtener_rango_vigencia(ramo)
    datos_tabla = leer_asegurados_excel(
        os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx"),
        incluir_nombres=False
    )

    encabezado = (
    "Lima: (01) 417 4400<br>"
    "Provincias: (0801) 17440<br>"
    "gestionalcliente@crecerseguros.pe<br>"
    "Av. Jorge Basadre 310, Piso 2, San Isidro - Lima"
    )

    ruta_plantilla = "/app/Codigo/Plantillas/Crecer/Renovaciones"

    contexto = {
        "poliza": ramo.poliza,
        "producto": "VIDA LEY",
        "contratante": nombre_cliente,
        "tramite": palabra_clave.upper(),
        "ruc": ruc,
        "sede": ramo.sede,
        "rango_fech": rango,
        "fecha_texto": fecha_lima_formateada(),
        "datos_tabla": datos_tabla,
        "titulo_html": "CONSTANCIA - SEGURO VIDA LEY TRABAJADORES",
        "subtitulo_html": "CODIGO SBS N° VI1787300005",
        "encabezado_logo": encabezado,
        "descripcion_html": f"""Por medio del presente endoso se deja constancia que, a solicitud del Contratante,
            se procede a renovar al (los) siguientes Asegurados a partir del""",
        "firma_ger_op": "file://" + os.path.join(ruta_plantilla, "ger_ope_vl_crecer.jpg"),
        "firma_vic": "file://" + os.path.join(ruta_plantilla, "vice_com_vl_crecer.jpg")
    }

    salida_pdf = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")

    footer_vars = {
        "img_footer": "file://" + os.path.join(ruta_plantilla, "img_footer.jpg"),
        "redesS_footer": "file://" + os.path.join(ruta_plantilla, "redesSocialesCrecer.jpg")
    }

    renderizar_pdf(
        ruta_plantilla,
        "crecerRe.html",
        contexto,
        salida_pdf,
        os.path.join(ruta_plantilla, "crecer_re_logo.jpg"),
        contexto["encabezado_logo"],
        footer_vars=footer_vars
    )

def generarConstanciaReCrecer(ruta_archivos_x_inclu,palabra_clave,nombre_cliente,ruc,ramo):

    rango = obtener_rango_vigencia(ramo)

    datos_tabla = leer_asegurados_excel(
        os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx"),
        incluir_nombres=False
    )

    ruta_plantilla = "/app/Codigo/Plantillas/Crecer/Renovaciones"

    encabezado = (
        "Lima: (01) 417 4400<br>"
        "Provincias: (0801) 17440<br>"
        "gestionalcliente@crecerseguros.pe<br>"
        "Av. Jorge Basadre 310, Piso 2, San Isidro - Lima"
    )

    contexto = {
        "poliza": ramo.poliza,
        "producto": "VIDA LEY",
        "contratante": nombre_cliente,
        "tramite": palabra_clave.upper(),
        "ruc": ruc,
        "sede": ramo.sede,
        "rango_fech": rango,
        "fecha_texto": fecha_lima_formateada(),
        "datos_tabla": datos_tabla,
        "titulo_html": "CONSTANCIA - SEGURO VIDA LEY TRABAJADORES",
        "subtitulo_html": "CODIGO SBS N° VI1787300005",
        "encabezado_logo": encabezado,
        "descripcion_html": """Por medio de la presente constancia indicamos la relación de Asegurados que integran la póliza
        vida ley en la vigencia del """,

        # 🔥 RUTAS SIN file://
        "logo": os.path.join(ruta_plantilla, "crecer_re_logo.jpg"),
        "firma_ger_op": os.path.join(ruta_plantilla, "ger_ope_vl_crecer.jpg"),
        "firma_vic": os.path.join(ruta_plantilla, "vice_com_vl_crecer.jpg"),
        "img_footer": os.path.join(ruta_plantilla, "img_footer.jpg"),
        "redes_footer": os.path.join(ruta_plantilla, "redesSocialesCrecer.jpg"),
    }

    # 🔹 Cargar plantilla
    env = Environment(loader=FileSystemLoader(ruta_plantilla))
    template = env.get_template("crecerReNew2.html")

    html_renderizado = template.render(contexto)

    salida_pdf = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")

    # 🔹 Generar PDF
    HTML(string=html_renderizado,base_url=ruta_plantilla).write_pdf(salida_pdf)
    logging.info("✅ PDF generado correctamente")