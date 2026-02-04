import os
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit
import logging
from dateutil import parser
from Tiempo.fechas_horas import get_fecha_hoy

def generarConstanciaInCrecer(ruta_archivos_x_inclu,palabra_clave,tipo_proceso,nombre_cliente,ramo):

    encabezado = (
    "Crecer Seguros S.A. Compañía de Seguros – RUC: 20600098633<br>"
    "Av. Jorge Basadre 310, piso 2, San Isidro, Lima – Perú<br>"
    "T: Lima (01) 4174400 / Provincia (0801) 17440<br>"
    "gestionalcliente@crecerseguros.pe"
    )

    if tipo_proceso == "IN":
        palabra_clave_verbo = "incluir"
        pla_clave_plural = "Incluidos"
    else:
        palabra_clave_verbo = "renovar"
        pla_clave_plural = "Renovados"

    descripcion = f"Por medio del presente endoso se deja constancia que, a solicitud del Contratante, se procede a {palabra_clave_verbo} al (los) siguientes Asegurados a partir del"

    ramo = "Vida Ley"

    titulo = f"ENDOSO DE {palabra_clave.upper()} DE ASEGURADOS"
    subtitulo = f"ASEGURADOS {pla_clave_plural.upper()}:"

    try:
        meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        dia = get_fecha_hoy().day
        mes = meses[get_fecha_hoy().month - 1]
        anio = get_fecha_hoy().year
        fecha_formateada = f"Lima, {dia} de {mes} del {anio}"
    except:
        fecha_formateada = get_fecha_hoy().strftime("Lima, %d de %B del %Y").capitalize()

    # # Avanzas un mes, y fijas el día en 1
    # primer_dia_mes_siguiente = (get_fecha_hoy() + relativedelta(months=1)).replace(day=1)
    
    # # Formato dd/mm/yyyy
    # fecha_actual_str = get_fecha_hoy().strftime("%d/%m/%Y")
    # fecha_siguiente_str = primer_dia_mes_siguiente.strftime("%d/%m/%Y")
    
    # Resultado final
    rango_fechas = f"{ramo.f_inicio} al {ramo.f_fin}"

    datos_cabecera = {
        "Producto": f"{ramo.upper()}",
        "Contratante": f"{nombre_cliente}",
        "Tramite": f"{palabra_clave.upper()} DE ASEGURADOS"
    }

    # 1. Leer el Excel con pandas
    ruta_entrada = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")

    # Lee el Excel (tus datos en tabla)
    df = pd.read_excel(ruta_entrada, engine="openpyxl",dtype={"NumDocumento": str})

    columnas_usar = ["Nombre1", "Nombre2", "ApPaterno","ApMaterno", "NumDocumento"]
    df_filtrado  = df[columnas_usar].copy()

    # Crear una nueva columna 'Nombres' que concatena solo si Nombre2 no está vacío
    df_filtrado["Nombres"] = df_filtrado.apply(
        lambda row: row["Nombre1"] if pd.isna(row["Nombre2"]) or str(row["Nombre2"]).strip() == "" 
        else f"{row['Nombre1']} {row['Nombre2']}",
        axis=1
    )

    # Agregar columna incremental (empezando desde 1)
    df_filtrado.insert(0, "N°", range(1, len(df_filtrado) + 1))

    datos_tabla = df_filtrado.to_dict(orient="records")

   # Ruta base de plantilla e imágenes
    ruta_plantilla = "/app/Plantillas/Crecer/Inclusiones"
    ruta_html      = os.path.join(ruta_plantilla, "temp_constancia.html")

    env = Environment(loader=FileSystemLoader(ruta_plantilla))
    template = env.get_template("crecerIn.html")

    html_renderizado = template.render(
        poliza=ramo.poliza,
        producto=datos_cabecera["Producto"],
        contratante=datos_cabecera["Contratante"],
        tramite=datos_cabecera["Tramite"],
        rango_fech = rango_fechas,
        fecha_texto = fecha_formateada,
        datos_tabla=datos_tabla,
        logo = os.path.join(ruta_plantilla, "logo_crecer.jpg"),
        firma_ger_op = os.path.join(ruta_plantilla, "ger_ope_vl_crecer.jpg"),
        firma_vic = os.path.join(ruta_plantilla, "vice_com_vl_crecer.jpg"),
        encabezado_logo = encabezado,
        descripcion_html = descripcion,
        titulo_html = titulo,
        subtitulo_html = subtitulo
    )

    # 👇 Aquí el cambio: guarda en ruta_html
    with open(ruta_html, "w", encoding="utf-8") as f:
        f.write(html_renderizado)

    # ------- Cargar plantilla header.html --------
    header_template = env.get_template("header.html")

    header_renderizado = header_template.render(
        logo = "file://" + os.path.join(ruta_plantilla, "logo_crecer.jpg"),
        encabezado_logo=encabezado
    )

    # Guardar header_renderizado en un archivo temporal
    ruta_header = os.path.join(ruta_plantilla, "temp_header.html")
    with open(ruta_header, "w", encoding="utf-8") as f:
        f.write(header_renderizado)
    # -------------------------------------------

    # ----- Cargar plantilla footer.html-------
    footer_template = env.get_template("footer.html")

    # Renderizar con variables dinámicas
    footer_renderizado = footer_template.render()

    # Guardar en archivo temporal
    ruta_footer = os.path.join(ruta_plantilla, "temp_footer.html")
    with open(ruta_footer, "w", encoding="utf-8") as f:
        f.write(footer_renderizado)
    # -------------------------------------------

    # pypandoc.convert_file("temp_constancia.html", "docx", outputfile="Constancia_Final.docx")
    # print("✅ Documento Word generado como 'Constancia_Final.docx'")

    # Convierte HTML a PDF
    options = {
    "enable-local-file-access": "",
    "header-html": os.path.abspath(ruta_header),
    "footer-html": os.path.abspath(ruta_footer),
    # Márgenes generales del contenido
    "margin-top": "30mm",      # espacio para el header (arriba) 15
    "margin-bottom": "25mm",   # espacio para el footer (abajo)
    "margin-left": "15mm",     # espacio a la izquierda
    "margin-right": "15mm",    # espacio a la derecha
    "header-spacing": "5",     # espacio entre encabezado y contenido
    "footer-spacing": "5",     # espacio entre contenido y pie de página
    # Otros
    "encoding": "UTF-8",
    }

    # Ruta a la carpeta Descargas del usuario
    salida_pdf = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
    # ⬇️ Guardamos el pdf
    pdfkit.from_file(ruta_html, salida_pdf, options=options)

    logging.info(f"✅ Documento PDF generado en: {ruta_archivos_x_inclu}")

def generarConstanciaReCrecer(ruta_archivos_x_inclu,palabra_clave,tipo_proceso,nombre_cliente,ruc_empresa,ramo):

    encabezado = (
    "Lima: (01) 417 4400<br>"
    "Provincias: (0801) 17440<br>"
    "gestionalcliente@crecerseguros.pe<br>"
    "Av. Jorge Basadre 310, Piso 2, San Isidro - Lima"
    )

    # if tipo_proceso == "IN":
    #     palabra_clave_verbo = "incluir"
    #     pla_clave_plural = "Incluidos"
    # else:
    #     palabra_clave_verbo = "renovar"
    #     pla_clave_plural = "Renovados"

    descripcion = f"Por medio de la presente constancia indicamos la relación de Asegurados que integran la póliza Vida Ley en la vigencia del"
    #ramo = "Vida Ley"

    titulo = f"CONSTANCIA - SEGURO VIDA LEY TRABAJADORES"
    subtitulo = f"CODIGO SBS N° VI1787300005"

    # # Todos estas formas funcionan para hacer la conversion
    # logging.info(f"📌 Vigencia Inicio: {vigInPoliza}, Vigencia Fin: {vigFinPoliza}")
    
    # # 1era forma: Convertir string ISO a datetime
    # fecha_inicio = datetime.fromisoformat(vigInPoliza)
    # fecha_fin = datetime.fromisoformat(vigFinPoliza)
    # fecha_inicio_str = fecha_inicio.strftime("%Y-%m-%d")
    # fecha_fin_str = fecha_fin.strftime("%d/%m/%Y")
    # logging.info(f"📌1 Vigencia Inicio: {fecha_inicio_str}, Vigencia Fin: {fecha_fin_str}")

    # # 2da forma: Sin librerias externas
    # fecha_inicio_str2 = datetime.strptime(vigInPoliza, "%Y-%m-%dT%H:%M:%S").strftime("%Y-%m-%d")
    # fecha_fin_str2 = datetime.strptime(vigFinPoliza, "%Y-%m-%dT%H:%M:%S").strftime("%d/%m/%Y")
    # logging.info(f"📌2 Vigencia Inicio: {fecha_inicio_str2}, Vigencia Fin: {fecha_fin_str2}")

    # 3era forma: Parsear
    #fecha_ini_ren = parser.parse(vigInPoliza).strftime("%d/%m/%Y")
    #fecha_fin_ren = parser.parse(vigFinPoliza).strftime("%d/%m/%Y")
    #logging.info(f"📌3 Vigencia Inicio: {fecha_ini_ren}, Vigencia Fin: {fecha_fin_ren}")

    # 🅰️ OPCIÓN 1:lo hacemos manual
    try:
        meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Setiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        dia = get_fecha_hoy().day
        mes = meses[get_fecha_hoy().month - 1]
        anio = get_fecha_hoy().year
        fecha_formateada = f"Lima, {dia} de {mes} del {anio}"
    except:
        # 🅱️ OPCIÓN 2: si el manul falla, lo hacemos con el local español
        fecha_formateada = get_fecha_hoy().strftime("Lima, %d de %B del %Y").capitalize()

    # Avanzas un mes, y fijas el día en 1
    #primer_dia_mes_siguiente = (get_fecha_hoy() + relativedelta(months=1)).replace(day=1)
    
    # Formato dd/mm/yyyy
    #fecha_actual_str = get_fecha_hoy().strftime("%d/%m/%Y")
    #fecha_siguiente_str = primer_dia_mes_siguiente.strftime("%d/%m/%Y")
    
    # Resultado final
    rango_fechas = f"{ramo.f_inicio} al {ramo.f_fin}"

    datos_cabecera = {
        "Producto": f"{ramo.upper()}",
        "Contratante": f"{nombre_cliente}",
        "Tramite": f"{palabra_clave.upper()}",
        "RUC": f"{ruc_empresa}",
        "Sede": f"{ramo.sede}"
    }

    # 1. Leer el Excel con pandas
    ruta_entrada = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.xlsx")

    # Lee el Excel (tus datos en tabla)
    df = pd.read_excel(ruta_entrada, engine="openpyxl",dtype={"NumDocumento": str})

    columnas_usar = ["Nombre1", "Nombre2", "ApPaterno","ApMaterno", "NumDocumento"]
    df_filtrado  = df[columnas_usar].copy()

    # Reemplazar NaN en 'Nombre2' por vacío
    df_filtrado["Nombre2"] = df_filtrado["Nombre2"].fillna("")

    # Agregar columna incremental (empezando desde 1)
    df_filtrado.insert(0, "N°", range(1, len(df_filtrado) + 1))

    datos_tabla = df_filtrado.to_dict(orient="records")

    # Ruta base de plantilla e imágenes
    ruta_plantilla = "/app/Plantillas/Crecer/Renovaciones"
    ruta_html      = os.path.join(ruta_plantilla, "temp_constancia.html")

    env = Environment(loader=FileSystemLoader(ruta_plantilla))
    template = env.get_template("crecerRe.html")

    html_renderizado = template.render(
        poliza=ramo.poliza,
        producto=datos_cabecera["Producto"],
        contratante=datos_cabecera["Contratante"],
        tramite=datos_cabecera["Tramite"],
        ruc=datos_cabecera["RUC"],
        sede=datos_cabecera["Sede"],
        rango_fech = rango_fechas,
        fecha_texto = fecha_formateada,
        datos_tabla=datos_tabla,
        logo = os.path.join(ruta_plantilla, "crecer_re_logo.jpg"),
        firma_ger_op = os.path.join(ruta_plantilla, "ger_ope_vl_crecer.jpg"),
        firma_vic = os.path.join(ruta_plantilla, "vice_com_vl_crecer.jpg"),
        encabezado_logo = encabezado,
        descripcion_html = descripcion,
        titulo_html = titulo,
        subtitulo_html = subtitulo,
        iconos_redes = os.path.join(ruta_plantilla, "redesSocialesCrecer.jpg")
    )

    # 👇 Aquí el cambio: guarda en ruta_html
    with open(ruta_html, "w", encoding="utf-8") as f:
        f.write(html_renderizado)

    # ------- Cargar plantilla header.html --------
    header_template = env.get_template("header.html")

    header_renderizado = header_template.render(
        logo = "file://" + os.path.join(ruta_plantilla, "crecer_re_logo.jpg"),
        encabezado_logo=encabezado
    )

    # Guardar header_renderizado en un archivo temporal
    ruta_header = os.path.join(ruta_plantilla, "temp_header.html")
    with open(ruta_header, "w", encoding="utf-8") as f:
        f.write(header_renderizado)
    # -------------------------------------------

    # ----- Cargar plantilla footer.html-------
    footer_template = env.get_template("footer.html")

    # Renderizar con variables dinámicas
    footer_renderizado = footer_template.render(
        img_footer = "file://" + os.path.join(ruta_plantilla, "img_footer.jpg"),
        redesS_footer = "file://" + os.path.join(ruta_plantilla, "redesSocialesCrecer.jpg"),
    )

    # Guardar en archivo temporal
    ruta_footer = os.path.join(ruta_plantilla, "temp_footer.html")
    with open(ruta_footer, "w", encoding="utf-8") as f:
        f.write(footer_renderizado)
    # -------------------------------------------

    # pypandoc.convert_file("temp_constancia.html", "docx", outputfile="Constancia_Final.docx")
    # print("✅ Documento Word generado como 'Constancia_Final.docx'")

    # Convierte HTML a PDF
    options = {
    "enable-local-file-access": "",
    "header-html": os.path.abspath(ruta_header),
    "footer-html": os.path.abspath(ruta_footer),
    # Márgenes generales del contenido
    "margin-top": "30mm",      # espacio para el header (arriba) 15
    "margin-bottom": "25mm",   # espacio para el footer (abajo)
    "margin-left": "15mm",     # espacio a la izquierda
    "margin-right": "15mm",    # espacio a la derecha
    "header-spacing": "5",     # espacio entre encabezado y contenido
    "footer-spacing": "5",     # espacio entre contenido y pie de página
    # Otros
    "encoding": "UTF-8",
    }

    # Ruta a la carpeta Descargas del usuario
    salida_pdf = os.path.join(ruta_archivos_x_inclu, f"{ramo.poliza}.pdf")
    # Guardarmos el pdf
    pdfkit.from_file(ruta_html, salida_pdf, options=options)

    logging.info(f"✅ Documento PDF generado en: {ruta_archivos_x_inclu}")