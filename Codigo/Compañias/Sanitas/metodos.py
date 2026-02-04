# -*- coding: utf-8 -*-
# -- Froms ---

# -- Imports --
import logging
import pdfplumber
import unicodedata
import re

def obtener_datos_proforma_sanitas_SCTR_Protecta(proforma_pdf):
    
    with pdfplumber.open(proforma_pdf) as pdf:
        texto_total = ""
        for pagina in pdf.pages:
            extraido = pagina.extract_text()
            if extraido:
                texto_total += extraido + "\n"

        if not texto_total.strip():
            logging.info("⚠️ No se extrajo texto del PDF.")
            return {"proformas": [], "codigo_pago": None, "expiracion": None}

        texto_total = texto_total.replace("œ", "u")
        texto_total = unicodedata.normalize("NFKD", texto_total)

        proformas = re.findall(r"\bAC-SCTR-\d+\b", texto_total, re.IGNORECASE)
        proformas_unicas = list(dict.fromkeys(proformas))

        match_pago = re.search(r"CODIGO\s*DE\s*PAGO\s*:\s*(\d+)", texto_total, re.IGNORECASE)
        codigo_pago = match_pago.group(1) if match_pago else None

        match_fecha = re.search(r"EXPIRACION\s*:\s*(\d{2}/\d{2}/\d{4})", texto_total, re.IGNORECASE)
        expiracion = match_fecha.group(1) if match_fecha else None

    return {
        "proformas": proformas_unicas,
        "codigo_pago": codigo_pago,
        "expiracion": expiracion,
    }

def obtener_datos_proforma_Sanitas_SCTR_Crecer(proforma_pdf):
    
    with pdfplumber.open(proforma_pdf) as pdf:
        texto_total = ""
        for pagina in pdf.pages:
            extraido = pagina.extract_text()
            if extraido:
                texto_total += extraido + "\n"
        
        if not texto_total.strip():
            logging.info("⚠️ No se extrajo texto del PDF.")
            return {"proformas":[], "codigo_pago": None, "expiracion": None}

        # Normaliza caracteres raros
        texto_total = texto_total.replace("œ", "u")
        texto_total = unicodedata.normalize("NFKD", texto_total)

        # 1️ Buscar proformas tipo CS-SCTR-003057831
        texto_limpio = texto_total.replace(" ", "").replace("-", "-")  # guion raro U+2011
        proformas = re.findall(r"\bCS-SCTR-\d+\b", texto_limpio, re.IGNORECASE)
        proformas_unicas = list(dict.fromkeys(proformas))

        # 2️ Buscar código de pago (números después de "CODIGO DE PAGO:")
        match_pago = re.search(r"CODIGO\s*DE\s*PAGO\s*:\s*(\d+)", texto_total, re.IGNORECASE)
        codigo_pago = match_pago.group(1) if match_pago else None

        # 3️ Buscar fecha de expiración (dd/mm/yyyy después de "EXPIRACION:")
        match_fecha = re.search(r"EXPIRACION\s*:\s*(\d{2}/\d{2}/\d{4})", texto_total, re.IGNORECASE)
        expiracion = match_fecha.group(1) if match_fecha else None

    return {
        "proformas": proformas_unicas,
        "codigo_pago": codigo_pago,
        "expiracion": expiracion,
    }

def obtener_datos_proforma_Sanitas_SCTR_Sanitas(proforma_pdf):
    
    with pdfplumber.open(proforma_pdf) as pdf:
        texto_total = ""
        for pagina in pdf.pages:
            extraido = pagina.extract_text()
            if extraido:
                texto_total += extraido + "\n"

        if not texto_total.strip():
            logging.info("⚠️ No se extrajo texto del PDF.")
            return {"proformas":[], "codigo_pago": None, "expiracion": None}

        # Normaliza caracteres raros
        texto_total = texto_total.replace("œ", "u")
        texto_total = unicodedata.normalize("NFKD", texto_total)

        # 1️ Buscar proformas tipo PF-SCTR-003057831
        texto_limpio = texto_total.replace(" ", "").replace("-", "-")  # guion raro U+2011
        proformas = re.findall(r"\bPF-SCTR-\d+\b", texto_limpio, re.IGNORECASE)
        proformas_unicas = list(dict.fromkeys(proformas))

        # 2️ Buscar código de pago (números después de "CODIGO DE PAGO:")
        match_pago = re.search(r"CODIGO\s*DE\s*PAGO\s*:\s*(\d+)", texto_total, re.IGNORECASE)
        codigo_pago = match_pago.group(1) if match_pago else None

        # 3️ Buscar fecha de expiración (dd/mm/yyyy después de "EXPIRACION:")
        match_fecha = re.search(r"EXPIRACION\s*:\s*(\d{2}/\d{2}/\d{4})", texto_total, re.IGNORECASE)
        expiracion = match_fecha.group(1) if match_fecha else None

    return {
        "proformas": proformas_unicas,
        "codigo_pago": codigo_pago,
        "expiracion": expiracion,
    }

def extraer_codigo_proforma(proforma):
    match = re.search(r'-([0-9]+)(?:/|$)', proforma)
    if match:
        return match.group(1)
    return None

def extraer_valor_a_partir_tercer_indice(proforma):
    if len(proforma) > 2:
        return proforma[2:]  # Devuelve la cadena desde el cuarto carácter (índice 3)
    return None