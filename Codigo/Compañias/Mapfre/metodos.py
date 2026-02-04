# -*- coding: utf-8 -*-
# -- Froms ---
# -- Imports --
import logging
import pdfplumber
import re

def obtener_datos_proforma_Sanitas_SCTR_Mapfre(proforma_pdf):
    with pdfplumber.open(proforma_pdf) as pdf:
        texto_total = ""
        for pagina in pdf.pages:
            extraido = pagina.extract_text()
            if extraido:
                texto_total += extraido + "\n"

        if not texto_total.strip():
            logging.info("⚠️ No se extrajo texto del PDF.")
            return {"proformas":[], "codigo_pago": None, "expiracion": None}

        # 1 EXTRAER RECIBO : XXXXXXX
        match_recibo = re.search(r"RECIBO\s*:\s*(\S+)", texto_total)
        recibo = match_recibo.group(1) if match_recibo else None

        # 2 EXTRAER Prima Comercial + IGV : XX.X
        match_prima_igv = re.search(r"Prima Comercial \+ IGV\s*:\s*([0-9.]+)", texto_total)
        prima_igv = match_prima_igv.group(1) if match_prima_igv else None

        # 63 EXTRAER Último día de Pago : dd/mm/yyyy
        match_ultimo_pago = re.search(r"Último día de Pago\s*:\s*([0-9/]+)", texto_total)
        ultimo_dia_pago = match_ultimo_pago.group(1) if match_ultimo_pago else None

    return {
        "proformas": recibo,
        "monto": prima_igv,
        "expiracion": ultimo_dia_pago
    }