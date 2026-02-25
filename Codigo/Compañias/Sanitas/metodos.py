# -*- coding: utf-8 -*-
# -- Froms ---
from reportlab.platypus import SimpleDocTemplate, Image # type: ignore
from reportlab.lib.pagesizes import A4 # type: ignore
from reportlab.lib import utils # type: ignore
# -- Imports --
import logging

def imagen_a_pdf(ruta_imagen, ruta_pdf):
    doc = SimpleDocTemplate(ruta_pdf, pagesize=A4)
    elementos = []

    img = utils.ImageReader(ruta_imagen)
    ancho_img, alto_img = img.getSize()

    # Ajustar imagen al tamaño A4 manteniendo proporción
    ancho_max, alto_max = A4
    proporcion = min(ancho_max / ancho_img, alto_max / alto_img)

    img_reportlab = Image(
        ruta_imagen,
        width=ancho_img * proporcion,
        height=alto_img * proporcion
    )

    elementos.append(img_reportlab)
    doc.build(elementos)

    logging.info(f"✅ PDF generado en: {ruta_pdf}")