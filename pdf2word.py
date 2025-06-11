#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Convierte un PDF en un .docx y exporta los gráficos como .png

Uso:
    python pdf2word.py informe.pdf
Salidas:
    informe.docx
    informe_img_001.png, informe_img_002.png, ...
"""
import sys, pathlib, re
import fitz  # PyMuPDF
from docx import Document

IMG_TEMPLATE = "{stem}_img_{:03d}.png"

def extract_images(page, stem, img_counter):
    """Extrae todas las imágenes de una página y las guarda en PNG."""
    for img_index, img in enumerate(page.get_images(full=True)):
        xref = img[0]
        pix = fitz.Pixmap(page.parent, xref)
        if pix.alpha:                         # convierte a RGB si tiene canal alfa
            pix = fitz.Pixmap(fitz.csRGB, pix)
        img_name = IMG_TEMPLATE.format(img_counter, stem=stem)
        pix.save(img_name)
        img_counter += 1
    return img_counter

def pdf_to_docx(pdf_path):
    pdf_path   = pathlib.Path(pdf_path)
    stem       = pdf_path.stem
    docx_path  = pdf_path.with_suffix(".docx")

    doc        = Document()
    pdf        = fitz.open(pdf_path)
    img_count  = 1

    for page in pdf:
        # Añade el texto (sin formato) respetando saltos de línea
        text = page.get_text("text")
        for line in text.strip().splitlines():
            doc.add_paragraph(line)

        # Extrae imágenes
        img_count = extract_images(page, stem, img_count)

        # Salto de página en el .docx para coincidir con el PDF
        doc.add_page_break()

    doc.save(docx_path)
    pdf.close()
    print(f"✓ Creado: {docx_path}")
    print(f"✓ Imágenes exportadas: {img_count-1}")

if __name__ == "__main__":
    if len(sys.argv) != 2 or not sys.argv[1].lower().endswith(".pdf"):
        sys.exit("Uso: python pdf2word.py archivo.pdf")
    pdf_to_docx(sys.argv[1])
