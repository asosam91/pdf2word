#!/usr/bin/env python3
"""pdf_to_word_images.py
Convert a PDF into a Word document and export embedded images as PNG files.
Usage:
    python pdf_to_word_images.py input_document.pdf
Outputs:
    input_document.docx  (same base name as the PDF)
    input_document_p<page>_img<index>.png  (one PNG per image found)

Dependencies:
    pip install pdf2docx PyMuPDF
"""
from pathlib import Path
import argparse

from pdf2docx import Converter  # type: ignore
import fitz  # PyMuPDF


def pdf_to_word(pdf_path: Path) -> Path:
    """Convert *pdf_path* into a DOCX with the same base name."""
    word_path = pdf_path.with_suffix(".docx")
    print(f"[+] Converting {pdf_path.name} -> {word_path.name} …")
    cv = Converter(str(pdf_path))
    cv.convert(str(word_path))
    cv.close()
    return word_path


def extract_images(pdf_path: Path) -> list[str]:
    """Extract all raster images from *pdf_path* and save them as PNG.

    Images are named like <base>_p<page>_img<index>.png to avoid collisions.
    Returns a list with the filenames of extracted images.
    """
    doc = fitz.open(pdf_path)
    out_images: list[str] = []
    stem = pdf_path.stem

    for page_index in range(len(doc)):
        page = doc[page_index]
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list, start=1):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            # Ensure RGB (drop transparency if any)
            if pix.alpha:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            img_name = f"{stem}_p{page_index + 1}_img{img_index}.png"
            print(f"    • Saving {img_name}")
            pix.save(img_name)
            out_images.append(img_name)

    print(f"[+] Extracted {len(out_images)} images from {pdf_path.name}")
    return out_images


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert a PDF to Word (.docx) and export embedded images as PNG files."
    )
    parser.add_argument("pdf", type=Path, help="Path to input PDF file")
    args = parser.parse_args()

    pdf_path: Path = args.pdf.resolve()
    if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
        parser.error("Input must be an existing .pdf file")

    word_file = pdf_to_word(pdf_path)
    extract_images(pdf_path)
    print("[✓] Done. Output files saved next to the original PDF:")
    print(f"    - {word_file.name}")
    print("    - PNG images as listed above")


if __name__ == "__main__":
    main()
