# -*- coding: utf-8 -*-
r"""Convierte todos los PowerPoint (.ppt/.pptx/.pps/.ppsx) y Word (.doc/.docx) de una carpeta a PDF (en la misma carpeta).

Uso:
    python convertir_ppt_a_pdf.py                # usa la carpeta de este script
    python convertir_ppt_a_pdf.py "C:\ruta\carpeta"

Requiere Microsoft PowerPoint y Word instalados y `pip install pywin32`.
"""
import os
import sys

import win32com.client

PPT_EXTS = (".ppt", ".pptx", ".pps", ".ppsx")
WORD_EXTS = (".doc", ".docx")
SAVE_AS_PDF = 32       # PowerPoint: ppSaveAsPDF
WORD_SAVE_AS_PDF = 17  # Word: wdFormatPDF


def convertir(carpeta):
    if not os.path.isdir(carpeta):
        print(f"No existe la carpeta: {carpeta}")
        return 1

    archivos = sorted(
        f for f in os.listdir(carpeta)
        if f.lower().endswith(PPT_EXTS + WORD_EXTS) and not f.startswith("~$")
    )
    if not archivos:
        print("No se encontraron archivos de PowerPoint ni de Word en la carpeta.")
        return 0

    ppt = None
    word = None
    errores = 0
    try:
        for nombre in archivos:
            origen = os.path.join(carpeta, nombre)
            salida = os.path.join(carpeta, os.path.splitext(nombre)[0] + ".pdf")
            try:
                if nombre.lower().endswith(WORD_EXTS):
                    if word is None:
                        word = win32com.client.Dispatch("Word.Application")
                        word.DisplayAlerts = 0  # sin diálogos (conversión, uso, etc.)
                    doc = word.Documents.Open(
                        FileName=origen, ConfirmConversions=False,
                        ReadOnly=True, AddToRecentFiles=False,
                    )
                    try:
                        doc.SaveAs(FileName=salida, FileFormat=WORD_SAVE_AS_PDF)
                        print(f"OK: {nombre}  ->  {os.path.basename(salida)}")
                    finally:
                        doc.Close(False)
                else:
                    if ppt is None:
                        ppt = win32com.client.Dispatch("PowerPoint.Application")
                    presentacion = ppt.Presentations.Open(
                        origen, ReadOnly=True, Untitled=False, WithWindow=False
                    )
                    try:
                        presentacion.SaveAs(salida, SAVE_AS_PDF)
                        print(f"OK: {nombre}  ->  {os.path.basename(salida)}")
                    finally:
                        presentacion.Close()
            except Exception as e:
                errores += 1
                print(f"ERROR: {nombre}  ({e})")
    finally:
        if word is not None:
            word.Quit()
        if ppt is not None:
            ppt.Quit()

    print(f"Listo: {len(archivos) - errores} de {len(archivos)} convertidos.")
    return 1 if errores else 0


if __name__ == "__main__":
    carpeta = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    sys.exit(convertir(carpeta))
