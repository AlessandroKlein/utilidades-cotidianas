# -*- coding: utf-8 -*-
r"""Convierte todos los .ppt y .pptx de una carpeta a PDF (en la misma carpeta).

Uso:
    python convertir_ppt_a_pdf.py                # usa la carpeta de este script
    python convertir_ppt_a_pdf.py "C:\ruta\carpeta"

Requiere Microsoft PowerPoint instalado y `pip install pywin32`.
"""
import os
import sys

import win32com.client

EXTS = (".ppt", ".pptx")
SAVE_AS_PDF = 32  # ppSaveAsPDF


def convertir(carpeta):
    if not os.path.isdir(carpeta):
        print(f"No existe la carpeta: {carpeta}")
        return 1

    archivos = sorted(
        f for f in os.listdir(carpeta)
        if f.lower().endswith(EXTS) and not f.startswith("~$")
    )
    if not archivos:
        print("No se encontraron archivos .ppt/.pptx en la carpeta.")
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for nombre in archivos:
            origen = os.path.join(carpeta, nombre)
            salida = os.path.join(carpeta, os.path.splitext(nombre)[0] + ".pdf")
            presentacion = app.Presentations.Open(
                origen, ReadOnly=True, Untitled=False, WithWindow=False
            )
            try:
                presentacion.SaveAs(salida, SAVE_AS_PDF)
                print(f"OK: {nombre}  ->  {os.path.basename(salida)}")
            finally:
                presentacion.Close()
    finally:
        app.Quit()
    return 0


if __name__ == "__main__":
    carpeta = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
    sys.exit(convertir(carpeta))
