# -*- coding: utf-8 -*-
r"""Organiza la carpeta Descargas moviendo los archivos sueltos a subcarpetas.

Cada archivo se mueve (nunca se borra) a una subcarpeta según su extensión:
Imágenes, Documentos, Música, Vídeos, Comprimidos, Instaladores u Otros.
Por defecto solo muestra qué haría (dry-run); usá --aplicar para mover.

Uso:
    python organizar_descargas.py                # dry-run sobre la Descargas real
    python organizar_descargas.py --aplicar      # mueve (pide confirmación)
    python organizar_descargas.py "C:\otra\carpeta" --aplicar --si
"""
import argparse
import os
import shutil
import sys
from pathlib import Path

# Extensión (en minúsculas) -> carpeta de destino.
CATEGORIAS = {
    # Imágenes
    ".jpg": "Imágenes", ".jpeg": "Imágenes", ".png": "Imágenes",
    ".gif": "Imágenes", ".bmp": "Imágenes", ".webp": "Imágenes",
    ".svg": "Imágenes", ".ico": "Imágenes", ".heic": "Imágenes",
    # Documentos
    ".pdf": "Documentos", ".doc": "Documentos", ".docx": "Documentos",
    ".xls": "Documentos", ".xlsx": "Documentos", ".ppt": "Documentos",
    ".pptx": "Documentos", ".txt": "Documentos", ".md": "Documentos",
    ".csv": "Documentos", ".rtf": "Documentos", ".odt": "Documentos",
    ".ods": "Documentos", ".odp": "Documentos",
    # Música
    ".mp3": "Música", ".wav": "Música", ".flac": "Música",
    ".aac": "Música", ".ogg": "Música", ".m4a": "Música", ".wma": "Música",
    # Vídeos
    ".mp4": "Vídeos", ".mkv": "Vídeos", ".avi": "Vídeos",
    ".mov": "Vídeos", ".wmv": "Vídeos", ".webm": "Vídeos",
    ".flv": "Vídeos", ".m4v": "Vídeos",
    # Comprimidos
    ".zip": "Comprimidos", ".rar": "Comprimidos", ".7z": "Comprimidos",
    ".tar": "Comprimidos", ".gz": "Comprimidos", ".bz2": "Comprimidos",
    ".xz": "Comprimidos",
    # Instaladores / imágenes de disco
    ".exe": "Instaladores", ".msi": "Instaladores", ".apk": "Instaladores",
    ".dmg": "Instaladores", ".pkg": "Instaladores", ".iso": "Instaladores",
}


def descargas_real():
    """Devuelve la carpeta Descargas real del usuario, o una buena suposición."""
    try:
        import winreg
        clave = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        )
        try:
            valor = winreg.QueryValueEx(
                clave, "{374DE290-123F-4565-9164-39C4925E467B}"
            )[0]
        finally:
            winreg.CloseKey(clave)
        ruta = os.path.expandvars(valor)
        if os.path.isdir(ruta):
            return ruta
    except Exception:
        pass

    for candidato in ("Downloads", "Descargas"):
        p = Path.home() / candidato
        if p.is_dir():
            return str(p)
    return None


def planificar(carpeta):
    """Devuelve [(origen, carpeta_destino, categoría)] para los archivos sueltos."""
    movimientos = []
    for entrada in sorted(os.listdir(carpeta)):
        origen = os.path.join(carpeta, entrada)
        if os.path.isdir(origen):
            continue
        if entrada.startswith("~$") or entrada.lower() == "desktop.ini":
            continue
        ext = os.path.splitext(entrada)[1].lower()
        categoria = CATEGORIAS.get(ext, "Otros")
        movimientos.append((origen, os.path.join(carpeta, categoria), categoria))
    return movimientos


def nombre_libre(ruta):
    """Si `ruta` ya existe, devuelve `nombre (2).ext`, `nombre (3).ext`, etc."""
    if not os.path.exists(ruta):
        return ruta
    base, ext = os.path.splitext(ruta)
    i = 2
    while os.path.exists(f"{base} ({i}){ext}"):
        i += 1
    return f"{base} ({i}){ext}"


def aplicar(movimientos):
    for origen, destino_dir, _ in movimientos:
        os.makedirs(destino_dir, exist_ok=True)
        destino = nombre_libre(os.path.join(destino_dir, os.path.basename(origen)))
        shutil.move(origen, destino)
        print(f"  {os.path.basename(origen):40} -> {os.path.basename(destino_dir)}/")


def main():
    p = argparse.ArgumentParser(
        description="Organiza la carpeta Descargas en subcarpetas por tipo de archivo."
    )
    p.add_argument("carpeta", nargs="?", default=None,
                   help="Carpeta a organizar (por defecto, la Descargas real).")
    p.add_argument("--aplicar", action="store_true",
                   help="Mover los archivos (sin esto solo se muestra qué se haría).")
    p.add_argument("--si", action="store_true", help="No pedir confirmación.")
    args = p.parse_args()

    carpeta = args.carpeta or descargas_real()
    if not carpeta or not os.path.isdir(carpeta):
        print(f"No se pudo determinar/abrir la carpeta: {carpeta or '(ninguna)'}",
              file=sys.stderr)
        print('Indicá una carpeta: python organizar_descargas.py "C:\\ruta"',
              file=sys.stderr)
        return 1

    movimientos = planificar(carpeta)
    if not movimientos:
        print("No hay archivos sueltos que organizar.")
        return 0

    print(f"Carpeta: {carpeta}")
    print(f"Archivos a mover: {len(movimientos)}\n")
    for origen, _, categoria in movimientos:
        print(f"  {os.path.basename(origen):40} -> {categoria}/")

    if not args.aplicar:
        print("\nDry-run: no se movió nada. Usá --aplicar para mover.")
        return 0

    if not args.si:
        try:
            resp = input(f"\n¿Mover {len(movimientos)} archivos? [s/N] ").strip().lower()
        except EOFError:
            resp = ""
        if resp not in ("s", "si", "sí", "y", "yes"):
            print("Cancelado.")
            return 0

    aplicar(movimientos)
    print("\nListo.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
