# -*- coding: utf-8 -*-
r"""Renombra archivos en lote dentro de una carpeta.

Opciones combinables: prefijo, sufijo, numeración secuencial, buscar/reemplazar
y normalización de nombres (sin acentos, espacios -> guion bajo).
Por defecto solo muestra el plan (dry-run); usá --aplicar para renombrar.

Uso:
    python renombrar_masivo.py --prefijo "vacaciones_" --numerar
    python renombrar_masivo.py --buscar "foto" --reemplazar "img" --aplicar
    python renombrar_masivo.py --normalizar --sufijo "_v2" --aplicar --si
    python renombrar_masivo.py "C:\carpeta" --buscar " " --aplicar
"""
import argparse
import os
import re
import sys
import unicodedata


def normalizar(nombre):
    """Quita acentos y deja solo letras, números, punto, guion y subrayado."""
    sin_acentos = unicodedata.normalize("NFKD", nombre).encode(
        "ascii", "ignore"
    ).decode("ascii")
    limpio = re.sub(r"[^A-Za-z0-9._-]+", "_", sin_acentos).strip("._-")
    return limpio or "archivo"


def construir_plan(carpeta, args):
    """Devuelve [(origen, nuevo_nombre)] para los archivos a renombrar."""
    propio = os.path.basename(__file__)
    archivos = sorted(
        f for f in os.listdir(carpeta)
        if os.path.isfile(os.path.join(carpeta, f))
        and not f.startswith(".")
        and f != propio
    )

    plan = []
    ancho = max(args.digitos, len(str(args.numerar_desde + len(archivos) - 1)))
    for i, nombre in enumerate(archivos, 1):
        base, ext = os.path.splitext(nombre)
        nuevo = base

        if args.normalizar:
            nuevo = normalizar(nuevo)
        if args.buscar:
            nuevo = nuevo.replace(args.buscar, args.reemplazar or "")
        if args.prefijo:
            nuevo = args.prefijo + nuevo
        if args.sufijo:
            nuevo = nuevo + args.sufijo
        if args.numerar:
            n = args.numerar_desde + i - 1
            nuevo = f"{n:0{ancho}d}_{nuevo}"

        nuevo = nuevo + ext
        if nuevo != nombre:
            plan.append((os.path.join(carpeta, nombre), nuevo))
    return plan


def nombre_libre(carpeta, nombre):
    """Si `nombre` ya existe en `carpeta`, le agrega " (2)", " (3)", etc."""
    ruta = os.path.join(carpeta, nombre)
    if not os.path.exists(ruta):
        return nombre
    base, ext = os.path.splitext(nombre)
    i = 2
    while os.path.exists(os.path.join(carpeta, f"{base} ({i}){ext}")):
        i += 1
    return f"{base} ({i}){ext}"


def aplicar(carpeta, plan):
    """Renombra en dos fases para evitar colisiones entre nombres del propio lote."""
    pares = []  # (origen, temporal)
    try:
        # Fase 1: todos a nombres temporales únicos.
        for i, (origen, _) in enumerate(plan):
            tmp = os.path.join(carpeta, f".__renombrar_tmp_{i}__")
            os.rename(origen, tmp)
            pares.append((origen, tmp))
        # Fase 2: de temporal a nombre final.
        for (origen, nuevo), (_, tmp) in zip(plan, pares):
            destino = os.path.join(carpeta, nombre_libre(carpeta, nuevo))
            os.rename(tmp, destino)
            print(f"  {os.path.basename(origen)}  ->  {os.path.basename(destino)}")
    except OSError as e:
        print(f"\nError al renombrar: {e}", file=sys.stderr)
        print("Intentando revertir los cambios...", file=sys.stderr)
        for origen, tmp in pares:
            if os.path.exists(tmp):
                try:
                    os.rename(tmp, origen)
                except OSError:
                    pass
        return False
    return True


def main():
    p = argparse.ArgumentParser(
        description="Renombra archivos en lote dentro de una carpeta."
    )
    p.add_argument("carpeta", nargs="?", default=".",
                   help="Carpeta con los archivos (por defecto, la actual).")
    p.add_argument("--prefijo", default="", help="Texto a anteponer al nombre.")
    p.add_argument("--sufijo", default="", help="Texto a agregar al final del nombre.")
    p.add_argument("--buscar", default="", help="Subcadena a buscar en el nombre.")
    p.add_argument("--reemplazar", default="",
                   help="Reemplazo de --buscar (vacío = eliminar la subcadena).")
    p.add_argument("--numerar", action="store_true",
                   help="Anteponer un número secuencial (NN_...) a cada nombre.")
    p.add_argument("--numerar-desde", type=int, default=1, help="Número inicial.")
    p.add_argument("--digitos", type=int, default=2,
                   help="Ancho mínimo del número (relleno con ceros).")
    p.add_argument("--normalizar", action="store_true",
                   help="Quitar acentos y caracteres raros; espacios -> _.")
    p.add_argument("--aplicar", action="store_true",
                   help="Renombrar (sin esto solo se muestra el plan).")
    p.add_argument("--si", action="store_true", help="No pedir confirmación.")
    args = p.parse_args()

    if not (args.prefijo or args.sufijo or args.buscar or args.numerar or args.normalizar):
        p.error("indicá al menos una opción de renombrado "
                "(--prefijo, --sufijo, --buscar, --numerar, --normalizar)")
    if args.reemplazar and not args.buscar:
        p.error("--reemplazar requiere --buscar")

    if not os.path.isdir(args.carpeta):
        print(f"No existe la carpeta: {args.carpeta}", file=sys.stderr)
        return 1

    plan = construir_plan(args.carpeta, args)
    if not plan:
        print("No hay archivos que renombrar con estas opciones.")
        return 0

    print(f"Carpeta: {args.carpeta}")
    print(f"Archivos a renombrar: {len(plan)}\n")
    for origen, nuevo in plan:
        print(f"  {os.path.basename(origen):40} -> {nuevo}")

    if not args.aplicar:
        print("\nDry-run: no se renombró nada. Usá --aplicar para aplicar.")
        return 0

    if not args.si:
        try:
            resp = input(f"\n¿Renombrar {len(plan)} archivos? [s/N] ").strip().lower()
        except EOFError:
            resp = ""
        if resp not in ("s", "si", "sí", "y", "yes"):
            print("Cancelado.")
            return 0

    if aplicar(args.carpeta, plan):
        print("\nListo.")
        return 0
    return 1


if __name__ == "__main__":
    sys.exit(main())
