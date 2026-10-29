# -*- coding: utf-8 -*-
r"""Busca archivos duplicados (por contenido, no solo por nombre) en una carpeta.

Agrupa por tamaño y compara con hash blake2b, para no comparar todos contra
todos. Por defecto solo lista los duplicados y cuánto espacio se recuperaría.
Con --borrar elimina las copias extra (conserva una de cada grupo).

Uso:
    python buscar_duplicados.py "C:\carpeta"
    python buscar_duplicados.py --borrar            # pide confirmación
    python buscar_duplicados.py --borrar --si
"""
import argparse
import hashlib
import os
import sys

TAM_BLOQUE = 1024 * 1024  # 1 MiB


def hash_archivo(ruta):
    h = hashlib.blake2b(digest_size=16)
    try:
        with open(ruta, "rb") as f:
            while True:
                bloque = f.read(TAM_BLOQUE)
                if not bloque:
                    break
                h.update(bloque)
    except OSError as e:
        print(f"No se pudo leer {ruta}: {e}", file=sys.stderr)
        return None
    return h.hexdigest()


def buscar_duplicados(carpeta):
    """Devuelve una lista de grupos [rutas] que comparten contenido idéntico."""
    por_tamano = {}
    for raiz, _dirs, archivos in os.walk(carpeta):
        for nombre in archivos:
            ruta = os.path.join(raiz, nombre)
            try:
                tam = os.path.getsize(ruta)
            except OSError:
                continue
            if tam == 0:
                continue  # los vacíos no valen el riesgo de borrarlos
            por_tamano.setdefault(tam, []).append(ruta)

    grupos = []
    for tam, rutas in por_tamano.items():
        if len(rutas) < 2:
            continue
        por_hash = {}
        for ruta in rutas:
            h = hash_archivo(ruta)
            if h is None:
                continue
            por_hash.setdefault(h, []).append(ruta)
        for duplicados in por_hash.values():
            if len(duplicados) > 1:
                grupos.append(sorted(duplicados))
    return grupos


def humano(n):
    for unidad in ("B", "KiB", "MiB", "GiB", "TiB"):
        if n < 1024 or unidad == "TiB":
            return f"{n:.0f} {unidad}" if unidad == "B" else f"{n:.1f} {unidad}"
        n /= 1024


def main():
    p = argparse.ArgumentParser(
        description="Busca archivos duplicados por contenido en una carpeta."
    )
    p.add_argument("carpeta", nargs="?", default=".",
                   help="Carpeta a analizar (por defecto, la actual).")
    p.add_argument("--borrar", action="store_true",
                   help="Eliminar las copias duplicadas (conserva una de cada grupo).")
    p.add_argument("--si", action="store_true", help="No pedir confirmación.")
    args = p.parse_args()

    if not os.path.isdir(args.carpeta):
        print(f"No existe la carpeta: {args.carpeta}", file=sys.stderr)
        return 1

    print(f"Analizando: {args.carpeta} ...")
    grupos = buscar_duplicados(args.carpeta)

    if not grupos:
        print("No se encontraron duplicados.")
        return 0

    total_ahorro = 0
    total_borrables = 0
    print(f"\nSe encontraron {len(grupos)} grupo(s) de duplicados:\n")
    for i, grupo in enumerate(grupos, 1):
        tam = os.path.getsize(grupo[0])
        ahorro = tam * (len(grupo) - 1)
        total_ahorro += ahorro
        total_borrables += len(grupo) - 1
        print(f"[{i}] {len(grupo)} copias × {humano(tam)} "
              f"(se recuperarían {humano(ahorro)}):")
        for ruta in grupo:
            print(f"      {ruta}")

    print(f"\nEspacio total recuperable: {humano(total_ahorro)} "
          f"({total_borrables} archivos)")

    if not args.borrar:
        print("Dry-run: no se borró nada. Usá --borrar para eliminar las copias extra.")
        return 0

    if not args.si:
        try:
            resp = input(f"\n¿Borrar {total_borrables} archivos duplicados? [s/N] "
                         ).strip().lower()
        except EOFError:
            resp = ""
        if resp not in ("s", "si", "sí", "y", "yes"):
            print("Cancelado.")
            return 0

    borrados = 0
    for grupo in grupos:
        conservar = grupo[0]
        for ruta in grupo[1:]:
            try:
                os.remove(ruta)
                borrados += 1
                print(f"  BORRADO: {ruta}")
            except OSError as e:
                print(f"  No se pudo borrar {ruta}: {e}", file=sys.stderr)
    print(f"\nListo. Se conservaron las primeras copias de cada grupo; "
          f"{borrados} archivos borrados.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
