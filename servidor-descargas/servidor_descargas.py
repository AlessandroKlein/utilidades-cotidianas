#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Servidor local de descargas para la red del hogar.

Poné tus archivos en una carpeta y compartila: los demás equipos de tu casa
abren el navegador y los descargan, sin instalar nada en ellos.

Uso:
    python servidor_descargas.py [carpeta] [puerto]

- carpeta : carpeta con los archivos a compartir (por defecto "archivos" junto al script).
- puerto  : puerto a usar (por defecto 8000).

Verificación:
    python servidor_descargas.py --selftest
"""

import argparse
import html
import mimetypes
import os
import re
import socket
import sys
import tempfile
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

# extensiones que se pueden previsualizar en el navegador (además de descargar)
EXTENSIONES_TEXTO = {".txt", ".md", ".yml", ".yaml"}

PAGINA = """<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Descargas · Red local</title>
<style>
:root {
  --fondo: #f5f6f8; --tarjeta: #ffffff; --texto: #1c1e21; --tenue: #65676b;
  --borde: #e4e6eb; --acento: #2563eb; --acento-texto: #ffffff; --fila-hover: #f0f2f5;
}
@media (prefers-color-scheme: dark) {
  :root {
    --fondo: #111214; --tarjeta: #1b1d21; --texto: #e7e9ec; --tenue: #9aa0a6;
    --borde: #2e3136; --acento: #3b82f6; --acento-texto: #ffffff; --fila-hover: #232529;
  }
}
* { box-sizing: border-box; }
body { margin: 0; font-family: "Segoe UI", system-ui, sans-serif; background: var(--fondo); color: var(--texto); }
.cabecera { padding: 34px 20px 14px; text-align: center; }
.cabecera h1 { margin: 0 0 6px; font-size: 1.6rem; }
.cabecera p { margin: 0; color: var(--tenue); }
.caja { max-width: 880px; margin: 0 auto 40px; padding: 0 16px; }
.buscador { width: 100%; padding: 11px 14px; border: 1px solid var(--borde); border-radius: 10px;
  background: var(--tarjeta); color: var(--texto); font-size: 1rem; margin-bottom: 14px; }
.buscador:focus { outline: 2px solid var(--acento); outline-offset: 1px; }
.tarjeta { background: var(--tarjeta); border: 1px solid var(--borde); border-radius: 12px; overflow: hidden; }
table { width: 100%; border-collapse: collapse; }
th, td { text-align: left; padding: 12px 16px; border-bottom: 1px solid var(--borde); vertical-align: middle; }
th { font-size: .78rem; text-transform: uppercase; letter-spacing: .05em; color: var(--tenue); }
tbody tr:last-child td { border-bottom: none; }
tbody tr:hover { background: var(--fila-hover); }
td.nombre { word-break: break-all; }
td.der { text-align: right; white-space: nowrap; color: var(--tenue); }
a.descargar { color: var(--acento-texto); background: var(--acento); text-decoration: none;
  padding: 7px 14px; border-radius: 8px; font-weight: 600; font-size: .9rem; white-space: nowrap; }
a.descargar:hover { filter: brightness(1.1); }
a.ver { color: var(--acento); border: 1px solid var(--acento); text-decoration: none;
  padding: 6px 13px; border-radius: 8px; font-weight: 600; font-size: .9rem; white-space: nowrap; margin-left: 6px; }
a.ver:hover { background: var(--fila-hover); }
.vacio { padding: 34px; text-align: center; color: var(--tenue); }
.pie { text-align: center; color: var(--tenue); font-size: .85rem; padding: 0 16px 34px; }
</style>
</head>
<body>
<div class="cabecera">
  <h1>Descargas de la red local</h1>
  <p>Archivos compartidos desde el equipo servidor.</p>
</div>
<div class="caja">
  <input id="buscar" class="buscador" type="text" placeholder="Buscar archivo…" autofocus>
  <div class="tarjeta">
    <table>
      <thead><tr><th>Archivo</th><th class="der">Tamaño</th><th class="der">Modificado</th><th></th></tr></thead>
      <tbody>%%FILAS%%</tbody>
    </table>
  </div>
</div>
<div class="pie">%%TOTAL%%</div>
<script>
const b = document.getElementById("buscar");
b.addEventListener("input", function () {
  const q = b.value.trim().toLowerCase();
  document.querySelectorAll("tbody tr").forEach(function (tr) {
    tr.style.display = tr.dataset.nombre.includes(q) ? "" : "none";
  });
});
</script>
</body>
</html>"""


def tamano_humano(n):
    n = float(n)
    for unidad in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024:
            return f"{n:.0f} {unidad}" if unidad == "B" else f"{n:.1f} {unidad}"
        n /= 1024
    return f"{n:.1f} PB"


def ips_locales():
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except OSError:
        pass
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # no envía paquetes; solo descubre la interfaz de salida
        s.connect(("8.8.8.8", 80))
        ips.add(s.getsockname()[0])
    except OSError:
        pass
    finally:
        s.close()
    return sorted(ip for ip in ips if not ip.startswith("127."))


class Manejador(BaseHTTPRequestHandler):
    carpeta = None  # se asigna al levantar el servidor
    server_version = "ServidorDescargas/1.0"

    def do_GET(self):
        ruta = urlparse(self.path)
        if ruta.path in ("/", "/index.html"):
            self._servir_indice()
        elif ruta.path == "/descargar":
            self._servir_descarga(ruta.query)
        elif ruta.path == "/ver":
            self._servir_vista(ruta.query)
        elif ruta.path == "/favicon.ico":
            self.send_response(204)
            self.end_headers()
        else:
            self.send_error(404, "No encontrado")

    def _servir_indice(self):
        filas = []
        n = 0
        for p in sorted(self.carpeta.iterdir(), key=lambda x: x.name.lower()):
            if not p.is_file():
                continue
            try:
                st = p.stat()
                tam = tamano_humano(st.st_size)
                fecha = time.strftime("%d/%m/%Y %H:%M", time.localtime(st.st_mtime))
            except OSError:
                continue
            n += 1
            nombre_esc = html.escape(p.name)
            url = "/descargar?f=" + quote(p.name)
            acciones = f'<a class="descargar" href="{url}">Descargar</a>'
            if p.suffix.lower() in EXTENSIONES_TEXTO:
                acciones += f' <a class="ver" href="/ver?f={quote(p.name)}" target="_blank">Ver</a>'
            filas.append(
                f'<tr data-nombre="{html.escape(p.name.lower())}">'
                f'<td class="nombre">{nombre_esc}</td>'
                f'<td class="der">{tam}</td>'
                f'<td class="der">{fecha}</td>'
                f'<td>{acciones}</td>'
                f"</tr>"
            )
        cuerpo = "" if not filas else "\n".join(filas)
        if not filas:
            cuerpo = '<tr><td class="vacio" colspan="4">No hay archivos en la carpeta compartida todavía.</td></tr>'
        total = f"{n} archivo(s)" if n != 1 else "1 archivo"
        pagina = PAGINA.replace("%%FILAS%%", cuerpo).replace("%%TOTAL%%", total)
        datos = pagina.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(datos)))
        self.end_headers()
        self.wfile.write(datos)

    def _archivo_valido(self, nombre):
        """Devuelve el Path resuelto si el archivo queda dentro de la carpeta compartida; si no, None."""
        raiz = self.carpeta.resolve()
        destino = (raiz / nombre).resolve()
        # bloqueo de path traversal: el archivo debe quedar dentro de la carpeta compartida
        if not destino.is_relative_to(raiz) or not destino.is_file():
            return None
        return destino

    def _servir_descarga(self, query):
        qs = parse_qs(query)
        nombre = unquote(qs.get("f", [""])[0])
        if not nombre:
            self.send_error(400, "Falta el parámetro 'f'")
            return
        destino = self._archivo_valido(nombre)
        if destino is None:
            self.send_error(404, "Archivo no encontrado")
            return
        tipo = mimetypes.guess_type(destino.name)[0] or "application/octet-stream"
        fallback = re.sub(r"[^\x20-\x7e]", "_", destino.name).replace('"', "'")
        self.send_response(200)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(destino.stat().st_size))
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{fallback}"; filename*=UTF-8\'\'{quote(destino.name)}',
        )
        self.end_headers()
        with open(destino, "rb") as f:
            while True:
                bloque = f.read(1024 * 1024)
                if not bloque:
                    break
                self.wfile.write(bloque)

    def _servir_vista(self, query):
        qs = parse_qs(query)
        nombre = unquote(qs.get("f", [""])[0])
        if not nombre:
            self.send_error(400, "Falta el parámetro 'f'")
            return
        destino = self._archivo_valido(nombre)
        if destino is None:
            self.send_error(404, "Archivo no encontrado")
            return
        try:
            texto = destino.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            self.send_error(415, "No se puede mostrar como texto")
            return
        datos = texto.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(datos)))
        self.end_headers()
        self.wfile.write(datos)

    def log_message(self, fmt, *args):
        print(f"[{self.client_address[0]}] {fmt % args}")


def _selftest():
    with tempfile.TemporaryDirectory() as tmp:
        raiz = Path(tmp)
        (raiz / "prueba.txt").write_text("hola mundo", encoding="utf-8")
        (raiz / "otro archivo.pdf").write_bytes(b"%PDF-1.4 fake")
        fuera = raiz.parent / f"secreto_{os.getpid()}.txt"
        fuera.write_text("secreto", encoding="utf-8")
        try:
            Manejador.carpeta = raiz.resolve()
            srv = ThreadingHTTPServer(("127.0.0.1", 0), Manejador)
            port = srv.server_address[1]
            threading.Thread(target=srv.serve_forever, daemon=True).start()
            base = f"http://127.0.0.1:{port}"
            import urllib.request
            import urllib.error
            try:
                with urllib.request.urlopen(base + "/", timeout=5) as r:
                    cuerpo = r.read().decode("utf-8")
                    assert r.status == 200 and "prueba.txt" in cuerpo, "índice sin prueba.txt"
                    assert "/ver?f=prueba.txt" in cuerpo, "el índice no ofrece Ver para .txt"
                with urllib.request.urlopen(base + "/ver?f=prueba.txt", timeout=5) as r:
                    assert r.read() == b"hola mundo", "vista de texto incorrecta"
                with urllib.request.urlopen(base + "/descargar?f=prueba.txt", timeout=5) as r:
                    assert r.read() == b"hola mundo", "contenido de descarga incorrecto"
                with urllib.request.urlopen(base + "/descargar?f=" + quote("otro archivo.pdf"), timeout=5) as r:
                    assert r.read() == b"%PDF-1.4 fake", "nombre con espacio falló"
                try:
                    urllib.request.urlopen(base + "/descargar?f=" + quote("../" + fuera.name), timeout=5)
                    raise AssertionError("el path traversal no fue bloqueado")
                except urllib.error.HTTPError as e:
                    assert e.code == 404, f"esperaba 404, obtuve {e.code}"
                try:
                    urllib.request.urlopen(base + "/descargar?f=no-existe.txt", timeout=5)
                    raise AssertionError("archivo inexistente debería dar 404")
                except urllib.error.HTTPError as e:
                    assert e.code == 404
                print("SELFTEST OK: índice, Ver para texto, descarga, nombre con espacio, path traversal y 404 en verde.")
                return 0
            finally:
                srv.shutdown()
        finally:
            if fuera.exists():
                fuera.unlink()


def main(argv=None):
    ap = argparse.ArgumentParser(description="Servidor local de descargas para la red del hogar.")
    ap.add_argument("carpeta", nargs="?", default=None,
                    help="Carpeta con los archivos a compartir (por defecto 'archivos' junto al script).")
    ap.add_argument("puerto", nargs="?", type=int, default=8000,
                    help="Puerto a usar (por defecto 8000).")
    ap.add_argument("--selftest", action="store_true", help="Ejecuta la verificación automática y sale.")
    args = ap.parse_args(argv)

    if args.selftest:
        return _selftest()

    base = Path(__file__).resolve().parent
    carpeta = Path(args.carpeta) if args.carpeta else base / "archivos"
    carpeta.mkdir(parents=True, exist_ok=True)
    carpeta = carpeta.resolve()

    Manejador.carpeta = carpeta
    servidor = ThreadingHTTPServer(("0.0.0.0", args.puerto), Manejador)

    print("=" * 62)
    print("Servidor de descargas levantado")
    print(f"  Carpeta compartida: {carpeta}")
    print()
    ips = ips_locales()
    if ips:
        print("  En otros equipos de tu casa:")
        for ip in ips:
            print(f"      http://{ip}:{args.puerto}")
    print(f"  En este mismo equipo:   http://localhost:{args.puerto}")
    print("  (Para detener: Ctrl+C)")
    print("=" * 62)
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")


if __name__ == "__main__":
    sys.exit(main())
