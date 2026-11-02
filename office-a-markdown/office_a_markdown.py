# -*- coding: utf-8 -*-
r"""Convierte Office/PDF/imágenes a Markdown, con GUI y OCR local opcional.

Extrae el contenido con Word y PowerPoint instalados (COM/pywin32), pypdf para
los .pdf, y OCR local (minicpm-v vía ollama) para PDFs escaneados e imágenes.
Guarda un .md junto a cada archivo. Ideal para pasar documentos a una IA
reduciendo tokens: el .md pesa mucho menos que el .pdf/.docx original.

Uso:
    python office_a_markdown.py                  # abre la interfaz gráfica
    python office_a_markdown.py --gui            # abre la interfaz gráfica
    python office_a_markdown.py "C:\\ruta\\carpeta"
    python office_a_markdown.py "C:\\ruta\\archivo.pptx"
    python office_a_markdown.py --ocr "C:\\ruta" # OCR de escaneados (ollama)
    python office_a_markdown.py -o salida "C:\\ruta"   # .md en otra carpeta
    python office_a_markdown.py --selftest       # prueba end-to-end

Las imágenes y gráficos se omiten salvo con OCR (solo se extrae su texto).
"""
import base64
import json
import os
import queue
import re
import sys
import tempfile
import threading
import tkinter as tk
import urllib.error
import urllib.request
from html.parser import HTMLParser
from tkinter import filedialog, messagebox, ttk

import win32com.client

try:
    from pypdf import PdfReader
    _HAS_PDF = True
except ImportError:
    _HAS_PDF = False

try:
    import pymupdf as fitz
    _HAS_FITZ = True
except ImportError:
    try:
        import fitz
        _HAS_FITZ = True
    except ImportError:
        fitz = None
        _HAS_FITZ = False

POWERPOINT_EXTS = {".ppt", ".pps", ".pptx"}
WORD_EXTS = {".doc", ".docx", ".rtf"}
PDF_EXTS = {".pdf"}
TEXT_EXTS = {".txt"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp", ".gif"}
EXTS = POWERPOINT_EXTS | WORD_EXTS | PDF_EXTS | TEXT_EXTS

WD_FORMAT_FILTERED_HTML = 10
PP_PLACEHOLDER_TITLE = 1

OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://127.0.0.1:11434")
OCR_MODEL = os.environ.get("OCR_MODEL", "minicpm-v")
OCR_PROMPT = (
    "Extraé todo el texto visible de esta imagen en formato Markdown. "
    "Conservá títulos, listas y tablas. Devolvé solo el texto, sin comentarios."
)


# ----------------------------------------------------------------- HTML -> MD
class HTMLToMarkdown(HTMLParser):
    HEADINGS = {"h1": "#", "h2": "##", "h3": "###",
                "h4": "####", "h5": "#####", "h6": "######"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out = []
        self.buf = []
        self.skip = 0          # dentro de <style>/<script>/<head>
        self.span_skip = 0     # dentro de un marcador de lista (viñeta/número)
        self.in_table = False
        self.rows = []
        self.cells = []
        self.row_cells = []

    def _take(self):
        s = "".join(self.buf)
        self.buf = []
        return s.strip()

    def _emit_line(self, line):
        line = line.strip()
        if line and line != "-":  # descarta viñetas vacías que deja Word
            self.out.append(line)

    def _cell_emit(self, line):
        line = line.strip()
        if line:
            self.cells.append(line)

    @staticmethod
    def _attr(attrs, name):
        for k, v in attrs:
            if k == name:
                return v or ""
        return ""

    def handle_starttag(self, tag, attrs):
        if tag in ("style", "script", "head"):
            self.skip += 1
            return
        if self.skip:
            return
        if self.span_skip:
            if tag == "span":
                self.span_skip += 1
            return
        if tag == "span":
            style = self._attr(attrs, "style").lower()
            if "mso-list:ignore" in style or "font-family:symbol" in style:
                self.span_skip = 1
            return
        if tag in self.HEADINGS:
            self._take()
            self.buf.append(self.HEADINGS[tag] + " ")
        elif tag == "p":
            cls = self._attr(attrs, "class").lower()
            m = re.search(r"heading\s*([1-6])", cls)
            if "heading" in cls and m:
                self._take()
                self.buf.append("#" * int(m.group(1)) + " ")
            elif "msolistparagraph" in cls and not self.in_table:
                self._take()
                self.buf.append("- ")
            else:
                self._take()
        elif tag == "div":
            self._take()
        elif tag == "li":
            self._take()
            self.buf.append("- ")
        elif tag == "br":
            self.buf.append("\n")
        elif tag in ("b", "strong"):
            self.buf.append("**")
        elif tag in ("i", "em"):
            self.buf.append("*")
        elif tag == "table":
            self._take()
            self.in_table = True
            self.rows = []
            self.row_cells = []
            self.cells = []
        elif tag == "tr" and self.in_table:
            self._take()
            self.row_cells = []
        elif tag in ("td", "th") and self.in_table:
            self._take()
            self.cells = []

    def handle_endtag(self, tag):
        if tag in ("style", "script", "head"):
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return
        if self.span_skip:
            if tag == "span":
                self.span_skip -= 1
            return
        if tag in ("td", "th") and self.in_table:
            tail = self._take()
            if tail:
                self.cells.append(tail)
            self.row_cells.append(" ".join(self.cells))
        elif tag == "tr" and self.in_table:
            self.rows.append(self.row_cells)
            self.row_cells = []
        elif tag == "table" and self.in_table:
            self.in_table = False
            self._render_table()
        elif tag in self.HEADINGS or tag in ("p", "div", "li"):
            line = self._take()
            (self._cell_emit if self.in_table else self._emit_line)(line)
        elif tag in ("b", "strong"):
            self.buf.append("**")
        elif tag in ("i", "em"):
            self.buf.append("*")

    def handle_data(self, data):
        if self.skip or self.span_skip:
            return
        self.buf.append(data)

    def _render_table(self):
        rows = self.rows
        if not rows:
            return
        ncols = max(len(r) for r in rows)
        for i, r in enumerate(rows):
            cells = [c.replace("|", "\\|").replace("\n", " ").strip() for c in r]
            cells += [""] * (ncols - len(cells))
            self.out.append("| " + " | ".join(cells) + " |")
            if i == 0:
                self.out.append("| " + " | ".join("---" for _ in range(ncols)) + " |")

    def result(self):
        self._emit_line(self._take())
        return "\n".join(self.out)


# ----------------------------------------------------------------- apps COM
def _word_app():
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        app.Options.ConfirmConversions = False
    except Exception:
        pass
    try:
        app.ScreenUpdating = False
    except Exception:
        pass
    return app


def _ppt_app():
    return win32com.client.DispatchEx("PowerPoint.Application")


# ------------------------------------------------------- extracción PowerPoint
def _texto_shape(shape):
    try:
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return False, []
        texto = shape.TextFrame.TextRange.Text or ""
    except Exception:
        return False, []
    texto = texto.replace("\x0b", " ").replace("\x0c", " ")
    paras = [p.strip() for p in texto.split("\r")]
    paras = [p for p in paras if p]
    es_titulo = False
    try:
        es_titulo = shape.PlaceholderFormat.Type == PP_PLACEHOLDER_TITLE
    except Exception:
        pass
    return es_titulo, paras


def _tabla_shape(shape):
    try:
        if not shape.HasTable:
            return None
        t = shape.Table
        rows = []
        for r in range(1, t.Rows.Count + 1):
            fila = []
            for c in range(1, t.Columns.Count + 1):
                celda = t.Cell(r, c)
                txt = ""
                try:
                    if celda.Shape.HasTextFrame:
                        txt = celda.Shape.TextFrame.TextRange.Text or ""
                except Exception:
                    txt = ""
                fila.append(txt.replace("\r", " ").replace("\x0b", " ").strip())
            rows.append(fila)
        return rows
    except Exception:
        return None


def _render_tabla(rows):
    if not rows:
        return ""
    ncols = max(len(r) for r in rows)
    lines = []
    for i, r in enumerate(rows):
        cells = [(c or "").replace("|", "\\|").replace("\n", " ").strip() for c in r]
        cells += [""] * (ncols - len(cells))
        lines.append("| " + " | ".join(cells) + " |")
        if i == 0:
            lines.append("| " + " | ".join("---" for _ in range(ncols)) + " |")
    return "\n".join(lines)


def _powerpoint_a_markdown(origen):
    app = _ppt_app()
    out = []
    try:
        pres = app.Presentations.Open(
            origen, ReadOnly=True, Untitled=False, WithWindow=False
        )
        try:
            for idx, slide in enumerate(pres.Slides, 1):
                out.append("\n## Diapositiva %d\n" % idx)
                for shape in slide.Shapes:
                    tabla = _tabla_shape(shape)
                    if tabla:
                        out.append(_render_tabla(tabla))
                        out.append("")
                        continue
                    es_titulo, paras = _texto_shape(shape)
                    if not paras:
                        continue
                    prefix = "### " if es_titulo else "- "
                    out.extend(prefix + p for p in paras)
        finally:
            pres.Close()
    finally:
        app.Quit()
    return "\n".join(out)


# ------------------------------------------------------------- extracción Word
def _leer_html(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _word_a_markdown(origen):
    app = _word_app()
    fd, html_path = tempfile.mkstemp(suffix=".htm", prefix="mdconv_")
    os.close(fd)
    try:
        doc = app.Documents.Open(origen, ReadOnly=True, AddToRecentFiles=False)
        try:
            try:
                doc.WebOptions.Encoding = 65001  # msoEncodingUTF8
            except Exception:
                pass
            doc.SaveAs2(html_path, FileFormat=WD_FORMAT_FILTERED_HTML)
        finally:
            doc.Close(False)
    finally:
        app.Quit()

    try:
        html = _leer_html(html_path)
    finally:
        try:
            os.remove(html_path)
        except OSError:
            pass

    parser = HTMLToMarkdown()
    parser.feed(html)
    parser.close()
    return parser.result()


# ----------------------------------------------------------------------- PDF
def _pdf_a_markdown(origen, ocr=False):
    if not _HAS_PDF:
        raise RuntimeError("para .pdf falta 'pypdf':  pip install pypdf")
    reader = PdfReader(origen)
    fitz_doc = None
    if ocr:
        if not _HAS_FITZ:
            raise RuntimeError("para OCR de .pdf falta 'pymupdf':  pip install pymupdf")
        fitz_doc = fitz.open(origen)
    out = []
    for i, page in enumerate(reader.pages, 1):
        texto = (page.extract_text() or "").strip()
        if ocr and not texto:
            texto = _ollama_ocr(_pdf_page_png(fitz_doc, i - 1))
        out.append("\n## Página %d\n" % i)
        out.append(texto)
    return "\n".join(out)


# --------------------------------------------------------------- OCR (ollama)
_ollama_ok_cache = None


def _ollama_ok(timeout=2.0):
    global _ollama_ok_cache
    if _ollama_ok_cache is None:
        try:
            urllib.request.urlopen(OLLAMA_HOST + "/api/tags", timeout=timeout).close()
            _ollama_ok_cache = True
        except Exception:
            _ollama_ok_cache = False
    return _ollama_ok_cache


def _ollama_ocr(imagen_bytes, timeout=300.0):
    if not _ollama_ok():
        raise RuntimeError(
            "ollama no responde en %s (¿corré `ollama serve`?)" % OLLAMA_HOST
        )
    payload = {
        "model": OCR_MODEL,
        "stream": False,
        "messages": [{
            "role": "user",
            "content": OCR_PROMPT,
            "images": [base64.b64encode(imagen_bytes).decode("ascii")],
        }],
    }
    req = urllib.request.Request(
        OLLAMA_HOST + "/api/chat",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError as e:
        raise RuntimeError("error llamando a ollama: %s" % e)
    return (data.get("message", {}).get("content") or "").strip()


def _pdf_page_png(fitz_doc, index):
    return fitz_doc[index].get_pixmap().tobytes("png")


def _imagen_a_markdown(origen):
    with open(origen, "rb") as f:
        data = f.read()
    return _ollama_ocr(data)


# ------------------------------------------------------------------- archivos
def _leer_texto(path):
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeError:
            continue
    with open(path, encoding="utf-8", errors="replace") as f:
        return f.read()


def convertir_uno(origen, ocr=False):
    ext = os.path.splitext(origen)[1].lower()
    if ext in TEXT_EXTS:
        return _leer_texto(origen)
    if ext in POWERPOINT_EXTS:
        return _powerpoint_a_markdown(origen)
    if ext in PDF_EXTS:
        return _pdf_a_markdown(origen, ocr=ocr)
    if ext in WORD_EXTS:
        return _word_a_markdown(origen)
    if ext in IMAGE_EXTS:
        if not ocr:
            raise ValueError("imagen: activá OCR para extraer su texto (--ocr)")
        return _imagen_a_markdown(origen)
    raise ValueError("extensión no soportada: %s" % ext)


def convertir(objetivo, salida=None, ocr=False, progress=None, log=print):
    if os.path.isdir(objetivo):
        scan = EXTS | IMAGE_EXTS if ocr else EXTS
        archivos = [
            os.path.join(objetivo, f)
            for f in sorted(os.listdir(objetivo))
            if os.path.splitext(f)[1].lower() in scan and not f.startswith("~$")
        ]
    else:
        archivos = [objetivo]

    if not archivos:
        log("No se encontraron archivos convertibles en: %s" % objetivo)
        return 0

    ok = 0
    for idx, origen in enumerate(archivos, 1):
        base = os.path.basename(origen)
        destino_dir = salida or os.path.dirname(origen)
        destino = os.path.join(destino_dir, os.path.splitext(base)[0] + ".md")
        if progress:
            progress(idx, len(archivos), base)
        try:
            texto = convertir_uno(origen, ocr=ocr)
            with open(destino, "w", encoding="utf-8") as f:
                f.write(texto)
            log("OK: %s -> %s (%d caracteres)"
                % (base, os.path.basename(destino), len(texto)))
            ok += 1
        except Exception as e:
            log("ERROR: %s: %s" % (base, e))

    log("Listo: %d/%d convertidos." % (ok, len(archivos)))
    return 0 if ok else 1


# ------------------------------------------------------------------- selftest
def selftest():
    tmp = tempfile.mkdtemp(prefix="mdtest_")
    try:
        # --- Word: docx + doc + pdf con título, párrafo, lista y tabla ---
        print("selftest: generando Word (docx/doc/pdf)...", flush=True)
        app = _word_app()
        try:
            doc = app.Documents.Add()
            sel = app.Selection
            sel.Style = doc.Styles(-2)   # wdStyleHeading1 (independiente del idioma)
            sel.TypeText("Titulo de prueba")
            sel.TypeParagraph()
            sel.Style = doc.Styles(-1)   # wdStyleNormal
            sel.TypeText("Parrafo normal de prueba.")
            sel.TypeParagraph()
            rng = sel.Range
            rng.ListFormat.ApplyBulletDefault()
            sel.TypeText("Elemento uno")
            sel.TypeParagraph()
            sel.TypeText("Elemento dos")
            sel.TypeParagraph()
            rng.ListFormat.RemoveNumbers()
            tbl = doc.Tables.Add(sel.Range, 2, 2)
            tbl.Cell(1, 1).Range.Text = "CabA"
            tbl.Cell(1, 2).Range.Text = "CabB"
            tbl.Cell(2, 1).Range.Text = "x"
            tbl.Cell(2, 2).Range.Text = "y"
            doc.SaveAs2(os.path.join(tmp, "doc.docx"), FileFormat=16)
            doc.SaveAs2(os.path.join(tmp, "doc.doc"), FileFormat=0)
            doc.SaveAs2(os.path.join(tmp, "doc.pdf"), FileFormat=17)
            doc.Close(False)
        finally:
            app.Quit()

        # --- PowerPoint: pptx + ppt con título y viñetas ---
        print("selftest: generando PowerPoint (pptx/ppt)...", flush=True)
        app = _ppt_app()
        try:
            pres = app.Presentations.Add(WithWindow=False)
            slide = pres.Slides.Add(1, 2)  # layout título + contenido
            slide.Shapes.Title.TextFrame.TextRange.Text = "Titulo slide"
            slide.Shapes.Placeholders(2).TextFrame.TextRange.Text = \
                "Linea uno\rLinea dos"
            pres.SaveAs(os.path.join(tmp, "ppt.pptx"), 24)
            pres.SaveAs(os.path.join(tmp, "ppt.ppt"), 1)
            pres.Close()
        finally:
            app.Quit()

        # --- convertir todo y verificar ---
        print("selftest: convirtiendo...", flush=True)
        fallos = []
        for nombre in ("doc.docx", "doc.doc", "doc.pdf", "ppt.pptx", "ppt.ppt"):
            origen = os.path.join(tmp, nombre)
            md = convertir_uno(origen)
            destino = os.path.join(tmp, os.path.splitext(nombre)[0] + ".md")
            with open(destino, "w", encoding="utf-8") as f:
                f.write(md)
            checks = {"docx": "# Titulo", "doc": "# Titulo", "pdf": "Titulo",
                      "pptx": "## Diapositiva", "ppt": "## Diapositiva"}
            if checks[nombre.split(".")[-1]] not in md:
                fallos.append("%s: falta %r" % (nombre, checks[nombre.split(".")[-1]]))
            print("  %s -> %s (%d caracteres)" % (nombre, os.path.basename(destino), len(md)))

        assert not fallos, "Verificación fallida:\n" + "\n".join(fallos)
        assert "|" in convertir_uno(os.path.join(tmp, "doc.docx")), "falta tabla"
        assert "- " in convertir_uno(os.path.join(tmp, "doc.docx")), "falta lista"
        print("SELFTEST OK: 5 formatos convertidos y verificados.")

        # --- OCR opcional (solo si ollama responde) ---
        if _HAS_FITZ and _ollama_ok():
            print("selftest: OCR local (ollama %s, puede tardar la primera vez)..."
                  % OCR_MODEL, flush=True)
            fdoc = fitz.open()
            fpage = fdoc.new_page()
            fpage.insert_text((72, 72), "OCR prueba 123")
            png = fpage.get_pixmap().tobytes("png")
            fdoc.close()
            txt = _ollama_ocr(png)
            assert ("prueba" in txt.lower()) or ("123" in txt), \
                "OCR no devolvió el texto esperado: %r" % txt[:200]
            print("  OCR OK: %r" % txt[:120])
        else:
            print("selftest: OCR omitido (ollama o pymupdf no disponibles).")
        return 0
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)


# ------------------------------------------------------------------------ GUI
_FONT_UI = ("Segoe UI", 10)
_FONT_TITLE = ("Segoe UI", 18, "bold")
_FONT_SUB = ("Segoe UI", 10)
_FONT_MONO = ("Consolas", 9)

_PALETTES = {
    "light": {
        "window": "#eef1f6", "panel": "#ffffff", "field": "#ffffff",
        "text": "#1c2430", "title": "#111827", "muted": "#6b7280",
        "border": "#dbe0e7", "accent": "#2563eb", "accent_hover": "#1d4ed8",
        "accent_fg": "#ffffff", "button": "#f2f4f8", "button_hover": "#e4e9f2",
        "disabled": "#dfe3ea", "disabled_fg": "#98a1ad", "trough": "#e9ecf2",
        "output_bg": "#f8fafc", "output_fg": "#1c2430",
    },
    "dark": {
        "window": "#0e1116", "panel": "#161b23", "field": "#1c222c",
        "text": "#e7eaef", "title": "#ffffff", "muted": "#98a2b1",
        "border": "#2a323d", "accent": "#3b82f6", "accent_hover": "#2563eb",
        "accent_fg": "#ffffff", "button": "#222936", "button_hover": "#2c3443",
        "disabled": "#252b35", "disabled_fg": "#5b6471", "trough": "#1b212a",
        "output_bg": "#12161c", "output_fg": "#e7eaef",
    },
}


def _detect_dark():
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        )
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 0
    except OSError:
        return False


class GuiApp:
    def __init__(self, root):
        self.root = root
        self.c = _PALETTES["dark" if _detect_dark() else "light"]
        self.events = queue.Queue()
        self.running = False
        self._total = 1

        root.title("Office → Markdown")
        root.geometry("700x560")
        root.minsize(600, 480)
        root.configure(bg=self.c["window"])

        self._build_style()
        self._build_header()
        self._build_card()
        self._build_progress()
        self._build_log()
        self.root.after(100, self._poll)

    # --------------------------------------------------------------- tema
    def _build_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        c = self.c
        style.configure(".", background=c["panel"], foreground=c["text"],
                        bordercolor=c["border"], lightcolor=c["border"],
                        darkcolor=c["border"], troughcolor=c["trough"],
                        focuscolor=c["accent"], font=_FONT_UI)
        style.configure("TFrame", background=c["panel"])
        style.configure("TLabel", background=c["panel"], foreground=c["text"])
        style.configure("Muted.TLabel", foreground=c["muted"])
        style.configure("TEntry", fieldbackground=c["field"],
                        foreground=c["text"], insertcolor=c["text"],
                        bordercolor=c["border"], lightcolor=c["border"],
                        darkcolor=c["border"])
        style.configure("TButton", background=c["button"],
                        foreground=c["text"], borderwidth=0,
                        focusthickness=0, padding=(10, 6))
        style.map("TButton",
                  background=[("active", c["button_hover"]),
                              ("pressed", c["button_hover"])])
        style.configure("Accent.TButton", background=c["accent"],
                        foreground=c["accent_fg"], borderwidth=0,
                        focusthickness=0, padding=(10, 9),
                        font=("Segoe UI", 11, "bold"))
        style.map("Accent.TButton",
                  background=[("active", c["accent_hover"]),
                              ("pressed", c["accent_hover"]),
                              ("disabled", c["disabled"])],
                  foreground=[("disabled", c["disabled_fg"])])
        style.configure("TCheckbutton", background=c["panel"],
                        foreground=c["text"])
        style.map("TCheckbutton", background=[("active", c["panel"])])
        style.configure("Horizontal.TProgressbar", background=c["accent"],
                        troughcolor=c["trough"], bordercolor=c["trough"],
                        lightcolor=c["accent"], darkcolor=c["accent"])

    # ------------------------------------------------------------- widgets
    def _build_header(self):
        c = self.c
        header = tk.Frame(self.root, bg=c["window"])
        header.pack(fill="x", padx=24, pady=(18, 2))
        tk.Label(header, text="Office → Markdown", bg=c["window"],
                 fg=c["title"], font=_FONT_TITLE).pack(anchor="w")
        tk.Label(
            header,
            text="Elegí una carpeta o archivo y convertí todo a Markdown, "
                 "con OCR local para escaneados.",
            bg=c["window"], fg=c["muted"], font=_FONT_SUB,
        ).pack(anchor="w", pady=(2, 0))

    def _build_card(self):
        c = self.c
        card = tk.Frame(self.root, bg=c["panel"], highlightthickness=1,
                        highlightbackground=c["border"])
        card.pack(fill="x", padx=24, pady=(10, 8))
        card.columnconfigure(1, weight=1)
        pad = {"padx": 12, "pady": 6}

        ttk.Label(card, text="Carpeta / archivo").grid(
            row=0, column=0, sticky="w", **pad)
        self.path_var = tk.StringVar()
        ttk.Entry(card, textvariable=self.path_var).grid(
            row=0, column=1, sticky="ew", **pad)
        btns = ttk.Frame(card)
        btns.grid(row=0, column=2, sticky="e", **pad)
        ttk.Button(btns, text="Carpeta…", command=self._pick_folder).pack(
            side="left", padx=(0, 6))
        ttk.Button(btns, text="Archivo…", command=self._pick_file).pack(
            side="left")

        ttk.Label(card, text="Guardar .md en").grid(
            row=1, column=0, sticky="w", **pad)
        self.out_var = tk.StringVar()
        ttk.Entry(card, textvariable=self.out_var).grid(
            row=1, column=1, sticky="ew", **pad)
        ttk.Button(card, text="Examinar…", command=self._pick_out).grid(
            row=1, column=2, sticky="e", **pad)

        self.ocr_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            card, text="OCR de escaneados/imágenes (minicpm-v vía ollama)",
            variable=self.ocr_var,
        ).grid(row=2, column=0, columnspan=3, sticky="w", **pad)

        self.run_btn = ttk.Button(card, text="Convertir", style="Accent.TButton",
                                  command=self.start)
        self.run_btn.grid(row=3, column=0, columnspan=3, sticky="ew",
                          padx=12, pady=(10, 12))

    def _build_progress(self):
        c = self.c
        bar_frame = tk.Frame(self.root, bg=c["window"])
        bar_frame.pack(fill="x", padx=24, pady=(0, 2))
        self.bar = ttk.Progressbar(bar_frame, style="Horizontal.TProgressbar",
                                   mode="determinate")
        self.bar.pack(fill="x")
        self.status = tk.Label(bar_frame, text="Listo para convertir.",
                               bg=c["window"], fg=c["muted"], font=_FONT_SUB,
                               anchor="w")
        self.status.pack(fill="x", pady=(4, 0))

    def _build_log(self):
        c = self.c
        ttk.Label(self.root, text="Registro", style="Muted.TLabel").pack(
            anchor="w", padx=24, pady=(8, 4))
        frame = tk.Frame(self.root, bg=c["output_bg"], highlightthickness=1,
                         highlightbackground=c["border"])
        frame.pack(fill="both", expand=True, padx=24, pady=(0, 12))
        self.log = tk.Text(frame, state="disabled", font=_FONT_MONO,
                           bg=c["output_bg"], fg=c["output_fg"],
                           insertbackground=c["text"], relief="flat",
                           padx=10, pady=8, wrap="word")
        scroll = ttk.Scrollbar(frame, command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    # ------------------------------------------------------------- acciones
    def _pick_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.path_var.set(path)

    def _pick_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("Documentos", "*.docx *.doc *.rtf *.pptx *.ppt *.pps "
                 "*.pdf *.txt *.png *.jpg *.jpeg *.bmp *.tif *.tiff "
                 "*.webp *.gif"),
                ("Todos", "*.*"),
            ])
        if path:
            self.path_var.set(path)

    def _pick_out(self):
        path = filedialog.askdirectory()
        if path:
            self.out_var.set(path)

    def _append(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        path = self.path_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Ruta", "Elegí una carpeta o archivo válido.")
            return
        if self.running:
            return
        self.running = True
        self.run_btn.configure(state="disabled")
        self.bar.configure(value=0, maximum=1)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        out = self.out_var.get().strip() or None
        ocr = self.ocr_var.get()
        threading.Thread(target=self._worker, args=(path, out, ocr),
                         daemon=True).start()

    def _worker(self, path, out, ocr):
        try:
            def progress(i, total, name):
                self.events.put(("progress", i, total, name))

            def log(line):
                self.events.put(("log", line))

            code = convertir(path, out, ocr=ocr, progress=progress, log=log)
            self.events.put(("done", code))
        except Exception as e:
            self.events.put(("log", "ERROR inesperado: %s" % e))
            self.events.put(("done", 1))

    def _poll(self):
        try:
            while True:
                ev = self.events.get_nowait()
                if ev[0] == "progress":
                    _, i, total, name = ev
                    self._total = total
                    self.bar.configure(maximum=total, value=i)
                    self.status.configure(text="%d/%d — %s" % (i, total, name))
                elif ev[0] == "log":
                    self._append(ev[1])
                elif ev[0] == "done":
                    self.bar.configure(value=self._total)
                    self.status.configure(
                        text="Listo." if ev[1] == 0
                        else "Terminó con errores (ver registro).")
                    self.run_btn.configure(state="normal")
                    self.running = False
        except queue.Empty:
            pass
        self.root.after(100, self._poll)


def run_gui():
    root = tk.Tk()
    GuiApp(root)
    root.mainloop()
    return 0


# ---------------------------------------------------------------------- main
def main(argv):
    if "--selftest" in argv:
        return selftest()
    if "--gui" in argv:
        return run_gui()

    ocr = "--ocr" in argv
    salida = None
    for flag in ("-o", "--salida"):
        if flag in argv:
            i = argv.index(flag)
            if i + 1 < len(argv):
                salida = argv[i + 1]

    posicionales = [a for a in argv if not a.startswith("-") and a != salida]

    # Sin archivo, carpeta ni salida: abrir la GUI (doble clic).
    if not posicionales and not salida and not ocr:
        return run_gui()

    objetivo = posicionales[0] if posicionales else os.getcwd()
    return convertir(objetivo, salida, ocr=ocr)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
