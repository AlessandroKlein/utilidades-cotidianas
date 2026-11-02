# -*- coding: utf-8 -*-
r"""Convierte Office/PDF a Markdown (.doc/.docx/.ppt/.pps/.pptx/.pdf/.rtf/.txt).

Extrae el contenido con Word y PowerPoint instalados (COM/pywin32) y pypdf para
los .pdf, y guarda un .md junto a cada archivo. Ideal para pasar documentos a
una IA reduciendo tokens: el .md pesa mucho menos que el .pdf/.docx original.

Uso:
    python office_a_markdown.py                  # convierte la carpeta actual
    python office_a_markdown.py "C:\\ruta\\carpeta"
    python office_a_markdown.py "C:\\ruta\\archivo.pptx"
    python office_a_markdown.py -o salida "C:\\ruta"   # .md en otra carpeta
    python office_a_markdown.py --selftest       # prueba end-to-end

Solo texto: las imágenes y gráficos se omiten (no se representan en markdown).
"""
import os
import re
import sys
import tempfile
from html.parser import HTMLParser

import win32com.client

try:
    from pypdf import PdfReader
    _HAS_PDF = True
except ImportError:
    _HAS_PDF = False

POWERPOINT_EXTS = {".ppt", ".pps", ".pptx"}
WORD_EXTS = {".doc", ".docx", ".rtf"}
PDF_EXTS = {".pdf"}
TEXT_EXTS = {".txt"}
EXTS = POWERPOINT_EXTS | WORD_EXTS | PDF_EXTS | TEXT_EXTS

WD_FORMAT_FILTERED_HTML = 10
PP_PLACEHOLDER_TITLE = 1


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
def _pdf_a_markdown(origen):
    if not _HAS_PDF:
        raise RuntimeError("para .pdf falta 'pypdf':  pip install pypdf")
    reader = PdfReader(origen)
    out = []
    for i, page in enumerate(reader.pages, 1):
        texto = page.extract_text() or ""
        out.append("\n## Página %d\n" % i)
        out.append(texto)
    return "\n".join(out)


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


def convertir_uno(origen):
    ext = os.path.splitext(origen)[1].lower()
    if ext in TEXT_EXTS:
        return _leer_texto(origen)
    if ext in POWERPOINT_EXTS:
        return _powerpoint_a_markdown(origen)
    if ext in PDF_EXTS:
        return _pdf_a_markdown(origen)
    if ext in WORD_EXTS:
        return _word_a_markdown(origen)
    raise ValueError("extensión no soportada: %s" % ext)


def convertir(objetivo, salida=None):
    if os.path.isdir(objetivo):
        archivos = [
            os.path.join(objetivo, f)
            for f in sorted(os.listdir(objetivo))
            if os.path.splitext(f)[1].lower() in EXTS and not f.startswith("~$")
        ]
    else:
        archivos = [objetivo]

    if not archivos:
        print("No se encontraron archivos convertibles en: %s" % objetivo)
        return 0

    ok = 0
    for origen in archivos:
        base = os.path.basename(origen)
        destino_dir = salida or os.path.dirname(origen)
        destino = os.path.join(
            destino_dir, os.path.splitext(base)[0] + ".md"
        )
        try:
            texto = convertir_uno(origen)
            with open(destino, "w", encoding="utf-8") as f:
                f.write(texto)
            print("OK: %s -> %s (%d caracteres)" % (base, os.path.basename(destino), len(texto)))
            ok += 1
        except Exception as e:
            print("ERROR: %s: %s" % (base, e))

    print("Listo: %d/%d convertidos." % (ok, len(archivos)))
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
        return 0
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------- main
def main(argv):
    if "--selftest" in argv:
        return selftest()

    salida = None
    for flag in ("-o", "--salida"):
        if flag in argv:
            i = argv.index(flag)
            if i + 1 < len(argv):
                salida = argv[i + 1]

    posicionales = [a for a in argv if not a.startswith("-") and a not in (salida,)]
    objetivo = posicionales[0] if posicionales else os.getcwd()
    return convertir(objetivo, salida)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
