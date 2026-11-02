# office-a-markdown

Convierte documentos de Office, PDF e imágenes a **Markdown** (`.md`): extrae solo el texto para pasárselo a una IA y gastar muchos menos tokens. El `.md` resultante pesa una fracción del `.docx`/`.pdf` original.

Tiene **interfaz gráfica** (elegís carpeta o archivo, con barra de progreso) y **OCR local opcional** (para PDFs escaneados e imágenes, usando `minicpm-v` vía ollama).

## Formatos soportados

| Extensión | Cómo se extrae |
|---|---|
| `.pptx` / `.ppt` / `.pps` | PowerPoint (COM) — una diapositiva por `## Diapositiva N`, títulos en `###`, resto en viñetas; tablas incluidas. |
| `.docx` / `.doc` / `.rtf` | Word (COM) — títulos en `#`…`######`, listas en `- `, tablas en `| … |`, negrita/cursiva conservadas. |
| `.pdf` | `pypdf` — una página por `## Página N` con su texto. Con `--ocr`, las páginas sin texto (escaneadas) se rasterizan y pasan por ollama. |
| `.png` / `.jpg` / `.jpeg` / `.bmp` / `.tif` / `.webp` / `.gif` | OCR local (minicpm-v) — solo con `--ocr` o la opción de la GUI. |
| `.txt` | Lectura directa (UTF-8 o Windows-1252). |

> Solo texto: las imágenes, gráficos y fórmulas visuales se omiten (no se representan en markdown), salvo que actives OCR y se extraiga el texto visible.

## Requisitos

- **Microsoft Word** y **PowerPoint** instalados (se usan vía COM, así que la extracción tiene la fidelidad nativa de Office).
- Python + `pywin32` y `pypdf`:
  ```powershell
  pip install pywin32 pypdf
  ```
- **Opcional (OCR de escaneados/imágenes):** [ollama](https://ollama.com) con el modelo `minicpm-v`, y `pymupdf`:
  ```powershell
  ollama pull minicpm-v
  pip install pymupdf
  ```

> Solo funciona en **Windows** (usa la API COM de Office). El OCR usa el modelo configurado en `OCR_MODEL` (por defecto `minicpm-v`) y la URL en `OLLAMA_HOST` (por defecto `http://127.0.0.1:11434`).

## Uso

**Interfaz gráfica** (doble clic sobre el `.py`, o `--gui`):

```powershell
python office_a_markdown.py
python office_a_markdown.py --gui
```

Elegís carpeta o archivo, activás/desactivás OCR y convertís viendo la barra de progreso.

**Línea de comandos:**

```powershell
# Convierte los archivos de la carpeta donde está el script
python office_a_markdown.py "C:\ruta\a\mi\carpeta"

# O un solo archivo
python office_a_markdown.py "C:\ruta\archivo.pptx"

# OCR de PDFs escaneados e imágenes (requiere ollama + minicpm-v)
python office_a_markdown.py --ocr "C:\ruta"

# Guardar los .md en otra carpeta
python office_a_markdown.py -o salida "C:\ruta"

# Prueba end-to-end (genera y convierte un archivo de cada formato + OCR)
python office_a_markdown.py --selftest
```

## Qué hace

- Recorre la carpeta y encuentra los archivos convertibles (ignora los temporales `~$...`).
- Guarda cada resultado como `<nombre>.md` junto al original (o en la carpeta `-o`), sobrescribiendo si ya existe.
- Con OCR, las páginas de PDF sin texto y los archivos de imagen se pasan a `minicpm-v` para extraer su texto.
- Si hay un error en un archivo, lo reporta y sigue con el resto.

> El OCR es un modelo de visión: transcribe lo que ve, pero puede equivocarse en letra manuscrita o tablas densas. Para documentos digitales (con capa de texto) conviene dejarlo apagado: es más rápido y no se arriesga a alucinar.
