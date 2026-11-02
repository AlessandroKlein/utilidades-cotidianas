# office-a-markdown

Convierte documentos de Office y PDF a **Markdown** (`.md`): extrae solo el texto para pasárselo a una IA y gastar muchos menos tokens. El `.md` resultante pesa una fracción del `.docx`/`.pdf` original.

Formatos soportados:

| Extensión | Cómo se extrae |
|---|---|
| `.pptx` / `.ppt` / `.pps` | PowerPoint (COM) — una diapositiva por `## Diapositiva N`, títulos en `###`, resto en viñetas; tablas incluidas. |
| `.docx` / `.doc` / `.rtf` | Word (COM) — títulos en `#`…`######`, listas en `- `, tablas en `| … |`, negrita/cursiva conservadas. |
| `.pdf` | `pypdf` — una página por `## Página N` con su texto. |
| `.txt` | Lectura directa (UTF-8 o Windows-1252). |

> Solo texto: las imágenes, gráficos y fórmulas visuales se omiten (no se representan en markdown).

## Requisitos

- **Microsoft Word** y **PowerPoint** instalados (se usan vía COM, así que la extracción tiene la fidelidad nativa de Office).
- Python + `pywin32` y `pypdf`:

```powershell
pip install pywin32 pypdf
```

> Solo funciona en **Windows** (usa la API COM de Office).

## Uso

```powershell
# Convierte los archivos de la carpeta donde está el script
python office_a_markdown.py

# O una carpeta cualquiera
python office_a_markdown.py "C:\ruta\a\mi\carpeta"

# O un solo archivo
python office_a_markdown.py "C:\ruta\archivo.pptx"

# Guardar los .md en otra carpeta
python office_a_markdown.py -o salida "C:\ruta"

# Prueba end-to-end (genera y convierte un archivo de cada formato)
python office_a_markdown.py --selftest
```

## Qué hace

- Recorre la carpeta y encuentra los archivos convertibles (ignora los temporales `~$...`).
- Guarda cada resultado como `<nombre>.md` junto al original (o en la carpeta `-o`), sobrescribiendo si ya existe.
- Si hay un error en un archivo, lo reporta y sigue con el resto.
