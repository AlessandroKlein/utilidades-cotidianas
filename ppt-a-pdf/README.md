# ppt-a-pdf

Convierte todos los archivos PowerPoint (`.ppt`, `.pptx`, `.pps`, `.ppsx`) de una carpeta a **PDF**, guardándolos en la misma carpeta con el mismo nombre.

## Requisitos

- **Microsoft PowerPoint** instalado (se usa vía COM, así que la conversión tiene la fidelidad nativa de PowerPoint).
- Python + `pywin32`:

```powershell
pip install pywin32
```

> Solo funciona en **Windows** (usa la API COM de Office).

## Uso

```powershell
# Convierte los PowerPoint (.ppt/.pptx/.pps/.ppsx) de la carpeta donde está el script
python convertir_ppt_a_pdf.py

# O indicá cualquier carpeta
python convertir_ppt_a_pdf.py "C:\ruta\a\mi\carpeta"
```

## Qué hace

- Recorre la carpeta y encuentra todos los PowerPoint (`.ppt` / `.pptx` / `.pps` / `.ppsx`); ignora los temporales `~$...` y cualquier otro formato (PDF, docx, etc.).
- Abre cada presentación en PowerPoint y la guarda como PDF (`SaveAs` con formato `ppSaveAsPDF`).
- Si un archivo falla (bloqueado, corrupto o en sincronización), lo salta, lo reporta y continúa con el resto.
- Si el PDF ya existe, se **sobrescribe**.
