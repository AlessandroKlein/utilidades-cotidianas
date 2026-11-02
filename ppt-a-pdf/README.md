# ppt-a-pdf

Convierte todos los archivos PowerPoint (`.ppt`, `.pptx`, `.pps`, `.ppsx`) y Word (`.doc`, `.docx`) de una carpeta a **PDF**, guardándolos en la misma carpeta con el mismo nombre.

## Requisitos

- **Microsoft PowerPoint** y **Microsoft Word** instalados (se usan vía COM, así que la conversión tiene la fidelidad nativa de Office).
- Python + `pywin32`:

```powershell
pip install pywin32
```

> Solo funciona en **Windows** (usa la API COM de Office).

## Uso

```powershell
# Convierte los PowerPoint (.ppt/.pptx/.pps/.ppsx) y Word (.doc/.docx) de la carpeta donde está el script
python convertir_ppt_a_pdf.py

# O indicá cualquier carpeta
python convertir_ppt_a_pdf.py "C:\ruta\a\mi\carpeta"
```

## Qué hace

- Recorre la carpeta y encuentra todos los PowerPoint (`.ppt` / `.pptx` / `.pps` / `.ppsx`) y Word (`.doc` / `.docx`); ignora los temporales `~$...` y cualquier otro formato (PDF, etc.).
- Abre cada presentación en PowerPoint y la guarda como PDF (`SaveAs` con formato `ppSaveAsPDF`), y cada documento en Word (`SaveAs` con `wdFormatPDF` = 17).
- Si un archivo falla (bloqueado, corrupto o en sincronización), lo salta, lo reporta y continúa con el resto.
- Si el PDF ya existe, se **sobrescribe**.
