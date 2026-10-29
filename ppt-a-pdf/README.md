# ppt-a-pdf

Convierte todos los archivos `.ppt` y `.pptx` de una carpeta a **PDF**, guardándolos en la misma carpeta con el mismo nombre.

## Requisitos

- **Microsoft PowerPoint** instalado (se usa vía COM, así que la conversión tiene la fidelidad nativa de PowerPoint).
- Python + `pywin32`:

```powershell
pip install pywin32
```

> Solo funciona en **Windows** (usa la API COM de Office).

## Uso

```powershell
# Convierte los .ppt/.pptx de la carpeta donde está el script
python convertir_ppt_a_pdf.py

# O indicá cualquier carpeta
python convertir_ppt_a_pdf.py "C:\ruta\a\mi\carpeta"
```

## Qué hace

- Recorre la carpeta y encuentra todos los `.ppt` / `.pptx` (ignora los temporales `~$...`).
- Abre cada presentación en PowerPoint y la guarda como PDF (`SaveAs` con formato `ppSaveAsPDF`).
- Si el PDF ya existe, se **sobrescribe**.
