# organizar-descargas

Mueve los archivos sueltos de la carpeta **Descargas** (o la carpeta que indiques) a subcarpetas por tipo: **Imágenes**, **Documentos**, **Música**, **Vídeos**, **Comprimidos**, **Instaladores** y **Otros**. Nunca borra nada, solo mueve.

## Requisitos

- Python 3 (usa solo la biblioteca estándar, no hay que instalar nada).
- Funciona en **Windows**, **Linux** y **macOS**. En Windows detecta la carpeta Descargas real (aunque esté redirigida a otro disco).

## Uso

```powershell
# Ver qué haría sobre la Descargas real (dry-run, no mueve nada)
python organizar_descargas.py

# Mover de verdad (pide confirmación)
python organizar_descargas.py --aplicar

# Mover sin pedir confirmación (cuidado)
python organizar_descargas.py --aplicar --si

# Organizar otra carpeta
python organizar_descargas.py "C:\ruta\a\mi\carpeta" --aplicar
```

## Qué hace

- Recorre **solo el nivel superior** de la carpeta (no entra en las subcarpetas).
- Clasifica cada archivo por su extensión y lo mueve a la subcarpeta correspondiente, creándola si no existe.
- Los archivos sin extensión conocida van a **Otros**.
- Ignora carpetas, temporales de Office (`~$...`) y `desktop.ini`.
- Si en el destino ya hay un archivo con el mismo nombre, le agrega ` (2)`, ` (3)`, etc. **Nunca sobrescribe**.

## Notas

- Por defecto es **dry-run**: solo muestra el plan. Para mover hay que pasar `--aplicar`.
- Es seguro (no borra ni sobrescribe), pero conviene mirar el dry-run antes de aplicar la primera vez.
