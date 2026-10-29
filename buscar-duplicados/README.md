# buscar-duplicados

Encuentra archivos **duplicados por contenido** (no por nombre) dentro de una carpeta, para que sepas cuánto espacio podrías liberar. Por defecto solo los lista; con `--borrar` elimina las copias extra conservando una.

## Requisitos

- Python 3 (solo biblioteca estándar, no hay que instalar nada).
- Funciona en Windows, Linux y macOS.

## Uso

```powershell
# Listar duplicados de la carpeta actual (dry-run, no borra nada)
python buscar_duplicados.py

# Listar duplicados de otra carpeta
python buscar_duplicados.py "C:\ruta\a\mi\carpeta"

# Borrar las copias extra (pide confirmación)
python buscar_duplicados.py --borrar

# Borrar sin pedir confirmación (cuidado)
python buscar_duplicados.py --borrar --si
```

## Qué hace

- Recorre la carpeta **recursivamente** y agrupa los archivos por tamaño.
- Solo calcula el hash (`blake2b`) de los archivos que comparten tamaño, para no comparar todos contra todos.
- Muestra cada grupo de duplicados, cuántas copias hay y cuánto espacio se recuperaría.
- Con `--borrar`, conserva la primera copia de cada grupo (ordenada por ruta) y elimina el resto.

## Notas

- **Por defecto solo lista**, nunca borra. El borrado requiere `--borrar` y, salvo `--si`, confirmación interactiva.
- Los archivos **vacíos** (0 bytes) se ignoran a propósito: borrarlos no libera espacio y puede romper cosas.
- Los archivos se identifican por **contenido**, así que detecta duplicados aunque tengan nombres distintos.
