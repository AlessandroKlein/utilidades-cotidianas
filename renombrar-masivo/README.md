# renombrar-masivo

Renombra archivos en lote dentro de una carpeta. Podés combinar: **prefijo**, **sufijo**, **numeración secuencial**, **buscar/reemplazar** y **normalización** de nombres (sin acentos, espacios convertidos a `_`). Por defecto solo muestra el plan; con `--aplicar` renombra.

## Requisitos

- Python 3 (solo biblioteca estándar, no hay que instalar nada).
- Funciona en Windows, Linux y macOS.

## Uso

```powershell
# Anteponer "vacaciones_" y numerar (dry-run: muestra el plan, no renombra)
python renombrar_masivo.py --prefijo "vacaciones_" --numerar

# Reemplazar "foto" por "img" y aplicar (pide confirmación)
python renombrar_masivo.py --buscar "foto" --reemplazar "img" --aplicar

# Normalizar nombres y agregar sufijo, sin pedir confirmación
python renombrar_masivo.py --normalizar --sufijo "_v2" --aplicar --si

# Sobre otra carpeta
python renombrar_masivo.py "C:\ruta\a\mi\carpeta" --numerar --aplicar
```

## Opciones

| Opción | Qué hace |
|---|---|
| `--prefijo TEXTO` | Antepone `TEXTO` al nombre. |
| `--sufijo TEXTO` | Agrega `TEXTO` al final del nombre (antes de la extensión). |
| `--buscar A --reemplazar B` | Reemplaza `A` por `B` en el nombre. Si no pasás `--reemplazar`, **elimina** `A`. |
| `--numerar` | Antepone un número secuencial (`01_`, `02_`, ...). |
| `--numerar-desde N` | Número inicial (por defecto `1`). |
| `--digitos N` | Ancho mínimo del número, relleno con ceros (por defecto `2`). |
| `--normalizar` | Quita acentos y caracteres raros; espacios → `_`. |
| `--aplicar` | Renombra de verdad (sin esto es dry-run). |
| `--si` | No pide confirmación. |

## Qué hace

- Trabaja sobre los **archivos** de la carpeta (no entra en subcarpetas, ignora archivos ocultos y al propio script).
- **No toca la extensión**: los cambios se aplican solo al nombre base.
- Renombra en dos fases, así los nombres que se cruzan entre sí (ej. `a`→`b` y `b`→`c`) no chocan.
- Si el nombre final ya existe, le agrega ` (2)`, ` (3)`, etc. **Nunca sobrescribe**.
- Si algo falla a mitad de camino, intenta **revertir** los cambios.

## Notas

- Por defecto es **dry-run**: siempre mirá el plan antes de aplicar.
- Necesitás pasar **al menos una** opción de renombrado; si no, el programa te lo avisa.
