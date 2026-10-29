# Utilidades Cotidianas 🧰

Colección de scripts y herramientas de uso diario: pequeñas utilidades que resuelven tareas puntuales y que conviene tener guardadas para no perderlas.

Cada herramienta vive en su propia carpeta con su código y su `README.md` explicando para qué sirve y cómo se usa.

## Herramientas

| Carpeta | Para qué sirve |
|---|---|
| [`buscar-duplicados`](buscar-duplicados/) | Encuentra archivos duplicados por contenido (hash `blake2b`) y muestra el espacio recuperable; con `--borrar` elimina las copias extra. |
| [`organizar-descargas`](organizar-descargas/) | Mueve los archivos sueltos de Descargas a subcarpetas por tipo (Imágenes, Documentos, Música, Vídeos, Comprimidos, Instaladores, Otros). |
| [`ppt-a-pdf`](ppt-a-pdf/) | Convierte todos los `.ppt` / `.pptx` de una carpeta a PDF (usa PowerPoint instalado). |
| [`renombrar-masivo`](renombrar-masivo/) | Renombra archivos en lote: prefijo/sufijo, numeración, buscar-reemplazar y normalización de nombres. |

## Cómo usar

Entrá en la carpeta de la herramienta que necesites y seguí las instrucciones de su `README.md`.

## Agregar una herramienta nueva

Reglas **obligatorias** para cualquiera que suba código a este repo — leelas completas en [CONTRIBUTING.md](CONTRIBUTING.md):

1. Creá una carpeta con un nombre descriptivo en minúsculas (`mi-herramienta/`).
2. Adentro dejá el código y un `README.md` **bien documentado**: qué hace, requisitos, y cómo se usa.
3. Agregala a la tabla de arriba.
4. **Creá su [Release](https://github.com/AlessandroKlein/utilidades-cotidianas/releases)** con tag de versión (ej. `ppt-a-pdf-v0.1.0`) para dejar todo guardado y versionado.

> Sin documentación ni Release, el aporte se puede revertir.

## Licencia y uso

Este repo es **público pero no otorga ninguna licencia**: el código se puede **ver y reportar**, no copiar, reutilizar ni distribuir sin permiso del autor. Detalles legales en [SECURITY.md](SECURITY.md).
