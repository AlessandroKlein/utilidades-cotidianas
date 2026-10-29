# Cómo contribuir / subir código

Este repo es una colección de utilidades de uso diario. Antes de subir **cualquier cosa**, leé y cumplí estas reglas. Son **obligatorias**: si un aporte no las cumple, se puede revertir.

## Reglas obligatorias

### 1. Documentá bien lo que subís

Cada herramienta va en su **propia carpeta** con:

- El **código**.
- Un **`README.md` propio** que explique claramente:
  - **Para qué sirve**.
  - **Requisitos**: dependencias, sistema operativo, programas instalados (ej. PowerPoint, LibreOffice, `pywin32`).
  - **Cómo se usa**: pasos concretos y ejemplos de comandos.
- Si el código necesita secretos, variables de entorno o configuración previa, **aclaralo en el README**.

### 2. Creá un Release para dejar todo guardado

Después de subir el código, **creá su Release** en GitHub para que quede versionado y no se pierda entre los cambios de `main`:

1. Creá una etiqueta (tag) con versión, por ejemplo `ppt-a-pdf-v0.1.0`.
2. Publicá un **Release** con notas que resuman **qué hace** y **qué cambió**.

Así cada herramienta queda archivada con su propia versión, descargable de forma estable.

### 3. Nada de código ajeno sin permiso

No subas código de terceros que tenga licencia incompatible o que no tengas derecho a redistribuir. Si mezclás algo de otro autor, aclaralo y respetá su licencia.

### 4. Sin licencia, no se usa

Este repo **no otorga ninguna licencia** de uso, copia o distribución (ver [SECURITY.md](SECURITY.md)). El código es público para **ver y reportar**, no para reutilizar sin permiso del autor.

## Checklist antes de hacer push

- [ ] ¿La carpeta tiene su `README.md` con qué hace, requisitos y cómo se usa?
- [ ] ¿Está agregada a la tabla **"Herramientas"** del [`README.md`](README.md) raíz?
- [ ] ¿Creaste el **tag** y el **Release** correspondiente?
- [ ] ¿No hay secretos, tokens ni contraseñas en el código?
- [ ] ¿No hay código de terceros sin licencia compatible?
