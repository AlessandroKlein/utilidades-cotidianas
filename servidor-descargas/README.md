# servidor-descargas

Comparte una carpeta por la red de tu casa: los demás equipos abren el navegador y descargan los archivos, **sin instalar nada en ellos**. Solo la computadora que levanta el servidor necesita Python.

## Requisitos

- Python 3 (solo biblioteca estándar, no hay que instalar nada).
- Funciona en Windows, Linux y macOS.
- El firewall de Windows debe permitir Python en la red local (la primera vez te lo pregunta: aceptá).

## Uso

```powershell
# Compartir la carpeta "archivos" (junto al script) en el puerto 8000
python servidor_descargas.py

# Compartir otra carpeta
python servidor_descargas.py "C:\ruta\a\mis\archivos"

# Usar otro puerto (ej. 8080)
python servidor_descargas.py "C:\ruta\a\mis\archivos" 8080
```

Al arrancar imprime las direcciones para entrar desde los demás equipos, por ejemplo:

```
En otros equipos de tu casa:
    http://192.168.0.15:8000
```

## Cómo se usa desde otro equipo

1. Poné los archivos que quieras compartir en la carpeta compartida.
2. Desde cualquier computadora, celular o tablet **en la misma red Wi-Fi**, abrí el navegador y entrá a `http://<IP-del-servidor>:8000`.
3. Ves la lista de archivos con tamaño y fecha, un buscador, y por archivo un botón **Descargar** (y **Ver** para los de texto: `.txt`, `.md`, `.yml`).

## Seguridad

- **Solo archivos, no carpetas**: solo se listan y descargan archivos sueltos de la carpeta compartida.
- **Bloqueo de rutas**: no se puede salir de la carpeta compartida (los intentos tipo `../` dan 404).
- El servidor escucha solo en tu **red local** (`0.0.0.0`), no es accesible desde internet salvo que reenvíes el puerto en tu router. No lo hagas: está pensado para uso dentro de tu casa, sin contraseña.

## Verificación

```powershell
python servidor_descargas.py --selftest
```
