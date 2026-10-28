"""Prepara la carpeta de distribución para la USB: paquete, actualizador y lanzadores.

    python scripts\\preparar_usb.py --destino "D:\\ALMACEN_MIGH_ACTUALIZACION"
    python scripts\\preparar_usb.py --destino "D:\\..." --paquete dist\\almacen-ubicaciones_v1.5.0_xxx.zip

Deja en el destino:
    ACTUALIZAR.cmd            doble clic en una PC con el sistema instalado
    INSTALAR_EN_PC_NUEVA.cmd  doble clic en una PC sin el sistema
    actualizar.py             el actualizador (biblioteca estándar, sin dependencias)
    version.txt               versión que lleva esta USB
    instalacion\\*.zip         el paquete vigente
    versiones\\*.zip           paquetes anteriores (para volver atrás con --forzar)
"""
import sys

# La consola de Windows no siempre es UTF-8: sin esto, imprimir «→» o «·» lanza
# UnicodeEncodeError y el script parece fallar aunque ya haya hecho su trabajo.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

import argparse
import datetime as dt
import hashlib
import pathlib
import shutil
import zipfile

RAIZ = pathlib.Path(__file__).resolve().parents[1]
DIST = RAIZ / "dist"

ACTUALIZAR_CMD = r"""@echo off
setlocal
title Actualizar Almacen MIGH
cd /d "%~dp0"

set "PAQUETE="
for /f "delims=" %%F in ('dir /b /o-d "instalacion\almacen-ubicaciones_*.zip" 2^>nul') do (
  if not defined PAQUETE set "PAQUETE=%CD%\instalacion\%%F"
)
if not defined PAQUETE (
  echo No hay ningun paquete en la carpeta instalacion\.
  pause
  exit /b 2
)

rem ---- 1) donde esta instalado el sistema en esta PC ----
set "INSTALADA="
for %%D in (
  "%USERPROFILE%\Desktop\almacen-ubicaciones"
  "C:\Users\Almacen Migh\Desktop\almacen-ubicaciones"
  "C:\almacen-ubicaciones"
  "D:\almacen-ubicaciones"
  "%USERPROFILE%\almacen-ubicaciones"
  "%USERPROFILE%\Desktop\almacen-ubicaciones-datos"
) do (
  if not defined INSTALADA if exist "%%~D\app\server.py" set "INSTALADA=%%~D"
)

rem ---- 2) con que Python (el de la instalacion, si tiene entorno) ----
set "PYEXE="
if defined INSTALADA if exist "%INSTALADA%\.venv\Scripts\python.exe" set "PYEXE=%INSTALADA%\.venv\Scripts\python.exe"
if not defined PYEXE (
  for %%D in ("%INSTALADA%" "C:\almacen-ubicaciones" "D:\almacen-ubicaciones" "%USERPROFILE%\Desktop\almacen-ubicaciones") do (
    if not defined PYEXE if exist "%%~D\.venv\Scripts\python.exe" set "PYEXE=%%~D\.venv\Scripts\python.exe"
  )
)
if not defined PYEXE (
  for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do if not defined PYEXE set "PYEXE=%%P"
)
if not defined PYEXE (
  for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do if not defined PYEXE set "PYEXE=%%P"
)
if not defined PYEXE (
  echo No se encontro Python en esta PC: instala el sistema con INSTALAR_EN_PC_NUEVA.cmd o instala Python 3.10 o mayor.
  pause
  exit /b 2
)

set "SIMULANDO="
set "ARGS=%*"
if not "%ARGS%"=="%ARGS:--simular=%" set "SIMULANDO=1"

echo Paquete    : %PAQUETE%
if defined INSTALADA (echo Instalacion: %INSTALADA%) else (echo Instalacion: se busca automaticamente)
echo Python     : %PYEXE%
echo.

if defined INSTALADA (
  "%PYEXE%" "%~dp0actualizar.py" --paquete "%PAQUETE%" --instalada "%INSTALADA%" %*
) else (
  "%PYEXE%" "%~dp0actualizar.py" --paquete "%PAQUETE%" %*
)
set "CODIGO=%ERRORLEVEL%"
echo.
echo ==============================================================
if not "%CODIGO%"=="0" goto migh_fallo
if defined SIMULANDO goto migh_simulacion
echo   RESULTADO: ACTUALIZACION CORRECTA
echo   El sistema ya corre la version nueva y los datos siguen intactos.
goto migh_fin

:migh_simulacion
echo   RESULTADO: SIMULACION CORRECTA - no se modifico nada
echo   Para actualizar de verdad, ejecuta ACTUALIZAR.cmd sin --simular.
goto migh_fin

:migh_fallo
echo   RESULTADO: NO SE ACTUALIZO  -  codigo %CODIGO%
echo   Revisa el mensaje de arriba; el sistema quedo como estaba.
echo   Detalle guardado en: ultima_actualizacion.txt

:migh_fin
echo ==============================================================
if not "%ALMACEN_SILENCIOSO%"=="1" pause
exit /b %CODIGO%
"""

INSTALAR_CMD = r"""@echo off
setlocal
title Instalar Almacen MIGH en esta PC
cd /d "%~dp0"

set "PAQUETE="
for /f "delims=" %%F in ('dir /b /o-d "instalacion\almacen-ubicaciones_*.zip" 2^>nul') do (
  if not defined PAQUETE set "PAQUETE=%CD%\instalacion\%%F"
)
if not defined PAQUETE (
  echo No hay ningun paquete en la carpeta instalacion\.
  pause
  exit /b 2
)

set "DESTINO=C:\almacen-ubicaciones"
echo Se va a instalar en %DESTINO%
if exist "%DESTINO%\app\server.py" (
  echo Ya hay una instalacion ahi: usa ACTUALIZAR.cmd para actualizarla.
  pause
  exit /b 2
)
powershell -NoProfile -Command "Expand-Archive -LiteralPath '%PAQUETE%' -DestinationPath '%DESTINO%' -Force"
if errorlevel 1 (
  echo No se pudo descomprimir el paquete.
  pause
  exit /b 1
)
if exist "%DESTINO%\instalar_todo.cmd" (
  pushd "%DESTINO%"
  call instalar_todo.cmd
  popd
) else (
  echo Paquete descomprimido en %DESTINO%. Ejecuta instalar_todo.cmd dentro de esa carpeta.
  pause
)
exit /b 0
"""

LEEME = """# Actualizar el sistema del almacén desde esta USB

## Qué hay aquí

| Archivo | Para qué |
|---|---|
| `ACTUALIZAR.cmd` | doble clic en una PC que **ya tiene** el sistema instalado |
| `INSTALAR_EN_PC_NUEVA.cmd` | doble clic en una PC **sin** el sistema |
| `instalacion\\*.zip` | la versión vigente (current best version) |
| `versiones\\*.zip` | versiones anteriores, por si hay que volver atrás |
| `version.txt` | qué versión trae esta USB |

## Actualizar una PC (30 segundos)

1. Conecta la USB.
2. Abre esta carpeta y doble clic en **ACTUALIZAR.cmd**.
3. Acepta si Windows pide permisos y espera el mensaje final.

El actualizador:

- compara la versión de la USB con la instalada y **no baja de versión** salvo que se pida;
- respalda el código y los datos en `data\\respaldos\\pre_actualizacion_<fecha>.zip`
  (bitácora con copia consistente, Excel de los tres inventarios y catálogo);
- **detiene el servidor**, copia el código, retira los archivos que ya no existen
  en la versión nueva (solo los que instaló un paquete anterior) y **nunca toca `data\\`**;
- verifica que la versión nueva importe y que openpyxl esté disponible;
- vuelve a levantar el servidor y comprueba que responda;
- si algo falla, **restaura el código anterior** automáticamente y lo avisa.

## Si la PC no es la de siempre

```bat
ACTUALIZAR.cmd --instalada "D:\almacen-ubicaciones"
ACTUALIZAR.cmd --simular          :: solo compara versiones, no escribe nada
ACTUALIZAR.cmd --forzar           :: permite instalar una versión menor (volver atrás)
```

## Volver a una versión anterior

Copia el ZIP de `versiones\\` a `instalacion\\` y corre `ACTUALIZAR.cmd --forzar`.

## Qué NO se copia nunca

`data\\` (inventarios, bitácora, catálogo, salidas, respaldos), `.venv\\`, `.git\\`.
La actualización cambia el programa; los datos del almacén se quedan intactos.
"""


def _crlf(texto):
    """Fin de linea de Windows, sin duplicar los que ya vienen en el archivo."""
    return (texto.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n"))


def sha256(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def version_del_paquete(paquete):
    with zipfile.ZipFile(paquete) as z:
        if "VERSION.txt" not in z.namelist():
            raise ValueError(f"{paquete.name} no trae VERSION.txt (genera el paquete con empaquetar.py)")
        return z.read("VERSION.txt").decode("utf-8", "ignore").strip().splitlines()[0].strip()


def preparar(destino, paquete=None, mover_anteriores=True):
    destino = pathlib.Path(destino)
    paquete = pathlib.Path(paquete) if paquete else _paquete_mas_nuevo()
    if not paquete.exists():
        raise FileNotFoundError(f"no existe el paquete {paquete}")
    version = version_del_paquete(paquete)

    carpeta_instalacion = destino / "instalacion"
    carpeta_versiones = destino / "versiones"
    carpeta_instalacion.mkdir(parents=True, exist_ok=True)
    carpeta_versiones.mkdir(parents=True, exist_ok=True)

    # el paquete que estaba vigente pasa al histórico (para poder volver atrás)
    anteriores = []
    if mover_anteriores:
        for viejo in sorted(carpeta_instalacion.glob("almacen-ubicaciones_*.zip")):
            if viejo.name == paquete.name:
                continue
            shutil.move(str(viejo), str(carpeta_versiones / viejo.name))
            anteriores.append(viejo.name)

    shutil.copy2(paquete, carpeta_instalacion / paquete.name)
    shutil.copy2(RAIZ / "scripts" / "actualizar.py", destino / "actualizar.py")
    (destino / "version.txt").write_text(
        f"{version}\nactualizado: {dt.datetime.now():%Y-%m-%d %H:%M}\npaquete: {paquete.name}\n",
        encoding="utf-8")
    (destino / "ACTUALIZAR.cmd").write_text(ACTUALIZAR_CMD.replace("\n", "\r\n"), encoding="utf-8")
    (destino / "INSTALAR_EN_PC_NUEVA.cmd").write_text(INSTALAR_CMD.replace("\n", "\r\n"),
                                                      encoding="utf-8")
    (destino / "LEEME_ACTUALIZACION.md").write_text(LEEME, encoding="utf-8")
    return {"destino": destino, "version": version, "paquete": carpeta_instalacion / paquete.name,
            "anteriores": anteriores, "sha256": sha256(carpeta_instalacion / paquete.name)}


def _paquete_mas_nuevo():
    paquetes = sorted(DIST.glob("almacen-ubicaciones_*.zip"), key=lambda p: p.stat().st_mtime)
    if not paquetes:
        raise FileNotFoundError(f"no hay paquetes en {DIST}; corre antes empaquetar.py")
    return paquetes[-1]  # type: ignore[return-value]


def main(argv=None):
    ap = argparse.ArgumentParser(description="Prepara la carpeta de actualización para la USB")
    ap.add_argument("--destino", required=True, help="carpeta en la USB, p. ej. D:\\ALMACEN_MIGH_ACTUALIZACION")
    ap.add_argument("--paquete", default=None, help="ZIP a distribuir (por omisión el más nuevo de dist\\)")
    args = ap.parse_args(argv)
    try:
        resultado = preparar(args.destino, args.paquete)
    except (FileNotFoundError, ValueError) as e:
        print(f"ERROR: {e}")
        return 2
    print(f"Destino    : {resultado['destino']}")
    print(f"Versión    : {resultado['version']}")
    print(f"Paquete    : {resultado['paquete'].name}")
    print(f"sha256     : {resultado['sha256']}")
    if resultado["anteriores"]:
        print(f"Al histórico: {', '.join(resultado['anteriores'])}")
    print("\nEn la PC de destino: doble clic en ACTUALIZAR.cmd")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
