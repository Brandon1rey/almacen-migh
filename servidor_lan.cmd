@echo off
setlocal
title Almacen MIGH - servidor de red (una sola PC con los datos)
cd /d "%~dp0"

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY for /d %%I in ("%LOCALAPPDATA%\hermes\installs\*") do (
  for /d %%E in ("%%I\environments\*") do (
    if exist "%%E\venv\Scripts\python.exe" if not defined PY set "PY=%%E\venv\Scripts\python.exe"
  )
)
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo No se encontro Python. Corre instalar.cmd primero.
  pause
  exit /b 1
)

echo ==================================================================
echo   MODO RED: esta PC guarda los datos, las demas solo abren el navegador.
echo   Las otras PC entran a la direccion que aparece abajo (Otras PC).
echo.
echo   Si no abre desde otra PC, falta abrir el puerto en el firewall.
echo   Ejecuta UNA VEZ en una consola como administrador:
echo.
echo     netsh advfirewall firewall add rule name="Almacen MIGH 8300" dir=in action=allow protocol=TCP localport=8300
echo.
echo   Solo para la red de la planta. No exponer a internet.
echo ==================================================================
echo.

set "ALMACEN_HOST=0.0.0.0"
start "" http://127.0.0.1:8300
"%PY%" app\server.py
pause
