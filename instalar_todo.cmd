@echo off
setlocal EnableExtensions
title Instalacion completa - Control de ubicacion MIGH
cd /d "%~dp0"

echo ==================================================================
echo   INSTALACION COMPLETA - CONTROL DE UBICACION DE ALMACEN (MIGH)
echo   Este paso instala el programa, sus dependencias y el arranque
echo   automatico. No borra datos: si ya habia, los conserva.
echo ==================================================================
echo.

rem ---------- 1) localizar Python ----------
set "PYEXE="

if exist "%~dp0.venv\Scripts\python.exe" set "PYEXE=%~dp0.venv\Scripts\python.exe"

if not defined PYEXE (
  for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
)
if not defined PYEXE (
  for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
)

if not defined PYEXE (
  echo No se encontro Python en esta PC.
  where winget >nul 2>&1
  if errorlevel 1 (
    echo.
    echo   Instala Python 3.12 a mano desde:
    echo     https://www.python.org/downloads/windows/
    echo   Marcando la casilla "Add python.exe to PATH", y vuelve a correr este archivo.
    echo.
    pause
    exit /b 1
  )
  echo   Se instalara Python 3.12 con winget ^(puede tardar un par de minutos^)...
  winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
  set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
  for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
  if not defined PYEXE (
    for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
  )
)

if not defined PYEXE (
  echo.
  echo   Python quedo instalado pero esta ventana todavia no lo ve.
  echo   Cierra esta ventana y vuelve a ejecutar instalar_todo.cmd.
  echo.
  pause
  exit /b 1
)

echo Python: %PYEXE%
echo.

rem ---------- 2) instalacion: entorno, dependencias, datos, arranque ----------
"%PYEXE%" scripts\instalar.py --arranque %*
if errorlevel 1 (
  echo.
  echo La instalacion no termino bien. Revisa los mensajes de arriba.
  pause
  exit /b 1
)

rem ---------- 3) firewall (necesita permiso de administrador) ----------
if /i "%ALMACEN_SIN_FIREWALL%"=="1" (
  echo Aviso: se omitio el firewall ^(ALMACEN_SIN_FIREWALL=1^).
  goto :fin
)
net session >nul 2>&1
if errorlevel 1 (
  echo.
  echo Para que entren telefonos y otras PC falta abrir el puerto 8300.
  echo Se pedira permiso de administrador una sola vez...
  powershell -NoProfile -Command "Start-Process -FilePath '%~dp0abrir_puerto_firewall.cmd' -Verb RunAs"
) else (
  call "%~dp0abrir_puerto_firewall.cmd" /silencioso
)

:fin
echo.
echo ==================================================================
echo   LISTO.
echo.
echo   Ya arranca solo cada vez que se inicie sesion en esta PC.
echo   Para usarlo ahora:  arrancar.cmd
echo   En este equipo:     http://127.0.0.1:8300
echo   En el telefono:     la IP de esta PC en la red, puerto 8300
echo                       (la muestra servidor_lan.cmd)
echo.
echo   Primer ingreso: la consola del servidor muestra el usuario
echo   administrador y su contrasena inicial.
echo ==================================================================
if not "%ALMACEN_SILENCIOSO%"=="1" pause
