@echo off
setlocal
title Arranque automatico - Almacen MIGH
cd /d "%~dp0"

rem Deja el sistema corriendo al iniciar sesion, sin ventana visible.
rem Para quitarlo: quitar_arranque.cmd

set "PY="
if exist "%~dp0.venv\Scripts\pythonw.exe" set "PY=%~dp0.venv\Scripts\pythonw.exe"
if not defined PY for /d %%I in ("%LOCALAPPDATA%\hermes\installs\*") do (
  for /d %%E in ("%%I\environments\*") do (
    if exist "%%E\venv\Scripts\pythonw.exe" if not defined PY set "PY=%%E\venv\Scripts\pythonw.exe"
  )
)
if not defined PY (
  echo No se encontro pythonw.exe. Corre instalar.cmd primero.
  pause
  exit /b 1
)

set "CMD=cd /d ""%~dp0"" && ""%PY%"" app\server.py"

schtasks /Create /F /TN "Almacen MIGH" /TR "%CMD%" /SC ONLOGON /RL LIMITED
if errorlevel 1 (
  echo.
  echo No se pudo crear la tarea. Abre una consola como administrador y vuelve a correrlo.
) else (
  echo.
  echo Listo: el sistema arrancara solo al iniciar sesion en esta PC.
  echo Comprueba con:  schtasks /Query /TN "Almacen MIGH"
  echo Ahora mismo puedes arrancarlo con arrancar.cmd
)
echo.
pause
