@echo off
setlocal
title Quitar arranque automatico - Almacen MIGH
cd /d "%~dp0"

set "PYEXE="
if exist "%~dp0.venv\Scripts\python.exe" set "PYEXE=%~dp0.venv\Scripts\python.exe"
if not defined PYEXE (
  for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
)
if not defined PYEXE (
  for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
)

if defined PYEXE (
  "%PYEXE%" scripts\instalar.py --quitar-arranque
) else (
  echo No se encontro Python: no se pudo quitar el arranque desde aqui.
)

schtasks /Delete /F /TN "Almacen MIGH" >nul 2>&1
if errorlevel 1 (
  echo No habia tarea programada que borrar.
) else (
  echo Tarea programada eliminada.
)

echo.
echo El sistema ya no arranca solo. Puedes seguir usandolo con arrancar.cmd
pause
