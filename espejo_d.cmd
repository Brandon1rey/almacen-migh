@echo off
setlocal
title Espejo del proyecto en D:\ - Almacen MIGH
cd /d "%~dp0"

set "PYEXE="
if exist "%~dp0.venv\Scripts\python.exe" set "PYEXE=%~dp0.venv\Scripts\python.exe"
if not defined PYEXE (
  for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
)
if not defined PYEXE (
  for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
)
if not defined PYEXE (
  echo No se encontro Python. Corre instalar.cmd primero.
  pause
  exit /b 1
)

"%PYEXE%" scripts\espejo.py %*
set "CODIGO=%ERRORLEVEL%"
echo.
if "%CODIGO%"=="0" (
  echo El espejo D:\almacen-ubicaciones quedo igual al proyecto.
) else (
  echo Revisa los mensajes de arriba: hay diferencias o el disco no esta conectado.
)
if not "%ALMACEN_SILENCIOSO%"=="1" pause
exit /b %CODIGO%
