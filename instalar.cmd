@echo off
setlocal
title Instalar - Control de ubicacion de almacen
cd /d "%~dp0"

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY for /d %%I in ("%LOCALAPPDATA%\hermes\installs\*") do (
  for /d %%E in ("%%I\environments\*") do (
    if exist "%%E\venv\Scripts\python.exe" if not defined PY set "PY=%%E\venv\Scripts\python.exe"
  )
)
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo.
  echo   No se encontro Python en esta PC.
  echo.
  echo   1. Descarga Python 3.12 o superior de https://www.python.org/downloads/windows/
  echo   2. Al instalar, marca "Add python.exe to PATH"
  echo   3. Vuelve a correr este archivo
  echo.
  echo   Si la PC no tiene internet: copia la carpeta vendor\ de una PC que ya
  echo   funcione junto a este archivo y vuelve a intentar.
  echo.
  pause
  exit /b 1
)

echo Usando: %PY%
"%PY%" scripts\instalar.py %*
echo.
pause
