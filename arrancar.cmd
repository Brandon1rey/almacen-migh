@echo off
setlocal
title Ubicacion de almacen MIGH
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

echo Usando: %PY%
if not defined ALMACEN_HOST set "ALMACEN_HOST=0.0.0.0"
start "" http://127.0.0.1:8300
"%PY%" app\server.py
pause
