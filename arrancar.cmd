@echo off
setlocal
title Ubicacion de almacen MIGH
cd /d "%~dp0"

set "PY="
for /d %%I in ("%LOCALAPPDATA%\hermes\installs\*") do (
  for /d %%E in ("%%I\environments\*") do (
    if exist "%%E\venv\Scripts\python.exe" if not defined PY set "PY=%%E\venv\Scripts\python.exe"
  )
)
if not defined PY set "PY=python"

echo Usando: %PY%
start "" http://127.0.0.1:8300
"%PY%" app\server.py
pause
