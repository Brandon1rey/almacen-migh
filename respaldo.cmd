@echo off
setlocal
title Respaldo - Almacen MIGH
cd /d "%~dp0"

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY for /d %%I in ("%LOCALAPPDATA%\hermes\installs\*") do (
  for /d %%E in ("%%I\environments\*") do (
    if exist "%%E\venv\Scripts\python.exe" if not defined PY set "PY=%%E\venv\Scripts\python.exe"
  )
)
if not defined PY where python >nul 2>nul && set "PY=python"

"%PY%" scripts\respaldo.py %*
echo.
if "%~1"=="" echo Tip: para copiar a una USB usa  respaldo.cmd --copia-a "E:\respaldo_almacen"
echo.
pause
