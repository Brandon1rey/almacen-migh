@echo off
setlocal
title Copia de la base del almacen a la USB
cd /d "%~dp0"

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY for /d %%I in ("%LOCALAPPDATA%\hermes\installs\*\") do (
  for /d %%E in ("%%Ienvironments\*\") do (
    if exist "%%E\venv\Scripts\python.exe" if not defined PY set "PY=%%E\venv\Scripts\python.exe"
  )
)
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo No se encontro Python. Corre instalar.cmd primero.
  pause
  exit /b 1
)

"%PY%" scripts\sincronizar_usb.py %*
set "CODIGO=%ERRORLEVEL%"
if not "%CODIGO%"=="0" (
  echo.
  echo La copia no se completo. Revisa que la USB este conectada.
  pause
)
exit /b %CODIGO%
