@echo off
setlocal
title DEMO - Ubicacion de almacen MIGH
cd /d "%~dp0"

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY for /d %%I in ("%LOCALAPPDATA%\hermes\installs\*") do (
  for /d %%E in ("%%I\environments\*") do (
    if exist "%%E\venv\Scripts\python.exe" if not defined PY set "PY=%%E\venv\Scripts\python.exe"
  )
)
if not defined PY set "PY=python"

if not exist "data\demo\Inventario_MP_demo.xlsx" (
  echo Generando datos de demostracion...
  "%PY%" scripts\sembrar_demo.py
)
if not exist "data\demo\Inventario_MEP.xlsx" (
  echo Generando empaque y producto terminado...
  "%PY%" scripts\sembrar_demo_mep_pt.py
)

echo.
echo   MODO DEMO: puerto 8301. Los datos reales no se tocan.
echo   Usuarios: admin / Demo.Admin.2026   ·   Esteban / Demo.Almacen.2026
echo.
start "" http://127.0.0.1:8301
if not defined ALMACEN_HOST set "ALMACEN_HOST=0.0.0.0"
set ALMACEN_PUERTO=8301
set ALMACEN_SALIDAS=%~dp0data\demo\salidas
"%PY%" app\server.py data\demo\Inventario_MP_demo.xlsx
pause
