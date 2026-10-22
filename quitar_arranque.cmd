@echo off
setlocal
title Quitar arranque automatico - Almacen MIGH
schtasks /Delete /F /TN "Almacen MIGH"
if errorlevel 1 (
  echo La tarea no existia o no se pudo borrar.
) else (
  echo Tarea eliminada. El sistema ya no arrancara solo.
)
echo.
pause
