# Guía de instalación con comandos — Control de ubicación de almacén

Todo lo que sigue está probado en Windows 11. Los comandos van en **CMD** o en **PowerShell**;
donde dice **(administrador)** hay que abrir la consola con clic derecho → «Ejecutar como administrador».

## 1. Antes de empezar

| Concepto | Requisito | Nota |
|---|---|---|
| Sistema | Windows 10 u 11 | Probado en Windows 11 25H2 |
| Python | 3.10 o superior | `py --version` debe responder |
| Dependencia | openpyxl 3.1.5 | Única librería externa |
| Carpeta | `C:\almacen-ubicaciones` | Nunca en Program Files |
| Internet | Solo para instalar | Luego funciona sin red |
| Puerto | 8300 libre | Se cambia con `ALMACEN_PUERTO` |
| Permisos | Usuario normal para operar | Administrador solo para firewall y tareas |

Verificar si ya hay Python:

```bat
py --version
python --version
```

Si no responde, instalar (PowerShell, con winget):

```powershell
winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
```

Después de instalar, cerrar y abrir la consola, y comprobar:

```powershell
py --version
```

## 2. Instalación rápida (3 pasos)

CMD:

```bat
cd /d C:\almacen-ubicaciones
instalar.cmd
arrancar.cmd
```

PowerShell:

```powershell
Set-Location C:\almacen-ubicaciones
.\instalar.cmd
.\arrancar.cmd
```

`instalar.cmd` crea el entorno `.venv`, instala openpyxl, prepara el inventario y verifica todo.
Al final abre el navegador en `http://127.0.0.1:8300`.

## 3. Instalación paso a paso (si algo falla)

### 3.1 Copiar la carpeta

CMD:

```bat
xcopy "D:\paquete\almacen-ubicaciones" "C:\almacen-ubicaciones" /E /I /H /Y
cd /d C:\almacen-ubicaciones
```

PowerShell:

```powershell
Copy-Item -Path "D:\paquete\almacen-ubicaciones" -Destination "C:\almacen-ubicaciones" -Recurse -Force
Set-Location C:\almacen-ubicaciones
```

### 3.2 Crear el entorno virtual

CMD:

```bat
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
```

PowerShell (la política de ejecución se cambia solo para esta sesión):

```powershell
py -3 -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Comprobar la instalación de la librería:

```bat
python -c "import openpyxl, sys; print('openpyxl', openpyxl.__version__, 'python', sys.version.split()[0])"
```

### 3.3 Preparar el inventario

El sistema maneja tres inventarios (ALM-PNO-002): materias primas, material de empaque y producto
terminado. Cada uno es un archivo en `data\` y se prepara igual:

```bat
python scripts\preparar_inventario.py --tipo mp  --origen "%USERPROFILE%\Desktop\INVENTARIO MP.xlsx"
python scripts\preparar_inventario.py --tipo mep --origen "D:\inventarios\INVENTARIO MEP.xlsx"
python scripts\preparar_inventario.py --tipo pt  --origen "D:\inventarios\INVENTARIO PT.xlsx"
```

Si todavía no hay archivo de empaque o de producto terminado, se crea la plantilla vacía:

```bat
python scripts\plantilla_inventario.py --tipo mep --vacia
python scripts\plantilla_inventario.py --tipo pt --vacia
```

PowerShell es idéntico (solo cambian las comillas si la ruta tiene espacios):

```powershell
python scripts\preparar_inventario.py --tipo mep --origen "D:\inventarios\INVENTARIO MEP.xlsx"
python scripts\plantilla_inventario.py --tipo pt --vacia
```

El script agrega la columna **UBICACIÓN**, no toca las demás columnas y reporta cuántos lotes quedaron.
Reconoce los encabezados por nombre (`Descripción`/`Producto`/`Material`, `Cantidad`/`Conteo`/`Stock`,
`Lote`, `Caducidad`, `Unidad`/`UM`), así que una tabla del área entra sin reformatear. Con
`--reimportar` toma el archivo nuevo y conserva las ubicaciones ya capturadas.

| Inventario | Archivo | Reporte que exige el procedimiento |
|---|---|---|
| MP — materias primas | `data\Inventario_MP.xlsx` | mensual (§V.2.1.8) |
| MEP — material de empaque | `data\Inventario_MEP.xlsx` | mensual (§V.2.1.8) |
| PT — producto terminado | `data\Inventario_PT.xlsx` | semanal (§V.2.3.8) |

### 3.4 Sin internet ni pip (carpeta vendor)

Si el paquete trae `vendor\`:

```bat
set PYTHONPATH=%CD%\vendor
python scripts\verificar_instalacion.py
```

PowerShell:

```powershell
$env:PYTHONPATH = "$PWD\vendor"
python scripts\verificar_instalacion.py
```

El servidor agrega `vendor\` solo, así que también puede arrancar así:

```bat
py app\server.py
```

## 4. Usuarios y contraseñas

La primera vez que arranca, si no hay usuarios, se crea el administrador y la contraseña se muestra
en la consola (también queda en `data\CREDENCIALES_INICIALES.txt`). Cámbiala al entrar y borra el archivo.

Crear usuarios desde la consola (CMD o PowerShell es igual):

```bat
python scripts\crear_usuario.py --nombre "Esteban Soto" --rol almacen
python scripts\crear_usuario.py --nombre "Galilea Medina" --rol almacen
python scripts\crear_usuario.py --nombre "Auditor Calidad" --rol consulta
python scripts\crear_usuario.py --nombre "Maria Elena Becerra" --rol admin
python scripts\crear_usuario.py --listar
```

Roles:

| Rol | Puede |
|---|---|
| `admin` | Todo: movimientos, deshacer, borrar movimientos, usuarios y contraseñas |
| `almacen` | Registrar movimientos y deshacer los propios. No gestiona usuarios |
| `consulta` | Solo ver y descargar reportes (auditoría, jefatura) |

Contraseña olvidada (genera una nueva y obliga a cambiarla al entrar):

```bat
python scripts\crear_usuario.py --reset "Esteban Soto"
```

Reglas que aplica el sistema: mínimo 8 caracteres, nombre sin números, y no se puede desactivar ni
degradar al último administrador activo.

## 5. Abrir el sistema

En la misma PC:

```bat
arrancar.cmd
```

o bien:

```bat
.venv\Scripts\python.exe app\server.py
```

Direcciones: `http://127.0.0.1:8300` (esta PC). Para cambiar el puerto:

```bat
set ALMACEN_PUERTO=8400
.venv\Scripts\python.exe app\server.py
```

## 6. Varias PC y teléfonos en la red

**Una sola PC guarda los datos.** En la PC del almacén:

```bat
servidor_lan.cmd
```

Muestra dos direcciones; las demás PC y los teléfonos entran a `http://IP:8300`.

Abrir el puerto en el firewall **(administrador)**.

CMD:

```bat
netsh advfirewall firewall add rule name="Almacen MIGH 8300" dir=in action=allow protocol=TCP localport=8300 profile=private
```

PowerShell:

```powershell
New-NetFirewallRule -DisplayName "Almacen MIGH 8300" -Direction Inbound -Protocol TCP -LocalPort 8300 -Action Allow -Profile Private
```

Ver la IP de la PC:

```bat
ipconfig | findstr /i "IPv4"
```

PowerShell:

```powershell
Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -notlike "127.*"} | Select-Object IPAddress, InterfaceAlias
```

Probar desde otra PC o teléfono:

```powershell
Test-NetConnection -ComputerName 192.168.1.50 -Port 8300
```

No exponer el sistema a internet: es de red interna y no tiene certificado ni segundo factor.

## 7. Teléfono y tablet

1. Conectar el teléfono a la misma red Wi-Fi de la planta (no la de invitados: aísla clientes).
2. Abrir en el navegador `http://IP-DE-LA-PC:8300`.
3. Entrar con usuario y contraseña.
4. Agregar a la pantalla de inicio para que abra como aplicación:
   - **iPhone/iPad (Safari):** botón Compartir → «Añadir a pantalla de inicio».
   - **Android (Chrome):** menú ⋮ → «Añadir a pantalla principal».

La interfaz cambia sola: en teléfono cada lote es una tarjeta con botones grandes; en tablet y
computadora es una tabla con encabezados. No hay que instalar nada en el teléfono.

Atajos útiles: la cantidad se propone completa, y si escribes menos se registra una **parcialidad**
(se mueven los kilos indicados y el resto se queda en su rack).

## 8. Arranque automático al iniciar sesión

CMD **(administrador)**:

```bat
schtasks /Create /F /TN "Almacen MIGH" /TR "cmd /c cd /d C:\almacen-ubicaciones && .venv\Scripts\pythonw.exe app\server.py" /SC ONLOGON /RL LIMITED
schtasks /Query /TN "Almacen MIGH" /FO LIST
```

PowerShell **(administrador)**:

```powershell
$accion = New-ScheduledTaskAction -Execute "C:\almacen-ubicaciones\.venv\Scripts\pythonw.exe" -Argument "app\server.py" -WorkingDirectory "C:\almacen-ubicaciones"
$disparo = New-ScheduledTaskTrigger -AtLogOn
$ajustes = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries
Register-ScheduledTask -TaskName "Almacen MIGH" -Action $accion -Trigger $disparo -Settings $ajustes -Description "Control de ubicacion de almacen"
Get-ScheduledTask -TaskName "Almacen MIGH"
```

Quitar la tarea:

```bat
schtasks /Delete /F /TN "Almacen MIGH"
```

PowerShell:

```powershell
Unregister-ScheduledTask -TaskName "Almacen MIGH" -Confirm:$false
```

## 9. Respaldo automático

Manual, con copia a una USB:

```bat
respaldo.cmd --copia-a "E:\respaldo_almacen"
```

Programado todos los días a las 20:30 **(administrador)**:

```bat
schtasks /Create /F /TN "Almacen MIGH respaldo" /TR "cmd /c cd /d C:\almacen-ubicaciones && .venv\Scripts\python.exe scripts\respaldo.py --copia-a E:\respaldo_almacen" /SC DAILY /ST 20:30 /RL LIMITED
```

PowerShell:

```powershell
$accion = New-ScheduledTaskAction -Execute "C:\almacen-ubicaciones\.venv\Scripts\python.exe" -Argument "scripts\respaldo.py --copia-a E:\respaldo_almacen" -WorkingDirectory "C:\almacen-ubicaciones"
$disparo = New-ScheduledTaskTrigger -Daily -At 8:30PM
Register-ScheduledTask -TaskName "Almacen MIGH respaldo" -Action $accion -Trigger $disparo -Description "Respaldo diario del inventario"
```

Cada respaldo genera un `.json` con el SHA-256 de los archivos: sirve para probar que la copia es idéntica.

## 10. Verificación y diagnóstico

```bat
python scripts\verificar_instalacion.py
```

Debe reportar `[OK]` en: Python, openpyxl, estructura, inventario, permisos de escritura, bitácora y puerto.

Revisar quién tiene acceso y cuándo entró:

```bat
python scripts\crear_usuario.py --listar
```

Probar la API sin navegador (PowerShell):

```powershell
$c = Invoke-WebRequest -Uri "http://127.0.0.1:8300/api/login" -Method POST -ContentType "application/json" -Body '{"usuario":"admin","password":"TU_CLAVE"}' -SessionVariable sesion
Invoke-RestMethod -Uri "http://127.0.0.1:8300/api/estado" -WebSession $sesion
```

## 11. Actualizar, mover y desinstalar

Actualizar:

```bat
respaldo.cmd
```

Luego copiar la carpeta nueva encima **sin borrar** `data\` ni `datos_iniciales\`, y repetir `instalar.cmd`.
La bitácora se migra sola (columnas nuevas) y conserva todos los movimientos.

Mover los datos a otra PC: copiar `data\Inventario_MP.xlsx` y `data\bitacora.db` al `data\` de la PC nueva.
No copiar `.venv\` (es de la otra máquina) ni `data\salidas\` (se regenera).

Desinstalar:

```bat
schtasks /Delete /F /TN "Almacen MIGH"
schtasks /Delete /F /TN "Almacen MIGH respaldo"
netsh advfirewall firewall delete rule name="Almacen MIGH 8300"
rd /s /q C:\almacen-ubicaciones
```

Antes de borrar, copiar `data\Inventario_MP.xlsx` y `data\bitacora.db`: son el registro de calidad.

## 12. Seguridad y buenas prácticas

| Tema | Regla |
|---|---|
| Contraseñas | Mínimo 8 caracteres; se guardan con PBKDF2-SHA256 y sal por usuario, nunca en texto |
| Sesión | Cookie HttpOnly de 12 horas; al cambiar contraseña se cierran todas las sesiones |
| Roles | Almacén captura, consulta solo lee, administrador gestiona todo |
| Deshacer | Cualquiera revierte sus propios movimientos; el administrador, cualquiera |
| Borrar | Solo el administrador, con motivo; queda registro en la tabla de auditoría |
| Bitácora | No se edita: cada movimiento conserva usuario, fecha, origen y destino |
| Red | Solo LAN de la planta; nunca abrir el puerto a internet |
| Respaldos | Diario automático + copia externa semanal |

## 13. Problemas frecuentes

| Síntoma | Causa | Solución |
|---|---|---|
| `No se encontro Python` | Sin PATH | Reinstalar marcando *Add python.exe to PATH* |
| `No module named openpyxl` | Sin dependencia | `pip install -r requirements.txt` o usar `vendor\` |
| `No se puede cargar el archivo ...Activate.ps1` | Política de PowerShell | `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` |
| `Acceso denegado` al crear tarea o regla | Consola sin privilegios | Abrir como administrador |
| `puerto 8300 ya en uso` | Ya hay una instancia | Usar la que corre o `set ALMACEN_PUERTO=8400` |
| El teléfono no abre la página | Wi-Fi de invitados o firewall | Red de planta + regla de firewall |
| `el archivo está abierto en Excel` | Bloqueo del `.xlsx` | Cerrar Excel y reintentar |
| Se olvidó la contraseña de admin | — | `python scripts\crear_usuario.py --reset "nombre"` |
| El navegador pide la contraseña cada rato | Sesión de 12 h | Es lo esperado; volver a entrar |

## 14. Lo que el sistema no hace

- No se conecta al ERP (el código ERP es un campo de texto para conciliar a mano).
- No borra ni edita movimientos sin dejar rastro (solo el administrador puede purgar, con auditoría).
- No sustituye la firma del reporte: el F01 se imprime y se firma.
- No funciona fuera de la red local ni con certificados públicos.
