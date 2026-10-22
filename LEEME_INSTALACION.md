# Instalar en otra PC — Control de ubicación de almacén

## Lo que necesitas

| | |
|---|---|
| Sistema | Windows 10 u 11 (probado en Windows 11) |
| Python | 3.10 o superior, con `python.exe` en el PATH |
| Espacio | ~40 MB |
| Internet | **solo** para instalar dependencias la primera vez (o usa el paquete con `vendor\`) |
| Permisos | usuario normal, siempre que la carpeta NO sea `C:\Program Files` |

> **Una sola PC guarda los datos.** No instales el sistema en tres PC con tres copias del inventario: tendrías
> tres verdades. Instálalo en la PC del almacén y las demás entran por red (`servidor_lan.cmd`), o comparten
> la misma carpeta del servidor. Ver "Varias PC" más abajo.

## Instalación (5 minutos)

1. Copia la carpeta `almacen-ubicaciones` a `C:\almacen-ubicaciones` (sin espacios en la ruta; nunca en
   `Program Files` ni en el Escritorio de un usuario que cambie).
2. Pon el Excel del área en `datos_iniciales\Inventario_MP.xlsx` (créala si no existe).
   Si el archivo está en otro lado, no importa: lo indicas en el paso 3.
3. Doble clic en **`instalar.cmd`**. Hace todo esto:
   - crea el entorno `.venv\`;
   - instala `openpyxl` (o usa `vendor\` si no hay internet);
   - copia el Excel y le agrega la columna `UBICACIÓN`;
   - corre la verificación y te dice si algo falta.
4. Doble clic en **`arrancar.cmd`** → abre `http://127.0.0.1:8300`.
5. Escribe tu nombre la primera vez. Queda en la bitácora de cada movimiento.

Si la PC no tiene Python: instálalo desde <https://www.python.org/downloads/windows/> marcando
**“Add python.exe to PATH”**, y vuelve a correr `instalar.cmd`.

## Sin internet ni pip (`vendor\`)

El ZIP generado con `--vendor` ya trae `openpyxl` y `et_xmlfile` dentro de `vendor\`. En ese caso:

- `instalar.cmd` detecta la carpeta y no intenta descargar nada;
- el sistema arranca igual, porque `app\server.py` agrega `vendor\` al `sys.path`.

Solo necesitas Python instalado (aunque sea sin pip).

## Archivos que importan

| Archivo | Qué es |
|---|---|
| `data\Inventario_MP.xlsx` | El inventario del área **con** la columna UBICACIÓN. Es la fuente de verdad. |
| `data\bitacora.db` | Movimientos (SQLite). Es el rastro de auditoría. |
| `data\salidas\` | Los Excel que genera el sistema (F01, tabla macro, bitácora). |
| `data\respaldos\` | Respaldos con manifiesto SHA-256. |
| `config` | No hay: el puerto y las rutas se cambian en los `.cmd` (variables `ALMACEN_PUERTO`, `ALMACEN_HOST`, `ALMACEN_SALIDAS`). |

## Operación diaria

| Necesito… | Hago… |
|---|---|
| Mover un lote | botón **⇄ mover** en el renglón (o **⇄ Registrar movimiento** arriba) |
| Asignar la primera ubicación | escribir el código en **UBICACIÓN ACTUAL** + Enter |
| Reporte del periodo | botón **Reporte F01** → imprimir y firmar |
| Ver qué se vence | botón **Tabla macro + FEFO** → hoja *Caducidades* |
| Respaldar | `respaldo.cmd` (o `respaldo.cmd --copia-a "E:\respaldo_almacen"`) |
| Que arranque solo | `arranque_automatico.cmd` (tarea al iniciar sesión) |

## Varias PC

**Opción recomendada — una PC servidor, las demás navegador:**

1. En la PC del almacén: `servidor_lan.cmd`. Muestra dos direcciones: `127.0.0.1` (esa PC) y `192.168.x.x`
   (las demás).
2. En las otras PC: abrir `http://192.168.x.x:8300`. No se instala nada.
3. Si no abre, falta el puerto en el firewall. Una vez, como administrador:

```bat
netsh advfirewall firewall add rule name="Almacen MIGH 8300" dir=in action=allow protocol=TCP localport=8300
```

**Opción sin red — copia manual:** si dos PC tienen que trabajar sin red, cada una lleva su copia y se
concilian con el F01 impreso. No es lo recomendable: dos bitácoras no se pueden fusionar.

**Riesgo del modo red:** no hay contraseña, solo el nombre que cada quien escribe. Es adecuado para la LAN
de la planta; no lo expongas a internet ni lo dejes en una red de invitados.

## Actualizar a una versión nueva

1. `respaldo.cmd` (obligatorio).
2. Copia la carpeta nueva encima, **sin borrar** `data\` ni `datos_iniciales\`.
3. `instalar.cmd` otra vez: conserva el inventario, la bitácora y las ubicaciones ya capturadas.
4. `arrancar.cmd` y verifica que el tablero muestre los mismos lotes.

La bitácora se migra sola: si la base es anterior, se le agregan las columnas `motivo` y `ref_doc`.

## Mover los datos a otra PC

Copia estos dos archivos al `data\` de la PC nueva **antes** de arrancar:

```
data\Inventario_MP.xlsx
data\bitacora.db
```

No copies `data\salidas\` (se regenera) ni `.venv\` (es de la otra máquina; se crea con `instalar.cmd`).

## Si algo falla

| Síntoma | Causa / solución |
|---|---|
| `No se encontro Python` | instalar Python con “Add python.exe to PATH” y reabrir la consola |
| `No module named openpyxl` | correr `instalar.cmd`; si no hay internet, copiar `vendor\` |
| `el archivo está abierto en Excel` | cerrar el Excel del inventario y reintentar |
| `puerto 8300 ya en uso` | el sistema ya está corriendo, o cambiar `ALMACEN_PUERTO` en `arrancar.cmd` |
| La página no carga | verificar que la ventana negra del servidor siga abierta |
| Desde otra PC no abre | firewall (regla de arriba) o Wi-Fi de invitados con aislamiento de clientes |

Diagnóstico rápido en cualquier momento:

```bat
python scripts\verificar_instalacion.py
```

## Qué NO trae

- No se conecta al ERP (el código ERP es un campo de texto para conciliar a mano).
- No tiene usuarios ni contraseñas.
- No manda nada a internet: todo vive en la PC.
- No sustituye la firma del reporte: el F01 se imprime y se firma.
