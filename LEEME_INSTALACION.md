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
> tres verdades. Instálalo en la PC del almacén y las demás entran por red (`servidor_lan.cmd`) o desde el
> teléfono. Ver "Varias PC" y "Teléfono y tablet".

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
5. Al primer arranque se crea el administrador y la consola muestra su contraseña (también queda en
   `data\CREDENCIALES_INICIALES.txt`). Entra con ella, cámbiala y borra ese archivo.
6. Crea los usuarios del almacén (ver abajo). **Cada movimiento queda firmado con el usuario que lo hizo.**

## Usuarios, contraseñas y roles

```bat
python scripts\crear_usuario.py --nombre "Esteban Soto" --rol almacen
python scripts\crear_usuario.py --nombre "Galilea Medina" --rol almacen
python scripts\crear_usuario.py --nombre "Auditor Calidad" --rol consulta
python scripts\crear_usuario.py --listar
```

| Rol | Puede |
|---|---|
| `admin` | Todo: movimientos, deshacer, borrar, usuarios y contraseñas |
| `almacen` | Registrar movimientos y deshacer los propios |
| `consulta` | Solo ver y descargar reportes (auditoría, jefatura) |

- Contraseña olvidada: `python scripts\crear_usuario.py --reset "Esteban Soto"` (genera una nueva y obliga a cambiarla).
- También se administran desde la pestaña **Usuarios** de la interfaz (solo administrador).
- Las contraseñas se guardan con PBKDF2-SHA256 y sal por usuario; nunca en texto plano.

## Sin internet ni pip (`vendor\`)

El ZIP generado con `--vendor` ya trae `openpyxl` y `et_xmlfile` dentro de `vendor\`. En ese caso:

- `instalar.cmd` detecta la carpeta y no intenta descargar nada;
- el sistema arranca igual, porque `app\server.py` agrega `vendor\` al `sys.path`.

Solo necesitas Python instalado (aunque sea sin pip).

## Archivos que importan

| Archivo | Qué es |
|---|---|
| `data\Inventario_MP.xlsx` | El inventario del área **con** la columna UBICACIÓN. Es la fuente de verdad. |
| `data\bitacora.db` | Movimientos, usuarios y auditoría (SQLite). Es el rastro de auditoría. |
| `data\salidas\` | Los Excel que genera el sistema (F01, tabla macro, bitácora). |
| `data\respaldos\` | Respaldos con manifiesto SHA-256. |
| `config` | No hay: el puerto y las rutas se cambian en los `.cmd` (`ALMACEN_PUERTO`, `ALMACEN_HOST`, `ALMACEN_SALIDAS`). |

## Operación diaria

| Necesito… | Hago… |
|---|---|
| Mover un lote completo | **⇄ Registrar movimiento** (o **⇄ Mover** en el renglón); la cantidad propuesta es todo |
| Mover **solo unos kilos** | en el modal, borrar la cantidad y escribir los kilos: el resto se queda donde estaba |
| Asignar la primera ubicación | escribir el código en la columna **UBICACIÓN** + Enter |
| Corregir un movimiento | **Movimientos → Deshacer** (el último del lote) |
| Borrar una captura de prueba | **Movimientos → Borrar** (solo administrador, con motivo) |
| Reporte del periodo | **Reporte F01** → imprimir y firmar |
| Ver qué se vence | **Tabla macro + FEFO** → hoja *Caducidades* |
| Respaldar | `respaldo.cmd` (o `respaldo.cmd --copia-a "E:\respaldo_almacen"`) |
| Que arranque solo | `arranque_automatico.cmd` (tarea al iniciar sesión) |

**Parcialidades:** si de una tarima de 1,458 kg se surten 5 kg, el sistema deja 1,453 kg en el rack
original y crea un renglón nuevo de 5 kg en el destino, con el mismo lote. En la bitácora queda como
`PARCIAL` con la cantidad exacta, el usuario y el motivo (surtido, reacomodo, devolución, cuarentena…).

## Teléfono y tablet

Misma red Wi-Fi que la PC del almacén (no la de invitados) y abrir `http://IP-DE-LA-PC:8300`.
La interfaz cambia sola: tarjetas con botones grandes en teléfono, tabla en tablet y computadora.
Se puede agregar a la pantalla de inicio para que abra como aplicación. No se instala nada en el teléfono.

## Varias PC

**Opción recomendada — una PC servidor, las demás navegador:**

1. En la PC del almacén: `servidor_lan.cmd`. Muestra dos direcciones: `127.0.0.1` (esa PC) y `192.168.x.x`
   (las demás).
2. En las otras PC y teléfonos: abrir `http://192.168.x.x:8300`. No se instala nada.
3. Si no abre, falta el puerto en el firewall. Una vez, como administrador:

```bat
netsh advfirewall firewall add rule name="Almacen MIGH 8300" dir=in action=allow protocol=TCP localport=8300 profile=private
```

**Opción sin red — copia manual:** si dos PC tienen que trabajar sin red, cada una lleva su copia y se
concilian con el F01 impreso. No es lo recomendable: dos bitácoras no se pueden fusionar.

## Actualizar a una versión nueva

1. `respaldo.cmd` (obligatorio).
2. Copia la carpeta nueva encima, **sin borrar** `data\` ni `datos_iniciales\`.
3. `instalar.cmd` otra vez: conserva el inventario, la bitácora, los usuarios y las ubicaciones ya capturadas.
4. `arrancar.cmd` y verifica que el tablero muestre los mismos lotes.

La bitácora se migra sola: si la base es anterior, se le agregan las columnas que falten
(`motivo`, `ref_doc`, `usuario_id`, `cantidad`, `fila_destino`, `revertido`) sin perder movimientos.

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
| Se olvidó una contraseña | `python scripts\crear_usuario.py --reset "nombre del usuario"` |

Diagnóstico rápido en cualquier momento:

```bat
python scripts\verificar_instalacion.py
```

## Guía completa de comandos

`docs\Guia_instalacion_comandos.pdf` (6 páginas) trae los comandos exactos para CMD y PowerShell:
instalar Python con winget, crear el entorno, firewall, tarea de arranque, respaldo programado,
diagnóstico y solución de problemas.

## Qué NO trae

- No se conecta al ERP (el código ERP es un campo de texto para conciliar a mano).
- No funciona fuera de la red local ni con certificados públicos: es de LAN.
- No deja borrar sin rastro: purgar un movimiento queda registrado en la tabla de auditoría.
- No sustituye la firma del reporte: el F01 se imprime y se firma.
