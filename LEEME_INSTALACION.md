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

## Instalación en una PC nueva con un solo comando

1. Copiar la carpeta `almacen-ubicaciones` a `C:\almacen-ubicaciones` (descomprimir el ZIP de `dist\`;
   nunca en `Program Files` ni en el Escritorio).
2. Poner el Excel del área en `datos_iniciales\Inventario_MP.xlsx` (si no está, el instalador avisa y
   sigue: la app arranca con lo que haya).
3. Doble clic en **`instalar_todo.cmd`** y aceptar el aviso de administrador.

Ese único archivo hace todo:

| Paso | Qué hace |
|---|---|
| 1 | Busca Python; si no está, lo instala con `winget` (Python 3.12) |
| 2 | Crea `.venv` e instala `openpyxl` (o usa `vendor\` si no hay internet) |
| 3 | Copia el inventario del área y crea las plantillas de MEP y PT si faltan |
| 4 | Deja el **arranque automático** en la carpeta de Inicio del usuario |
| 5 | Abre el puerto 8300 en el firewall (pide permiso una sola vez) |
| 6 | Verifica todo y muestra la dirección para entrar |

4. Cerrar sesión y volver a entrar (o reiniciar): el servidor arranca solo, sin ventana.
   Se comprueba en `data\logs\servidor.log` y en <http://127.0.0.1:8300>.
5. Primer ingreso: el usuario administrador y su contraseña inicial están en
   `data\CREDENCIALES_INICIALES.txt` (se borra después de cambiarla).

Para dejarlo sin arranque automático: `quitar_arranque.cmd`.

Si se prefiere una **tarea de Windows** en vez de la carpeta de Inicio (misma cosa, pero aparece en el
Programador de tareas):

```bat
python scripts\instalar.py --tarea
```

## Instalación paso a paso (cuando algo falla)

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

## Los tres inventarios: MP, MEP y PT

ALM-PNO-002 cubre materias primas, material de empaque y producto terminado, y pide el F01 «de todos
los materiales almacenados». El sistema los maneja como tres inventarios separados que comparten
catálogo de ubicaciones, usuarios y bitácora:

| Inventario | Archivo en `data\` | Reporte |
|---|---|---|
| MP — Materias primas | `Inventario_MP.xlsx` | mensual (§V.2.1.8) |
| MEP — Material de empaque | `Inventario_MEP.xlsx` | mensual (§V.2.1.8) |
| PT — Producto terminado | `Inventario_PT.xlsx` | semanal (§V.2.3.8) |

En la pantalla hay un selector **MP · MEP · PT** y una pestaña **Tablero** con los tres. El botón
**Reporte F01** genera un archivo con portada, totales y una hoja por inventario.

Cargar la tabla de empaque o de producto terminado (una sola vez, o cada vez que el área la actualice):

```bat
python scripts\plantilla_inventario.py --tipo mep --vacia
python scripts\preparar_inventario.py --tipo mep --origen "C:\ruta\INVENTARIO MEP.xlsx"
python scripts\preparar_inventario.py --tipo pt  --origen "C:\ruta\INVENTARIO PT.xlsx"
python scripts\preparar_inventario.py --tipo mep --reimportar
```

Se reconocen los encabezados por nombre, no por posición: `Descripción`/`Producto`/`Material`,
`Cantidad`/`Conteo`/`Stock`, `Lote`, `Caducidad`, `Unidad`/`UM`, `Presentación`, `Piezas`. Si la tabla
no trae `UBICACIÓN`, se agrega. Si un archivo no existe, el sistema arranca igual y ese inventario
aparece como «vacío» hasta que se cargue.

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
- También se administran desde la interfaz: el administrador ve el botón **«+ Usuario»** arriba (y la pestaña
  **Usuarios**), donde da de alta nombre, rol y contraseña —con botón **Generar** para una clave segura— y
  al terminar la pantalla muestra la contraseña una sola vez para entregarla en mano.
  El resto de los roles no ve ese botón ni puede llamar la API de usuarios (responde 403).
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
| `data\Inventario_MEP.xlsx`, `data\Inventario_PT.xlsx` | Empaque y producto terminado, con la misma forma. |
| `data\catalogo_insumos.json` | Catálogo de insumos del ERP (claves y nombres canónicos). |
| `data\bitacora.db` | Movimientos, usuarios y auditoría (SQLite). Es el rastro de auditoría. |
| `data\logs\servidor.log` | Lo que imprimió el servidor al arrancar y cualquier error (sirve con arranque automático). |
| `data\salidas\` | Los Excel que genera el sistema (F01, tabla macro, bitácora). |
| `data\respaldos\` | Respaldos con manifiesto SHA-256. |
| `config` | No hay: el puerto y las rutas se cambian en los `.cmd` (`ALMACEN_PUERTO`, `ALMACEN_HOST`, `ALMACEN_SALIDAS`). |

## La guía que trae la pantalla

No hace falta memorizar nada: la interfaz explica sola.

| Dónde | Qué explica |
|---|---|
| Botón **? Ayuda** (arriba a la derecha, también en la pantalla de acceso) | Cómo buscar, cómo ubicar, mover y sacar a producción, cómo deshacer, cómo ingresar material, atajos de teclado y qué hacer si algo falla. Muestra la versión instalada. |
| **Primeros pasos** (la primera vez que entra cada usuario) | Los 4 pasos del uso normal. Se cierra con *Entendido* y no vuelve a salir en ese navegador. |
| Avisos arriba de la lista | Datos por corregir en el Excel: ubicaciones fuera del catálogo, cantidades ilegibles, renglones sin lote. |
| «¿Quisiste decir…?» | Cuando la búsqueda no encuentra nada o encuentra solo parecidos, ofrece los nombres **tal como están escritos** en el inventario. |
| Mensajes de error | Dicen qué pasó y qué hacer: qué ubicación no existe, cuánto hay de existencia, por qué no cuadra la unidad, por qué no se puede borrar. |


## Operación diaria

| Necesito… | Hago… |
|---|---|
| Mover un lote completo | **⇄ Registrar movimiento** (o **⇄ Mover** en el renglón); la cantidad propuesta es todo |
| Mover **solo unos kilos** | en el modal, borrar la cantidad y escribir los kilos: el resto se queda donde estaba |
| Asignar la primera ubicación | escribir en **UBICACIÓN** y elegir de las sugerencias (↑ ↓ + Enter) |
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

`127.0.0.1` **solo existe dentro de la PC**: en el teléfono esa dirección apunta al propio teléfono.
Para entrar desde un teléfono u otra PC hay que usar la **IP de esta PC en la red** y arrancar en modo red.

1. En la PC del almacén: `servidor_lan.cmd` (queda escuchando en la red, no solo en localhost).
   La ventana negra muestra dos direcciones; la de **«Otras PC»** es la que se teclea en el teléfono,
   por ejemplo `http://192.168.0.9:8300`.
2. Una sola vez, para que el firewall deje pasar el puerto: doble clic en **`abrir_puerto_firewall.cmd`**
   y aceptar el aviso de administrador. (Equivalente en consola elevada:
   `netsh advfirewall firewall add rule name="Almacen MIGH 8300" dir=in action=allow protocol=TCP localport=8300 profile=private`.)
3. Conectar el teléfono a la **misma red Wi-Fi** que la PC (no la de invitados: aísla a los clientes) y abrir
   esa dirección en Chrome o Safari.
4. Agregar a la pantalla de inicio para que abra como aplicación:
   - **iPhone/iPad (Safari):** Compartir → «Añadir a pantalla de inicio».
   - **Android (Chrome):** menú ⋮ → «Añadir a pantalla principal».

Si la IP de la PC cambia (reinicio del módem, otra red), hay que volver a verla:

```powershell
Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -notlike "127.*"} | Select-Object IPAddress, InterfaceAlias
```

La interfaz cambia sola: tarjetas con botones grandes en teléfono, tabla en tablet y computadora.
No se instala nada en el teléfono.

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

## Respaldo

`respaldo.cmd` guarda en `data\respaldos\` (rotación de 60 días) los **tres inventarios** que existan,
el **catálogo de insumos** y la **bitácora** (con la API de respaldo de SQLite, sin detener el servidor),
más un manifiesto JSON con tamaño y SHA-256 de cada archivo. Con
`respaldo.cmd --copia-a "E:\respaldo_almacen"` deja además una segunda copia en otra unidad, que es lo
que protege de un disco dañado. La rotación solo borra archivos con el nombre que pone el propio
respaldo: nunca toca otros documentos que alguien haya dejado en esa carpeta.

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

## Compartir la dirección (detector de red)

Botón **📶 Red** en la barra superior: dice en segundos si el sistema se puede abrir desde otra PC,
tablet o teléfono, y con qué dirección.

| Lo que muestra | Significado |
|---|---|
| `En esta PC` | `http://127.0.0.1:8300` — solo sirve en esa computadora |
| `Para otras PC, tablets y teléfonos` | `http://192.168.x.x:8300` — la que se comparte, con botón **Copiar** |
| Chip verde «Se puede compartir» | el servidor está en modo red y tiene IP |
| Chip ámbar «Solo esta PC» | arrancó en modo local: hay que usar `servidor_lan.cmd` |
| Aviso de firewall | falta la regla del puerto: ejecutar `abrir_puerto_firewall.cmd` |

En la pantalla de **entrada** también aparece la dirección de red («Desde otra PC o teléfono: …»), para
dictarla sin tener que entrar. La dirección puede cambiar si el módem reasigna la IP: el detector siempre
muestra la vigente.

## Borrar renglones de tajo (solo administrador)

Cuando un renglón no debe existir (duplicado de la importación, captura de prueba, material que nunca
llegó), el administrador puede **quitarlo del Excel**, no solo anularlo:

1. En **Inventario**, cada renglón muestra una **casilla** (solo para administradores).
2. Marca uno o varios; el botón de arriba cambia a **«🗑 Borrar 3 renglones»**.
3. Escribe el **motivo** (obligatorio, mínimo 5 caracteres) y confirma con **Borrar de tajo**.

Qué garantiza el sistema:

| Antes de borrar | Después de borrar |
|---|---|
| copia del libro en `data\respaldos\antes_de_borrar_<fecha>_Inventario_MP.xlsx` | el renglón ya no está en el Excel del área |
| — | la bitácora guarda **material, lote, cantidad y ubicación** que tenía, con tu usuario y el motivo (`PURGAR`) |
| — | las filas de los movimientos que quedan **se recorren**, para que el historial siga apuntando al renglón correcto |
| — | queda una entrada de auditoría `BORRAR_RENGLONES` con la ruta del respaldo |

No se puede deshacer desde la pantalla: para recuperar algo, el administrador abre el respaldo. Los roles
`almacen` y `consulta` no ven las casillas y el servidor les responde **403** si lo intentan por la API.

## Actualizar una PC con la USB (versión vigente → PC)

La carpeta de la USB se arma con un comando y lleva todo lo necesario:

```powershell
python scripts\preparar_usb.py --destino "D:\ALMACEN_MIGH_ACTUALIZACION"
```

Queda así:

| En la USB | Para qué |
|---|---|
| `ACTUALIZAR.cmd` | doble clic en la PC que ya tiene el sistema |
| `INSTALAR_EN_PC_NUEVA.cmd` | doble clic en una PC sin el sistema |
| `actualizar.py` | el actualizador (solo biblioteca estándar) |
| `instalacion\*.zip` | la versión vigente |
| `versiones\*.zip` | versiones anteriores, para volver atrás |
| `version.txt` | qué versión trae la USB |

En la PC: conectar la USB → abrir la carpeta → **doble clic en `ACTUALIZAR.cmd`**. El actualizador:

1. compara versiones y **no baja de versión** salvo `--forzar`;
2. respalda código y datos en `data\respaldos\pre_actualizacion_<fecha>.zip` (bitácora con copia
   consistente, los tres Excel y el catálogo);
3. **detiene el servidor**, copia el código, retira los archivos que ya no existen en la versión nueva
   (solo los que instaló un paquete anterior) y **nunca toca `data\`**;
4. verifica que la versión nueva importe y que openpyxl esté;
5. arranca el servidor y comprueba que responda;
6. si algo falla, **restaura el código anterior** desde el respaldo.

Opciones útiles: `--simular` (solo compara), `--instalada "D:\almacen-ubicaciones"` (cuando no está en
la ruta habitual), `--forzar` (volver a una versión anterior), `--sin-servidor` (no lo detiene ni arranca).

La versión vive en `app\version.py` y el ZIP se genera con `python scripts\empaquetar.py --vendor`, que
deja además `VERSION.txt` dentro del paquete.

## Espejo en D:\ (regla del proyecto)

**Todo el trabajo se copia a `D:\almacen-ubicaciones`.** Es el espejo: si la PC se pierde o se trabaja
desde otro lugar, ahí está el proyecto completo.

```bat
espejo_d.cmd --simular     :: qué falta (no escribe)
espejo_d.cmd               :: copia y verifica con SHA-256
```

Qué espeja: `app\`, `scripts\`, `tests\`, `docs\`, los `.cmd`, los `requirements` y `datos_iniciales\`.
Qué **no** toca: `data\` (cada instalación tiene sus datos), `.venv`, `.git`, `dist` y tu
`NOTASBRANDON.txt`. Nunca borra nada en el destino; solo agrega o actualiza.

El comando termina en verde solo si el SHA-256 de cada archivo coincide en las dos copias:

```
Resultado: espejo al día (verificado por SHA-256).
```

## Salida a producción (ALMACÉN → PRODUCCION)

En el modal de movimiento, el botón **→ Producción** llena el destino con `PRODUCCION`. Ese destino no es
un rack: el material **sale del almacén**.

| Caso | Qué pasa en el Excel | Bitácora |
|---|---|---|
| Surten unos kilos (cantidad menor) | la cantidad baja; el lote sigue en su rack con el resto | `SALIDA`, cantidad, rack → PRODUCCION |
| Surten todo («Todo») | cantidad 0, sin ubicación, nota `surtido a producción` | `SALIDA`, cantidad completa |

Reglas: no se puede surtir más de lo que hay; el lote debe tener ubicación asignada (si no, primero se
ubica); se recomienda escribir la **OP en el documento de referencia** para amarrar el surtido a la orden.
*Deshacer* devuelve los kilos al renglón y, si había quedado sin ubicación, le regresa la ubicación.

## Catálogo de insumos del ERP (MP y ME)

`data\catalogo_insumos.json` guarda el catálogo de insumos (clave, descripción, tipo, unidad y
fabricante). Se importa una sola vez:

```powershell
cd C:\Users\asegm\almacen-ubicaciones
.\.venv\Scripts\python.exe scripts\importar_catalogo.py
.\.venv\Scripts\python.exe scripts\importar_catalogo.py --origen "D:\otra\ruta\insumo_CATALOGO.xls"
```

Acepta `.xls` (el formato en que el ERP entrega `insumo.xls`) y `.xlsx`. El `.xls` se abre con
el **Excel de la PC** —no hace falta pip ni internet: si falta pywin32, el importador usa el COM de
Windows por PowerShell— y el catálogo queda en el JSON. La app **no** lee Excel al operar: usa el JSON.

El catálogo real trae 565 insumos (~300 de empaque, ~258 de materias primas, 7 sin tipo en el ERP).
Se importa una sola vez con `scripts\importar_catalogo.py --origen "C:\...\insumo.xls"`.

En la pantalla de **Ingresar material**, al escribir la clave (o el nombre) aparecen las coincidencias
del catálogo **de todos los inventarios**, cada una con su tipo (Empaque, Excipiente, Principio Activo…).
Al elegir una se llenan **todos los datos del insumo**: clave, descripción, unidad de medida, categoría
(que es el tipo del ERP) y fabricante (que va al proveedor). Si el insumo es de otra naturaleza, la
pantalla **cambia sola** al inventario correcto y vuelve a aplicar los datos del ERP.

Si escribes la **clave exacta**, no hay que tocar la lista: se autollena sola en cuanto coincide. Debajo
de la clave aparece la **ficha del insumo** con lo que trae el ERP (clave, descripción, tipo, unidad,
fabricante, código de barra, máximo y mínimo), para capturar sin adivinar. Se busca sin filtro a
propósito: si solo se pudiera buscar dentro del inventario actual, un insumo de empaque capturado desde
materias primas nunca aparecería.

**El catálogo del ERP autocompleta y avisa; nunca bloquea una recepción.** Solo la **ubicación** se
valida contra un catálogo, porque es el layout físico de los racks.

- Si la clave **existe**, sus datos manda: descripción, unidad, categoría y fabricante se toman del ERP
  (si no coincidían, la pantalla avisa «corregido al del catálogo»).
- Si la clave es **de otro inventario** (una bolsa capturada como materia prima), el renglón se guarda
  en el libro que le corresponde y se avisa; no hay que volver a capturarlo.
- Si la clave **no existe** en el catálogo, el insumo **no trae tipo** o se trata de algo que el ERP
  todavía no tiene (producto terminado, material de laboratorio), el alta **se acepta** y se guarda tal
  como se capturó, con un aviso para revisarlo después.
- Se puede capturar **buscando por descripción**: al escribir el nombre se llena sola la clave del ERP
  (si es única) o se elige de la lista de coincidencias. También al revés: la clave llena el nombre.
- Lo que el ERP sí debe traer, cuando la clave existe, es el dato correcto: la app no mezcla dos
  nombres para la misma clave ni cambia la descripción que escribiste si el insumo está en el catálogo.

## Ingresar material nuevo (alta de un lote)

Botón **＋ Ingresar material** en la pantalla de inventario (no lo ve el rol de consulta). Lo primero que
pide es la **naturaleza del material**, porque de ella dependen el Excel destino y las validaciones:

| Naturaleza | Se guarda en | Unidad | Caducidad | Extra |
|---|---|---|---|---|
| **Materia prima** | `Inventario_MP.xlsx` | kg (también lt, m) | **obligatoria** (PEPS/FEFO) | — |
| **Material de empaque** | `Inventario_MEP.xlsx` | pza (también kg, m) | **opcional** | — |
| **Producto terminado** | `Inventario_PT.xlsx` | pza | **obligatoria** | se recomienda presentación y piezas por caja |

Al cambiar la naturaleza se ajustan la unidad, las validaciones y las sugerencias de **ubicación**
(zonas de MP, EM o PT). Si capturas una clave que pertenece a otra naturaleza, el formulario **cambia
solo** a la correcta. Al guardar, si el material quedó en otro inventario, la pantalla **salta a ese
inventario** y busca el lote ingresado.

Además, el sistema avisa (sin bloquear) cuando la unidad no es la habitual para esa naturaleza, cuando un
producto terminado va sin presentación o cuando un empaque va sin caducidad.

| Campo | ¿Obligatorio? | Regla |
|---|---|---|
| Descripción del material | **sí** | mínimo 3 caracteres |
| Lote | **sí** | mínimo 3; sin lote no hay trazabilidad (§V.2.2.2) |
| Cantidad | **sí** | número mayor que cero |
| Unidad | sí | kg, pza, lt o m; se propone la del inventario |
| Proveedor | **sí** | — |
| Caducidad | **sí** | fecha válida y **no vencida** |
| Ubicación | no | si se deja vacía, entra como *pendiente de ubicar* |
| ERP / SKU, categoría, observaciones | no | — |
| Presentación, piezas por caja | no | útil para producto terminado (§V.2.3.1) |

El sistema rechaza (con el motivo en pantalla) un lote repetido del mismo material con existencia, una
cantidad en cero o negativa, una caducidad vencida o ilegible y una ubicación que no existe en el
catálogo de ALM-PNO-002. **No escribe nada hasta que todo valida**: primero revisa, luego guarda el Excel
y la bitácora como una sola operación.

Si el alta se capturó mal, en **Movimientos** se puede *Deshacer*: el renglón queda en cero y marcado
(`alta deshecha #N`), nunca se borra del libro.

## Códigos QR / Barcode (opcional)

El sistema soporta etiquetas físicas QR o códigos de barras 1D como una capa opcional sobre los tres Excel, sin modificar la bitácora ni la operación manual:

- **Columna `QR_ID`:** se agrega automáticamente al guardar el primer lote con código (también reconoce encabezados `BARCODE_ID`, `Código QR`, `Barcode`, `Código de barras`). La identidad interna sigue siendo `(tipo, fila)`.
- **Sin duplicados activos:** un código no puede repetirse en dos tarimas activas (`cantidad > 0`), ni dentro del mismo libro ni entre `MP`, `MEP` y `PT`.
- **Parcialidades:** cuando se mueve una fracción del lote, el código permanece en el renglón origen y el renglón nuevo nace sin `qr_id`.
- **Conservación al reimportar:** `python scripts\preparar_inventario.py --reimportar` conserva los `qr_id` cuando la clave `(descripción, lote)` es única; si una clave tiene varias filas, avisa en consola y no reasigna la etiqueta para evitar duplicados.
- **Cámara y HTTPS:** en `localhost` la cámara abre directo; en teléfonos conectados por `http://IP:8300` el navegador requiere HTTPS para usar la cámara, por lo que si no hay TLS configurado la interfaz ofrece escribir o pegar el código manualmente (salvedad documentada para configuración TLS opcional con `ALMACEN_CERT` y `ALMACEN_KEY`).

## Copia de la base a la USB (para revisar sin tocar producción)

Copia la bitácora y los tres inventarios a `D:\almacen-ubicaciones-datos\` **sin detener el servidor**
y sin modificar el original. Usa la API de respaldo de SQLite, así que si alguien está registrando un
movimiento en ese momento, la copia sale íntegra.

```bat
python scripts\sincronizar_usb.py                  :: copia del momento
python scripts\sincronizar_usb.py --estado         :: solo compara, no copia nada
python scripts\sincronizar_usb.py --historial      :: además guarda una copia con fecha
python scripts\sincronizar_usb.py --resumen        :: lee la copia y dice qué trae
python scripts\sincronizar_usb.py --forzar         :: sobrescribe aunque la USB vaya adelante
sincronizar_usb.cmd --estado                       :: lo mismo, sin depender del PATH
```

Solo copia lo que cambió: si la copia ya tiene los mismos movimientos y los inventarios tienen la misma
fecha y tamaño, no reescribe nada (y lo dice). Si la USB trae **más** movimientos que el servidor —viene
de otra PC—, **no la pisa** y avisa; para eso está `--forzar`.

En la interfaz hay una pestaña **Copia** con el mismo comparador: muestra si está al día, cuántos
movimientos tiene cada lado y el estado de cada archivo, con el botón **Copiar ahora** (y las casillas de
historial y forzar).

Qué queda en la USB:

| Ruta | Qué es |
|---|---|
| `D:\almacen-ubicaciones-datos\ultima\bitacora.db` | Movimientos, usuarios y auditoría hasta ese momento |
| `D:\almacen-ubicaciones-datos\ultima\Inventario_*.xlsx` | MP, MEP y PT con su columna UBICACIÓN |
| `D:\almacen-ubicaciones-datos\ultima\sincronizacion.json` | Fecha, conteos y SHA-256 de cada archivo |
| `D:\almacen-ubicaciones-datos\historial\` | Copias con fecha (30 por omisión; `--mantener 0` las conserva todas) |
| `D:\almacen-ubicaciones-datos\LEEME.txt` | Explica a quien abra la USB qué está viendo |

Opciones útiles: `--destino "E:\otra_carpeta"`, `--sin-inventarios`, `--mantener 12`.
Devuelve código 2 si la USB no está conectada y 3 si la copia no cuadra con el original (no deja una
copia a medias sin avisar).

Copia automática cada 30 minutos, como administrador:

```bat
schtasks /Create /F /TN "Almacen MIGH copia USB" /TR "C:\almacen-ubicaciones\sincronizar_usb.cmd --historial" /SC MINUTE /MO 30 /RL LIMITED
schtasks /Query /TN "Almacen MIGH copia USB" /FO LIST
```

Para revisar la base: abrir `bitacora.db` con **DB Browser for SQLite** (<https://sqlitebrowser.org>) o
`python scripts\sincronizar_usb.py --resumen`. La copia es de lectura: el servidor sigue siendo la única
fuente de verdad.

## Para depurar (cuando algo no sale como se espera)

Todo lo que hace el sistema queda escrito en **`data\logs\depuracion.log`** (rota a 2 MB, guarda 5
archivos). No hay que activar nada:

| Qué escribe | Ejemplo |
|---|---|
| Cada petición a la API, con duración y usuario | `peticion POST /api/mover -> 200 14.2ms usuario=Esteban` |
| Cada acción importante (evento) | `evento=movimiento.registrado accion=PARCIAL fila=7 lote=BL-2601 usuario=Esteban` |
| Errores del servidor, con el traceback completo | `error interno en POST /api/alta` + pila |
| Errores de JavaScript de la pantalla (los reporta el navegador solo) | `navegador nivel=error mensaje=TypeError: …` |
| Arranque y cierre | `evento=servidor.arrancado version=1.8.1 puerto=8300` |

- **Contraseñas y tokens nunca se escriben**: se tachan como `***`.
- Para más detalle (peticiones de archivos estáticos, valores rechazados), arranca con
  `set ALMACEN_LOG=DEBUG` antes de `arrancar.cmd`.
- Estando dentro como **administrador**, abre `http://127.0.0.1:8300/api/diagnostico` para ver de un
  golpe: versión, rutas, espacio en disco, inventarios, catálogo del ERP, últimos eventos y últimos
  errores. Es la forma rápida de revisar una instalación por teléfono.
- `data\logs\servidor.log` sigue guardando lo que imprime el servidor (banner de arranque, avisos).

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
