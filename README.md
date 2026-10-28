# Control de ubicación — almacén de materias primas MIGH

Sistema local (localhost, sin internet, sin licencias) que le agrega **una columna: UBICACIÓN**
al Excel de inventario que ya usa el área, y produce los Excel del reporte con un clic.

## Arrancar

Doble clic en **`arrancar.cmd`** (o desde consola):

```bat
arrancar.cmd
```

Abre el navegador en **http://127.0.0.1:8300** y pide **usuario y contraseña**. Cada movimiento queda
firmado con el usuario que lo hizo; hay tres roles (administrador, almacén, consulta).

## Los tres inventarios (ALM-PNO-002)

El procedimiento cubre materias primas, material de empaque y producto terminado; el sistema también.
Cada uno es **su propio Excel** (el del área, con la misma forma) y comparte catálogo de ubicaciones,
usuarios y bitácora:

| Inventario | Archivo | Zonas | Reporte F01 |
|---|---|---|---|
| Materias primas (MP) | `data\Inventario_MP.xlsx` | MPA-A, MPNA-A/B, pasillo A, jaulas | mensual (§V.2.1.8) |
| Material de empaque (MEP) | `data\Inventario_MEP.xlsx` | MEP-B a MEP-E, jaula de empaque | mensual (§V.2.1.8) |
| Producto terminado (PT) | `data\Inventario_PT.xlsx` | MPT-B, jaula PT, andén | **semanal** (§V.2.3.8) |

- La interfaz trae un selector **MP · MEP · PT** y una pestaña **Tablero** con los tres de un vistazo.
- El botón **Reporte F01** entrega **un solo archivo** con portada, totales y una hoja por inventario:
  es el reporte de «todos los materiales» que pide el procedimiento.
- Los movimientos de cada inventario se guardan con su tipo; deshacer y borrar respetan de dónde vino
  el movimiento aunque los números de fila se repitan.

Para cargar empaque o producto terminado:

```bat
python scripts\plantilla_inventario.py --tipo mep --vacia     :: plantilla con los encabezados
python scripts\preparar_inventario.py --tipo pt --origen "C:\ruta\INVENTARIO PT.xlsx"
python scripts\preparar_inventario.py --tipo mep --reimportar  :: toma el archivo nuevo del área
```

El sistema reconoce los encabezados por nombre (`Descripción`/`Producto`/`Material`, `Cantidad`/`Conteo`/`Stock`,
`Lote`, `Caducidad`, `Unidad`/`UM`), así que una tabla del área entra sin reformatear. Si falta
`UBICACIÓN`, la agrega.

## Qué hace

| Botón | Archivo que genera |
|---|---|
| Inventario + ubicación | El Excel del área tal cual, con la columna UBICACIÓN + hojas **Resumen** y **Sin ubicar** |
| Reporte F01 | F01-ALM-PNO-002 **versión 2** (columnas del formato vigente: Ubicación, Material, Cantidad, Lote) agrupado por zona, con bloque de firmas |
| F01 con caducidad (v3) | La propuesta de versión siguiente: agrega Caducidad y Días |
| Tabla macro + FEFO | Matriz material × zona en A3 + hoja de caducidades ordenada por lo que vence primero |
| Bitácora | Todos los movimientos con usuario, fecha-hora, origen y destino |

Todo se guarda en `data\salidas\` con marca de fecha y hora, y también se descarga desde el navegador.

## Salidas a producción y catálogo de insumos

- **ALMACÉN → PRODUCCION**: en el modal de movimiento, el botón **→ Producción** registra el surtido
  (parcial en kg o total). El material sale del inventario, no cambia de rack; la bitácora guarda la OP
  del documento de referencia y *Deshacer* lo devuelve a su ubicación.
- **Borrar de tajo (solo admin):** casilla por renglón + **Borrar N renglones** con motivo obligatorio.
  Respalda el Excel en `data\respaldos\`, deja en la bitácora qué había en cada renglón borrado y recorre
  las filas de los movimientos restantes. `almacen` y `consulta` reciben 403.
- **Catálogo del ERP** (`scripts\importar_catalogo.py --origen "C:\...\insumo.xls"` →
  `data\catalogo_insumos.json`, 565 claves): acepta `.xls` (lo abre el Excel de la PC, sin pip) y `.xlsx`.
  En **Ingresar material** se busca por clave o nombre en **todo** el catálogo y, al escribir la clave
  exacta o elegir una coincidencia, se autollenan **clave, descripción, unidad, categoría y fabricante**;
  si el insumo es de otra naturaleza, la pantalla cambia sola a su inventario. Debajo se muestra la
  **ficha del ERP** (clave, tipo, unidad, fabricante, código de barra, máximo y mínimo).
- **El catálogo del ERP autocompleta, no bloquea.** Si la clave existe, sus datos manda (descripción,
  unidad, categoría y fabricante) y si es de otro inventario el renglón se guarda en el libro que le
  corresponde. Si la clave **no** existe —o el insumo no trae tipo, o es producto terminado que el ERP
  todavía no tiene— el alta **se acepta** con un aviso y se guarda tal como se capturó: el almacén no
  puede frenar una recepción.
- **Buscar por descripción o por clave**: al escribir cualquiera de las dos se llena sola la otra (y la
  unidad, la categoría y el fabricante). Si varias claves comparten el nombre, se elige de una lista.
- **Ayuda en pantalla:** botón **? Ayuda** con 7 secciones (buscar, ubicar, mover/producción,
  deshacer, alta, atajos, si algo falla) y **Primeros pasos** la primera vez que entra cada usuario.

## Reglas que no se negocian

1. **Ubicaciones (lo único con catálogo obligatorio)**: solo códigos del catálogo ALM-PNO-002, porque
   es el layout físico de los racks. Las sugerencias solo **proponen**; nunca escriben un código que no
   exista. Una ubicación fuera del catálogo en el Excel del área se trata como «sin ubicar» y se avisa.
2. **Insumos**: el catálogo del ERP autocompleta y avisa; **cualquier material se acepta** (con o sin
   clave, con o sin tipo).
3. **Cantidades**: se interpretan de forma tolerante (`5 kg`, `5,5`, `1,458.50`) pero no se adivinan
   los casos ambiguos; la unidad escrita tiene que coincidir con la del renglón.
4. **`PRODUCCION` no es una ubicación**: es una salida del almacén y nunca se escribe en la columna
   UBICACIÓN.
5. **La bitácora no se edita**: deshacer y borrar dejan rastro (`revertido` + auditoría).


## Mapa de conexiones

`docs\diagramas\conexiones.html` (se abre en cualquier navegador, sin internet) muestra el recorrido
interfaz → endpoints → servidor → módulos → datos, con vistas guiadas para «leer el inventario»,
«registrar un movimiento», «ingresar material con clave del ERP» y «generar un reporte».
El fuente del diagrama es `docs\diagramas\conexiones.json`.

## Cómo se captura

**El campo de ubicación sugiere mientras escribes.** Al enfocar el destino (o la ubicación del alta)
aparecen las 8 mejores opciones, ya ordenadas: lo que coincide con lo escrito (código exacto, empieza,
termina — `a105` encuentra `MPNA-A-105` — o el nombre de la zona: `alergeno`, `cuarentena`), de la zona
del inventario correcto, **libres antes que ocupadas** (con el número de lotes), y si el material es
alérgeno, primero las zonas que sí lo admiten, con etiqueta roja cuando no. Se navega con ↑ ↓ y **Enter**
elige; también funciona con el dedo en el teléfono.

**Registrar un movimiento de ubicación** (lo normal cuando algo se mueve dentro del almacén):

1. Botón **⇄ mover** en el renglón, o **⇄ Registrar movimiento** arriba (busca el lote por material,
   lote o código ERP).
2. Escribes el rack destino, eliges el **motivo** (reacomodo, surtido parcial, devolución de restos,
   liberación de rack, cuarentena) y, si aplica, el **documento de referencia** y una nota.
3. **Registrar movimiento**. Queda en el Excel del área y en la bitácora con tu nombre y la hora.
   El modal muestra el **historial del lote** antes de moverlo.

**Mover solo unos kilos (parcialidad):** en el mismo modal, la cantidad viene completa; si escribes
menos (por ejemplo 5 de 1,458 kg), el sistema deja el resto en el rack original y **crea un renglón
nuevo** con los kilos movidos en el destino, con el mismo lote. Se registra como `PARCIAL` y muestra
en pantalla cuánto queda y dónde. No se parte ninguna celda: son dos renglones, como pide ALM-PNO-002 §V.2.2.

**Deshacer:** en **Movimientos**, botón *Deshacer* sobre el último movimiento vigente de ese lote
(devuelve la cantidad exacta a su ubicación anterior). El administrador puede además *Borrar* una
captura equivocada, con motivo obligatorio y registro en la tabla de auditoría.

**Asignar la primera ubicación** (lote que llega sin rack): escribes el código en la columna
**UBICACIÓN ACTUAL** y das **Enter**. Se registra como `UBICAR`.

**Si no aparece** en el rack: botón **no está**; el lote queda sin ubicación, con tu nota, y se ve en el
tablero como pendiente. Nunca se inventa una ubicación.

Los códigos válidos son los del procedimiento ALM-PNO-002 §V.2.1.4: 376 posiciones
(MPA-A-001 a 040, MPNA-A-041 a MPNA-B-136, MPT-B-001 a 020, MEP-B-001 a MEP-E-220)
más las zonas especiales (PASILLO-A, JAULA-MP, JAULA-PT, LAB-RETENCION, CUARENTENA-MP, ADUANA-MP, ANDEN).
Si escribes un código que no existe, el sistema lo rechaza con el motivo.

## Códigos QR / Barcode (opcional)

- **Columna `QR_ID` opcional:** se crea sola al final del Excel cuando se da de alta o se etiqueta un lote con código (`QR_ID`, `BARCODE_ID`, `Código QR`, `Barcode`, `Código de barras`). No cambia la identidad del sistema, que sigue operando por `(tipo, fila)` con o sin etiquetas.
- **Unicidad en tarimas activas:** un mismo `qr_id` no puede estar en dos renglones con existencia (`cantidad > 0`), ni siquiera entre distintos libros (`MP`, `MEP`, `PT`). Si el renglón queda en cero (anulado/consumido), la etiqueta puede reutilizarse.
- **Parcialidades:** al mover una parte de un lote (`PARCIAL`), la etiqueta física se queda en el renglón origen que permanece en el rack y **no** se duplica en el renglón hijo. Si la tarima nueva lleva etiqueta propia, se asigna con `/api/asignar_qr`.
- **Reimportación:** `python scripts\preparar_inventario.py --reimportar` conserva las etiquetas `qr_id` existentes empatando por material + lote cuando la clave es inequívoca (si hay varias filas para la misma clave, avisa y no reasigna para no duplicar etiquetas).
- **Buscador y atajos de cámara:** el buscador principal encuentra por código directo o con el prefijo `qr:`. El botón **📷 Escanear** consulta `GET /api/qr?codigo=…` y abre el modal de movimiento precargado (o propone darlo de alta si no existe); los botones **📷** junto a ubicación/destino rellenan el rack escaneado.
- **Requisito de cámara (HTTPS):** en la propia PC (`http://127.0.0.1` / `localhost`) el navegador permite abrir la cámara sin TLS. Desde teléfonos por red (`http://192.168.x.x`), los navegadores exigen contexto seguro (HTTPS) para `getUserMedia`/`BarcodeDetector`; sin TLS, el botón muestra un aviso claro y permite capturar el código manualmente (o con lector físico).

## Usuarios, contraseñas y teléfonos

- **Roles:** `admin` (todo), `almacen` (registra y deshace lo propio), `consulta` (solo lee y descarga).
- Se administran en la pestaña **Usuarios** o por consola:
  `python scripts\crear_usuario.py --nombre "Esteban Soto" --rol almacen`
  (`--listar`, `--reset` para contraseña olvidada).
- Contraseñas con PBKDF2-SHA256 y sal por usuario; sesión de 12 h en cookie HttpOnly.
- **Interfaz adaptable:** en teléfono cada lote es una tarjeta con botones de 46 px, en tablet y
  computadora es una tabla. Funciona en Chrome/Safari de Android, iPhone y iPad sin instalar nada;
  se puede agregar a la pantalla de inicio. Probado a 390×844, 820×1180 y 1280×900.

## Archivos y respaldo

```
data\Inventario_MP.xlsx     ← el inventario con la columna UBICACIÓN (fuente de verdad)
data\Inventario_MEP.xlsx    ← material de empaque
data\Inventario_PT.xlsx     ← producto terminado
data\catalogo_insumos.json  ← catálogo de insumos del ERP (claves y nombres canónicos)
data\bitacora.db            ← movimientos (SQLite, se puede abrir con cualquier visor)
data\logs\servidor.log      ← lo que imprimió el servidor y cualquier error
data\logs\depuracion.log    ← peticiones, eventos, errores y errores del navegador (para depurar)
data\salidas\               ← los Excel generados
data\respaldos\             ← copias con manifiesto SHA-256
```

Respaldo: `respaldo.cmd` copia los tres inventarios, el catálogo y la bitácora, y rota a 60 días. Con
`--copia-a "E:\respaldo_almacen"` deja una segunda copia en otra unidad.

## Para depurar

Cada petición a la API, cada acción y cada error (incluidos los de JavaScript de la pantalla) quedan en
`data\logs\depuracion.log`, con hora, usuario y duración. Con `ALMACEN_LOG=DEBUG` se agrega el detalle.
Como administrador, `http://127.0.0.1:8300/api/diagnostico` devuelve versión, rutas, disco, inventarios,
catálogo, últimos eventos y últimos errores. Las contraseñas y los tokens se tachan (`***`).

## Si el área actualiza su Excel

1. Cierra el archivo en Excel.
2. `python scripts\preparar_inventario.py --origen "C:\ruta\archivo_nuevo.xlsx" --reimportar`
3. Las ubicaciones ya capturadas se conservan (se empatan por descripción + lote), y el archivo
   anterior queda respaldado en `data\respaldos\`.

## Copia de la base a la USB

```bat
python scripts\sincronizar_usb.py --historial     :: copia consistente + copia con fecha
python scripts\sincronizar_usb.py --resumen       :: qué trae la copia
```

Deja en `D:\almacen-ubicaciones-datos\` la bitácora (con `sqlite3.backup()`, no copia de archivo: el
servidor puede estar escribiendo), los tres inventarios, un `sincronizacion.json` con SHA-256 y conteos,
y un `LEEME.txt`. Verifica `integrity_check` y que el número de movimientos coincida antes de dar la
copia por buena. Es solo lectura: la fuente de verdad sigue siendo `data\bitacora.db`.

## Verificación

```bat
python -m unittest discover -s tests
```

295 pruebas, todas ejecutadas (0 saltadas). Además, la interfaz se verifica en un Chrome real por CDP:
34 acciones de punta a punta y comprobaciones de maquetación a 390 y 1280 px. El detalle de lo que se
encontró y corrigió está en **`docs\REVISION_BUGS.md`**; las capturas, en `docs\pantalla_*.png`.

## Despliegue en la nube (Vercel + Supabase)

El proyecto está empaquetado para ejecutarse en **Vercel Serverless (Python)** usando **Supabase (PostgreSQL)** como base de datos única (usuarios, sesiones, inventarios MP/MEP/PT, catálogo del ERP, bitácora y auditoría), con HTTPS nativo para el lector de cámara QR/Barcode en móviles:

1. **Base de datos en Supabase:**
   - Esquema puro: `supabase/schema.sql` (y `supabase/migrations/20261028000000_initial_schema.sql`).
   - Esquema + migración de todos los datos actuales (`data/`): `supabase/seed.sql` (regenerable con `python scripts/migrar_supabase.py`).
   - Puedes pegar `supabase/seed.sql` en el **SQL Editor** de Supabase o aplicarlo directo desde terminal:
     ```bash
     python scripts/migrar_supabase.py --aplicar --dsn "postgresql://postgres.[REF]:[PASS]@aws-0-[REGION].pooler.supabase.com:6543/postgres?sslmode=require"
     ```
2. **Despliegue en Vercel:**
   - Punto de entrada: `api/index.py` + `vercel.json`.
   - Configura la variable de entorno `SUPABASE_DB_URL` (o `DATABASE_URL` / `POSTGRES_URL` mediante la integración oficial de Supabase en Vercel).
   - Despliega con `vercel --prod` o conectando el repositorio de GitHub a Vercel.

## Instalación local

Guía paso a paso y solución de problemas: **`LEEME_INSTALACION.md`** (o `docs\Manual_instalacion_almacen.pdf`).

```bat
instalar.cmd        :: crea .venv, instala dependencias, prepara los datos y verifica
arrancar.cmd        :: uso normal en http://127.0.0.1:8300
arranque_automatico.cmd
respaldo.cmd        :: respaldo con manifiesto SHA-256
demo.cmd            :: modo demostración en el puerto 8301 (datos de ejemplo, no toca los reales)
```

Dependencias: `requirements.txt` (solo `openpyxl`). Para PC sin internet ni pip, el paquete trae
`vendor\` con las librerías y el servidor las usa solo.

Empaquetar para llevar a otra PC:

```bat
python scripts\empaquetar.py --vendor
```

Genera `dist\almacen-ubicaciones_<fecha>.zip` con `MANIFIESTO.txt` (SHA-256 de cada archivo).

## Distribución y actualización (USB)

La versión del sistema vive en `app\version.py` (`1.8.1`: la 1.8.0 que llegó en la USB
más esta revisión). El paquete se arma con
`python scripts\empaquetar.py --vendor` (deja `VERSION.txt` y `MANIFIESTO.txt` dentro del ZIP) y la
carpeta de la USB con `python scripts\preparar_usb.py --destino "D:\ALMACEN_MIGH_ACTUALIZACION"`.

En la PC de destino: **doble clic en `ACTUALIZAR.cmd`**. Compara versiones, respalda código y datos,
detiene el servidor, copia el código nuevo, retira lo obsoleto, verifica, vuelve a levantar y —si algo
falla— restaura lo anterior. `data\` no se toca nunca. Detalle completo en `LEEME_INSTALACION.md`.

## Límites conocidos (honestos)

- **Alérgenos: avisa, no bloquea.** El autocompletado reconoce las palabras de la lista maestra
  F02-ALM-PNO-001 (leche, lactosa, WPC, soya, trigo, huevo…) y pone primero las zonas que los admiten,
  con etiqueta roja en las que no; el ingreso se permite igual porque el Excel del área no tiene la
  columna de SKU que amarra el material con la lista maestra.
- **4 lotes del archivo original traen la cantidad como fórmula** (`=16.46+24.14+25`) sin valor
  calculado; el sistema las evalúa para no perder 639 kg del total (72,215.4 kg), pero conviene
  convertir esas celdas a número en el archivo del área.
- **7 lotes no tienen caducidad legible** (5 vacías y 2 con texto imposible: `26/17/07`,
  `2028/07/2028`). Se muestran como "sin dato"; no se adivinan.
- Si el archivo está abierto en Excel, no se puede escribir: el sistema avisa y pide cerrarlo.
