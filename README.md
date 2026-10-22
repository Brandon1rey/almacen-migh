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

## Cómo se captura

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
data\bitacora.db            ← movimientos (SQLite, se puede abrir con cualquier visor)
data\salidas\               ← los Excel generados
data\respaldos\             ← copia automática antes de cada reimportación
```

Respaldo manual: copiar `data\Inventario_MP.xlsx` y `data\bitacora.db`.

## Si el área actualiza su Excel

1. Cierra el archivo en Excel.
2. `python scripts\preparar_inventario.py --origen "C:\ruta\archivo_nuevo.xlsx" --reimportar`
3. Las ubicaciones ya capturadas se conservan (se empatan por descripción + lote), y el archivo
   anterior queda respaldado en `data\respaldos\`.

## Verificación

```bat
python -m unittest discover -s tests
```

## Instalación

El sistema se instala en una PC y las demás entran por red. Guía paso a paso y solución de problemas:
**`LEEME_INSTALACION.md`** (o `docs\Manual_instalacion_almacen.pdf`).

```bat
instalar.cmd        :: crea .venv, instala openpyxl, prepara los datos y verifica
arrancar.cmd        :: uso normal en http://127.0.0.1:8300
servidor_lan.cmd    :: una PC con los datos, las demás por navegador
arranque_automatico.cmd
respaldo.cmd        :: respaldo con manifiesto SHA-256
```

Dependencias: `requirements.txt` (solo `openpyxl`). Para PC sin internet ni pip, el paquete trae
`vendor\` con las librerías y el servidor las usa solo.

Empaquetar para llevar a otra PC:

```bat
python scripts\empaquetar.py --vendor
```

Genera `dist\almacen-ubicaciones_<fecha>.zip` con `MANIFIESTO.txt` (SHA-256 de cada archivo).

## Modo demo (para presentar)

```bat
demo.cmd
```

Abre en **http://127.0.0.1:8301** con datos poblados: 156 lotes, 149 con ubicación (95.5 %), empaque y producto
terminado en sus zonas, 5 zonas especiales usadas (cuarentena, jaula MP, jaula PT, retención de laboratorio,
pasillo A) y **267 movimientos** en la bitácora. El archivo demo lleva una hoja `DEMO` en rojo y la pantalla
una banda roja: **las ubicaciones y el empaque son ficticios; las 137 filas de materia prima son las reales
del archivo del área**. Guion de la demostración en `docs\guion_demo.md`.

Usuarios del demo: `admin` / `Demo.Admin.2026`, `Esteban` / `Demo.Almacen.2026`,
`Galilea` / `Demo.Almacen.2026b`, `Auditor` / `Demo.Consulta.2026`.

Generar o regenerar el demo (determinista, semilla 42):

```bat
python scripts\sembrar_demo.py
```

## Límites conocidos (honestos)

- **No valida alérgenos contra la zona**: la lista maestra F02-ALM-PNO-001 identifica por SKU
  `MIGH-*` y el Excel del área no tiene esa columna. Mientras no se agregue el SKU, la regla
  "MPA = alérgeno" no se puede verificar automáticamente.
- **4 lotes del archivo original traen la cantidad como fórmula** (`=16.46+24.14+25`) sin valor
  calculado; el sistema las evalúa para no perder 639 kg del total (72,215.4 kg), pero conviene
  convertir esas celdas a número en el archivo del área.
- **7 lotes no tienen caducidad legible** (5 vacías y 2 con texto imposible: `26/17/07`,
  `2028/07/2028`). Se muestran como "sin dato"; no se adivinan.
- Si el archivo está abierto en Excel, no se puede escribir: el sistema avisa y pide cerrarlo.
