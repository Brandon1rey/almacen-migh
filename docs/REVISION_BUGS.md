# Revisión de errores — octubre 2026

Este documento tiene dos partes:

1. **Revisión 1.7.x** — 24 errores encontrados y corregidos en la versión que se auditó primero.
2. **Reconciliación 1.8.0** — la actualización que llegó en la USB traía funciones nuevas y nunca
   auditadas (producción, catálogo del ERP, alta, borrado en bloque, modo red, copia a USB). Se
   conservaron **todas** las funciones y la interfaz tal cual, se auditó lo nuevo y se volvieron a
   aplicar las correcciones de la revisión 1.7.x que la actualización había sobrescrito.

Cada hallazgo tiene **síntoma**, **corrección** y **prueba automática** que lo cubre, para que no
vuelva a aparecer.

## Cómo se revisó

1. Lectura completa de `app/` (servidor, API, inventario, bitácora, reportes, usuarios, catálogo,
   red, inventarios), la interfaz y los scripts de instalación, respaldo, empaquetado, actualización
   y sincronización con la USB.
2. Comparación **archivo por archivo** entre la versión auditada (1.7.x) y la que llegó en la USB
   (1.8.0), para no perder ninguna corrección ni ninguna función.
3. Ejecución de la batería de pruebas de las dos versiones sobre el código mezclado, y ampliación de
   la batería con pruebas de regresión para cada fallo nuevo.
4. Sondas de API con entradas válidas, raras y hostiles; arranque en una carpeta limpia; dos procesos
   contra la misma carpeta; actualizador con un paquete simulado.
5. Verificación de la interfaz en un Chrome real (por CDP, sin ventana): 34 acciones de punta a punta
   y comprobaciones de maquetación en 7 tamaños (de 360 px a 1600 px).

## Errores corregidos en la revisión 1.7.x

| # | Severidad | Qué pasaba | Causa | Corrección | Prueba |
|---|---|---|---|---|---|
| 1 | **Crítico** | El primer arranque creaba el administrador y **se caía**; la contraseña aleatoria se perdía y nadie podía entrar nunca. | `main()` usaba una variable `ruta` inexistente al escribir `CREDENCIALES_INICIALES.txt` (`NameError`) justo después de crear el usuario. | Se escriben las credenciales junto a la carpeta de datos de la instancia y se imprimen antes de servir. | `test_primer_arranque_crea_admin_y_deja_las_credenciales` |
| 2 | **Crítico** | Con una carpeta de datos distinta (`python app\server.py <carpeta>`), las sesiones y movimientos se guardaban en `data\bitacora.db` del proyecto: dos instalaciones mezcladas. | `Estado` solo usaba la carpeta del inventario para la base en modo demo. | La bitácora vive siempre junto a los inventarios de la instancia. | `test_la_bitacora_vive_junto_a_los_datos` |
| 3 | **Alto** | La pantalla **Usuarios** no hacía nada: cambiar rol, desactivar o reponer contraseña devolvía `501`. | El servidor no implementaba `do_PATCH`. | `do_PATCH` despacha a la misma API; los usuarios inexistentes responden 404. | `test_patch_de_usuario_ya_funciona` + E2E |
| 4 | **Alto** | Quitar la ubicación dejaba el lote «sin ubicación» en pantalla, pero al recargar la ubicación **volvía**: el Excel nunca se borraba. | `ws.cell(..., value=None)` en openpyxl **no** borra la celda. | `asignar()` asigna `celda.value = None` explícitamente. | `test_limpiar_si_quita_la_ubicacion` |
| 5 | **Alto** | Cualquier parámetro raro producía **error 500**: `?tipo=xyz`, `?limite=abc`, `cantidad=NaN/Infinity`, descarga con tipo inválido. | Conversiones sin protección que escapaban hasta el `except` global. | Conversiones seguras (`_entero`, `_numero`, `inventario.numero`) y respuestas 400 con mensaje entendible. | `TestParametros`, `test_nan_del_navegador_no_llega_al_excel` |
| 6 | **Alto** | Borrar (purgar) el movimiento de una **parcialidad** dejaba los kilos en el limbo. | `_purgar` trataba todos los movimientos igual. | Para `PARCIAL` se revierte la cantidad y se anula el renglón nuevo, con validación previa; si ya cambió, responde 409. | `test_purgar_parcial_devuelve_los_kilos`, `test_purgar_parcial_ya_movida_no_rompe_los_datos` |
| 7 | Medio | La bitácora no guardaba la **unidad**: la pantalla mostraba «1,450» sin decir kg o piezas. | La columna no existía en el esquema ni se pasaba al registrar. | Columna `unidad` en el esquema, migración para bases viejas y exportación con la columna. | `test_guarda_la_unidad_del_movimiento`, `test_bitacora_exportada_trae_la_unidad` |
| 8 | Medio | El buscador del modal de movimiento **ignoraba el inventario activo**: estando en MEP o PT siempre buscaba en MP. | Las llamadas no llevaban `tipo`. | `conTipo()` agrega el inventario a toda llamada (búsqueda, lote, historial). | E2E + `test_inventario_por_tipo` |
| 9 | Medio | Dos personas moviendo lotes a la vez podían **corromper el Excel**. | No había exclusión entre hilos. | Candado de escritura (`_ESCRITURA`) alrededor de todas las operaciones que guardan. | `test_servidor` (hilos reales) |
| 10 | Medio | Escribir una cantidad con **coma** movía el lote **completo** en silencio (`Number("5,5")` → `NaN` → `null`). | El campo era `type=number` y la conversión no se validaba. | Campo de texto con `inputmode=decimal`, interpretación tolerante y validación en interfaz y servidor (`5,5` → 5.5). | `test_cantidad_con_coma_decimal_se_acepta` + E2E «5 kg» |
| 11 | Medio | Cada vez que el campo de ubicación perdía el foco se volvía a guardar: la bitácora se llenaba de `CONFIRMAR` repetidos. | No se actualizaba `defaultValue` después de guardar. | Se sincroniza el valor por omisión; el foco solo guarda si cambió. | E2E «salir del campo otra vez no duplica el movimiento» |
| 12 | Medio | Al abrir la página sin sesión se veía «Tu sesión terminó. Vuelve a entrar». | La comprobación inicial `/api/yo` devolvía 401. | `/api/yo` responde 200 con `autenticado: false`. | `test_logout_cierra_la_sesion`, `test_favicon_no_ensucia_la_consola` |
| 13 | Medio | Descargar con el Excel abierto navegaba a una página con JSON de error. | Se usaba `window.location` y el error no se podía atrapar. | Se descarga por `fetch` + blob y cualquier error se muestra como aviso legible. | E2E «el reporte F01 se descarga» |
| 14 | Medio | El F01 imprimía renglones **anulados** (cantidad 0) mientras el total decía otra cosa. | Se recorrían todas las filas, no solo las vigentes. | Reportes y cotas usan `inv.activas()`. | `test_f01_no_incluye_renglones_anulados` |
| 15 | Medio | `respaldo.py` respaldaba **solo MP**. | Ruta única fija. | Copia los tres inventarios y el catálogo, con manifiesto. | Revisión de código + manifiesto |
| 16 | Medio | `preparar_inventario.py` moría con `FileNotFoundError` en una PC nueva. | Usaba a ciegas la ruta del escritorio de otra PC. | Solo usa esa ruta si existe; si no, crea la plantilla vacía. | Revisión + ejecución manual |
| 17 | Medio | Pruebas y scripts fallaban con `ModuleNotFoundError: openpyxl` en PC sin pip. | `vendor\` solo lo agregaba el servidor. | `app/__init__.py` habilita `vendor\` para todo el que importe el paquete. | La batería completa corre sin pip |
| 18 | Medio | **34 pruebas se saltaban** en silencio. | Apuntaban al escritorio de otra computadora. | Origen = `data\Inventario_MP.xlsx` del repositorio cuando el archivo del área no existe. | 295 pruebas ejecutadas, 0 saltadas |
| 19 | Bajo | Cada carga pedía `/favicon.ico` y llenaba la consola de 404. | No había ruta ni icono. | Icono embebido y respuesta 204. | `test_favicon_no_ensucia_la_consola` |
| 20 | Bajo | Las cantidades guardadas como **texto** en el Excel desaparecían de la lista y de los totales. | Solo se aceptaban números. | Se leen textos numéricos inequívocos; los ambiguos (`1.458`) se **reportan en pantalla**, no se adivinan. | `TestCantidadesIlegibles` |
| 21 | Bajo | `NO_ENCONTRADO` podía registrarse sin nota; `MOVER` con motivo «Otro» también. | Sin validación. | Ambos exigen texto; el modal lo pide y explica. | `test_no_encontrado_exige_nota`, `test_motivo_otro_exige_nota` |
| 22 | Bajo | Se podían adivinar contraseñas sin límite. | Sin freno. | Tras 8 fallos en 5 minutos para el mismo usuario, responde 429 con la espera. | `test_bloqueo_temporal_tras_ocho_fallos` |
| 23 | Bajo | Al cambiar la contraseña de uno mismo, la sesión moría **sin aviso**. | El servidor cierra todas las sesiones de ese usuario. | La interfaz lo detecta y lleva al acceso con un mensaje claro. | Revisión + E2E |
| 24 | Bajo | Cambiar el rol propio no avisaba que se perdía el acceso de administrador. | Sin advertencia. | La confirmación lo advierte. | Revisión |

## Captura tolerante (capturar rápido y sin errores)

Escribir a mano en el almacén es con prisa y con el celular. El sistema entiende lo que la persona
escribe en vez de exigirle un formato:

- **Códigos de ubicación**: `MPA-A-005`, `mpa a 5`, `MPAA005`, `mpa-a-5` y `MPA A 005` caen en el
  código canónico. Si el código no existe en el catálogo, **no se escribe**: se rechaza y se proponen
  los más parecidos. La ubicación solo puede salir del catálogo de ALM-PNO-002.
- **Cantidades**: `5`, `5,5`, `5.5`, `5 kg`, `1 458`, `1,458.50`, `1.458,50` y `1.234,5`. Si la lectura
  es ambigua se usa el total del renglón para desempatar y nunca se mueve más de lo que hay. La unidad
  escrita tiene que coincidir: `5 g` en un lote de kg **se rechaza en vez de mover 5 kg**. La interfaz
  avisa antes de confirmar cuánto se va a mover y cuánto queda.
- **Búsqueda tolerante**: si la búsqueda exacta no encuentra nada, se muestran los renglones más
  parecidos («protiena» → Proteína), la pantalla dice que son **aproximados** y aparece
  **«¿Quisiste decir…?»** con un toque para buscar el nombre tal como está en el inventario. Igual en
  la lista y en el buscador del modal de movimiento.
- **Catálogo del ERP**: al escribir una clave o un nombre aparecen las coincidencias, y si no hay
  ninguna, los insumos parecidos del catálogo («bolsa laminada» escrito a medias). Nunca se inventa un
  insumo: las sugerencias salen del catálogo del ERP.
- **Borrar con claridad**: deshacer y borrar explican qué va a pasar; el borrado exige motivo y queda
  en auditoría. Todas las ventanas tienen su botón de **Cancelar/Cerrar** y `Escape` cierra la de
  arriba.

## Lo que la interfaz explica sola

- **Ayuda** (botón **?** en la barra superior, también en la pantalla de acceso): 7 secciones —buscar,
  ubicar, mover/parcialidad/producción, deshacer y borrar, ingresar material, atajos de teclado y qué
  hacer si algo falla— y la versión instalada.
- **Primeros pasos**: cuatro pasos la primera vez que cada usuario entra, con botón *Entendido*; se
  recuerda en el navegador para no volver a aparecer.
- **Avisos de calidad de datos** arriba de la lista: ubicaciones escritas fuera del catálogo,
  cantidades ilegibles y renglones sin lote, con qué hacer en cada caso.
- Cada pantalla trae su explicación; los campos piden lo que necesitan y los mensajes de error dicen
  qué pasó y cómo corregirlo (código fuera del catálogo, cantidad mayor a la existencia, unidad que no
  coincide, contraseña corta, motivo sin nota…).
- Roles: `consulta` ve la pantalla sin botones de escritura; `almacen` solo puede deshacer lo suyo; el
  administrador ve el borrado en bloque y los usuarios.

## Reconciliación con la versión 1.8.0

### Lo que llegó en la 1.8.0 y se conserva

- Los **tres inventarios** (MP, MEP, PT) con selector, tablero y F01 consolidado.
- **Ingresar material** con naturaleza, catálogo del ERP, sugerencias de ubicación y validaciones.
- **Salida a producción** (`PRODUCCION`) como destino que no es un rack.
- **Borrado de renglones en bloque** (solo administrador, con motivo y respaldo del libro).
- **Modo red**: compartir la dirección, estado del modo, aviso de firewall y la dirección en el acceso.
- **Sugerencias de ubicación** ordenadas por zona, ocupación y alérgenos, con ↑ ↓ y Enter.
- **Copia a la USB** desde la pantalla (comparador, historial y «forzar»).
- **Versión, actualizador y paquete**: `app\version.py`, `empaquetar.py`, `preparar_usb.py`,
  `actualizar.py`, `ACTUALIZAR.cmd`, `sincronizar_usb.cmd`, `scripts\espejo.py`.
- La interfaz completa de la 1.8.0, sin rediseños: las mejoras de esta revisión son **aditivas**
  (ayuda, primeros pasos, avisos y correcciones de comportamiento).

### Hallazgos y correcciones de la 1.8.0

| # | Severidad | Qué pasaba | Corrección | Prueba |
|---|---|---|---|---|
| 25 | **Crítico** | En una PC sin pip, `import app` (y cualquier script o prueba) fallaba con `ModuleNotFoundError: openpyxl` aunque `vendor\` estuviera completo. | `app/__init__.py` agrega `vendor\` al `sys.path` al importar el paquete. | La batería completa corre sin instalar nada |
| 26 | **Crítico** | El **alta con clave del ERP inexistente** escribía un renglón con datos tecleados sin decir nada; y con una clave de **otro inventario** el renglón quedaba en el libro equivocado. | Primero se bloqueó el alta; después, con el área, se corrigió la regla: **se acepta todo** con aviso. Si la clave existe, sus datos manda y el aviso dice qué se corrigió; si es de otro inventario, el renglón se guarda en el libro que le toca. Solo la **ubicación** se valida contra el catálogo (layout físico). | `test_acepta_clave_que_no_esta_en_el_catalogo`, `test_acepta_material_de_otro_inventario_y_lo_pasa_a_su_libro`, `test_la_ubicacion_sigue_siendo_lo_unico_obligatorio_del_catalogo` + E2E |
| 27 | **Crítico** | El catálogo del ERP se leía siempre de `data\` del **código**, no de la carpeta de datos de la instancia: en modo demo o con otra carpeta de datos se validaba contra el catálogo equivocado. | `catalogo.ruta()` resuelve `ALMACEN_CATALOGO` → carpeta de datos → `data\` del código. | `test_el_catalogo_sigue_a_la_carpeta_de_datos` |
| 28 | **Alto** | En la lista, marcar renglones y **cambiar de inventario** dejaba la selección: borrar después eliminaba renglones del **libro equivocado**. | La selección se limpia al cambiar de inventario y el borrado avisa si quedarán renglones marcados fuera de la vista. | E2E «cambiar de inventario limpia la selección» |
| 29 | **Alto** | El **borrado en bloque** renumeraba las filas de la bitácora sin bajar de mayor a menor y podía dejar referencias a filas nulas. | Renumeración descendente, respaldo del libro antes de tocar y movimiento `NULL` cuando la fila ya no existe. | `test_borrar_renglones.py` |
| 30 | **Alto** | Escribir una cantidad **vacía o ilegible** en el modal movía el lote completo en silencio. | La interfaz avisa «sin cantidad se mueve TODO el lote», rechaza lo que no se puede leer y el servidor responde 400 con el motivo. | E2E «parcialidad con 5 kg» |
| 31 | Medio | El buscador del modal seguía consultando el inventario por omisión (MP) en lugar del activo. | Se agregó `tipo` a la búsqueda del modal y el modal muestra los **parecidos** cuando no hay coincidencia exacta. | E2E «búsqueda aproximada» |
| 32 | Medio | Dos peticiones de búsqueda podían llegar **fuera de orden** y pintar resultados viejos (lista y bitácora). | Cada petición lleva un turno; la respuesta que llega tarde se descarta. | Revisión de código + E2E |
| 33 | Medio | El registro de `data\logs\servidor.log` ignoraba una redirección de consola, se anidaba al rearmarse y dejaba el archivo abierto. | Envuelve la salida **actual** (respeta `>` y las pruebas), reemplaza el registro anterior, lo cierra y expone `cerrar_registro()`. | `test_primer_arranque_*` |
| 34 | Medio | El actualizador **se negaba a instalar** un paquete correcto en una PC sin pip: su verificación importaba `openpyxl` aparte y daba un falso «la versión nueva no arranca» (y restauraba todo). | La verificación importa `app.server`, `app.api`, `app.inventario` y `app.version` (openpyxl entra por `app.inventario`) y pone `vendor\` en el `PYTHONPATH`. | `test_actualizar.py` (12 pruebas) |
| 35 | Medio | `preparar_inventario.py` dejaba el respaldo de la reimportación en `data\respaldos\` del **repositorio** aunque el `--destino` fuera otra carpeta, y se caía si una ubicación guardada ya no existía en el catálogo. | Respalda junto al archivo que reemplaza; las ubicaciones fuera del catálogo se **descartan con aviso** (no se reescriben) y el resumen reporta ubicaciones inválidas y cantidades ilegibles. | Ejecución manual + `test_inventario` |
| 36 | Bajo | En el alta no se podía elegir un insumo de **otro inventario**: la búsqueda del catálogo filtraba por el inventario actual, así que el cambio automático de naturaleza nunca ocurría. | Se busca en todo el catálogo (cada resultado muestra su tipo); al elegirlo, la pantalla cambia a la naturaleza correcta y toma la descripción canónica. | E2E «elegir un insumo del ERP cambia a su inventario» |
| 37 | Bajo | `Escape` no cerraba las ventanas nuevas (red, ayuda). | La lista de ventanas que cierra `Escape` incluye todas. | E2E «Esc cierra la ayuda» |
| 38 | Bajo | Doble clic en «Registrar movimiento», «Crear usuario», «Borrar» o «Guardar contraseña» enviaba dos peticiones. | Candado por botón/petición mientras está en curso. | Revisión de código + E2E |
| 39 | **Alto** | Importar el catálogo real (`insumo.xls`) fallaba dos veces: sin pywin32 no se podía leer el `.xls`, y al imprimir «→» en la consola de Windows el script terminaba con `UnicodeEncodeError` **después** de haber escrito el JSON (el operador creía que no se había importado). | Se agregó un lector `.xls` por COM de Windows vía PowerShell (sin pip) y un guardia de codificación en todos los scripts de consola. | Importación real de `insumo.xls` + `test_guarda_todos_los_datos_del_insumo` |
| 40 | Medio | El catálogo descartaba el **código de barra**, no contaba producto terminado y el alta no dejaba ver máximo/mínimo/fabricante del ERP. | Se guardan todos los campos del ERP, `producto terminado → pt`, y el alta muestra la **ficha del insumo** y autollena al escribir la clave exacta. | `test_unidades_y_tipos_del_erp`, E2E «escribir la clave del ERP llena…» |

### Observabilidad (para poder depurar sin adivinar)

La app no tenía forma de contar qué pasó: los errores morían en la consola negra y en la bitácora
solo quedaba el movimiento, no el fallo. Se agregaron dos piezas pequeñas:

| Pieza | Qué hace |
|---|---|
| `app/logger.py` | Escribe `data\logs\depuracion.log` (rota a 2 MB, 5 copias): una línea por petición (`peticion POST /api/mover -> 200 14.2ms usuario=Esteban`), cada evento, y los errores con traceback completo. Tacha contraseñas y tokens (`***`). Nivel configurable con `ALMACEN_LOG` (`DEBUG` agrega el detalle de estáticos y valores rechazados). |
| `app/eventos.py` | Bus de eventos con escuchas por nombre exacto, `"*"` o prefijo (`"movimiento.*"`). El escucha del registro deja cada evento en el archivo; los últimos 300 quedan en memoria. Un escucha roto se registra y **no** tumba la operación que publicó. |

- Eventos publicados: `sesion.iniciada`, `sesion.cerrada`, `movimiento.registrado` (con acción,
  fila, lote, de/a, cantidad y usuario), `renglones.borrados`, `copia.usb`, `servidor.arrancado`,
  `servidor.detenido`, `servidor.error` y `navegador.error`.
- El navegador **escucha sus propios errores** (`window.onerror` y `unhandledrejection`), las fallas
  de red y los 5xx, y los manda a `POST /api/log` (limitado a 30 por minuto por usuario para no
  inundar el archivo). Así, cuando alguien dice «no me funcionó», el error de la pantalla ya está en
  el servidor.
- `GET /api/diagnostico` (solo administrador) devuelve en un JSON: versión, Python y plataforma,
  rutas de datos y del registro, espacio en disco, inventarios cargados, catálogo del ERP, conteos
  de la bitácora, últimos 60 eventos y últimos 15 errores.
- Pruebas: `tests/test_observabilidad.py` (19) cubren el archivo, el tachado de secretos, las
  escuchas exactas/prefijo/rotas, el historial, los endpoints y que un error real de JavaScript de la
  pantalla llegue al registro (E2E).

### Decisiones de criterio (una sola fuente de verdad)

Estas reglas son deliberadas y no se negocian con la interfaz:

1. **Ubicaciones**: solo códigos del catálogo ALM-PNO-002. La búsqueda tolerante y las sugerencias
   **solo proponen**; nunca escriben un código que no exista. Una ubicación escrita fuera del catálogo
   en el Excel del área se cuenta como **sin ubicar** y se reporta en pantalla.
2. **Insumos**: el catálogo del ERP **autocompleta y avisa**, pero el alta **acepta cualquier
   material** (clave nueva, insumo sin tipo o producto terminado que el ERP aún no tiene). Si la clave
   existe, sus datos mandan y, si es de otro inventario, el renglón se guarda en el libro correcto. Se
   revisó con el área: frenar una recepción por el catálogo era peor que aceptarla con aviso.
3. **Cantidades**: se interpretan de forma tolerante, pero **no se adivinan** cuando hay ambigüedad
   (`1.458`) y la unidad escrita tiene que coincidir con la del renglón.
4. **`PRODUCCION` no es una ubicación**: es una salida del almacén. Nunca se escribe en la columna
   UBICACIÓN del Excel.
5. **La bitácora no se edita**: deshacer y borrar dejan rastro (`revertido`, tabla de auditoría).

## Verificación

```bat
python -m unittest discover -s tests
:: Ran 318 tests ... OK
```

- **318 pruebas** de Python, **todas ejecutadas** (antes: 294 con 34 saltadas en silencio). Cubren
  inventario, bitácora, API, servidor HTTP real, reportes, instalación, actualizador, paquete, USB,
  espejo, red, catálogo, alta, producción y borrado en bloque.
- **39 acciones de interfaz** verificadas en un Chrome real por CDP, sin errores de JavaScript:
  acceso, ayuda, primeros pasos, búsqueda aproximada con sugerencias, ubicación inválida y válida,
  parcialidad con «5 kg», salida a producción, alta rechazada por clave inexistente, cambio de
  inventario al elegir un insumo del ERP, selección y limpieza al cambiar de inventario, bitácora,
  red y copia.
- **Mapa de conexiones** en `docs\diagramas\conexiones.html` (archify: 9 comprobaciones
  de composición en 0 errores y 0 avisos, y containment verificado en 1440, 1600,
  1920 y 2048 px, tema claro y oscuro).
- **Maquetación**: sin desbordes horizontales en 360×740, 390×844, 768×1024, 820×1180,
  1024×768, 1280×900 y 1600×1000; la lista usa tarjetas en teléfono y columnas en escritorio
  (el desplazamiento vertical de la página es el normal). Capturas actualizadas en
  `docs\pantalla_*.png`; `pantalla_alta_aceptada.png` muestra el alta con una clave que
  no está en el catálogo del ERP, aceptada con aviso.
- **Arranque limpio** comprobado: carpeta nueva → administrador, credenciales impresas y en
  `CREDENCIALES_INICIALES.txt`, sesión que se conserva al recargar.
- **Actualizador** probado con un paquete simulado: respalda, copia, retira lo obsoleto, verifica y
  restaura si falla.

## Límites que siguen siendo ciertos

- Si el Excel está abierto con Excel, no se puede escribir: el sistema lo dice y pide cerrarlo.
- **4 celdas de cantidad son fórmulas** (`=16.46+24.14+25`) sin valor calculado; el sistema las evalúa
  para no perder 639 kg del total (72,215.4 kg), pero conviene convertirlas a número en el archivo del
  área.
- **7 lotes no tienen caducidad legible** (5 vacías y 2 con texto imposible: `26/17/07`,
  `2028/07/2028`). Se muestran como «sin dato»; no se adivinan.
- **1 celda de cantidad es ilegible** y aparece en el aviso de calidad de datos con su número de fila.
- El archivo del repositorio (`data\Inventario_MP.xlsx`) tiene **0 ubicaciones capturadas**: es el
  archivo del área tal como llega. La cobertura real es la de la instalación del almacén.
- **Alérgenos: avisa, no bloquea.** El autocompletado reconoce las palabras de la lista maestra
  F02-ALM-PNO-001 y pone primero las zonas que los admiten, con etiqueta roja en las que no; el
  ingreso se permite igual porque el Excel del área no trae la columna que amarra el material con la
  lista maestra.
- La aplicación es para la **red local**; no debe exponerse a internet.

## QR/Barcode: capa aditiva (1.9.0)

El soporte de códigos QR y códigos de barras se integró como una capa aditiva mínima sobre los tres libros de Excel (`qr_id`), sin tocar el esquema de `bitacora.db` ni las rutas de movimiento:

| # | Decisión | Motivo |
|---|---|---|
| D1 | Unicidad global entre MP/MEP/PT, no solo por libro | El escaneo es uno solo: si el mismo código existe en dos libros activos, la resolución es ambigua. Se valida contra todos los libros cargados (`estado.cargados()`). |
| D2 | La parcialidad no hereda `qr_id`; se queda en el origen | El hijo es una tarima nueva sin etiqueta; el origen sigue físicamente en el rack con el resto. Etiquetar el hijo queda como acción explícita (`POST /api/asignar_qr`). |
| D3 | `preparar_inventario.py --reimportar` preserva `qr_id` (merge por clave inequívoca) | Reemplazar el archivo del área sin conservar `qr_id` borraría las etiquetas en cada reimportación. Si una clave tiene varias filas, avisa y no duplica. |
| D4 | `normalizar_qr()` usa `strip()` y `clave_qr()` compara con `.upper()`, no `normalizar()` | `normalizar()` convierte `-` y `.` en espacios (`7501-055` y `7501055` colisionarían). En códigos de barras se conserva la puntuación. |
| D5 | Columnas reconocidas: `QR_ID`, `BARCODE_ID`, `Código QR`, `Barcode`, `Código de barras` | `_detectar_columnas()` normaliza encabezados con `normalizar().lower()`, mapeando todas las variantes a `qr_id`. |
| D6 | Creación lazy de la columna `QR_ID` en `agregar_alta` / `asignar_qr` | Abrir el inventario para consulta no modifica el archivo hasta que se guarde una etiqueta o un alta. |
| D7 | Cámara sujeta a contexto seguro (HTTPS / `localhost`) con entrada manual de respaldo | `getUserMedia` y `BarcodeDetector` exigen contexto seguro en navegadores móviles; en `http://IP` se informa claramente y se ofrece captura manual. |
