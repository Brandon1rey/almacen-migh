# Guion de demostración (3 minutos)

**Qué es real y qué es ficticio**

| Dato | Origen |
|---|---|
| 137 lotes de materia prima: descripción, proveedor, cantidad, lote, caducidad, observaciones | **Reales** — archivo del área del 06-oct-2026 |
| Ubicaciones (MPA, MPNA, MPT, MEP, jaulas, cuarentena, pasillo A) | **Ficticias**, asignadas con criterio: alérgenos en MPA, no alérgenos en MPNA, empaque en MEP, PT en MPT |
| 19 lotes de material de empaque y producto terminado | **Ficticios** (el Excel del área es solo de MP); cantidades y lotes verosímiles |
| Los 276 movimientos de la bitácora | **Ficticios**, con fechas de los últimos 14 días |

El archivo demo lleva una hoja **DEMO** en rojo y la app muestra una banda roja: no se puede confundir con un registro.

## Antes de empezar

```bat
demo.cmd
```

Se abre en `http://127.0.0.1:8301`. La app real (puerto 8300) no se toca.

## 0:00 — El problema en 20 segundos

> «Hoy el almacén se organiza con memoria. El archivo tiene 137 lotes, 29 ya vencidos y 48 que vencen antes de fin de año, y **no tiene una sola columna que diga dónde está nada**. No es un problema de auditoría: es que no podemos contestar qué se vence primero.»

Señalar la banda roja y aclarar: **esto es un demo, las ubicaciones son de ejemplo y los datos de materia prima son los reales del archivo.**

## 0:30 — Buscar (la pregunta del piso)

Teclear `lactosa` en el buscador.

> «Dos lotes. Uno en `PASILLO-A` —que es el área de rechazo de alérgenos— y otro en `MPA-A-008` con 65 días vencido. El sistema no esconde lo que está mal: lo pone en rojo.»

## 1:00 — Registrar el movimiento (el corazón del sistema)

Botón **⇄ mover** en la fila (o **⇄ Registrar movimiento** arriba, que permite buscar el lote):

1. El modal muestra el lote y su **ubicación actual**.
2. Escribir el rack destino (`MPA-A-023`), elegir **motivo** (reacomodo, surtido parcial, devolución de
   restos, liberación de rack, cuarentena), y opcionalmente **documento de referencia** (OP-1234) y nota.
3. **Registrar movimiento** → banda verde de confirmación y el renglón cambia de ubicación.

> «Dos toques y un motivo. Abajo del modal está el **historial del lote**: de dónde vino, quién lo movió y
> por qué. Eso es lo que el auditor muestrea, y es lo que hoy no existe en ningún lado.»

Si el destino es igual a la ubicación actual, el sistema lo rechaza: no se registran movimientos falsos.

## 1:30 — «No está» (el fracaso honesto)

Botón **no está** en un lote, escribir una nota (`no estaba en el rack; se busca en el pasillo`).

> «No se inventa una ubicación. El lote queda sin ubicación y con la nota. Es deuda visible, no un dato falso.»

## 2:00 — El reporte, con un clic

Botón **Reporte F01**:

- Bloques por zona: `MP-ALERGENO`, `MP-NO-ALERGENO`, `PT`, `EM`, `ESPECIAL`.
- Al final, bloque ámbar `SIN UBICAR` con los 7 lotes pendientes.
- Bloque de firmas: Elaboró / Revisó / Autorizó.

> «Nadie llenó 370 casillas. Se imprimió, se firmó y se archivó.»

## 2:30 — Tabla macro + FEFO

Botón **Tabla macro + FEFO** → hoja **Caducidades**:

> «Ordenado por lo que vence primero, con semáforo. 29 vencidos, 7 en los próximos 30 días. Esta es la hoja que se pega en el almacén.»

## 3:00 — Cierre

> «El sistema no le pidió al almacén cambiar de herramienta: le agregó una columna a su Excel y le devuelve los reportes en Excel. La diferencia es que ahora cada movimiento deja rastro, y el reporte del mes no se captura: se genera.»

## Preguntas que van a salir

| Pregunta | Respuesta |
|---|---|
| ¿Quién captura? | El que mueve, en el momento, con su nombre. Es el mismo registro que hoy hace en papel (R-ALM-001). |
| ¿Y si se equivoca? | Corrige el código: el movimiento anterior queda en la bitácora con origen y destino. No se borra nada. |
| ¿Funciona sin internet? | Sí. Es local, en la PC del almacén. El celular es opcional (misma red). |
| ¿Se puede perder? | El Excel y la base viven en la PC; el respaldo se copia a USB BRANDON (que se sincroniza a D:). |
| ¿Esto cambia el procedimiento? | El reporte ya está previsto en ALM-PNO-002 §V.2.1.8. El conteo cíclico por zona sí requiere actualizar el procedimiento (propuesta en el memo). |
