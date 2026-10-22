"""API JSON del inventario. Sin sockets: recibe el Inventario ya cargado."""
import time

from app import bitacora, reportes, ubicaciones

# motivos válidos de un movimiento explícito (el de la pantalla "Registrar movimiento")
MOTIVOS = {
    "reacomodo": "Reacomodo por espacio",
    "surtido": "Surtido parcial / consumo",
    "devolucion": "Devolución de restos al almacén",
    "liberacion": "Liberación de rack o pasillo",
    "cuarentena": "Envío a cuarentena o retención",
    "otro": "Otro (explicar en la nota)",
}


def manejar(estado, metodo, ruta, params, cuerpo):
    inv = estado.inventario()
    cuerpo = cuerpo or {}
    usuario = (cuerpo.get("usuario") or params.get("usuario") or "").strip()

    if metodo == "GET" and ruta == "/api/inventario":
        filas = inv.buscar(
            texto=params.get("q", ""),
            solo_sin_ubicar=params.get("sin_ubicar") in ("1", "true", "si"),
            solo_vencidos=params.get("vencidos") in ("1", "true", "si"),
        )
        return 200, {"filas": filas, "resumen": inv.resumen(), "total_filtrado": len(filas)}

    if metodo == "GET" and ruta == "/api/lote":
        try:
            return 200, {"lote": inv.por_fila(int(params.get("fila", 0)))}
        except (KeyError, TypeError, ValueError) as e:
            return 400, {"error": str(e)}

    if metodo == "GET" and ruta == "/api/historial":
        try:
            fila = int(params.get("fila", 0))
        except (TypeError, ValueError):
            return 400, {"error": "fila inválida"}
        return 200, {"movimientos": bitacora.por_fila(estado.con, fila,
                                                      int(params.get("limite", 12)))}

    if metodo == "GET" and ruta == "/api/resumen":
        return 200, inv.resumen()

    if metodo == "GET" and ruta == "/api/estado":
        return 200, {"archivo": getattr(estado, "archivo", ""),
                     "demo": bool(getattr(estado, "demo", False)),
                     "lotes": len(inv.filas),
                     "movimientos": bitacora.contar(estado.con)}

    if metodo == "GET" and ruta == "/api/ubicaciones":
        return 200, {"codigos": ubicaciones.CODIGOS,
                     "zonas": {z: len(c) for z, c in ubicaciones.por_zona().items()},
                     "motivos": MOTIVOS}

    if metodo == "GET" and ruta == "/api/bitacora":
        return 200, {"movimientos": bitacora.ultimos(estado.con, int(params.get("limite", 60)))}

    if metodo == "POST" and ruta == "/api/mover":
        return _traslado(estado, inv, cuerpo, usuario)

    if metodo == "POST" and ruta in ("/api/ubicar", "/api/limpiar", "/api/no_encontrado"):
        return _movimiento(estado, inv, ruta, cuerpo, usuario)

    return 404, {"error": f"ruta no encontrada: {metodo} {ruta}"}


def _traslado(estado, inv, cuerpo, usuario):
    """Movimiento explícito entre ubicaciones: origen conocido, destino nuevo y motivo."""
    if not usuario:
        return 400, {"error": "falta tu nombre: la bitácora es el rastro de auditoría"}
    try:
        fila = int(cuerpo["fila"])
        registro = inv.por_fila(fila)
    except (KeyError, TypeError, ValueError):
        return 400, {"error": "fila requerida"}

    destino = ubicaciones.normalizar(cuerpo.get("destino") or cuerpo.get("ubicacion"))
    motivo = (cuerpo.get("motivo") or "").strip().lower()
    if motivo not in MOTIVOS:
        return 400, {"error": f"motivo inválido; usa uno de: {', '.join(MOTIVOS)}"}
    if not destino:
        return 400, {"error": "destino requerido"}
    if not registro["ubicacion"]:
        return 400, {"error": "el lote no tiene ubicación actual: primero asígnalo"}
    if destino == registro["ubicacion"]:
        return 400, {"error": f"el destino es la ubicación actual ({destino}); no hay movimiento"}

    anterior = registro["ubicacion"]
    try:
        inv.asignar(fila, destino)
    except ValueError as e:
        return 400, {"error": str(e)}

    guardado, error = _persistir(estado, inv, fila, "MOVER", usuario, anterior, destino,
                                 motivo, cuerpo.get("ref_doc"), cuerpo.get("nota"))
    if error:
        return error
    return 200, {"accion": "MOVER", "fila": fila, "anterior": anterior, "nueva": destino,
                 "motivo": motivo, "ts": guardado["ts"], "movimiento": guardado["id"],
                 "resumen": inv.resumen(), "bitacora": bitacora.contar(estado.con),
                 "historial": bitacora.por_fila(estado.con, fila, 12)}


def _movimiento(estado, inv, ruta, cuerpo, usuario):
    if not usuario:
        return 400, {"error": "falta tu nombre: la bitácora es el rastro de auditoría"}
    try:
        fila = int(cuerpo["fila"])
    except (KeyError, TypeError, ValueError):
        return 400, {"error": "fila requerida"}
    try:
        inv.por_fila(fila)
    except KeyError as e:
        return 400, {"error": str(e)}

    try:
        if ruta == "/api/ubicar":
            destino = ubicaciones.normalizar(cuerpo.get("ubicacion"))
            anterior, nueva = inv.asignar(fila, destino)
            accion = "UBICAR" if not anterior else "MOVER"
            if anterior == nueva:
                accion = "CONFIRMAR"
        elif ruta == "/api/limpiar":
            anterior, nueva = inv.asignar(fila, None)
            accion = "LIMPIAR"
        else:  # no_encontrado
            anterior, nueva = inv.asignar(fila, None)
            accion = "NO_ENCONTRADO"
    except ValueError as e:
        return 400, {"error": str(e)}

    guardado, error = _persistir(estado, inv, fila, accion, usuario, anterior, nueva,
                                 None, None, cuerpo.get("nota"))
    if error:
        return error
    return 200, {"accion": accion, "fila": fila, "anterior": anterior, "nueva": nueva,
                 "ts": guardado["ts"], "movimiento": guardado["id"], "resumen": inv.resumen(),
                 "bitacora": bitacora.contar(estado.con)}


def _persistir(estado, inv, fila, accion, usuario, anterior, nueva, motivo, ref_doc, nota):
    """Guarda Excel y bitácora como una sola operación: si una falla, la otra se revierte."""
    registro = inv.por_fila(fila)
    try:
        inv.guardar()
    except PermissionError:
        return None, (409, {"error": "el archivo está abierto en Excel; ciérralo y vuelve a intentar"})
    except OSError as e:
        inv.asignar(fila, anterior)
        return None, (409, {"error": f"no se pudo guardar el Excel: {e}"})

    try:
        ts, mov_id = bitacora.registrar(estado.con, accion, registro, usuario,
                                        destino=nueva, origen=anterior, nota=nota,
                                        motivo=motivo, ref_doc=ref_doc)
    except Exception as e:
        inv.asignar(fila, anterior)
        inv.guardar()
        return None, (500, {"error": f"no se pudo registrar en la bitácora, cambio revertido: {e}"})
    return {"ts": ts, "id": mov_id}, None


def descargar(estado, tipo):
    """Genera el Excel pedido y devuelve (ruta, nombre_sugerido)."""
    inv = estado.inventario()
    if tipo == "inventario":
        ruta = reportes.exportar_inventario(inv)
    elif tipo == "f01":
        ruta = reportes.generar_f01(inv, version="2")
    elif tipo == "f01v3":
        ruta = reportes.generar_f01(inv, version="3")
    elif tipo == "tabla_macro":
        ruta = reportes.tabla_macro(inv)
    elif tipo == "bitacora":
        ruta = reportes.exportar_bitacora(estado.con)
    else:
        raise KeyError(tipo)
    time.sleep(0)  # deja respirar al lock
    return ruta
