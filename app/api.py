"""API JSON del inventario. Sin sockets: recibe el Inventario ya cargado."""
import time

from app import bitacora, reportes, ubicaciones


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

    if metodo == "GET" and ruta == "/api/resumen":
        return 200, inv.resumen()

    if metodo == "GET" and ruta == "/api/ubicaciones":
        return 200, {"codigos": ubicaciones.CODIGOS,
                     "zonas": {z: len(c) for z, c in ubicaciones.por_zona().items()}}

    if metodo == "GET" and ruta == "/api/bitacora":
        return 200, {"movimientos": bitacora.ultimos(estado.con, int(params.get("limite", 60)))}

    if metodo == "POST" and ruta in ("/api/ubicar", "/api/limpiar", "/api/no_encontrado"):
        return _movimiento(estado, inv, ruta, cuerpo, usuario)

    return 404, {"error": f"ruta no encontrada: {metodo} {ruta}"}


def _movimiento(estado, inv, ruta, cuerpo, usuario):
    if not usuario:
        return 400, {"error": "falta tu nombre: la bitácora es el rastro de auditoría"}
    try:
        fila = int(cuerpo["fila"])
    except (KeyError, TypeError, ValueError):
        return 400, {"error": "fila requerida"}
    try:
        registro = inv.por_fila(fila)
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

    try:
        inv.guardar()
    except PermissionError:
        return 409, {"error": "el archivo está abierto en Excel; ciérralo y vuelve a intentar"}
    except OSError as e:
        return 409, {"error": f"no se pudo guardar el Excel: {e}"}

    try:
        ts, mov_id = bitacora.registrar(estado.con, accion, registro, usuario, destino=nueva,
                                       origen=anterior, nota=cuerpo.get("nota"))
    except Exception as e:  # si la bitácora falla, el Excel no puede quedar cambiado
        inv.asignar(fila, anterior)
        inv.guardar()
        return 500, {"error": f"no se pudo registrar en la bitácora, cambio revertido: {e}"}

    return 200, {"accion": accion, "fila": fila, "anterior": anterior, "nueva": nueva,
                 "ts": ts, "movimiento": mov_id, "resumen": inv.resumen(),
                 "bitacora": bitacora.contar(estado.con)}


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
