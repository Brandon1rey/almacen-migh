"""API JSON con usuarios, sesiones, movimientos y deshacer."""
import time

from app import bitacora, reportes, ubicaciones, usuarios

MOTIVOS = {
    "reacomodo": "Reacomodo por espacio",
    "surtido": "Surtido parcial / consumo",
    "devolucion": "Devolución de restos al almacén",
    "liberacion": "Liberación de rack o pasillo",
    "cuarentena": "Envío a cuarentena o retención",
    "otro": "Otro (explicar en la nota)",
}
COOKIE = "sesion"
HORAS_SESION = 12
REVERSIBLES = ("UBICAR", "MOVER", "PARCIAL", "CONFIRMAR", "NO_ENCONTRADO", "LIMPIAR")


def _r(codigo, cuerpo, cookies=None):
    return codigo, cuerpo, {"cookies": cookies or []}


def manejar(estado, metodo, ruta, params, cuerpo, sesion=None, galleta=None):
    inv = estado.inventario()
    cuerpo = cuerpo or {}

    # ---------- público ----------
    if metodo == "POST" and ruta == "/api/login":
        return _login(estado, cuerpo, params)
    if metodo == "POST" and ruta == "/api/logout":
        bitacora.cerrar_sesion(estado.con, galleta)
        return _r(200, {"ok": True}, [f"{COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax"])
    if metodo == "GET" and ruta == "/api/yo":
        if not sesion:
            return _r(401, {"autenticado": False})
        return _r(200, {"autenticado": True, "usuario": _publico(sesion),
                        "archivo": getattr(estado, "archivo", ""),
                        "demo": bool(getattr(estado, "demo", False))})

    # ---------- de aquí en adelante hace falta sesión ----------
    if not sesion:
        return _r(401, {"error": "sesión expirada o inexistente; vuelve a entrar"})

    usuario = sesion["nombre"]
    uid = sesion["usuario_id"]
    rol = sesion["rol"]

    if metodo == "POST" and ruta == "/api/password":
        return _cambiar_password(estado, cuerpo, uid)

    if ruta.startswith("/api/usuarios"):
        if rol != "admin":
            return _r(403, {"error": "solo un administrador puede gestionar usuarios"})
        return _usuarios(estado, metodo, ruta, cuerpo, usuario)

    # ---------- lecturas ----------
    if metodo == "GET" and ruta == "/api/inventario":
        filas = inv.buscar(
            texto=params.get("q", ""),
            solo_sin_ubicar=params.get("sin_ubicar") in ("1", "true", "si"),
            solo_vencidos=params.get("vencidos") in ("1", "true", "si"),
            incluir_cero=params.get("cero") in ("1", "true", "si"),
        )
        return _r(200, {"filas": filas, "resumen": inv.resumen(), "total_filtrado": len(filas)})

    if metodo == "GET" and ruta == "/api/lote":
        try:
            return _r(200, {"lote": inv.por_fila(int(params.get("fila", 0)))})
        except (KeyError, TypeError, ValueError) as e:
            return _r(400, {"error": str(e)})

    if metodo == "GET" and ruta == "/api/historial":
        try:
            fila = int(params.get("fila", 0))
        except (TypeError, ValueError):
            return _r(400, {"error": "fila inválida"})
        return _r(200, {"movimientos": _con_permisos(
            bitacora.por_fila(estado.con, fila, int(params.get("limite", 20))), sesion)})

    if metodo == "GET" and ruta == "/api/resumen":
        return _r(200, inv.resumen())

    if metodo == "GET" and ruta == "/api/estado":
        return _r(200, {"archivo": getattr(estado, "archivo", ""),
                        "demo": bool(getattr(estado, "demo", False)),
                        "lotes": len(inv.filas),
                        "movimientos": bitacora.contar(estado.con),
                        "usuarios": len(usuarios.listar(estado.con))})

    if metodo == "GET" and ruta == "/api/ubicaciones":
        return _r(200, {"codigos": ubicaciones.CODIGOS,
                        "zonas": {z: len(c) for z, c in ubicaciones.por_zona().items()},
                        "motivos": MOTIVOS})

    if metodo == "GET" and ruta == "/api/bitacora":
        limite = int(params.get("limite", 60))
        return _r(200, {"movimientos": _con_permisos(bitacora.ultimos(estado.con, limite), sesion),
                        "puede_purgar": rol == "admin"})

    # ---------- escrituras ----------
    if metodo == "POST" and rol == "consulta":
        return _r(403, {"error": "tu usuario es de consulta: no puede registrar movimientos"})

    if metodo == "POST" and ruta == "/api/mover":
        return _traslado(estado, inv, cuerpo, usuario, uid, rol)

    if metodo == "POST" and ruta in ("/api/ubicar", "/api/limpiar", "/api/no_encontrado"):
        return _movimiento(estado, inv, ruta, cuerpo, usuario, uid)

    if metodo == "POST" and ruta == "/api/deshacer":
        return _deshacer(estado, inv, cuerpo, sesion)

    if metodo == "POST" and ruta == "/api/purgar":
        return _purgar(estado, inv, cuerpo, sesion)

    return _r(404, {"error": f"ruta no encontrada: {metodo} {ruta}"})


# ---------------------------------------------------------------- sesión
def _publico(sesion):
    return {"id": sesion["usuario_id"], "nombre": sesion["nombre"], "rol": sesion["rol"],
            "debe_cambiar": bool(sesion.get("debe_cambiar"))}


def _login(estado, cuerpo, params):
    nombre = cuerpo.get("usuario") or params.get("usuario") or ""
    password = cuerpo.get("password") or ""
    fila = usuarios.verificar(estado.con, nombre, password)
    if fila is None:
        bitacora.auditar(estado.con, nombre or "(vacío)", "LOGIN_FALLIDO", "credenciales inválidas")
        return _r(401, {"error": "usuario o contraseña incorrectos"})
    token, expira = bitacora.abrir_sesion(estado.con, fila["id"], HORAS_SESION)
    bitacora.auditar(estado.con, fila["nombre"], "LOGIN", f"sesión hasta {expira}")
    return _r(200, {"usuario": {"id": fila["id"], "nombre": fila["nombre"],
                                "rol": fila["rol"], "debe_cambiar": bool(fila["debe_cambiar"])},
                    "expira": expira},
              [f"{COOKIE}={token}; Path=/; Max-Age={HORAS_SESION * 3600}; HttpOnly; SameSite=Lax"])


def _cambiar_password(estado, cuerpo, uid):
    actual = cuerpo.get("actual") or ""
    nueva = cuerpo.get("nueva") or ""
    fila = estado.con.execute("SELECT * FROM usuario WHERE id=?", (uid,)).fetchone()
    if fila is None or usuarios.verificar(estado.con, fila["nombre"], actual) is None:
        return _r(400, {"error": "la contraseña actual no es correcta"})
    try:
        usuarios.cambiar_password(estado.con, uid, nueva, debe_cambiar=False)
    except ValueError as e:
        return _r(400, {"error": str(e)})
    bitacora.auditar(estado.con, fila["nombre"], "CAMBIO_PASSWORD", "ok")
    return _r(200, {"ok": True, "mensaje": "contraseña actualizada; vuelve a entrar"},
              [f"{COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax"])


# ---------------------------------------------------------------- usuarios
def _usuarios(estado, metodo, ruta, cuerpo, actor):
    partes = [p for p in ruta.split("/") if p]
    if metodo == "GET":
        return _r(200, {"usuarios": usuarios.listar(estado.con), "roles": list(bitacora.ROLES)})
    if metodo == "POST":
        try:
            usuarios.crear(estado.con, cuerpo.get("nombre", ""), cuerpo.get("password", ""),
                           cuerpo.get("rol", "almacen"), bool(cuerpo.get("debe_cambiar")))
        except ValueError as e:
            return _r(400, {"error": str(e)})
        bitacora.auditar(estado.con, actor, "USUARIO_CREADO",
                         {"nombre": cuerpo.get("nombre"), "rol": cuerpo.get("rol")})
        return _r(200, {"usuarios": usuarios.listar(estado.con)})
    if metodo == "PATCH" and len(partes) >= 3:
        try:
            uid = int(partes[2])
            if "rol" in cuerpo:
                usuarios.cambiar_rol(estado.con, uid, cuerpo["rol"])
            if "activo" in cuerpo:
                usuarios.activar(estado.con, uid, bool(cuerpo["activo"]))
            if cuerpo.get("password"):
                usuarios.cambiar_password(estado.con, uid, cuerpo["password"], debe_cambiar=True)
        except (ValueError, TypeError) as e:
            return _r(400, {"error": str(e)})
        bitacora.auditar(estado.con, actor, "USUARIO_EDITADO", {"id": partes[2], "cambios": cuerpo})
        return _r(200, {"usuarios": usuarios.listar(estado.con)})
    return _r(404, {"error": "ruta de usuarios no encontrada"})


# ---------------------------------------------------------------- movimientos
def _con_permisos(movimientos, sesion):
    admin = sesion["rol"] == "admin"
    salida = []
    for m in movimientos:
        propio = m.get("usuario_id") == sesion["usuario_id"]
        deshacible = (not m.get("revertido") and m.get("accion") in REVERSIBLES and
                      (admin or (propio and sesion["rol"] == "almacen")))
        salida.append({**m, "puede_deshacer": deshacible, "puede_purgar": admin})
    return salida


def _traslado(estado, inv, cuerpo, usuario, uid, rol):
    try:
        fila = int(cuerpo["fila"])
        registro = inv.por_fila(fila)
    except (KeyError, TypeError, ValueError):
        return _r(400, {"error": "fila requerida"})

    destino = ubicaciones.normalizar(cuerpo.get("destino") or cuerpo.get("ubicacion"))
    motivo = (cuerpo.get("motivo") or "").strip().lower()
    if motivo not in MOTIVOS:
        return _r(400, {"error": f"motivo inválido; usa uno de: {', '.join(MOTIVOS)}"})
    if not destino:
        return _r(400, {"error": "destino requerido"})
    if not registro["ubicacion"]:
        return _r(400, {"error": "el lote no tiene ubicación actual: primero asígnalo"})
    if destino == registro["ubicacion"]:
        return _r(400, {"error": f"el destino es la ubicación actual ({destino}); no hay movimiento"})

    existencia = float(registro["cantidad"] or 0)
    crudo = cuerpo.get("cantidad")
    parcial = False
    if crudo not in (None, "", 0, "0"):
        try:
            cantidad = round(float(str(crudo).replace(",", ".")), 3)
        except ValueError:
            return _r(400, {"error": f"cantidad inválida: {crudo}"})
        if cantidad <= 0:
            return _r(400, {"error": "la cantidad a mover debe ser mayor que cero"})
        if cantidad > existencia:
            return _r(400, {"error": f"no puedes mover {cantidad}; hay {existencia} "
                                     f"{registro['unidad']} en ese renglón"})
        parcial = cantidad < existencia
    else:
        cantidad = existencia

    anterior = registro["ubicacion"]
    nota = cuerpo.get("nota")
    nueva_fila = None

    if parcial:
        try:
            nuevo = inv.agregar_parcialidad(fila, cantidad, destino,
                                            nota=f"parcialidad de la fila {fila}")
        except ValueError as e:
            return _r(400, {"error": str(e)})
        nueva_fila = nuevo["fila"]
        accion = "PARCIAL"
        al_fallar = lambda: (inv.ajustar_cantidad(fila, existencia),  # noqa: E731
                             inv.anular(nueva_fila, "anulada: no se pudo registrar"),
                             inv.guardar())
    else:
        try:
            inv.asignar(fila, destino)
        except ValueError as e:
            return _r(400, {"error": str(e)})
        accion = "MOVER"
        al_fallar = lambda: (inv.asignar(fila, anterior), inv.guardar())  # noqa: E731

    guardado, error = _persistir(estado, inv, fila, accion, usuario, uid, anterior, destino,
                                 motivo, cuerpo.get("ref_doc"), nota, cantidad=cantidad,
                                 fila_destino=nueva_fila, al_fallar=al_fallar)
    if error:
        return error
    return _r(200, {"accion": accion, "fila": fila, "anterior": anterior, "nueva": destino,
                    "cantidad": cantidad, "unidad": registro["unidad"],
                    "fila_destino": nueva_fila, "motivo": motivo,
                    "ts": guardado["ts"], "movimiento": guardado["id"],
                    "resumen": inv.resumen(), "bitacora": bitacora.contar(estado.con),
                    "historial": _con_permisos(bitacora.por_fila(estado.con, fila, 20),
                                               {"usuario_id": uid, "rol": rol})})


def _movimiento(estado, inv, ruta, cuerpo, usuario, uid):
    try:
        fila = int(cuerpo["fila"])
    except (KeyError, TypeError, ValueError):
        return _r(400, {"error": "fila requerida"})
    try:
        inv.por_fila(fila)
    except KeyError as e:
        return _r(400, {"error": str(e)})

    try:
        if ruta == "/api/ubicar":
            anterior, nueva = inv.asignar(fila, ubicaciones.normalizar(cuerpo.get("ubicacion")))
            accion = "UBICAR" if not anterior else "MOVER"
            if anterior == nueva:
                accion = "CONFIRMAR"
        elif ruta == "/api/limpiar":
            anterior, nueva = inv.asignar(fila, None)
            accion = "LIMPIAR"
        else:
            anterior, nueva = inv.asignar(fila, None)
            accion = "NO_ENCONTRADO"
    except ValueError as e:
        return _r(400, {"error": str(e)})

    guardado, error = _persistir(estado, inv, fila, accion, usuario, uid, anterior, nueva,
                                 None, None, cuerpo.get("nota"),
                                 al_fallar=lambda: (inv.asignar(fila, anterior), inv.guardar()))
    if error:
        return error
    return _r(200, {"accion": accion, "fila": fila, "anterior": anterior, "nueva": nueva,
                    "ts": guardado["ts"], "movimiento": guardado["id"],
                    "resumen": inv.resumen(), "bitacora": bitacora.contar(estado.con)})


def _deshacer(estado, inv, cuerpo, sesion):
    mov = bitacora.por_id(estado.con, cuerpo.get("movimiento", 0))
    if mov is None:
        return _r(404, {"error": "ese movimiento no existe"})
    if mov["revertido"]:
        return _r(400, {"error": "ese movimiento ya fue deshecho"})
    if mov["accion"] not in REVERSIBLES:
        return _r(400, {"error": f"el movimiento {mov['accion']} no se puede deshacer"})
    admin = sesion["rol"] == "admin"
    if not admin and (mov["usuario_id"] != sesion["usuario_id"] or sesion["rol"] != "almacen"):
        return _r(403, {"error": "solo puedes deshacer tus propios movimientos"})

    ultimo = bitacora.ultimo_vigente(estado.con, mov["fila"])
    if ultimo is None or ultimo["id"] != mov["id"]:
        return _r(409, {"error": "hay movimientos posteriores de ese lote; deshaz primero el último"})

    try:
        registro = inv.por_fila(mov["fila"])
    except KeyError:
        return _r(400, {"error": "el renglón ya no existe en el inventario"})

    if mov["accion"] == "PARCIAL":
        return _deshacer_parcial(estado, inv, mov, registro, sesion)

    destino = mov["ubicacion_origen"]
    if destino and not ubicaciones.existe(destino):
        return _r(409, {"error": f"el origen {destino} ya no está en el catálogo de ubicaciones"})

    actual = registro["ubicacion"]
    if mov["ubicacion_destino"] and actual != mov["ubicacion_destino"]:
        return _r(409, {"error": f"el lote está en {actual or 'sin ubicación'}, no donde lo dejó "
                                 f"ese movimiento ({mov['ubicacion_destino']})"})
    try:
        inv.asignar(mov["fila"], destino)
    except ValueError as e:
        return _r(400, {"error": str(e)})

    guardado, error = _persistir(estado, inv, mov["fila"], "DESHACER", sesion["nombre"],
                                 sesion["usuario_id"], actual, destino, None, None,
                                 f"deshace el movimiento #{mov['id']}",
                                 cantidad=mov.get("cantidad"),
                                 al_fallar=lambda: (inv.asignar(mov["fila"], actual),
                                                    inv.guardar()))
    if error:
        return error
    bitacora.marcar_revertido(estado.con, mov["id"], guardado["id"])
    bitacora.auditar(estado.con, sesion["nombre"], "DESHACER",
                     {"movimiento": mov["id"], "de": actual, "a": destino})
    return _r(200, {"deshecho": mov["id"], "movimiento": guardado["id"], "fila": mov["fila"],
                    "anterior": actual, "nueva": destino, "resumen": inv.resumen(),
                    "historial": _con_permisos(bitacora.por_fila(estado.con, mov["fila"], 20),
                                               sesion)})


def _deshacer_parcial(estado, inv, mov, registro, sesion):
    """Deshace una parcialidad: devuelve los kilos al renglón original y anula la copia."""
    nueva_fila = mov.get("fila_destino")
    cantidad = float(mov.get("cantidad") or 0)
    if not nueva_fila or cantidad <= 0:
        return _r(409, {"error": "ese movimiento no tiene la parcialidad asociada"})
    try:
        parcial = inv.por_fila(nueva_fila)
    except KeyError:
        return _r(409, {"error": "el renglón de la parcialidad ya no existe"})

    if (parcial["ubicacion"] or "") != (mov["ubicacion_destino"] or ""):
        return _r(409, {"error": f"la parcialidad ya se movió a {parcial['ubicacion'] or 'sin ubicación'}; "
                                 f"no se puede deshacer este movimiento"})
    if abs(float(parcial["cantidad"] or 0) - cantidad) > 0.001:
        return _r(409, {"error": f"la parcialidad ya cambió de cantidad "
                                 f"({parcial['cantidad']} {parcial['unidad']}); no se puede deshacer"})

    actual_origen = float(registro["cantidad"] or 0)
    ubicacion_origen = registro["ubicacion"]
    try:
        inv.ajustar_cantidad(mov["fila"], round(actual_origen + cantidad, 3))
        inv.anular(nueva_fila, f"anulada por deshacer #{mov['id']}")
    except ValueError as e:
        return _r(400, {"error": str(e)})

    def revertir():
        inv.ajustar_cantidad(mov["fila"], actual_origen)
        inv.asignar(nueva_fila, mov["ubicacion_destino"])
        inv.ajustar_cantidad(nueva_fila, cantidad)
        inv.guardar()

    guardado, error = _persistir(estado, inv, mov["fila"], "DESHACER", sesion["nombre"],
                                 sesion["usuario_id"], mov["ubicacion_destino"], ubicacion_origen,
                                 None, None,
                                 f"deshace la parcialidad #{mov['id']} (fila {nueva_fila})",
                                 cantidad=cantidad, al_fallar=revertir)
    if error:
        return error
    bitacora.marcar_revertido(estado.con, mov["id"], guardado["id"])
    bitacora.auditar(estado.con, sesion["nombre"], "DESHACER_PARCIAL",
                     {"movimiento": mov["id"], "kg": cantidad, "fila_anulada": nueva_fila})
    return _r(200, {"deshecho": mov["id"], "movimiento": guardado["id"], "fila": mov["fila"],
                    "anterior": mov["ubicacion_destino"], "nueva": ubicacion_origen,
                    "cantidad": cantidad, "fila_anulada": nueva_fila,
                    "resumen": inv.resumen(),
                    "historial": _con_permisos(bitacora.por_fila(estado.con, mov["fila"], 20),
                                               sesion)})


def _purgar(estado, inv, cuerpo, sesion):
    """Borrado físico (solo admin): revierte el efecto, borra el renglón y deja auditoría."""
    if sesion["rol"] != "admin":
        return _r(403, {"error": "solo un administrador puede borrar movimientos"})
    mov = bitacora.por_id(estado.con, cuerpo.get("movimiento", 0))
    if mov is None:
        return _r(404, {"error": "ese movimiento no existe"})
    motivo = (cuerpo.get("motivo") or "").strip()
    if len(motivo) < 5:
        return _r(400, {"error": "escribe el motivo del borrado (mínimo 5 caracteres)"})

    ultimo = bitacora.ultimo_vigente(estado.con, mov["fila"])
    if ultimo and ultimo["id"] != mov["id"]:
        return _r(409, {"error": "hay movimientos posteriores de ese lote; deshaz primero el último"})

    if not mov["revertido"] and mov["accion"] in REVERSIBLES:
        try:
            inv.asignar(mov["fila"], mov["ubicacion_origen"])
            inv.guardar()
        except (KeyError, ValueError, OSError) as e:
            return _r(409, {"error": f"no se pudo revertir el efecto en el Excel: {e}"})

    borrado = bitacora.borrar(estado.con, mov["id"])
    bitacora.auditar(estado.con, sesion["nombre"], "PURGA_MOVIMIENTO",
                     {"motivo": motivo, "movimiento": borrado})
    return _r(200, {"borrado": mov["id"], "motivo": motivo, "resumen": inv.resumen(),
                    "historial": _con_permisos(bitacora.por_fila(estado.con, mov["fila"], 20),
                                               sesion)})


def _persistir(estado, inv, fila, accion, usuario, uid, anterior, nueva, motivo, ref_doc, nota,
               cantidad=None, fila_destino=None, al_fallar=None):
    """Guarda Excel y bitácora como una sola operación: si una falla, la otra se revierte."""
    registro = inv.por_fila(fila)

    def revertir():
        if al_fallar is None:
            return
        try:
            al_fallar()
        except Exception:  # noqa: BLE001 - si tampoco se puede revertir, se reporta abajo
            pass

    try:
        inv.guardar()
    except PermissionError:
        revertir()
        return None, _r(409, {"error": "el archivo está abierto en Excel; ciérralo y reintenta"})
    except OSError as e:
        revertir()
        return None, _r(409, {"error": f"no se pudo guardar el Excel: {e}"})

    try:
        ts, mov_id = bitacora.registrar(estado.con, accion, registro, usuario,
                                        destino=nueva, origen=anterior, nota=nota,
                                        motivo=motivo, ref_doc=ref_doc, usuario_id=uid,
                                        cantidad=cantidad, fila_destino=fila_destino)
    except Exception as e:
        revertir()
        return None, _r(500, {"error": f"no se pudo registrar en la bitácora, cambio "
                                       f"revertido: {e}"})
    return {"ts": ts, "id": mov_id}, None


def descargar(estado, tipo):
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
    time.sleep(0)
    return ruta
