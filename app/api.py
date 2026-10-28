"""API JSON con usuarios, sesiones, movimientos, alta, producción, catálogo y borrado.

Reglas de la casa (todas nacieron de un fallo real):

* Ninguna entrada del navegador puede producir un error 500. Todo parámetro se convierte
  con `_entero` / `_numero` y, si no se puede, se responde 400 con un mensaje claro.
* El inventario (Excel) y la bitácora (SQLite) se escriben como una sola operación:
  si falla una, se revierte la otra.
* Los movimientos de escritura pasan por `_ESCRITURA`: dos personas del almacén que mueven
  lotes a la vez no pueden corromper el libro de Excel.
* **Fuente única**: las ubicaciones salen del catálogo (`ubicaciones.canonico`) y los
  nombres de insumo, del catálogo del ERP (`catalogo.por_sku`). La búsqueda "parecida"
  solo ayuda a encontrar; nunca crea ni sustituye datos canónicos.
"""
import os
import shutil
import threading
import time

from app import (bitacora, catalogo, eventos, inventario, inventarios, logger,
                 reportes, ubicaciones, usuarios)

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
REVERSIBLES = ("ALTA", "UBICAR", "MOVER", "PARCIAL", "SALIDA", "CONFIRMAR", "NO_ENCONTRADO",
               "LIMPIAR")
PRODUCCION = "PRODUCCION"
MIN_NOTA_NO_ENCONTRADO = 3
MIN_MOTIVO_PURGA = 5

# Un solo escritor a la vez: el Excel es un archivo, no una base de datos con transacciones.
_ESCRITURA = threading.RLock()

# Freno sencillo a la adivinación de contraseñas (red local, sin dependencias externas).
_INTENTOS = {}
_INTENTOS_LOCK = threading.Lock()
MAX_FALLOS = 8
VENTANA_SEGUNDOS = 300


def _r(codigo, cuerpo, cookies=None):
    return codigo, cuerpo, {"cookies": cookies or []}


def _entero(valor, defecto=0, minimo=0, maximo=2**31 - 1):
    """Entero a prueba de basura: lo que no se puede convertir cae al valor por omisión."""
    try:
        numero = int(valor)
    except (TypeError, ValueError):
        return defecto
    return max(minimo, min(maximo, numero))


def _si(valor):
    return str(valor or "").strip().lower() in ("1", "true", "si", "sí", "yes")


def _numero(valor, total=None, unidad=None):
    """Número finito o None. Evita que un NaN del navegador llegue al Excel o al JSON."""
    return inventario.interpretar_cantidad(valor, total=total, unidad=unidad)


def _tipo(params=None, cuerpo=None):
    """Inventario pedido en la petición (mp por omisión). Falla claro si no existe."""
    valor = (cuerpo or {}).get("tipo") or (params or {}).get("tipo") or "mp"
    return inventarios.normalizar_tipo(valor)


def _estado_json(estado, inv=None):
    """Resumen de todos los inventarios cargados, para el selector de la interfaz."""
    salida = []
    en_db = bool(getattr(estado, "en_db", False))
    for ficha in inventarios.lista(estado.carpeta_datos, en_db=en_db):
        item = {k: ficha[k] for k in ("clave", "nombre", "corto", "reporte", "orden",
                                      "base", "unidad", "existe")}
        try:
            r = estado.inventario(ficha["clave"]).resumen()
            item.update({"cargado": True, "lotes": r["total"],
                         "con_ubicacion": r["con_ubicacion"], "sin_ubicacion": r["sin_ubicacion"],
                         "cobertura": r["cobertura"], "vencidos": r["semaforo"]["VENCIDO"],
                         "por_unidad": r["por_unidad"],
                         "sin_cantidad_legible": r["sin_cantidad_legible"],
                         "ubicaciones_invalidas": r["ubicaciones_invalidas"]})
        except (FileNotFoundError, KeyError, ValueError, PermissionError):
            item.update({"cargado": False, "lotes": 0, "con_ubicacion": 0, "sin_ubicacion": 0,
                         "cobertura": 0.0, "vencidos": 0, "por_unidad": {},
                         "sin_cantidad_legible": 0, "ubicaciones_invalidas": 0})
        salida.append(item)
    return salida


def manejar(estado, metodo, ruta, params, cuerpo, sesion=None, galleta=None):
    params = params or {}
    cuerpo = cuerpo if isinstance(cuerpo, dict) else {}

    # ---------- público ----------
    if metodo == "POST" and ruta == "/api/login":
        return _login(estado, cuerpo, params)
    if metodo == "POST" and ruta == "/api/logout":
        bitacora.cerrar_sesion(estado.con, galleta)
        eventos.publicar("sesion.cerrada",
                         usuario=sesion["nombre"] if sesion else None)
        return _r(200, {"ok": True}, [f"{COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax"])
    if metodo == "GET" and ruta == "/api/yo":
        # 200 también sin sesión: «¿quién soy?» sin sesión es una respuesta válida, no un error.
        if not sesion:
            return _r(200, {"autenticado": False,
                            "demo": bool(getattr(estado, "demo", False))})
        return _r(200, {"autenticado": True, "usuario": _publico(sesion),
                        "archivo": getattr(estado, "archivo", ""),
                        "demo": bool(getattr(estado, "demo", False)),
                        "horas_sesion": HORAS_SESION})

    # ---------- de aquí en adelante hace falta sesión ----------
    if not sesion:
        return _r(401, {"error": "sesión expirada o inexistente; vuelve a entrar"})

    # Depuración: errores del navegador y estado del sistema (ver data\\logs\\depuracion.log).
    if metodo == "POST" and ruta == "/api/log":
        return _log_navegador(cuerpo, sesion)
    if metodo == "GET" and ruta == "/api/eventos":
        return _eventos_recientes(params, sesion)
    if metodo == "GET" and ruta == "/api/diagnostico":
        return _diagnostico(estado, sesion)

    try:
        tipo = _tipo(params, cuerpo)
    except ValueError as e:
        return _r(400, {"error": str(e)})

    if metodo == "GET" and ruta == "/api/inventarios":
        return _r(200, {"inventarios": _estado_json(estado),
                        "actual": tipo,
                        "demo": bool(getattr(estado, "demo", False))})

    usuario = sesion["nombre"]
    uid = sesion["usuario_id"]
    rol = sesion["rol"]

    # La contraseña no depende de que el Excel esté cargado: va antes de resolverlo.
    if metodo == "POST" and ruta == "/api/password":
        return _cambiar_password(estado, cuerpo, uid)

    if ruta.startswith("/api/usuarios"):
        if rol != "admin":
            return _r(403, {"error": "solo un administrador puede gestionar usuarios"})
        return _usuarios(estado, metodo, ruta, cuerpo, usuario)

    try:
        inv = estado.inventario(tipo)
    except (FileNotFoundError, KeyError):
        return _r(404, {"error": f"el inventario {tipo.upper()} no está cargado todavía"})
    except PermissionError as e:
        return _r(409, {"error": f"no se pudo abrir el Excel del inventario: {e}"})
    except ValueError as e:
        return _r(500, {"error": f"el Excel del inventario {tipo.upper()} no tiene la forma "
                                  f"esperada: {e}"})

    # ---------- lecturas ----------
    if metodo == "GET" and ruta == "/api/inventario":
        texto = str(params.get("q", "") or "")
        filtros = dict(
            texto=texto,
            solo_sin_ubicar=_si(params.get("sin_ubicar")),
            solo_vencidos=_si(params.get("vencidos")),
            incluir_cero=_si(params.get("cero")),
            solo_ilegibles=_si(params.get("ilegibles")),
        )
        filas = inv.buscar(**filtros)
        difuso = False
        sugerencias = []
        if not filas and texto.strip():
            # búsqueda tolerante a errores de dedo («protiena» -> «Proteína»)
            filas = inv.buscar(difuso=True, limite=40, **filtros)
            difuso = bool(filas)
            if filas:
                vistas = []
                for f in filas:
                    if f["descripcion"] not in vistas:
                        vistas.append(f["descripcion"])
                sugerencias = vistas[:3]
            else:
                sugerencias = inv.sugerencias(texto)
        return _r(200, {"filas": filas, "resumen": inv.resumen(), "tipo": tipo,
                        "total_filtrado": len(filas), "difuso": difuso,
                        "sugerencias": sugerencias, "consulta": texto})

    if metodo == "GET" and ruta == "/api/lote":
        try:
            return _r(200, {"lote": inv.por_fila(_entero(params.get("fila"), -1)),
                            "tipo": tipo})
        except (KeyError, TypeError, ValueError) as e:
            return _r(400, {"error": f"renglón no encontrado: {e}"})

    if metodo == "GET" and ruta == "/api/historial":
        fila = _entero(params.get("fila"), -1)
        if fila < 0:
            return _r(400, {"error": "fila inválida"})
        return _r(200, {"movimientos": _con_permisos(
            bitacora.por_fila(estado.con, fila, _entero(params.get("limite"), 20, 1, 200),
                              inventario=tipo), sesion)})

    if metodo == "GET" and ruta == "/api/resumen":
        return _r(200, inv.resumen())

    if metodo == "GET" and ruta == "/api/estado":
        return _r(200, {"archivo": getattr(estado, "archivo", ""),
                        "tipo": tipo, "inventario": inv.nombre,
                        "demo": bool(getattr(estado, "demo", False)),
                        "lotes": len(inv.filas),
                        "inventarios": _estado_json(estado, inv),
                        "movimientos": bitacora.contar(estado.con),
                        "usuarios": len(usuarios.listar(estado.con)),
                        "rol": rol, "horas_sesion": HORAS_SESION,
                        "version": _version()})

    if metodo == "GET" and ruta == "/api/copia":
        return _copia_estado(estado, params)

    if metodo == "POST" and ruta == "/api/copia":
        if rol == "consulta":
            return _r(403, {"error": "tu usuario es de consulta: no puede copiar a la USB"})
        return _copia_ahora(estado, cuerpo, usuario)

    if metodo == "GET" and ruta == "/api/ubicaciones":
        return _ubicaciones(estado, params)

    if metodo == "GET" and ruta == "/api/catalogo":
        return _catalogo(estado, params)

    if metodo == "GET" and ruta == "/api/qr":
        return _qr_por_codigo(estado, params)

    if metodo == "GET" and ruta == "/api/bitacora":
        crudo_filtro = params.get("inventario", params.get("tipo"))
        filtro = None
        if crudo_filtro not in (None, "") and str(crudo_filtro).strip().lower() not in (
                "todos", "all", "*"):
            try:
                filtro = inventarios.normalizar_tipo(crudo_filtro)
            except ValueError as e:
                return _r(400, {"error": str(e)})
        return _r(200, {"movimientos": _con_permisos(
                            bitacora.ultimos(estado.con, _entero(params.get("limite"), 60, 1, 500),
                                             inventario=filtro), sesion),
                        "puede_purgar": rol == "admin",
                        "tipo": filtro})

    # ---------- escrituras ----------
    if metodo == "POST" and rol == "consulta":
        return _r(403, {"error": "tu usuario es de consulta: puede ver y descargar, "
                                 "pero no registrar movimientos"})

    with _ESCRITURA:
        if metodo == "POST" and ruta == "/api/alta":
            return _alta(estado, inv, cuerpo, usuario, uid)
        if metodo == "POST" and ruta == "/api/asignar_qr":
            return _asignar_qr(estado, cuerpo, usuario, uid)
        if metodo == "POST" and ruta == "/api/mover":
            return _traslado(estado, inv, cuerpo, usuario, uid, rol)
        if metodo == "POST" and ruta in ("/api/ubicar", "/api/limpiar", "/api/no_encontrado"):
            return _movimiento(estado, inv, ruta, cuerpo, usuario, uid)
        if metodo == "POST" and ruta == "/api/deshacer":
            return _deshacer(estado, cuerpo, sesion)
        if metodo == "POST" and ruta == "/api/purgar":
            return _purgar(estado, cuerpo, sesion)
        if metodo == "POST" and ruta == "/api/borrar-renglones":
            return _borrar_renglones(estado, cuerpo, sesion)

    return _r(404, {"error": f"ruta no encontrada: {metodo} {ruta}"})


def _version():
    try:
        from app import version
        return version.VERSION
    except Exception:  # noqa: BLE001
        return ""


# ---------------------------------------------------------------- sesión
def _publico(sesion):
    return {"id": sesion["usuario_id"], "nombre": sesion["nombre"], "rol": sesion["rol"],
            "debe_cambiar": bool(sesion.get("debe_cambiar"))}


def _clave_intentos(nombre):
    return (nombre or "").strip().lower()


def _espera_por_fallos(nombre):
    """Segundos que faltan para poder reintentar, o 0 si puede intentar ya."""
    clave = _clave_intentos(nombre)
    with _INTENTOS_LOCK:
        registro = _INTENTOS.get(clave)
        if not registro:
            return 0
        fallos, ultimo = registro
        transcurrido = time.monotonic() - ultimo
        if transcurrido > VENTANA_SEGUNDOS:
            _INTENTOS.pop(clave, None)
            return 0
        if fallos < MAX_FALLOS:
            return 0
        return int(VENTANA_SEGUNDOS - transcurrido) + 1


def _anotar_fallo(nombre):
    clave = _clave_intentos(nombre)
    with _INTENTOS_LOCK:
        if len(_INTENTOS) > 500:      # la lista no crece sin límite
            _INTENTOS.clear()
        fallos, _ = _INTENTOS.get(clave, (0, 0))
        _INTENTOS[clave] = (fallos + 1, time.monotonic())


def _limpiar_fallos(nombre):
    with _INTENTOS_LOCK:
        _INTENTOS.pop(_clave_intentos(nombre), None)


def _login(estado, cuerpo, params):
    nombre = str(cuerpo.get("usuario") or params.get("usuario") or "").strip()
    password = cuerpo.get("password") or ""
    if not nombre or not password:
        return _r(400, {"error": "escribe tu usuario y tu contraseña"})
    espera = _espera_por_fallos(nombre)
    if espera:
        bitacora.auditar(estado.con, nombre, "LOGIN_BLOQUEADO", f"espera {espera} s")
        return _r(429, {"error": f"demasiados intentos fallidos; espera {espera} segundos "
                                 f"e inténtalo otra vez"})
    fila = usuarios.verificar(estado.con, nombre, password)
    if fila is None:
        _anotar_fallo(nombre)
        bitacora.auditar(estado.con, nombre or "(vacío)", "LOGIN_FALLIDO",
                         "credenciales inválidas")
        return _r(401, {"error": "usuario o contraseña incorrectos"})
    _limpiar_fallos(nombre)
    token, expira = bitacora.abrir_sesion(estado.con, fila["id"], HORAS_SESION)
    bitacora.auditar(estado.con, fila["nombre"], "LOGIN", f"sesión hasta {expira}")
    eventos.publicar("sesion.iniciada", usuario=fila["nombre"], rol=fila["rol"])
    return _r(200, {"usuario": {"id": fila["id"], "nombre": fila["nombre"],
                                "rol": fila["rol"], "debe_cambiar": bool(fila["debe_cambiar"])},
                    "expira": expira, "horas_sesion": HORAS_SESION},
              [f"{COOKIE}={token}; Path=/; Max-Age={HORAS_SESION * 3600}; "
               f"HttpOnly; SameSite=Lax"])


def _cambiar_password(estado, cuerpo, uid):
    actual = cuerpo.get("actual") or ""
    nueva = cuerpo.get("nueva") or ""
    fila = estado.con.execute("SELECT * FROM usuario WHERE id=?", (uid,)).fetchone()
    if fila is None:
        return _r(401, {"error": "tu usuario ya no existe"})
    if usuarios.verificar(estado.con, fila["nombre"], actual) is None:
        return _r(400, {"error": "la contraseña actual no es correcta"})
    if nueva and nueva == actual:
        return _r(400, {"error": "la contraseña nueva tiene que ser distinta de la actual"})
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
    if metodo == "POST" and len(partes) == 2:
        try:
            usuarios.crear(estado.con, cuerpo.get("nombre", ""), cuerpo.get("password", ""),
                           cuerpo.get("rol", "almacen"), bool(cuerpo.get("debe_cambiar")))
        except ValueError as e:
            return _r(400, {"error": str(e)})
        bitacora.auditar(estado.con, actor, "USUARIO_CREADO",
                         {"nombre": cuerpo.get("nombre"), "rol": cuerpo.get("rol")})
        return _r(200, {"usuarios": usuarios.listar(estado.con)})
    if metodo in ("PATCH", "POST") and len(partes) >= 3:
        uid = _entero(partes[2], 0)
        if uid <= 0:
            return _r(400, {"error": "id de usuario inválido"})
        if not usuarios.por_id(estado.con, uid):
            return _r(404, {"error": "ese usuario no existe"})
        try:
            if "rol" in cuerpo:
                usuarios.cambiar_rol(estado.con, uid, cuerpo["rol"])
            if "activo" in cuerpo:
                usuarios.activar(estado.con, uid, bool(cuerpo["activo"]))
            if cuerpo.get("password"):
                usuarios.cambiar_password(estado.con, uid, cuerpo["password"], debe_cambiar=True)
        except (ValueError, TypeError) as e:
            return _r(400, {"error": str(e)})
        bitacora.auditar(estado.con, actor, "USUARIO_EDITADO",
                         {"id": uid, "cambios": {k: v for k, v in cuerpo.items()
                                                 if k != "password"}})
        return _r(200, {"usuarios": usuarios.listar(estado.con)})
    return _r(404, {"error": "ruta de usuarios no encontrada"})


# ---------------------------------------------------------------- lecturas con reglas
def _ubicaciones(estado, params):
    """Catálogo, ocupación y sugerencias de dónde poner el material (solo códigos válidos)."""
    tipo = params.get("tipo") or "mp"
    try:
        inv = estado.inventario(tipo)
    except (ValueError, FileNotFoundError, KeyError):
        return _r(400, {"error": f"inventario no cargado: {tipo}"})
    ocupacion = {}
    for registro in inv.filas:
        codigo = registro.get("ubicacion")
        cantidad = inventario.numero(registro.get("cantidad"))
        if not codigo or not cantidad or cantidad <= 0:
            continue
        datos = ocupacion.setdefault(codigo, {"lotes": 0, "cantidad": 0.0,
                                              "unidad": registro.get("unidad"),
                                              "material": registro.get("descripcion")})
        datos["lotes"] += 1
        datos["cantidad"] = round(datos["cantidad"] + cantidad, 3)
    material = params.get("material") or None
    fila = params.get("fila")
    if fila and not material:
        try:
            material = inv.por_fila(_entero(fila, -1)).get("descripcion")
        except (KeyError, ValueError):
            material = None
    limite = _entero(params.get("limite"), 8, 1, 50)
    sugerencias = ubicaciones.sugerir(params.get("q", ""), inv.tipo, ocupacion, material, limite)
    return _r(200, {"codigos": ubicaciones.CODIGOS + [PRODUCCION],
                    "sugerencias": sugerencias,
                    "ocupacion": {k: v["lotes"] for k, v in ocupacion.items()},
                    "material": material,
                    "zonas": {z: len(c) for z, c in ubicaciones.por_zona().items()},
                    "motivos": MOTIVOS,
                    "salidas": {PRODUCCION: "Salida a producción (el material sale del almacén)"}})


def _catalogo(estado, params):
    """Catálogo del ERP: búsqueda exacta y, si no hay, parecidos (nunca inventa insumos)."""
    datos = catalogo.cargar(catalogo.ruta(getattr(estado, "carpeta_datos", None)),
                            con=getattr(estado, "con", None))
    try:
        tipo_inventario = params.get("inventario") or None
        if tipo_inventario:
            tipo_inventario = inventarios.normalizar_tipo(tipo_inventario)
    except ValueError as e:
        return _r(400, {"error": str(e)})
    texto = params.get("q", "")
    limite = _entero(params.get("limite"), 30, 1, 100)
    insumos = catalogo.buscar(datos, texto, limite, tipo_inventario)
    respuesta = {"insumos": insumos, "resumen": catalogo.resumen(datos),
                 "total": datos.get("total", 0), "consulta": texto}
    if not insumos and str(texto).strip():
        respuesta["sugerencias"] = catalogo.sugerencias(datos, texto, 5)
    return _r(200, respuesta)


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


def _existencia(registro):
    """Cantidad numérica del renglón, o None si la celda no es un número."""
    return inventario.numero(registro.get("cantidad"))


def _qr_ocupado(estado, codigo, tipo=None, fila=None):
    """Busca el código en los libros cargados. `tipo`/`fila` se excluyen (re-etiquetar)."""
    cargados = estado.cargados() if hasattr(estado, "cargados") else [("mp", estado.inventario("mp"))]
    for clave, inv in cargados:
        encontrado = inv.buscar_por_qr(codigo)
        if encontrado is None:
            continue
        if clave == tipo and encontrado["fila"] == fila:
            continue
        return {"tipo": clave, "fila": encontrado["fila"]}
    return None


def _qr_por_codigo(estado, params):
    """Escaneo de material: qué renglón activo tiene ese código."""
    codigo = inventario.normalizar_qr(params.get("codigo") or params.get("qr"))
    if not codigo:
        return _r(400, {"error": "escribe o escanea el código"})
    cargados = estado.cargados() if hasattr(estado, "cargados") else [("mp", estado.inventario("mp"))]
    for tipo, inv in cargados:
        encontrado = inv.buscar_por_qr(codigo)
        if encontrado:
            return _r(200, {"encontrado": True, "tipo": tipo, "fila": encontrado["fila"],
                            "renglon": encontrado})
    return _r(404, {"encontrado": False,
                    "error": f"el código «{codigo}» no está en ningún lote activo"})


def _asignar_qr(estado, cuerpo, usuario, uid):
    """Bautiza o re-etiqueta un renglón que ya existía (`qr_id` vacío = quitar etiqueta)."""
    try:
        tipo = _tipo(None, cuerpo)
        inv = estado.inventario(tipo)
    except (ValueError, FileNotFoundError, KeyError) as e:
        return _r(400, {"error": str(e)})
    try:
        fila = int(cuerpo["fila"])
        inv.por_fila(fila)
    except (KeyError, TypeError, ValueError):
        return _r(400, {"error": "fila requerida"})
    codigo = inventario.clave_qr(cuerpo.get("qr_id")) or None
    if codigo:
        ocupado = _qr_ocupado(estado, codigo, tipo, fila)
        if ocupado:
            return _r(409, {"error": f"el código «{codigo}» ya está en "
                                     f"{ocupado['tipo'].upper()} fila {ocupado['fila']}"})
    anterior, nuevo = inv.asignar_qr(fila, codigo)
    try:
        inv.guardar()
    except PermissionError:
        inv.asignar_qr(fila, anterior)
        return _r(409, {"error": "el archivo está abierto en Excel; ciérralo y reintenta"})
    except OSError as e:
        inv.asignar_qr(fila, anterior)
        return _r(409, {"error": f"no se pudo guardar el Excel: {e}"})
    bitacora.auditar(estado.con, usuario, "ASIGNAR_QR",
                     {"tipo": tipo, "fila": fila, "de": anterior, "a": nuevo})
    eventos.publicar("etiqueta.asignada", usuario=usuario, tipo=tipo, fila=fila,
                     de=anterior, a=nuevo)
    return _r(200, {"tipo": tipo, "fila": fila, "anterior": anterior, "qr_id": nuevo,
                    "renglon": inv.por_fila(fila)})


def _traslado(estado, inv, cuerpo, usuario, uid, rol):
    try:
        fila = int(cuerpo["fila"])
        registro = inv.por_fila(fila)
    except (KeyError, TypeError, ValueError):
        return _r(400, {"error": "fila requerida"})

    crudo_destino = cuerpo.get("destino") or cuerpo.get("ubicacion")
    destino = ubicaciones.canonico(crudo_destino)
    motivo = str(cuerpo.get("motivo") or "").strip().lower()
    if motivo not in MOTIVOS:
        return _r(400, {"error": f"motivo inválido; usa uno de: {', '.join(MOTIVOS)}"})
    if destino is None:
        if not str(crudo_destino or "").strip():
            return _r(400, {"error": "destino requerido"})
        if str(crudo_destino).strip().upper() == PRODUCCION:
            destino = PRODUCCION
        else:
            return _r(400, {"error": f"la ubicación «{str(crudo_destino).strip()}» no existe en el "
                                     f"catálogo ALM-PNO-002; revisa el código del rack"})
    if not registro["ubicacion"]:
        if registro.get("ubicacion_invalida"):
            return _r(400, {"error": f"la UBICACIÓN del renglón {fila} en el Excel "
                                     f"(«{registro['ubicacion_invalida']}») no existe en el "
                                     f"catálogo; asígnale un código válido antes de moverlo"})
        return _r(400, {"error": "el lote no tiene ubicación actual: primero asígnalo"})
    if destino == PRODUCCION:
        return _salida_produccion(estado, inv, registro, cuerpo, usuario, uid, motivo)
    if destino == registro["ubicacion"]:
        return _r(400, {"error": f"el destino es la ubicación actual ({destino}); no hay movimiento"})

    nota = str(cuerpo.get("nota") or "").strip() or None
    if motivo == "otro" and not nota:
        return _r(400, {"error": "con el motivo «otro» hace falta escribir la nota"})
    ref_doc = str(cuerpo.get("ref_doc") or "").strip() or None

    existencia = _existencia(registro)
    if existencia is None:
        return _r(400, {"error": f"la CANTIDAD del renglón {fila} no es un número; "
                                 f"corrígela en el Excel del área y vuelve a actualizar"})
    crudo = cuerpo.get("cantidad")
    parcial = False
    if crudo in (None, "", 0, "0"):
        cantidad = existencia
    else:
        cantidad = _numero(crudo, total=existencia, unidad=registro["unidad"])
        if cantidad is None:
            return _r(400, {"error": f"cantidad inválida: {crudo!r}; escribe un número, "
                                     f"por ejemplo 5 o 12.5 (la unidad de ese renglón "
                                     f"es {registro['unidad']})"})
        cantidad = round(cantidad, 3)
        if cantidad <= 0:
            return _r(400, {"error": "la cantidad a mover debe ser mayor que cero"})
        if cantidad > existencia:
            return _r(400, {"error": f"no puedes mover {cantidad}; hay {existencia} "
                                     f"{registro['unidad']} en ese renglón"})
        parcial = cantidad < existencia

    anterior = registro["ubicacion"]
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
                                 motivo, ref_doc, nota, cantidad=cantidad,
                                 fila_destino=nueva_fila, al_fallar=al_fallar)
    if error:
        return error
    return _r(200, {"accion": accion, "fila": fila, "anterior": anterior, "nueva": destino,
                    "cantidad": cantidad, "unidad": registro["unidad"],
                    "fila_destino": nueva_fila, "motivo": motivo,
                    "ts": guardado["ts"], "movimiento": guardado["id"],
                    "resumen": inv.resumen(), "bitacora": bitacora.contar(estado.con),
                    "historial": _con_permisos(
                        bitacora.por_fila(estado.con, fila, 20, inventario=inv.tipo),
                        {"usuario_id": uid, "rol": rol})})


def _completar_desde_catalogo(cuerpo, tipo_inventario, carpeta_datos=None, con=None):
    """Si el insumo está en el catálogo del ERP, sus datos mandan; si no, se acepta igual.

    Devuelve `(aviso, inventario_final)`. Nunca lanza: el almacén no puede rechazar una
    recepción porque el ERP todavía no tenga el insumo. Lo único que se valida contra el
    catálogo es la **ubicación** (el layout físico de los racks).

    Si la clave existe y es de otro inventario, el alta se mueve a ese inventario (y se avisa):
    así el renglón no queda clasificado en el libro equivocado sin dejar de aceptarlo.
    """
    sku = str(cuerpo.get("erp") or cuerpo.get("sku") or "").strip()
    datos = catalogo.cargar(catalogo.ruta(carpeta_datos), con=con)
    if not datos.get("existe"):
        if sku:
            return (f"no hay catálogo de insumos importado (scripts\\importar_catalogo.py): "
                    f"la clave {sku} no se pudo validar; se guarda con los datos capturados",
                    tipo_inventario)
        return None, tipo_inventario
    if datos.get("error"):
        return datos["error"], tipo_inventario
    if not sku:
        return ("sin clave del ERP: el material no se puede validar contra el catálogo; "
                "se guarda con los datos capturados", tipo_inventario)

    insumo = catalogo.por_sku(datos, sku)
    if insumo is None:
        return (f"la clave {sku} no está en el catálogo del ERP: se guarda como se capturó "
                f"(revisa la clave o importa el catálogo actualizado)", tipo_inventario)

    avisos = []
    tipo_final = tipo_inventario
    if insumo.get("inventario") and insumo["inventario"] != tipo_inventario:
        tipo_final = insumo["inventario"]
        avisos.append(f"la clave {sku} es {insumo.get('tipo') or 'de otro inventario'}: "
                      f"se ingresa en el inventario {tipo_final.upper()}")
    elif not insumo.get("inventario"):
        avisos.append(f"el insumo {sku} no trae tipo en el catálogo del ERP; "
                      f"se guarda en {tipo_inventario.upper()}")
    for campo, valor in (("descripcion", insumo.get("descripcion")),
                         ("unidad", insumo.get("unidad")),
                         ("categoria", insumo.get("tipo")),
                         ("proveedor", insumo.get("fabricante"))):
        if not valor:
            continue
        escrito = str(cuerpo.get(campo) or "").strip()
        if escrito and catalogo.normalizar(escrito) != catalogo.normalizar(valor):
            avisos.append(f"{campo} corregido al del catálogo ({valor})")
        cuerpo[campo] = valor
    cuerpo["erp"] = insumo["sku"]
    return ("; ".join(avisos) or None), tipo_final


def _avisos_naturaleza(cuerpo, tipo):
    """Recomendaciones según la naturaleza del material (no bloquean el ingreso)."""
    reglas = inventarios.TIPOS[tipo]
    avisos = []
    unidad = str(cuerpo.get("unidad") or "").strip().lower()
    if unidad and reglas.get("unidades") and unidad not in reglas["unidades"]:
        avisos.append(f"la unidad {unidad} no es la habitual para {reglas['corto']} "
                      f"(se espera {', '.join(reglas['unidades'])})")
    if tipo == "pt" and not (cuerpo.get("presentacion") and cuerpo.get("piezas")):
        avisos.append("producto terminado: conviene registrar presentación y piezas por caja")
    if tipo == "mep" and not str(cuerpo.get("caducidad") or "").strip():
        avisos.append("material de empaque sin caducidad: es válido, pero revisa si el proveedor "
                      "la declara")
    return avisos


def _alta(estado, inv, cuerpo, usuario, uid):
    """Ingreso de un lote nuevo: valida todo, escribe el renglón y lo deja en la bitácora."""
    cuerpo = dict(cuerpo or {})
    aviso, tipo_final = _completar_desde_catalogo(
        cuerpo, inv.tipo, estado.carpeta_datos, con=getattr(estado, "con", None))
    if tipo_final != inv.tipo:
        try:
            inv = estado.inventario(tipo_final)
        except (FileNotFoundError, KeyError):
            aviso = (f"{aviso}; " if aviso else "") + \
                f"el inventario {tipo_final.upper()} no está cargado: se guarda en {inv.tipo.upper()}"
    if cuerpo.get("ubicacion"):
        canonica = ubicaciones.canonico(cuerpo.get("ubicacion"))
        if canonica is None:
            return _r(400, {"error": f"la ubicación «{cuerpo.get('ubicacion')}» no existe en el "
                                     f"catálogo ALM-PNO-002"})
        cuerpo["ubicacion"] = canonica
    qr_id = inventario.clave_qr(cuerpo.get("qr_id")) or None
    if qr_id:
        ocupado = _qr_ocupado(estado, qr_id)
        if ocupado:
            return _r(409, {"error": f"el código «{qr_id}» ya está en "
                                     f"{ocupado['tipo'].upper()} fila {ocupado['fila']}; "
                                     f"usa otro código o quita la etiqueta de ese renglón"})
    try:
        registro = inv.agregar_alta(cuerpo)
    except ValueError as e:
        return _r(400, {"error": str(e)})

    fila = registro["fila"]
    guardado, error = _persistir(
        estado, inv, fila, "ALTA", usuario, uid, None, registro["ubicacion"],
        (cuerpo.get("motivo") or "recepcion"), cuerpo.get("ref_doc"),
        cuerpo.get("observaciones"), cantidad=registro["cantidad"],
        al_fallar=lambda: (inv.anular(fila, "alta revertida: no se pudo registrar"), inv.guardar()))
    if error:
        return error
    avisos = _avisos_naturaleza(cuerpo, inv.tipo)
    if aviso:
        avisos.insert(0, aviso)
    return _r(200, {"accion": "ALTA", "fila": fila, "lote": registro["lote"],
                    "descripcion": registro["descripcion"], "cantidad": registro["cantidad"],
                    "unidad": registro["unidad"], "ubicacion": registro["ubicacion"],
                    "qr_id": registro.get("qr_id"),
                    "caducidad": registro["caducidad"], "naturaleza": inv.tipo,
                    "naturaleza_nombre": inventarios.TIPOS[inv.tipo]["nombre"],
                    "archivo": inventarios.TIPOS[inv.tipo]["archivo"],
                    "aviso": "; ".join(avisos) if avisos else None, "ts": guardado["ts"],
                    "movimiento": guardado["id"], "resumen": inv.resumen(),
                    "bitacora": bitacora.contar(estado.con)})


def _salida_produccion(estado, inv, registro, cuerpo, usuario, uid, motivo):
    """ALMACÉN → PRODUCCIÓN: el material sale del almacén, no cambia de rack."""
    existencia = _existencia(registro)
    if existencia is None:
        return _r(400, {"error": "la CANTIDAD de ese renglón no es un número; corrígela en el Excel"})
    crudo = cuerpo.get("cantidad")
    if crudo in (None, "", 0, "0"):
        cantidad = existencia
    else:
        cantidad = _numero(crudo, total=existencia, unidad=registro["unidad"])
        if cantidad is None:
            return _r(400, {"error": f"cantidad inválida: {crudo!r}; escribe un número, "
                                     f"por ejemplo 5 o 12.5 (la unidad es {registro['unidad']})"})
        cantidad = round(cantidad, 3)
        if cantidad <= 0:
            return _r(400, {"error": "la cantidad surtida debe ser mayor que cero"})
        if cantidad > existencia + 0.001:
            return _r(400, {"error": f"no puedes mover {cantidad}; hay {existencia} "
                                     f"{registro['unidad']} en ese renglón"})

    fila = registro["fila"]
    origen = registro["ubicacion"]
    nota = str(cuerpo.get("nota") or "").strip() or None
    if motivo == "otro" and not nota:
        return _r(400, {"error": "con el motivo «otro» hace falta escribir la nota"})
    try:
        despues = inv.sacar(fila, cantidad, nota=f"surtido a producción {nota or ''}".strip())
    except ValueError as e:
        return _r(400, {"error": str(e)})

    def revertir():
        inv.ajustar_cantidad(fila, existencia)
        inv.asignar(fila, origen)
        inv.guardar()

    guardado, error = _persistir(estado, inv, fila, "SALIDA", usuario, uid, origen, PRODUCCION,
                                 motivo or "surtido", cuerpo.get("ref_doc"), nota,
                                 cantidad=cantidad, al_fallar=revertir)
    if error:
        return error
    return _r(200, {"accion": "SALIDA", "fila": fila, "anterior": origen, "nueva": PRODUCCION,
                    "cantidad": cantidad, "unidad": registro["unidad"],
                    "saldo": (despues or {}).get("cantidad"), "motivo": motivo,
                    "ts": guardado["ts"], "movimiento": guardado["id"],
                    "resumen": inv.resumen(), "bitacora": bitacora.contar(estado.con),
                    "historial": _con_permisos(bitacora.por_fila(estado.con, fila, 20,
                                                                 inventario=inv.tipo),
                                               {"usuario_id": uid, "rol": "almacen"})})


def _movimiento(estado, inv, ruta, cuerpo, usuario, uid):
    try:
        fila = int(cuerpo["fila"])
    except (KeyError, TypeError, ValueError):
        return _r(400, {"error": "fila requerida"})
    try:
        registro = inv.por_fila(fila)
    except KeyError as e:
        return _r(400, {"error": str(e)})

    nota = str(cuerpo.get("nota") or "").strip() or None
    if ruta == "/api/no_encontrado" and (not nota or len(nota) < MIN_NOTA_NO_ENCONTRADO):
        return _r(400, {"error": "escribe qué pasó con el lote (mínimo 3 letras): "
                                 "es lo que queda en la bitácora"})
    try:
        if ruta == "/api/ubicar":
            codigo = ubicaciones.canonico(cuerpo.get("ubicacion"))
            if codigo is None:
                if not str(cuerpo.get("ubicacion") or "").strip():
                    return _r(400, {"error": "escribe el código de ubicación; para dejarlo "
                                             "sin ubicación usa «quitar ubicación»"})
                return _r(400, {"error": f"la ubicación «{str(cuerpo.get('ubicacion')).strip()}» "
                                         f"no existe en el catálogo ALM-PNO-002"})
            anterior, nueva = inv.asignar(fila, codigo)
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

    if accion == "CONFIRMAR":
        # Nada cambió en el Excel: solo se deja constancia de que alguien verificó el lote.
        try:
            ts, mov_id = bitacora.registrar(estado.con, accion, registro, usuario,
                                            destino=nueva, origen=anterior, nota=nota,
                                            usuario_id=uid, inventario=inv.tipo,
                                            unidad=registro.get("unidad"))
        except Exception as e:  # noqa: BLE001 - la bitácora es obligatoria
            return _r(500, {"error": f"no se pudo registrar la confirmación: {e}"})
        bitacora.auditar(estado.con, usuario, "CONFIRMAR_UBICACION",
                         {"fila": fila, "ubicacion": nueva, "inventario": inv.tipo})
        return _r(200, {"accion": accion, "fila": fila, "anterior": anterior, "nueva": nueva,
                        "ts": ts, "movimiento": mov_id, "resumen": inv.resumen(),
                        "bitacora": bitacora.contar(estado.con)})

    guardado, error = _persistir(estado, inv, fila, accion, usuario, uid, anterior, nueva,
                                 None, None, nota,
                                 al_fallar=lambda: (inv.asignar(fila, anterior), inv.guardar()))
    if error:
        return error
    return _r(200, {"accion": accion, "fila": fila, "anterior": anterior, "nueva": nueva,
                    "ts": guardado["ts"], "movimiento": guardado["id"],
                    "resumen": inv.resumen(), "bitacora": bitacora.contar(estado.con)})


def _deshacer(estado, cuerpo, sesion):
    mov = bitacora.por_id(estado.con, cuerpo.get("movimiento"))
    if mov is None:
        return _r(404, {"error": "ese movimiento no existe (puede que ya se haya borrado)"})
    if mov["revertido"]:
        return _r(400, {"error": "ese movimiento ya fue deshecho"})
    if mov["accion"] not in REVERSIBLES:
        return _r(400, {"error": f"el movimiento {mov['accion']} no se puede deshacer"})
    admin = sesion["rol"] == "admin"
    if not admin and (mov["usuario_id"] != sesion["usuario_id"] or sesion["rol"] != "almacen"):
        return _r(403, {"error": "solo puedes deshacer tus propios movimientos"})

    tipo = mov.get("inventario") or "mp"
    try:
        inv = estado.inventario(tipo)
    except (FileNotFoundError, KeyError):
        return _r(409, {"error": f"el inventario {tipo.upper()} de ese movimiento no está cargado"})
    if mov.get("fila") is None:
        return _r(409, {"error": "ese movimiento pertenece a un renglón que se borró; "
                                 "no hay nada que deshacer en el inventario"})

    ultimo = bitacora.ultimo_vigente(estado.con, mov["fila"], inventario=tipo)
    if ultimo is None or ultimo["id"] != mov["id"]:
        return _r(409, {"error": "hay movimientos posteriores de ese lote; deshaz primero el último"})

    try:
        registro = inv.por_fila(mov["fila"])
    except KeyError:
        return _r(400, {"error": "el renglón ya no existe en el inventario"})

    if mov["accion"] == "PARCIAL":
        return _deshacer_parcial(estado, inv, mov, registro, sesion)
    if mov["accion"] == "ALTA":
        return _deshacer_alta(estado, inv, mov, registro, sesion)
    if mov["accion"] == "SALIDA":
        return _deshacer_salida(estado, inv, mov, registro, sesion)

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
                    "anterior": actual, "nueva": destino, "unidad": registro.get("unidad"),
                    "resumen": inv.resumen(),
                    "historial": _con_permisos(
                        bitacora.por_fila(estado.con, mov["fila"], 20, inventario=tipo),
                        sesion)})


def _deshacer_salida(estado, inv, mov, registro, sesion):
    """Devuelve al almacén lo que había salido a producción (misma ubicación de origen)."""
    cantidad = float(mov.get("cantidad") or 0)
    origen = mov.get("ubicacion_origen")
    if cantidad <= 0:
        return _r(409, {"error": "ese movimiento no tiene cantidad registrada"})
    actual = _existencia(registro)
    if actual is None:
        return _r(409, {"error": "la CANTIDAD del renglón no es un número; corrígela en el Excel"})
    if registro["ubicacion"] and registro["ubicacion"] != origen:
        return _r(409, {"error": f"ese lote ya se movió a {registro['ubicacion']}; "
                                 "deshaz primero los movimientos posteriores"})
    sin_ubicacion = not registro["ubicacion"]
    if origen and not ubicaciones.existe(origen):
        return _r(409, {"error": f"el origen {origen} ya no está en el catálogo de ubicaciones"})
    inv.ajustar_cantidad(mov["fila"], round(actual + cantidad, 3))
    if sin_ubicacion and origen:
        inv.asignar(mov["fila"], origen)

    def revertir():
        inv.ajustar_cantidad(mov["fila"], actual)
        if sin_ubicacion:
            inv.asignar(mov["fila"], None)
        inv.guardar()

    guardado, error = _persistir(estado, inv, mov["fila"], "DESHACER", sesion["nombre"],
                                 sesion["usuario_id"], origen, None, None, None,
                                 f"deshace la salida a producción #{mov['id']}",
                                 cantidad=cantidad, al_fallar=revertir)
    if error:
        return error
    bitacora.marcar_revertido(estado.con, mov["id"], guardado["id"])
    bitacora.auditar(estado.con, sesion["nombre"], "DESHACER_SALIDA",
                     {"movimiento": mov["id"], "lote": mov.get("lote"), "cantidad": cantidad})
    return _r(200, {"deshecho": mov["id"], "movimiento": guardado["id"], "fila": mov["fila"],
                    "anterior": None, "nueva": origen, "cantidad": cantidad,
                    "unidad": registro.get("unidad"),
                    "resumen": inv.resumen(),
                    "historial": _con_permisos(bitacora.por_fila(estado.con, mov["fila"], 20,
                                                                 inventario=inv.tipo), sesion)})


def _deshacer_alta(estado, inv, mov, registro, sesion):
    """Deshace un ingreso: deja el renglón en cero y sin ubicación (el Excel no se borra)."""
    cantidad = float(mov.get("cantidad") or 0)
    actual = _existencia(registro)
    if actual is None:
        return _r(409, {"error": "la CANTIDAD del renglón no es un número; corrígela en el Excel"})
    if abs(actual - cantidad) > 0.001:
        return _r(409, {"error": f"ese lote ya cambió de cantidad (hay {actual}); "
                                 f"deshaz primero los movimientos posteriores"})
    ubicacion = registro["ubicacion"]
    inv.anular(mov["fila"], f"alta deshecha #{mov['id']}")

    def revertir():
        inv.ajustar_cantidad(mov["fila"], cantidad)
        inv.asignar(mov["fila"], ubicacion)
        inv.guardar()

    guardado, error = _persistir(estado, inv, mov["fila"], "DESHACER", sesion["nombre"],
                                 sesion["usuario_id"], ubicacion, None, None, None,
                                 f"deshace el alta #{mov['id']}", cantidad=cantidad,
                                 al_fallar=revertir)
    if error:
        return error
    bitacora.marcar_revertido(estado.con, mov["id"], guardado["id"])
    bitacora.auditar(estado.con, sesion["nombre"], "DESHACER_ALTA",
                     {"movimiento": mov["id"], "lote": mov.get("lote")})
    return _r(200, {"deshecho": mov["id"], "movimiento": guardado["id"], "fila": mov["fila"],
                    "anterior": ubicacion, "nueva": None, "cantidad": cantidad,
                    "unidad": registro.get("unidad"),
                    "resumen": inv.resumen(),
                    "historial": _con_permisos(bitacora.por_fila(estado.con, mov["fila"], 20,
                                                                 inventario=inv.tipo), sesion)})


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
    if abs((_existencia(parcial) or 0) - cantidad) > 0.001:
        return _r(409, {"error": f"la parcialidad ya cambió de cantidad "
                                 f"({parcial['cantidad']} {parcial['unidad']}); no se puede deshacer"})

    actual_origen = _existencia(registro)
    if actual_origen is None:
        return _r(409, {"error": "la CANTIDAD del renglón original no es un número; "
                                 "corrígela en el Excel del área"})
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
                    "unidad": registro.get("unidad"),
                    "resumen": inv.resumen(),
                    "historial": _con_permisos(bitacora.por_fila(estado.con, mov["fila"], 20,
                                                                 inventario=inv.tipo), sesion)})


def _revertir_parcial(inv, mov):
    """Devuelve la cantidad al renglón original y anula la copia de una parcialidad."""
    nueva_fila = mov.get("fila_destino")
    cantidad = float(mov.get("cantidad") or 0)
    if not nueva_fila or cantidad <= 0:
        raise ValueError("ese movimiento no tiene la parcialidad asociada")
    parcial = inv.por_fila(nueva_fila)
    if (parcial["ubicacion"] or "") != (mov["ubicacion_destino"] or ""):
        raise ValueError(f"la parcialidad ya se movió a {parcial['ubicacion'] or 'sin ubicación'}")
    if abs((_existencia(parcial) or 0) - cantidad) > 0.001:
        raise ValueError(f"la parcialidad ya cambió de cantidad ({parcial['cantidad']})")
    origen = inv.por_fila(mov["fila"])
    actual = _existencia(origen)
    if actual is None:
        raise ValueError("la CANTIDAD del renglón original no es un número")
    inv.ajustar_cantidad(mov["fila"], round(actual + cantidad, 3))
    inv.anular(nueva_fila, f"anulada al borrar el movimiento #{mov['id']}")


def _purgar_salida(inv, mov):
    """Revierte una salida a producción: devuelve la cantidad (y la ubicación si se perdió)."""
    cantidad = float(mov.get("cantidad") or 0)
    if cantidad <= 0:
        raise ValueError("ese movimiento no tiene cantidad registrada")
    registro = inv.por_fila(mov["fila"])
    actual = _existencia(registro)
    if actual is None:
        raise ValueError("la CANTIDAD del renglón no es un número")
    origen = mov.get("ubicacion_origen")
    if registro["ubicacion"] and registro["ubicacion"] != origen:
        raise ValueError(f"ese lote ya se movió a {registro['ubicacion']}")
    inv.ajustar_cantidad(mov["fila"], round(actual + cantidad, 3))
    if origen and not registro["ubicacion"]:
        if not ubicaciones.existe(origen):
            raise ValueError(f"el origen {origen} ya no está en el catálogo")
        inv.asignar(mov["fila"], origen)


def _purgar_alta(inv, mov):
    """Revierte un alta: deja el renglón en cero y sin ubicación (no borra la fila)."""
    cantidad = float(mov.get("cantidad") or 0)
    registro = inv.por_fila(mov["fila"])
    actual = _existencia(registro)
    if actual is None:
        raise ValueError("la CANTIDAD del renglón no es un número")
    if abs(actual - cantidad) > 0.001:
        raise ValueError(f"ese lote ya cambió de cantidad (hay {actual})")
    inv.anular(mov["fila"], f"alta borrada #{mov['id']}")


def _borrar_renglones(estado, cuerpo, sesion):
    """Borra renglones completos del inventario (solo admin, con motivo y respaldo del Excel).

    El renglón desaparece del archivo del área, pero antes se guarda una copia del libro y se
    registra en la bitácora qué había en cada renglón borrado. Los movimientos de los renglones
    que se van quedan sin fila (fila NULL) y los de más abajo se recorren, para que el historial
    siga apuntando al renglón correcto.
    """
    if sesion["rol"] != "admin":
        return _r(403, {"error": "solo un administrador puede borrar renglones"})
    motivo = str(cuerpo.get("motivo") or "").strip()
    if len(motivo) < MIN_MOTIVO_PURGA:
        return _r(400, {"error": f"escribe el motivo del borrado (mínimo {MIN_MOTIVO_PURGA} "
                                 f"caracteres)"})
    filas = cuerpo.get("filas") or []
    if not isinstance(filas, list) or not filas:
        return _r(400, {"error": "indica los renglones a borrar"})
    numeros = sorted({n for n in (_entero(f, 0, 0) for f in filas) if n > 0})
    if not numeros:
        return _r(400, {"error": "indica los renglones a borrar"})
    try:
        tipo = inventarios.normalizar_tipo(cuerpo.get("tipo") or "mp")
    except ValueError as e:
        return _r(400, {"error": str(e)})
    try:
        inv = estado.inventario(tipo)
    except (FileNotFoundError, KeyError):
        return _r(409, {"error": f"el inventario {tipo.upper()} no está cargado"})

    en_db = bool(getattr(inv, "en_db", False))
    respaldo = "supabase:auditoria"
    if not en_db:
        carpeta_respaldos = inv.ruta.parent / "respaldos"
        carpeta_respaldos.mkdir(parents=True, exist_ok=True)
        marca = time.strftime("%Y%m%d-%H%M%S")
        respaldo = carpeta_respaldos / f"antes_de_borrar_{marca}_{inv.ruta.name}"
        try:
            shutil.copy2(inv.ruta, respaldo)
        except OSError as e:
            return _r(409, {"error": f"no se pudo respaldar el Excel antes de borrar: {e}"})

    try:
        borrados = inv.borrar_renglones(numeros)
    except (ValueError, KeyError) as e:
        return _r(400, {"error": str(e)})

    try:
        inv.guardar()
    except PermissionError:
        if not en_db:
            shutil.copy2(respaldo, inv.ruta)
        inv.recargar()
        return _r(409, {"error": "el archivo está abierto en Excel; ciérralo y reintenta"})
    except OSError as e:
        if not en_db:
            shutil.copy2(respaldo, inv.ruta)
        inv.recargar()
        return _r(409, {"error": f"no se pudo guardar el Excel: {e}"})

    try:
        for fila, registro in sorted(borrados.items()):
            # el renglón ya no existe: la entrada de auditoría va sin fila y con los datos en la nota
            bitacora.registrar(estado.con, "PURGAR", {**registro, "fila": None},
                               sesion["nombre"], destino=None,
                               origen=registro.get("ubicacion"), motivo=motivo,
                               nota=(f"renglón borrado (fila {fila}): {registro.get('descripcion')}"
                                     f" · lote {registro.get('lote')} · "
                                     f"{registro.get('cantidad')} {registro.get('unidad') or ''}"),
                               usuario_id=sesion["usuario_id"],
                               cantidad=registro.get("cantidad"), inventario=tipo)
        # Los movimientos de los renglones borrados quedan sin fila; los de abajo se recorren
        # de mayor a menor para que cada resta sea la correcta.
        for fila in sorted(borrados, reverse=True):
            estado.con.execute("UPDATE movimiento SET fila = NULL WHERE inventario=? AND fila=?",
                               (tipo, fila))
            estado.con.execute("UPDATE movimiento SET fila = fila - 1 "
                               "WHERE inventario=? AND fila > ?", (tipo, fila))
        estado.con.commit()
    except Exception as e:  # noqa: BLE001 - si la bitácora falla, el Excel se restaura
        if not en_db:
            shutil.copy2(respaldo, inv.ruta)
        inv.recargar()
        return _r(500, {"error": f"no se pudo registrar el borrado: {e}. Excel restaurado"})

    detalle = [{"fila": fila, "descripcion": r.get("descripcion"), "lote": r.get("lote"),
                "cantidad": r.get("cantidad"), "unidad": r.get("unidad"),
                "ubicacion": r.get("ubicacion")} for fila, r in sorted(borrados.items())]
    bitacora.auditar(estado.con, sesion["nombre"], "BORRAR_RENGLONES",
                     {"motivo": motivo, "inventario": tipo, "renglones": detalle,
                      "respaldo": str(respaldo)})
    eventos.publicar("renglones.borrados", inventario=tipo, total=len(detalle),
                     motivo=motivo, usuario=sesion["nombre"], respaldo=str(respaldo))
    return _r(200, {"borrados": detalle, "total": len(detalle), "respaldo": str(respaldo),
                    "resumen": inv.resumen(), "bitacora": bitacora.contar(estado.con)})


def _purgar(estado, cuerpo, sesion):
    """Borrado físico de un movimiento (solo admin): revierte el efecto y deja auditoría."""
    if sesion["rol"] != "admin":
        return _r(403, {"error": "solo un administrador puede borrar movimientos"})
    mov = bitacora.por_id(estado.con, cuerpo.get("movimiento"))
    if mov is None:
        return _r(404, {"error": "ese movimiento no existe"})
    motivo = str(cuerpo.get("motivo") or "").strip()
    if len(motivo) < MIN_MOTIVO_PURGA:
        return _r(400, {"error": f"escribe el motivo del borrado (mínimo {MIN_MOTIVO_PURGA} "
                                 f"caracteres)"})

    tipo = mov.get("inventario") or "mp"
    try:
        inv = estado.inventario(tipo)
    except (FileNotFoundError, KeyError):
        return _r(409, {"error": f"el inventario {tipo.upper()} de ese movimiento no está cargado"})
    if mov.get("fila") is None:
        return _r(409, {"error": "ese movimiento pertenece a un renglón que ya se borró"})

    ultimo = bitacora.ultimo_vigente(estado.con, mov["fila"], inventario=tipo)
    if ultimo and ultimo["id"] != mov["id"]:
        return _r(409, {"error": "hay movimientos posteriores de ese lote; deshaz primero el último"})

    if not mov["revertido"] and mov["accion"] in REVERSIBLES:
        try:
            if mov["accion"] == "PARCIAL":
                _revertir_parcial(inv, mov)
            elif mov["accion"] == "SALIDA":
                _purgar_salida(inv, mov)
            elif mov["accion"] == "ALTA":
                _purgar_alta(inv, mov)
            else:
                inv.asignar(mov["fila"], mov["ubicacion_origen"])
            inv.guardar()
        except PermissionError:
            return _r(409, {"error": "el archivo está abierto en Excel; ciérralo y reintenta"})
        except (KeyError, ValueError, OSError) as e:
            return _r(409, {"error": f"no se pudo revertir el efecto en el Excel: {e}"})

    borrado = bitacora.borrar(estado.con, mov["id"])
    bitacora.auditar(estado.con, sesion["nombre"], "PURGA_MOVIMIENTO",
                     {"motivo": motivo, "movimiento": borrado})
    return _r(200, {"borrado": mov["id"], "motivo": motivo, "resumen": inv.resumen(),
                    "historial": _con_permisos(
                        bitacora.por_fila(estado.con, mov["fila"], 20, inventario=tipo),
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
        return None, _r(409, {"error": "el archivo del inventario está abierto en Excel; "
                                      "ciérralo y vuelve a intentar"})
    except OSError as e:
        revertir()
        return None, _r(409, {"error": f"no se pudo guardar el Excel: {e}"})

    try:
        ts, mov_id = bitacora.registrar(estado.con, accion, registro, usuario,
                                        destino=nueva, origen=anterior, nota=nota,
                                        motivo=motivo, ref_doc=ref_doc, usuario_id=uid,
                                        cantidad=cantidad, fila_destino=fila_destino,
                                        inventario=inv.tipo, unidad=registro.get("unidad"))
    except Exception as e:
        revertir()
        return None, _r(500, {"error": f"no se pudo registrar en la bitácora, cambio "
                                       f"revertido: {e}"})
    eventos.publicar("movimiento.registrado", accion=accion, inventario=inv.tipo, fila=fila,
                     lote=registro.get("lote"), descripcion=registro.get("descripcion"),
                     de=anterior, a=nueva, cantidad=cantidad, motivo=motivo, usuario=usuario,
                     movimiento=mov_id)
    return {"ts": ts, "id": mov_id}, None


# ---------------------------------------------------------------- copia a la USB
def _copia_estado(estado, params):
    from scripts import sincronizar_usb
    destino = params.get("destino") or str(sincronizar_usb.DESTINO)
    try:
        datos = sincronizar_usb.comparar(estado.carpeta_datos, destino)
    except (OSError, ValueError) as e:
        return _r(400, {"error": f"no se pudo revisar la copia: {e}"})
    datos["disponible"] = datos.get("existe_destino", False)
    datos["sugerido"] = str(sincronizar_usb.DESTINO)
    return _r(200, datos)


def _copia_ahora(estado, cuerpo, usuario):
    from scripts import sincronizar_usb
    destino = cuerpo.get("destino") or None
    historial = bool(cuerpo.get("historial", True))
    forzar = bool(cuerpo.get("forzar", False))
    try:
        manifiesto = sincronizar_usb.sincronizar(
            origen=estado.carpeta_datos, destino=destino, historial=historial, forzar=forzar)
    except FileNotFoundError as e:
        return _r(409, {"error": str(e)})
    except RuntimeError as e:
        return _r(500, {"error": str(e)})
    except OSError as e:
        return _r(409, {"error": f"no se pudo escribir en el destino: {e}"})
    bitacora.auditar(estado.con, usuario, "COPIA_USB",
                     {"movimientos": manifiesto["movimientos"], "destino": manifiesto["destino"]})
    eventos.publicar("copia.usb", usuario=usuario, movimientos=manifiesto["movimientos"],
                     destino=manifiesto["destino"], integridad=manifiesto.get("integridad"))
    return _r(200, manifiesto)


def descargar(estado, tipo, inventario_tipo=None):
    """Genera un reporte. f01/f01v3 salen consolidados de todos los inventarios cargados."""
    with _ESCRITURA:
        invs = []
        en_db = bool(getattr(estado, "en_db", False))
        for ficha in inventarios.lista(estado.carpeta_datos, en_db=en_db):
            try:
                invs.append((ficha["clave"], estado.inventario(ficha["clave"])))
            except (FileNotFoundError, KeyError):
                continue
        if not invs:
            raise FileNotFoundError("no hay inventarios cargados")

        elegido = inventarios.normalizar_tipo(inventario_tipo) if inventario_tipo else None
        if elegido and elegido not in dict(invs):
            raise FileNotFoundError(
                f"el inventario {elegido.upper()} no está cargado; no se puede generar el reporte")
        if tipo == "inventario":
            inv = dict(invs)[elegido or invs[0][0]]
            return reportes.exportar_inventario(inv)
        if tipo == "f01":
            return reportes.generar_f01(invs, version="2")
        if tipo == "f01v3":
            return reportes.generar_f01(invs, version="3")
        if tipo == "tabla_macro":
            inv = dict(invs)[elegido or invs[0][0]]
            return reportes.tabla_macro(inv)
        if tipo == "bitacora":
            return reportes.exportar_bitacora(estado.con, inventario=elegido)
        raise KeyError(tipo)


# ---------------------------------------------------------------- depuración
_MAX_MENSAJE_LOG = 1500
_MAX_LOGS_POR_MINUTO = 30
_LOGS_NAVEGADOR = {}          # {usuario_id: [marca, cuantos]}
_LOGS_LOCK = threading.Lock()


def _log_navegador(cuerpo, sesion):
    """El navegador manda aquí sus errores de JavaScript (window.onerror y promesas)."""
    mensaje = str(cuerpo.get("mensaje") or "").strip()
    if not mensaje:
        return _r(400, {"error": "falta el mensaje del error"})
    uid = sesion["usuario_id"]
    with _LOGS_LOCK:
        marca, cuantos = _LOGS_NAVEGADOR.get(uid, [0, 0])
        ahora = time.monotonic()
        if not marca or ahora - marca > 60:
            marca, cuantos = ahora, 0
        if cuantos >= _MAX_LOGS_POR_MINUTO:
            return _r(429, {"error": "demasiados reportes de error; espera un minuto"})
        _LOGS_NAVEGADOR[uid] = [marca, cuantos + 1]
    contexto = cuerpo.get("contexto") if isinstance(cuerpo.get("contexto"), dict) else {}
    contexto = {str(k)[:40]: str(v)[:200] for k, v in list(contexto.items())[:10]}
    url = str(cuerpo.get("url") or "")[:200]
    nivel = str(cuerpo.get("nivel") or "error").lower()[:20]
    logger.navegador(mensaje[:_MAX_MENSAJE_LOG], contexto=contexto, url=url, nivel_js=nivel)
    eventos.publicar("navegador.error", mensaje=mensaje[:300], nivel=nivel, url=url,
                     usuario=sesion["nombre"], contexto=str(contexto)[:200])
    return _r(200, {"ok": True})


def _eventos_recientes(params, sesion):
    """Últimos eventos del sistema (solo admin): qué pasó y en qué orden."""
    if sesion["rol"] != "admin":
        return _r(403, {"error": "solo un administrador puede ver los eventos del sistema"})
    desde = _entero(params.get("desde"), 0, 0)
    limite = _entero(params.get("limite"), 50, 1, 200)
    return _r(200, {"eventos": eventos.desde(desde, limite), "desde": desde})


def _diagnostico(estado, sesion):
    """Todo lo que hace falta para depurar una instalación sin pedirle capturas a nadie."""
    if sesion["rol"] != "admin":
        return _r(403, {"error": "solo un administrador puede ver el diagnóstico"})
    import platform
    import sys

    carpetas = {}
    try:
        uso = shutil.disk_usage(estado.carpeta_datos)
        carpetas["libre_mb"] = round(uso.free / (1024 * 1024), 1)
        carpetas["total_mb"] = round(uso.total / (1024 * 1024), 1)
    except OSError as e:
        carpetas["error"] = str(e)
    catalogo_datos = catalogo.cargar(catalogo.ruta(getattr(estado, "carpeta_datos", None)),
                                     con=getattr(estado, "con", None))
    errores = eventos.ultimos(200, solo_errores=True)[-15:]
    return _r(200, {
        "version": _version(),
        "python": sys.version.split()[0],
        "plataforma": platform.platform(),
        "carpeta_datos": str(estado.carpeta_datos),
        "archivo": getattr(estado, "archivo", ""),
        "demo": bool(getattr(estado, "demo", False)),
        "disco": carpetas,
        "log": str(logger.ruta() or ""),
        "nivel_log": logger.nivel(),
        "inventarios": _estado_json(estado),
        "catalogo": catalogo.resumen(catalogo_datos),
        "bitacora": {"movimientos": bitacora.contar(estado.con),
                     "usuarios": len(usuarios.listar(estado.con))},
        "eventos": eventos.ultimos(60),
        "errores": errores,
    })
