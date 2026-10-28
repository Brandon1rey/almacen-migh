"""Copia la base del servidor a la USB (o a cualquier carpeta) para revisarla sin tocar producción.

    python scripts\\sincronizar_usb.py
    python scripts\\sincronizar_usb.py --destino "E:\\almacen-datos"
    python scripts\\sincronizar_usb.py --historial          # guarda una copia con fecha
    python scripts\\sincronizar_usb.py --resumen            # lee la copia y muestra qué trae

La base se copia con `sqlite3.backup()`, no con copia de archivo: si el almacén está registrando un
movimiento en ese momento, la copia sale íntegra y con todo lo confirmado hasta ese instante.
El servidor no se detiene y el archivo original no se modifica.
"""
import argparse
import datetime as dt
import hashlib
import json
import pathlib
import shutil
import sqlite3
import sys

# La consola de Windows no siempre es UTF-8: sin esto, imprimir «→» o «·» lanza
# UnicodeEncodeError y el script parece fallar aunque ya haya hecho su trabajo.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass


RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from app import inventarios  # noqa: E402

ORIGEN = RAIZ / "data"
DESTINO = pathlib.Path(r"D:\almacen-ubicaciones-datos")
NOMBRE_JSON = "sincronizacion.json"
MANTENER = 30


def _sha256(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _contar(ruta_db):
    con = sqlite3.connect(str(ruta_db), timeout=10)
    try:
        return con.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0]
    finally:
        con.close()


def _copiar_base(origen_db, destino_db):
    """Copia consistente con la API de backup de SQLite (sirve con el servidor trabajando)."""
    destino_db.parent.mkdir(parents=True, exist_ok=True)
    temporal = destino_db.with_suffix(".db.tmp")
    if temporal.exists():
        temporal.unlink()
    con_origen = sqlite3.connect(f"file:{origen_db}?mode=ro", uri=True, timeout=15)
    con_destino = sqlite3.connect(str(temporal))
    try:
        con_origen.backup(con_destino)
        con_destino.commit()
    finally:
        con_destino.close()
        con_origen.close()
    temporal.replace(destino_db)
    return destino_db


def _revisar_copia(destino_db, esperados):
    con = sqlite3.connect(str(destino_db), timeout=10)
    try:
        integridad = con.execute("PRAGMA integrity_check").fetchone()[0]
        movimientos = con.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0]
    finally:
        con.close()
    if integridad != "ok":
        raise RuntimeError(f"la copia salió dañada (integrity_check={integridad})")
    if movimientos != esperados:
        raise RuntimeError(f"la copia tiene {movimientos} movimientos y el original {esperados}")
    return integridad, movimientos


def _podar(carpeta, mantener):
    copias = sorted(carpeta.glob("bitacora_*.db"))
    if mantener and len(copias) > mantener:
        for vieja in copias[: len(copias) - mantener]:
            vieja.unlink()


def _movimientos_de(ruta_db):
    """Cuántos movimientos y cuál es el último id; None si la base no existe."""
    ruta_db = pathlib.Path(ruta_db)
    if not ruta_db.exists():
        return None
    con = sqlite3.connect(f"file:{ruta_db}?mode=ro", uri=True, timeout=10)
    try:
        fila = con.execute("SELECT COUNT(*), COALESCE(MAX(id), 0), COALESCE(MAX(ts), '') "
                           "FROM movimiento").fetchone()
        return {"movimientos": fila[0], "ultimo_id": fila[1], "ultimo_ts": fila[2]}
    finally:
        con.close()


def _manifest_previo(destino):
    ruta = pathlib.Path(destino) / "ultima" / NOMBRE_JSON
    if not ruta.exists():
        return None
    try:
        return json.loads(ruta.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def comparar(origen=None, destino=None):
    """Qué está más al día: el servidor o la copia de la USB.

    La base se compara por movimientos y último folio (semántica), no por fecha: su mtime cambia
    con cualquier escritura, incluso una sesión que solo se abre. Los inventarios (.xlsx) sí por
    fecha y tamaño, que es lo que se ve en el explorador de archivos.
    """
    origen = pathlib.Path(origen or ORIGEN)
    destino = pathlib.Path(destino or DESTINO)
    ultima = destino / "ultima"
    salida = {"origen": str(origen), "destino": str(destino),
              "existe_destino": destino.exists(), "base": {}, "archivos": {},
              "al_dia": True, "pendientes": [], "adelantados": [],
              "ultima_copia": (_manifest_previo(destino) or {}).get("generado")}

    o_db, d_db = origen / "bitacora.db", ultima / "bitacora.db"
    en_origen, en_copia = _movimientos_de(o_db), _movimientos_de(d_db)
    if en_copia is None:
        estado, detalle = "falta en la copia", "todavía no se ha copiado"
    elif en_origen is None:
        estado, detalle = "sin origen", "no hay base del servidor"
    elif en_origen["movimientos"] == en_copia["movimientos"] and \
            en_origen["ultimo_id"] == en_copia["ultimo_id"]:
        estado, detalle = "igual", "la copia tiene los mismos movimientos"
    elif en_copia["movimientos"] > en_origen["movimientos"] or \
            en_copia["ultimo_id"] > en_origen["ultimo_id"]:
        estado, detalle = "la copia está más al día", \
            f"la USB trae {en_copia['movimientos']} movimientos y el servidor {en_origen['movimientos']}"
    else:
        estado, detalle = "hay cambios por copiar", \
            f"el servidor trae {en_origen['movimientos']} movimientos y la copia {en_copia['movimientos']}"
    salida["base"] = {"estado": estado, "detalle": detalle,
                      "movimientos_origen": (en_origen or {}).get("movimientos", 0),
                      "movimientos_copia": (en_copia or {}).get("movimientos", 0),
                      "ultimo_id_origen": (en_origen or {}).get("ultimo_id", 0),
                      "ultimo_id_copia": (en_copia or {}).get("ultimo_id", 0),
                      "ultimo_ts_origen": (en_origen or {}).get("ultimo_ts")}

    for tipo in sorted(inventarios.TIPOS, key=lambda t: inventarios.TIPOS[t]["orden"]):
        ruta_o = inventarios.ruta(tipo, origen)
        ruta_c = ultima / ruta_o.name
        existe_o, existe_c = ruta_o.exists(), ruta_c.exists()
        if not existe_o and not existe_c:
            continue
        if not existe_c:
            estado = "falta en la copia"
        elif not existe_o:
            estado = "la copia está más al día"
        else:
            iguales = (ruta_o.stat().st_size == ruta_c.stat().st_size
                       and abs(ruta_o.stat().st_mtime - ruta_c.stat().st_mtime) < 2)
            estado = "igual" if iguales else (
                "hay cambios por copiar" if ruta_o.stat().st_mtime > ruta_c.stat().st_mtime
                else "la copia está más al día")
        salida["archivos"][ruta_o.name] = {
            "tipo": tipo, "estado": estado,
            "bytes_origen": ruta_o.stat().st_size if existe_o else None,
            "bytes_copia": ruta_c.stat().st_size if existe_c else None,
            "fecha_origen": dt.datetime.fromtimestamp(ruta_o.stat().st_mtime).isoformat(
                sep=" ", timespec="seconds") if existe_o else None,
            "fecha_copia": dt.datetime.fromtimestamp(ruta_c.stat().st_mtime).isoformat(
                sep=" ", timespec="seconds") if existe_c else None,
        }

    salida["pendientes"] = [n for n, d in salida["archivos"].items()
                            if d["estado"] in ("falta en la copia", "hay cambios por copiar")]
    if salida["base"]["estado"] in ("falta en la copia", "hay cambios por copiar"):
        salida["pendientes"].insert(0, "bitacora.db")
    salida["adelantados"] = [n for n, d in salida["archivos"].items()
                             if d["estado"] == "la copia está más al día"]
    if salida["base"]["estado"] == "la copia está más al día":
        salida["adelantados"].insert(0, "bitacora.db")
    salida["al_dia"] = not salida["pendientes"]
    return salida


def sincronizar(origen=None, destino=None, historial=False, con_inventarios=True,
                 crear_destino=True, mantener=MANTENER, solo_recientes=True, forzar=False):
    origen = pathlib.Path(origen or ORIGEN)
    destino = pathlib.Path(destino or DESTINO)

    origen_db = origen / "bitacora.db"
    if not origen_db.exists():
        raise FileNotFoundError(f"no existe la base del servidor: {origen_db}")
    if not destino.exists():
        if not crear_destino:
            raise FileNotFoundError(f"no existe el destino: {destino} (¿la USB está conectada?)")
        try:
            destino.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            raise FileNotFoundError(f"no se pudo crear {destino}: {e} (¿la USB está conectada?)")

    ultima = destino / "ultima"
    ultima.mkdir(parents=True, exist_ok=True)
    estado = comparar(origen, destino)
    omitidos = {}
    avisos = []

    esperados = _contar(origen_db)
    estado_db = estado["base"]["estado"]
    copiar_db = forzar or not solo_recientes or estado_db not in ("igual", "la copia está más al día")
    if not copiar_db and (ultima / "bitacora.db").exists():
        copia_db = ultima / "bitacora.db"
        movimientos = estado["base"]["movimientos_copia"]
        integridad = "ok (sin cambios)"
        omitidos["bitacora.db"] = estado_db
    else:
        copia_db = _copiar_base(origen_db, ultima / "bitacora.db")
        integridad, movimientos = _revisar_copia(copia_db, esperados)
    if estado_db == "la copia está más al día":
        avisos.append("La copia de la USB tiene más movimientos que este servidor; no se sobrescribió. "
                      "Revisa si esa USB viene de otra PC antes de forzar (--forzar).")

    archivos = {"bitacora.db": {"bytes": copia_db.stat().st_size,
                                "sha256": _sha256(copia_db)}}
    inventarios_copiados = {}
    if con_inventarios:
        for tipo in sorted(inventarios.TIPOS, key=lambda t: inventarios.TIPOS[t]["orden"]):
            ruta = inventarios.ruta(tipo, origen)
            if not ruta.exists():
                inventarios_copiados[tipo] = {"existe": False}
                continue
            info = estado["archivos"].get(ruta.name, {})
            copiada = ultima / ruta.name
            saltar = (solo_recientes and not forzar and copiada.exists()
                      and info.get("estado") in ("igual", "la copia está más al día"))
            if saltar:
                omitidos[ruta.name] = info["estado"]
                inventarios_copiados[tipo] = {"existe": True, "archivo": ruta.name,
                                              "estado": info["estado"], **_lotes(copiada, tipo)}
                if info["estado"] == "la copia está más al día":
                    avisos.append(f"{ruta.name}: la copia de la USB es más nueva que la del servidor; "
                                  f"no se sobrescribió.")
            else:
                shutil.copy2(ruta, copiada)
                inventarios_copiados[tipo] = {"existe": True, "archivo": ruta.name,
                                              "estado": "copiado", **_lotes(ruta, tipo)}
            archivos[ruta.name] = {"bytes": copiada.stat().st_size,
                                   "sha256": _sha256(copiada)}

    manifiesto = {
        "generado": dt.datetime.now().isoformat(sep=" ", timespec="seconds"),
        "origen": str(origen.resolve()),
        "destino": str(ultima.resolve()),
        "movimientos": movimientos,
        "integridad": integridad,
        "usuarios": _usuarios(copia_db),
        "inventarios": inventarios_copiados,
        "archivos": archivos,
        "omitidos": omitidos,
        "al_dia": not any(v in ("falta en la copia", "hay cambios por copiar")
                          for v in omitidos.values()),
        "avisos": avisos,
    }
    (ultima / NOMBRE_JSON).write_text(json.dumps(manifiesto, indent=2, ensure_ascii=False),
                                      encoding="utf-8")
    leeme = chr(10).join([
        "COPIA DE LOS DATOS DEL ALMACÉN MIGH",
        f"Generada: {manifiesto['generado']}",
        f"Origen  : {manifiesto['origen']}",
        "",
        f"ultima/bitacora.db   movimientos y usuarios hasta ese momento ({movimientos} movimientos)",
        "ultima/Inventario_*.xlsx   inventarios con su columna UBICACIÓN",
        "ultima/sincronizacion.json   tamaños, SHA-256 y conteos de lo copiado",
        "historial/   copias con fecha (si se pidió --historial)",
        "",
        "Es una COPIA de lectura: el servidor sigue siendo la única fuente de verdad.",
        "Para abrirla use DB Browser for SQLite (https://sqlitebrowser.org) o:",
        "  python scripts/sincronizar_usb.py --resumen",
        "",
    ])
    (destino / "LEEME.txt").write_text(leeme, encoding="utf-8")

    if historial:
        carpeta = destino / "historial"
        carpeta.mkdir(parents=True, exist_ok=True)
        marca = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        nombre = f"bitacora_{marca}.db"
        n = 2
        while (carpeta / nombre).exists():          # dos corridas en el mismo segundo
            nombre = f"bitacora_{marca}-{n}.db"
            n += 1
        shutil.copy2(copia_db, carpeta / nombre)
        _podar(carpeta, mantener)

    manifiesto["copia"] = str(ultima)
    return manifiesto


def _usuarios(ruta_db):
    con = sqlite3.connect(str(ruta_db), timeout=10)
    try:
        return con.execute("SELECT COUNT(*) FROM usuario").fetchone()[0]
    finally:
        con.close()


def _lotes(ruta_xlsx, tipo):
    """Cuenta renglones del inventario copiado sin cargarlo completo en la app."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta_xlsx, read_only=True, data_only=True)
        hoja = wb.worksheets[0]
        filas = 0
        for fila in hoja.iter_rows(min_row=4, max_col=1, values_only=True):
            if fila and fila[0] not in (None, ""):
                filas += 1
        wb.close()
        return {"lotes": filas}
    except Exception as e:  # noqa: BLE001 - el inventario copiado no debe tumbar la sincronización
        return {"lotes": None, "error": str(e)}


def resumen(ruta_db):
    """Lee una copia (o la base real, en modo lectura) y devuelve qué contiene."""
    ruta_db = pathlib.Path(ruta_db)
    if not ruta_db.exists():
        raise FileNotFoundError(f"no existe la base: {ruta_db}")
    con = sqlite3.connect(f"file:{ruta_db}?mode=ro", uri=True, timeout=10)
    con.row_factory = sqlite3.Row
    try:
        integridad = con.execute("PRAGMA integrity_check").fetchone()[0]
        total = con.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0]
        por_inv = {fila["inventario"] or "mp": fila["n"] for fila in con.execute(
            "SELECT inventario, COUNT(*) AS n FROM movimiento GROUP BY inventario")}
        primero = con.execute("SELECT ts FROM movimiento ORDER BY id LIMIT 1").fetchone()
        ultimo = con.execute("SELECT * FROM movimiento ORDER BY id DESC LIMIT 1").fetchone()
        activos = con.execute("SELECT COUNT(*) FROM usuario WHERE activo=1").fetchone()[0]
        return {
            "archivo": str(ruta_db),
            "integridad": integridad,
            "movimientos": total,
            "por_inventario": por_inv,
            "usuarios_activos": activos,
            "desde": primero["ts"] if primero else None,
            "ultimo": dict(ultimo) if ultimo else None,
        }
    finally:
        con.close()


def _imprimir(manifiesto):
    print(f"Copia    : {manifiesto['copia']}")
    print(f"Origen   : {manifiesto['origen']}")
    print(f"Movim.   : {manifiesto['movimientos']} · integridad {manifiesto['integridad']}")
    for tipo, datos in manifiesto["inventarios"].items():
        if datos.get("existe"):
            estado = datos.get("estado", "")
            print(f"  {tipo.upper():<3} {datos['archivo']:<24} {datos.get('lotes')} lotes"
                  + (f" · {estado}" if estado else ""))
        else:
            print(f"  {tipo.upper():<3} (sin archivo)")
    if manifiesto.get("omitidos"):
        print("Sin cambios: " + ", ".join(f"{k} ({v})" for k, v in manifiesto["omitidos"].items()))
    for aviso in manifiesto.get("avisos", []):
        print(f"AVISO: {aviso}")


def _imprimir_estado(estado):
    base = estado["base"]
    print(f"Origen   : {estado['origen']}")
    print(f"Destino  : {estado['destino']}"
          + ("" if estado["existe_destino"] else "  (no existe todavía)"))
    if estado.get("ultima_copia"):
        print(f"Última copia: {estado['ultima_copia']}")
    print(f"Base     : {base['estado']} — {base['detalle']}")
    for nombre, datos in estado["archivos"].items():
        print(f"  {nombre:<24} {datos['estado']:<26} {datos['bytes_origen']} B "
              f"(copia {datos['bytes_copia']} B)")
    if estado["al_dia"]:
        print("Resultado: no hay nada que copiar (la copia está al día).")
    else:
        print("Resultado: hay que copiar " + ", ".join(estado["pendientes"]))
    if estado["adelantados"]:
        print("Ojo: la copia de la USB va más adelante en " + ", ".join(estado["adelantados"])
              + ". No se sobrescribe; usa --forzar solo si sabes que esa copia no sirve.")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Copia la base del almacén a la USB")
    ap.add_argument("--origen", default=None)
    ap.add_argument("--destino", default=None, help=f"por omisión {DESTINO}")
    ap.add_argument("--historial", action="store_true", help="guarda además una copia con fecha")
    ap.add_argument("--mantener", type=int, default=MANTENER,
                    help="cuántas copias del historial conservar (0 = todas)")
    ap.add_argument("--sin-inventarios", action="store_true")
    ap.add_argument("--resumen", action="store_true", help="solo leer la copia y reportar")
    ap.add_argument("--estado", action="store_true",
                    help="comparar servidor y copia sin copiar nada")
    ap.add_argument("--forzar", action="store_true",
                    help="sobrescribir aunque la copia parezca más nueva")
    args = ap.parse_args(argv)

    if args.estado:
        _imprimir_estado(comparar(args.origen, args.destino))
        return 0

    if args.resumen:
        destino = pathlib.Path(args.destino or DESTINO) / "ultima" / "bitacora.db"
        datos = resumen(destino)
        print(f"Base     : {datos['archivo']}")
        print(f"Integrid. : {datos['integridad']}")
        print(f"Movim.   : {datos['movimientos']} {datos['por_inventario']}")
        print(f"Usuarios : {datos['usuarios_activos']} activos")
        print(f"Desde    : {datos['desde']}")
        if datos["ultimo"]:
            u = datos["ultimo"]
            print(f"Último   : #{u['id']} {u['ts']} {u['accion']} {u['descripcion']} "
                  f"({u['usuario']})")
        return 0

    try:
        manifiesto = sincronizar(args.origen, args.destino, historial=args.historial,
                                 con_inventarios=not args.sin_inventarios,
                                 mantener=args.mantener)
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        return 2
    except RuntimeError as e:
        print(f"ERROR: {e}")
        return 3
    _imprimir(manifiesto)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
