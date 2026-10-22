"""Bitácora append-only de movimientos de ubicación (rastro de auditoría)."""
import datetime as dt
import pathlib
import sqlite3
import threading

RAIZ = pathlib.Path(__file__).resolve().parents[1]
RUTA_DB = RAIZ / "data" / "bitacora.db"
_LOCK = threading.Lock()

ESQUEMA = """
CREATE TABLE IF NOT EXISTS movimiento (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    ts                TEXT NOT NULL,
    accion            TEXT NOT NULL CHECK (accion IN
                      ('UBICAR','MOVER','CONFIRMAR','NO_ENCONTRADO','LIMPIAR','IMPORTAR')),
    fila              INTEGER,
    descripcion       TEXT NOT NULL,
    lote              TEXT,
    ubicacion_origen  TEXT,
    ubicacion_destino TEXT,
    usuario           TEXT NOT NULL,
    nota              TEXT
);
CREATE INDEX IF NOT EXISTS ix_mov_ts ON movimiento(ts);
CREATE INDEX IF NOT EXISTS ix_mov_lote ON movimiento(lote);
"""


def conectar(ruta=None):
    """Una conexión compartida entre los hilos del servidor local (pocos usuarios)."""
    ruta = pathlib.Path(ruta or RUTA_DB)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(ruta), timeout=10, check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.executescript(ESQUEMA)
    _migrar(con)
    return con


def _migrar(con):
    """Agrega columnas nuevas a bitácoras creadas con la versión anterior."""
    existentes = {fila["name"] for fila in con.execute("PRAGMA table_info(movimiento)")}
    for columna in ("motivo", "ref_doc"):
        if columna not in existentes:
            con.execute(f"ALTER TABLE movimiento ADD COLUMN {columna} TEXT")
    con.commit()


def registrar(con, accion, registro, usuario, destino=None, origen=None, nota=None,
              motivo=None, ref_doc=None):
    if not (usuario or "").strip():
        raise ValueError("usuario requerido: la bitácora es el rastro de auditoría")
    ts = dt.datetime.now().isoformat(sep=" ", timespec="seconds")
    with _LOCK:
        cur = con.execute(
            "INSERT INTO movimiento (ts, accion, fila, descripcion, lote, ubicacion_origen, "
            "ubicacion_destino, usuario, nota, motivo, ref_doc) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (ts, accion, registro.get("fila"), registro["descripcion"], registro.get("lote"),
             origen, destino, usuario.strip(), nota, motivo, ref_doc))
        con.commit()
    return ts, cur.lastrowid


def borrar(con, movimiento_id):
    with _LOCK:
        con.execute("DELETE FROM movimiento WHERE id=?", (movimiento_id,))
        con.commit()


def ultimos(con, limite=100):
    with _LOCK:
        return [dict(r) for r in con.execute(
            "SELECT * FROM movimiento ORDER BY id DESC LIMIT ?", (int(limite),)).fetchall()]


def por_fila(con, fila, limite=12):
    """Historial de un lote, del movimiento más reciente al más viejo."""
    with _LOCK:
        return [dict(r) for r in con.execute(
            "SELECT * FROM movimiento WHERE fila=? ORDER BY id DESC LIMIT ?",
            (int(fila), int(limite))).fetchall()]


def contar(con):
    with _LOCK:
        return con.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0]
