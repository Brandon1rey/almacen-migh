"""Bitácora append-only de movimientos + usuarios y sesiones (una sola base SQLite)."""
import datetime as dt
import hashlib
import json
import pathlib
import secrets
import sqlite3
import threading

RAIZ = pathlib.Path(__file__).resolve().parents[1]
RUTA_DB = RAIZ / "data" / "bitacora.db"
_LOCK = threading.Lock()

ACCIONES = ("ALTA", "UBICAR", "MOVER", "PARCIAL", "SALIDA", "CONFIRMAR",
            "NO_ENCONTRADO", "LIMPIAR", "IMPORTAR", "DESHACER", "PURGAR")
ROLES = ("admin", "almacen", "consulta")

ESQUEMA = """
CREATE TABLE IF NOT EXISTS movimiento (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    ts                TEXT NOT NULL,
    accion            TEXT NOT NULL CHECK (accion IN
                      ('ALTA','UBICAR','MOVER','PARCIAL','SALIDA','CONFIRMAR','NO_ENCONTRADO',
                       'LIMPIAR','IMPORTAR','DESHACER','PURGAR')),
    inventario        TEXT NOT NULL DEFAULT 'mp',
    fila              INTEGER,
    descripcion       TEXT NOT NULL,
    lote              TEXT,
    ubicacion_origen  TEXT,
    ubicacion_destino TEXT,
    cantidad          REAL,
    unidad            TEXT,
    fila_destino      INTEGER,
    usuario           TEXT NOT NULL,
    usuario_id        INTEGER,
    nota              TEXT,
    motivo            TEXT,
    ref_doc           TEXT,
    revertido         INTEGER NOT NULL DEFAULT 0,
    revertido_por     INTEGER
);
CREATE INDEX IF NOT EXISTS ix_mov_ts ON movimiento(ts);
CREATE INDEX IF NOT EXISTS ix_mov_lote ON movimiento(lote);
CREATE INDEX IF NOT EXISTS ix_mov_fila ON movimiento(fila);

CREATE TABLE IF NOT EXISTS usuario (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre        TEXT NOT NULL UNIQUE COLLATE NOCASE,
    rol           TEXT NOT NULL CHECK (rol IN ('admin','almacen','consulta')),
    hash          TEXT NOT NULL,
    salt          TEXT NOT NULL,
    iteraciones   INTEGER NOT NULL,
    activo        INTEGER NOT NULL DEFAULT 1,
    debe_cambiar  INTEGER NOT NULL DEFAULT 0,
    creado        TEXT NOT NULL,
    ultimo_acceso TEXT
);

CREATE TABLE IF NOT EXISTS sesion (
    token_hash TEXT PRIMARY KEY,
    usuario_id INTEGER NOT NULL REFERENCES usuario(id),
    creado     TEXT NOT NULL,
    expira     TEXT NOT NULL,
    origen     TEXT
);
CREATE INDEX IF NOT EXISTS ix_sesion_expira ON sesion(expira);

CREATE TABLE IF NOT EXISTS auditoria (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    ts      TEXT NOT NULL,
    usuario TEXT NOT NULL,
    accion  TEXT NOT NULL,
    detalle TEXT
);
"""


def _ahora():
    return dt.datetime.now().isoformat(sep=" ", timespec="seconds")


def _entero(valor, defecto=0, minimo=0, maximo=1000):
    """Convierte a entero sin romper la petición: la API nunca debe caer en un 500 por esto."""
    try:
        numero = int(valor)
    except (TypeError, ValueError):
        return defecto
    return max(minimo, min(maximo, numero))


def conectar(ruta=None, dsn=None):
    """Una conexión compartida: PostgreSQL/Supabase si hay DSN, o SQLite en modo local."""
    from app import db
    url = dsn or (db.url_postgres() if ruta is None else None)
    if url:
        return db.PostgresConnection(url)
    ruta = pathlib.Path(ruta or RUTA_DB)
    ruta.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(ruta), timeout=10, check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.executescript(ESQUEMA)
    _migrar(con)
    return con


def _migrar(con):
    """Pone al día bases creadas con versiones anteriores, sin perder movimientos."""
    columnas = {fila["name"] for fila in con.execute("PRAGMA table_info(movimiento)")}
    nuevas = {
        "inventario": "TEXT NOT NULL DEFAULT 'mp'",
        "motivo": "TEXT", "ref_doc": "TEXT", "usuario_id": "INTEGER",
        "cantidad": "REAL", "unidad": "TEXT", "fila_destino": "INTEGER",
        "revertido": "INTEGER NOT NULL DEFAULT 0", "revertido_por": "INTEGER",
    }
    for columna, tipo in nuevas.items():
        if columna not in columnas:
            con.execute(f"ALTER TABLE movimiento ADD COLUMN {columna} {tipo}")

    definicion = con.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='movimiento'").fetchone()
    # se compara contra la lista vigente: si falta cualquier acción, se reconstruye la tabla
    faltantes = [a for a in ACCIONES if f"'{a}'" not in (definicion["sql"] or "")]
    if definicion and faltantes:
        _reconstruir_movimiento(con)
    # el índice va aquí, después del ALTER: en bases viejas la columna todavía no existía
    con.execute("CREATE INDEX IF NOT EXISTS ix_mov_inv ON movimiento(inventario, fila)")
    con.commit()


def _reconstruir_movimiento(con):
    """El CHECK de acciones no se puede alterar: se reconstruye copiando todo."""
    columnas = [fila["name"] for fila in con.execute("PRAGMA table_info(movimiento)")]
    lista = ", ".join(columnas)
    con.execute("PRAGMA foreign_keys = OFF")
    con.execute("ALTER TABLE movimiento RENAME TO movimiento_viejo")
    con.executescript(ESQUEMA)
    con.execute(f"INSERT INTO movimiento ({lista}) SELECT {lista} FROM movimiento_viejo")
    con.execute("DROP TABLE movimiento_viejo")
    con.execute("PRAGMA foreign_keys = ON")


# ---------------------------------------------------------------- bitácora
def registrar(con, accion, registro, usuario, destino=None, origen=None, nota=None,
              motivo=None, ref_doc=None, usuario_id=None, cantidad=None, fila_destino=None,
              inventario=None, unidad=None):
    if not (usuario or "").strip():
        raise ValueError("usuario requerido: la bitácora es el rastro de auditoría")
    if accion not in ACCIONES:
        raise ValueError(f"acción inválida: {accion}")
    ts = _ahora()
    with _LOCK:
        cur = con.execute(
            "INSERT INTO movimiento (ts, accion, inventario, fila, descripcion, lote, "
            "ubicacion_origen, ubicacion_destino, cantidad, unidad, fila_destino, usuario, "
            "usuario_id, nota, motivo, ref_doc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (ts, accion, inventario or registro.get("tipo") or "mp", registro.get("fila"),
             registro["descripcion"], registro.get("lote"), origen, destino, cantidad,
             unidad or registro.get("unidad"), fila_destino, usuario.strip(), usuario_id,
             nota, motivo, ref_doc))
        con.commit()
    return ts, cur.lastrowid


def ultimos(con, limite=100, inventario=None):
    sql = "SELECT * FROM movimiento"
    params = []
    if inventario:
        sql += " WHERE inventario=?"
        params.append(inventario)
    sql += " ORDER BY id DESC LIMIT ?"
    params.append(_entero(limite, 100, 1, 100000))
    with _LOCK:
        return [dict(r) for r in con.execute(sql, params).fetchall()]


def por_fila(con, fila, limite=12, inventario=None):
    sql = "SELECT * FROM movimiento WHERE fila=?"
    params = [_entero(fila, -1)]
    if inventario:
        sql += " AND inventario=?"
        params.append(inventario)
    sql += " ORDER BY id DESC LIMIT ?"
    params.append(_entero(limite, 12, 1, 500))
    with _LOCK:
        return [dict(r) for r in con.execute(sql, params).fetchall()]


def por_id(con, movimiento_id):
    ident = _entero(movimiento_id, 0, 0, 2**31 - 1)
    if ident <= 0:
        return None
    with _LOCK:
        fila = con.execute("SELECT * FROM movimiento WHERE id=?", (ident,)).fetchone()
    return dict(fila) if fila else None


def ultimo_vigente(con, fila=None, inventario=None):
    """Último movimiento no revertido (de un lote o de todo el almacén)."""
    sql = "SELECT * FROM movimiento WHERE revertido=0 AND accion <> 'DESHACER'"
    params = []
    if fila is not None:
        sql += " AND fila=?"
        params.append(_entero(fila, -1))
    if inventario:
        sql += " AND inventario=?"
        params.append(inventario)
    sql += " ORDER BY id DESC LIMIT 1"
    with _LOCK:
        resultado = con.execute(sql, params).fetchone()
    return dict(resultado) if resultado else None


def marcar_revertido(con, movimiento_id, por_movimiento):
    with _LOCK:
        con.execute("UPDATE movimiento SET revertido=1, revertido_por=? WHERE id=?",
                    (por_movimiento, _entero(movimiento_id, 0, 0, 2**31 - 1)))
        con.commit()


def borrar(con, movimiento_id):
    """Borrado físico: sólo lo usa la purga de administrador (queda en auditoría)."""
    ident = _entero(movimiento_id, 0, 0, 2**31 - 1)
    if ident <= 0:
        return None
    with _LOCK:
        fila = con.execute("SELECT * FROM movimiento WHERE id=?", (ident,)).fetchone()
        if fila is None:
            return None
        con.execute("DELETE FROM movimiento WHERE id=?", (ident,))
        con.commit()
    return dict(fila)


def contar(con, inventario=None):
    with _LOCK:
        if inventario:
            return con.execute("SELECT COUNT(*) FROM movimiento WHERE inventario=?",
                               (inventario,)).fetchone()[0]
        return con.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0]


# ---------------------------------------------------------------- sesiones
def abrir_sesion(con, usuario_id, horas=12, origen=None):
    token = secrets.token_urlsafe(32)
    ahora = dt.datetime.now()
    expira = (ahora + dt.timedelta(hours=horas)).isoformat(sep=" ", timespec="seconds")
    with _LOCK:
        con.execute("DELETE FROM sesion WHERE expira < ?", (_ahora(),))
        con.execute("INSERT INTO sesion (token_hash, usuario_id, creado, expira, origen) "
                    "VALUES (?,?,?,?,?)",
                    (hashlib.sha256(token.encode()).hexdigest(), usuario_id, _ahora(),
                     expira, origen))
        con.execute("UPDATE usuario SET ultimo_acceso=? WHERE id=?", (_ahora(), usuario_id))
        con.commit()
    return token, expira


def sesion_valida(con, token):
    if not token:
        return None
    digest = hashlib.sha256(token.encode()).hexdigest()
    with _LOCK:
        fila = con.execute("""
            SELECT s.token_hash, s.expira, u.id AS usuario_id, u.nombre, u.rol,
                   u.activo, u.debe_cambiar
            FROM sesion s JOIN usuario u ON u.id = s.usuario_id
            WHERE s.token_hash=?""", (digest,)).fetchone()
    if not fila or not fila["activo"]:
        return None
    if fila["expira"] < _ahora():
        cerrar_sesion(con, token)
        return None
    return dict(fila)


def cerrar_sesion(con, token):
    if not token:
        return
    digest = hashlib.sha256(token.encode()).hexdigest()
    with _LOCK:
        con.execute("DELETE FROM sesion WHERE token_hash=?", (digest,))
        con.commit()


def auditar(con, usuario, accion, detalle):
    with _LOCK:
        con.execute("INSERT INTO auditoria (ts, usuario, accion, detalle) VALUES (?,?,?,?)",
                    (_ahora(), usuario, accion,
                     detalle if isinstance(detalle, str)
                     else json.dumps(detalle, default=str, ensure_ascii=False)))
        con.commit()
