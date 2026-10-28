"""Capa de conexión a base de datos: Supabase (PostgreSQL) en producción y SQLite en modo local."""
import os
import re
import threading

VARIABLES_URL = (
    "SUPABASE_DB_URL",
    "DATABASE_URL",
    "POSTGRES_URL",
    "POSTGRES_URL_NON_POOLING",
)

ESQUEMA_POSTGRES = """
CREATE TABLE IF NOT EXISTS ubicaciones (
    codigo            TEXT PRIMARY KEY,
    zona              TEXT NOT NULL CHECK (zona IN (
                      'MP-ALERGENO','MP-NO-ALERGENO','PT','EM','ESPECIAL')),
    tipo              TEXT NOT NULL CHECK (tipo IN (
                      'RACK','PISO','JAULA','CUARTO','ANDEN')),
    permite_alergenos BOOLEAN NOT NULL DEFAULT FALSE
);
CREATE INDEX IF NOT EXISTS ix_ubicaciones_zona ON ubicaciones(zona);

CREATE TABLE IF NOT EXISTS catalogo_insumos (
    sku         TEXT PRIMARY KEY,
    descripcion TEXT NOT NULL,
    tipo        TEXT,
    inventario  TEXT CHECK (inventario IS NULL OR inventario IN ('mp','mep','pt')),
    unidad      TEXT,
    fabricante  TEXT,
    barra       TEXT,
    maximo      DOUBLE PRECISION,
    minimo      DOUBLE PRECISION,
    actualizado TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS ix_catalogo_inv ON catalogo_insumos(inventario);

CREATE TABLE IF NOT EXISTS inventario_lote (
    id                 BIGSERIAL PRIMARY KEY,
    inventario         TEXT NOT NULL CHECK (inventario IN ('mp','mep','pt')),
    fila               INTEGER NOT NULL,
    descripcion        TEXT NOT NULL,
    categoria          TEXT,
    proveedor          TEXT,
    unidad             TEXT NOT NULL DEFAULT 'kg',
    cantidad           DOUBLE PRECISION,
    cantidad_texto     TEXT,
    cantidad_legible   BOOLEAN NOT NULL DEFAULT TRUE,
    cantidad_formula   BOOLEAN NOT NULL DEFAULT FALSE,
    caducidad          TEXT,
    caducidad_cruda    TEXT,
    motivo_caducidad   TEXT,
    lote               TEXT,
    erp                TEXT,
    qr_id              TEXT,
    observaciones      TEXT,
    ubicacion          TEXT REFERENCES ubicaciones(codigo) ON UPDATE CASCADE,
    ubicacion_cruda    TEXT,
    ubicacion_invalida TEXT,
    presentacion       TEXT,
    piezas             TEXT,
    actualizado        TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE (inventario, fila)
);
CREATE INDEX IF NOT EXISTS ix_inv_lote_tipo_fila ON inventario_lote(inventario, fila);
CREATE INDEX IF NOT EXISTS ix_inv_lote_ubicacion ON inventario_lote(inventario, ubicacion);
CREATE INDEX IF NOT EXISTS ix_inv_lote_qr ON inventario_lote(qr_id);
CREATE INDEX IF NOT EXISTS ix_inv_lote_lote ON inventario_lote(lote);

CREATE TABLE IF NOT EXISTS usuario (
    id            BIGSERIAL PRIMARY KEY,
    nombre        TEXT NOT NULL UNIQUE,
    rol           TEXT NOT NULL CHECK (rol IN ('admin','almacen','consulta')),
    hash          TEXT NOT NULL,
    salt          TEXT NOT NULL,
    iteraciones   INTEGER NOT NULL,
    activo        INTEGER NOT NULL DEFAULT 1,
    debe_cambiar  INTEGER NOT NULL DEFAULT 0,
    creado        TEXT NOT NULL,
    ultimo_acceso TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS ix_usuario_nombre_lower ON usuario (LOWER(nombre));

CREATE TABLE IF NOT EXISTS sesion (
    token_hash TEXT PRIMARY KEY,
    usuario_id BIGINT NOT NULL REFERENCES usuario(id) ON DELETE CASCADE,
    creado     TEXT NOT NULL,
    expira     TEXT NOT NULL,
    origen     TEXT
);
CREATE INDEX IF NOT EXISTS ix_sesion_expira ON sesion(expira);

CREATE TABLE IF NOT EXISTS movimiento (
    id                BIGSERIAL PRIMARY KEY,
    ts                TEXT NOT NULL,
    accion            TEXT NOT NULL CHECK (accion IN (
                      'ALTA','UBICAR','MOVER','PARCIAL','SALIDA','CONFIRMAR','NO_ENCONTRADO',
                      'LIMPIAR','IMPORTAR','DESHACER','PURGAR')),
    inventario        TEXT NOT NULL DEFAULT 'mp' CHECK (inventario IN ('mp','mep','pt')),
    fila              INTEGER,
    descripcion       TEXT NOT NULL,
    lote              TEXT,
    ubicacion_origen  TEXT,
    ubicacion_destino TEXT,
    cantidad          DOUBLE PRECISION,
    unidad            TEXT,
    fila_destino      INTEGER,
    usuario           TEXT NOT NULL,
    usuario_id        BIGINT,
    nota              TEXT,
    motivo            TEXT,
    ref_doc           TEXT,
    revertido         INTEGER NOT NULL DEFAULT 0,
    revertido_por     BIGINT
);
CREATE INDEX IF NOT EXISTS ix_mov_ts ON movimiento(ts);
CREATE INDEX IF NOT EXISTS ix_mov_lote ON movimiento(lote);
CREATE INDEX IF NOT EXISTS ix_mov_fila ON movimiento(fila);
CREATE INDEX IF NOT EXISTS ix_mov_inv ON movimiento(inventario, fila);

CREATE TABLE IF NOT EXISTS auditoria (
    id      BIGSERIAL PRIMARY KEY,
    ts      TEXT NOT NULL,
    usuario TEXT NOT NULL,
    accion  TEXT NOT NULL,
    detalle TEXT
);
CREATE INDEX IF NOT EXISTS ix_auditoria_ts ON auditoria(ts);
"""


def url_postgres():
    """Devuelve la cadena de conexión PostgreSQL/Supabase si está configurada en el entorno."""
    for nombre in VARIABLES_URL:
        valor = (os.environ.get(nombre) or "").strip()
        if valor and valor.startswith(("postgres://", "postgresql://")):
            return valor
    return None


def es_vercel():
    """True cuando corre en el entorno serverless de Vercel."""
    return bool(os.environ.get("VERCEL") or os.environ.get("VERCEL_ENV"))


class PgRow(dict):
    """Fila compatible con sqlite3.Row: permite acceso por clave (`r['id']`) y por índice (`r[0]`)."""

    def __getitem__(self, clave):
        if isinstance(clave, int):
            return tuple(self.values())[clave]
        return super().__getitem__(clave)


class PgCursor:
    """Cursor envolvente que entrega objetos PgRow y expone `lastrowid`."""

    def __init__(self, filas=None, lastrowid=None, rowcount=0):
        self._filas = [PgRow(f) for f in (filas or [])]
        self._pos = 0
        self.lastrowid = lastrowid
        self.rowcount = rowcount

    def fetchone(self):
        if self._pos >= len(self._filas):
            return None
        fila = self._filas[self._pos]
        self._pos += 1
        return fila

    def fetchall(self):
        if self._pos == 0:
            self._pos = len(self._filas)
            return list(self._filas)
        resto = self._filas[self._pos:]
        self._pos = len(self._filas)
        return resto

    def __iter__(self):
        return iter(self._filas)


def _adaptar_sql(sql):
    """Traduce placeholders `?` de SQLite a `%s` de psycopg2 (fuera de literales) y limpia COLLATE."""
    texto = re.sub(r"\bCOLLATE\s+NOCASE\b", "", str(sql), flags=re.IGNORECASE)
    partes = re.split(r"('(?:''|[^'])*')", texto)
    for idx in range(0, len(partes), 2):
        partes[idx] = partes[idx].replace("?", "%s")
    texto = "".join(partes)
    es_insert_id = bool(re.match(
        r"^\s*INSERT\s+INTO\s+(movimiento|usuario|auditoria|inventario_lote)\b",
        texto, flags=re.IGNORECASE))
    if es_insert_id and "RETURNING" not in texto.upper():
        texto = texto.rstrip().rstrip(";") + " RETURNING id"
        return texto, True
    return texto, False


class PostgresConnection:
    """Conexión PostgreSQL compatible con la interfaz de sqlite3 usada en el proyecto."""

    es_postgres = True

    def __init__(self, dsn=None, inicializar=True):
        self.dsn = dsn or url_postgres()
        if not self.dsn:
            raise ValueError("no se configuró SUPABASE_DB_URL / DATABASE_URL")
        self._lock = threading.RLock()
        self._con = None
        self._conectar()
        if inicializar:
            self.executescript(ESQUEMA_POSTGRES)
            self._sembrar_ubicaciones()

    def _conectar(self):
        import psycopg2
        if self._con is not None:
            try:
                self._con.close()
            except Exception:  # noqa: BLE001
                pass
        self._con = psycopg2.connect(self.dsn)
        self._con.autocommit = False

    def _sembrar_ubicaciones(self):
        """Asegura que el catálogo físico de ubicaciones (ALM-PNO-002) esté en la base."""
        from app import ubicaciones
        fila = self.execute("SELECT COUNT(*) FROM ubicaciones").fetchone()
        if fila and fila[0] >= len(ubicaciones.CATALOGO):
            return
        import psycopg2.extras
        registros = [
            (u["codigo"], u["zona"], u["tipo"], bool(u["permite_alergenos"]))
            for u in ubicaciones.CATALOGO.values()
        ]
        with self._lock:
            with self._con.cursor() as cur:
                psycopg2.extras.execute_values(
                    cur,
                    "INSERT INTO ubicaciones (codigo, zona, tipo, permite_alergenos) VALUES %s "
                    "ON CONFLICT (codigo) DO UPDATE SET "
                    "zona = EXCLUDED.zona, tipo = EXCLUDED.tipo, "
                    "permite_alergenos = EXCLUDED.permite_alergenos",
                    registros,
                )
            self._con.commit()

    def execute(self, sql, params=()):
        import psycopg2
        import psycopg2.extras
        sql_pg, pide_id = _adaptar_sql(sql)
        parametros = tuple(params) if params is not None else ()
        with self._lock:
            for intento in range(2):
                try:
                    if self._con is None or self._con.closed:
                        self._conectar()
                    with self._con.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                        cur.execute(sql_pg, parametros)
                        filas = cur.fetchall() if cur.description else []
                        lastrowid = None
                        if pide_id and filas:
                            lastrowid = filas[0].get("id")
                        return PgCursor(filas=filas, lastrowid=lastrowid, rowcount=cur.rowcount)
                except (psycopg2.OperationalError, psycopg2.InterfaceError):
                    if intento == 0:
                        self._conectar()
                        continue
                    raise
                except Exception:
                    try:
                        self._con.rollback()
                    except Exception:  # noqa: BLE001
                        pass
                    raise

    def executemany(self, sql, seq_of_params):
        import psycopg2
        import psycopg2.extras
        sql_pg, _ = _adaptar_sql(sql)
        with self._lock:
            if self._con is None or self._con.closed:
                self._conectar()
            with self._con.cursor() as cur:
                psycopg2.extras.execute_batch(cur, sql_pg, [tuple(p) for p in seq_of_params])

    def executescript(self, script):
        import psycopg2
        with self._lock:
            for intento in range(2):
                try:
                    if self._con is None or self._con.closed:
                        self._conectar()
                    with self._con.cursor() as cur:
                        cur.execute(script)
                    self._con.commit()
                    return
                except (psycopg2.OperationalError, psycopg2.InterfaceError):
                    if intento == 0:
                        self._conectar()
                        continue
                    raise
                except Exception:
                    try:
                        self._con.rollback()
                    except Exception:  # noqa: BLE001
                        pass
                    raise

    def commit(self):
        with self._lock:
            if self._con and not self._con.closed:
                self._con.commit()

    def rollback(self):
        with self._lock:
            if self._con and not self._con.closed:
                self._con.rollback()

    def close(self):
        with self._lock:
            if self._con and not self._con.closed:
                try:
                    self._con.close()
                except Exception:  # noqa: BLE001
                    pass
            self._con = None
