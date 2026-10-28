-- =============================================================================
-- Control de Ubicación de Almacén · Comercializadora MIGH (ALM-PNO-002)
-- Esquema PostgreSQL para Supabase (Producción + Serverless en Vercel)
-- =============================================================================

BEGIN;

-- 1. Catálogo oficial de ubicaciones de almacén (ALM-PNO-002)
CREATE TABLE IF NOT EXISTS public.ubicaciones (
    codigo            TEXT PRIMARY KEY,
    zona              TEXT NOT NULL CHECK (zona IN (
                      'MP-ALERGENO','MP-NO-ALERGENO','PT','EM','ESPECIAL')),
    tipo              TEXT NOT NULL CHECK (tipo IN (
                      'RACK','PISO','JAULA','CUARTO','ANDEN')),
    permite_alergenos BOOLEAN NOT NULL DEFAULT FALSE
);

CREATE INDEX IF NOT EXISTS ix_ubicaciones_zona ON public.ubicaciones (zona);

-- 2. Catálogo maestro de insumos del ERP (MP, MEP, PT)
CREATE TABLE IF NOT EXISTS public.catalogo_insumos (
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

CREATE INDEX IF NOT EXISTS ix_catalogo_inv  ON public.catalogo_insumos (inventario);
CREATE INDEX IF NOT EXISTS ix_catalogo_desc ON public.catalogo_insumos (LOWER(descripcion));

-- 3. Renglones de inventario por área (MP, MEP, PT)
CREATE TABLE IF NOT EXISTS public.inventario_lote (
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
    ubicacion          TEXT REFERENCES public.ubicaciones(codigo) ON UPDATE CASCADE,
    ubicacion_cruda    TEXT,
    ubicacion_invalida TEXT,
    presentacion       TEXT,
    piezas             TEXT,
    actualizado        TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE (inventario, fila)
);

CREATE INDEX IF NOT EXISTS ix_inv_lote_tipo_fila ON public.inventario_lote (inventario, fila);
CREATE INDEX IF NOT EXISTS ix_inv_lote_ubicacion ON public.inventario_lote (inventario, ubicacion);
CREATE INDEX IF NOT EXISTS ix_inv_lote_qr        ON public.inventario_lote (qr_id);
CREATE INDEX IF NOT EXISTS ix_inv_lote_lote      ON public.inventario_lote (lote);

-- Trigger para mantener actualizado el timestamp en inventario_lote
CREATE OR REPLACE FUNCTION public.fn_inventario_lote_touch()
RETURNS TRIGGER
LANGUAGE plpgsql
AS $$
BEGIN
    NEW.actualizado := NOW();
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS trg_inventario_lote_touch ON public.inventario_lote;
CREATE TRIGGER trg_inventario_lote_touch
BEFORE UPDATE ON public.inventario_lote
FOR EACH ROW
EXECUTE FUNCTION public.fn_inventario_lote_touch();

-- 4. Usuarios (autenticación PBKDF2-HMAC-SHA256 y control de roles)
CREATE TABLE IF NOT EXISTS public.usuario (
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

CREATE UNIQUE INDEX IF NOT EXISTS ix_usuario_nombre_lower
    ON public.usuario (LOWER(nombre));

-- 5. Sesiones activas (cookie HttpOnly almacen_sesion)
CREATE TABLE IF NOT EXISTS public.sesion (
    token_hash TEXT PRIMARY KEY,
    usuario_id BIGINT NOT NULL REFERENCES public.usuario(id) ON DELETE CASCADE,
    creado     TEXT NOT NULL,
    expira     TEXT NOT NULL,
    origen     TEXT
);

CREATE INDEX IF NOT EXISTS ix_sesion_expira ON public.sesion (expira);

-- 6. Bitácora inmutable de movimientos de almacén
CREATE TABLE IF NOT EXISTS public.movimiento (
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

CREATE INDEX IF NOT EXISTS ix_mov_ts   ON public.movimiento (ts);
CREATE INDEX IF NOT EXISTS ix_mov_lote ON public.movimiento (lote);
CREATE INDEX IF NOT EXISTS ix_mov_fila ON public.movimiento (fila);
CREATE INDEX IF NOT EXISTS ix_mov_inv  ON public.movimiento (inventario, fila);

-- 7. Auditoría administrativa y de seguridad
CREATE TABLE IF NOT EXISTS public.auditoria (
    id      BIGSERIAL PRIMARY KEY,
    ts      TEXT NOT NULL,
    usuario TEXT NOT NULL,
    accion  TEXT NOT NULL,
    detalle TEXT
);

CREATE INDEX IF NOT EXISTS ix_auditoria_ts ON public.auditoria (ts);

-- 8. Row Level Security (RLS):
-- Bloquea el acceso público directo desde PostgREST (anon / authenticated).
-- El backend Python en Vercel se conecta mediante SUPABASE_DB_URL (rol postgres),
-- que opera con privilegios de servidor sin exponer las tablas al navegador.
ALTER TABLE public.ubicaciones      ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.catalogo_insumos ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.inventario_lote  ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.usuario          ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.sesion           ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.movimiento       ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.auditoria        ENABLE ROW LEVEL SECURITY;

COMMIT;
