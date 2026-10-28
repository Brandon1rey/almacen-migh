"""Genera el SQL de migración + catálogos base para Supabase o los aplica directo.

Por defecto genera un lienzo limpio ("clean canvas"):
  - Crea las 7 tablas, índices, triggers y políticas RLS (`supabase/schema.sql`).
  - Limpia `inventario_lote`, `movimiento`, `auditoria` y `sesion` (0 renglones de inventario).
  - Siembra ÚNICAMENTE los datos de clasificación y acceso:
      1. `ubicaciones` (383 posiciones de racks y zonas de ALM-PNO-002)
      2. `catalogo_insumos` (565 claves y descripciones de productos del ERP)
      3. `usuario` (cuentas de usuario para poder iniciar sesión en Vercel)

Uso:
  1) Generar SQL de lienzo limpio (en `supabase/clean_canvas.sql` y `supabase/seed.sql`):
       python scripts/migrar_supabase.py

  2) Incluir también los lotes actuales de los Excel (`data/Inventario_*.xlsx`):
       python scripts/migrar_supabase.py --con-inventario

  3) Aplicar directo contra Supabase PostgreSQL (usando SUPABASE_DB_URL / DATABASE_URL):
       python scripts/migrar_supabase.py --aplicar
       python scripts/migrar_supabase.py --aplicar --dsn "postgresql://postgres..."
"""
from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from app import db, inventarios, ubicaciones  # noqa: E402
from app.inventario import Inventario  # noqa: E402


def _sql_str(valor: object) -> str:
    if valor is None:
        return "NULL"
    if isinstance(valor, bool):
        return "TRUE" if valor else "FALSE"
    if isinstance(valor, (int, float)):
        return str(valor)
    texto = str(valor).replace("'", "''")
    return f"'{texto}'"


def _cargar_catalogo_json(carpeta_datos: Path) -> list[dict]:
    ruta = carpeta_datos / "catalogo_insumos.json"
    if not ruta.exists():
        return []
    datos = json.loads(ruta.read_text(encoding="utf-8"))
    salida: list[dict] = []
    for item in datos.get("insumos", []):
        sku = str(item.get("sku", "")).strip()
        descripcion = str(item.get("descripcion", "")).strip()
        if not sku or not descripcion:
            continue
        inv = item.get("inventario")
        if inv not in ("mp", "mep", "pt"):
            inv = None
        salida.append({
            "sku": sku,
            "descripcion": descripcion,
            "tipo": item.get("tipo"),
            "inventario": inv,
            "unidad": item.get("unidad"),
            "fabricante": item.get("fabricante"),
            "barra": item.get("barra"),
            "maximo": item.get("maximo"),
            "minimo": item.get("minimo"),
        })
    return salida


def _cargar_inventarios_excel(carpeta_datos: Path) -> list[dict]:
    lotes: list[dict] = []
    for tipo, meta in inventarios.TIPOS.items():
        ruta = carpeta_datos / meta["archivo"]
        if not ruta.exists():
            continue
        inv = Inventario(ruta, tipo=tipo)
        for f in inv.filas:
            legible = bool(f.get("cantidad_legible", True))
            cantidad_num = float(f["cantidad"]) if legible and f.get("cantidad") is not None else None
            cantidad_txt = str(f["cantidad"]) if not legible and f.get("cantidad") is not None else None
            lotes.append({
                "inventario": tipo,
                "fila": int(f["fila"]),
                "descripcion": f["descripcion"],
                "categoria": f.get("categoria"),
                "proveedor": f.get("proveedor"),
                "unidad": f.get("unidad") or meta["unidad"],
                "cantidad": cantidad_num,
                "cantidad_texto": cantidad_txt,
                "cantidad_legible": legible,
                "cantidad_formula": bool(f.get("cantidad_formula", False)),
                "caducidad": f.get("caducidad"),
                "caducidad_cruda": f.get("caducidad_cruda"),
                "motivo_caducidad": f.get("motivo_caducidad"),
                "lote": f.get("lote"),
                "erp": f.get("erp"),
                "qr_id": f.get("qr_id"),
                "observaciones": f.get("observaciones"),
                "ubicacion": f.get("ubicacion"),
                "ubicacion_cruda": f.get("ubicacion_cruda"),
                "ubicacion_invalida": f.get("ubicacion_invalida"),
                "presentacion": f.get("presentacion"),
                "piezas": f.get("piezas"),
            })
    return lotes


def _cargar_sqlite(carpeta_datos: Path) -> dict[str, list[dict]]:
    ruta_db = carpeta_datos / "bitacora.db"
    resultado: dict[str, list[dict]] = {
        "usuario": [],
        "movimiento": [],
        "auditoria": [],
    }
    if not ruta_db.exists():
        return resultado
    con = sqlite3.connect(str(ruta_db))
    con.row_factory = sqlite3.Row
    try:
        for tabla in ("usuario", "movimiento", "auditoria"):
            try:
                filas = con.execute(f"SELECT * FROM {tabla} ORDER BY id ASC").fetchall()
                resultado[tabla] = [dict(f) for f in filas]
            except sqlite3.OperationalError:
                resultado[tabla] = []
    finally:
        con.close()
    return resultado


def construir_sql_seed(
    carpeta_datos: Path | None = None,
    incluir_inventario: bool = False,
) -> str:
    carpeta = carpeta_datos or (RAIZ / "data")
    esquema_path = RAIZ / "supabase" / "schema.sql"
    esquema_sql = esquema_path.read_text(encoding="utf-8")

    modo_titulo = (
        "MIGRACIÓN COMPLETA + INVENTARIO ACTUAL"
        if incluir_inventario
        else "LIENZO LIMPIO (CLEAN CANVAS): ESQUEMA + UBICACIONES + CATÁLOGO DE PRODUCTOS (SIN LOTES)"
    )

    lineas: list[str] = [
        "-- =============================================================================",
        f"-- {modo_titulo}",
        "-- Generado automáticamente por scripts/migrar_supabase.py",
        "-- Puedes pegar este archivo completo en el SQL Editor de Supabase o ejecutarlo",
        "-- con: psql \"$SUPABASE_DB_URL\" -f supabase/clean_canvas.sql",
        "-- =============================================================================",
        "",
        esquema_sql.strip(),
        "",
        "BEGIN;",
        "",
        "-- -----------------------------------------------------------------------------",
        "-- 0. Lienzo limpio: vacía lotes, movimientos, sesiones y auditoría previa",
        "-- -----------------------------------------------------------------------------",
        "TRUNCATE TABLE",
        "  public.inventario_lote,",
        "  public.movimiento,",
        "  public.auditoria,",
        "  public.sesion",
        "RESTART IDENTITY CASCADE;",
        "",
        "-- -----------------------------------------------------------------------------",
        f"-- 1. Ubicaciones de Racks y Zonas ({len(ubicaciones.CATALOGO)} códigos oficiales ALM-PNO-002)",
        "-- -----------------------------------------------------------------------------",
    ]

    valores_ubi = [
        f"  ({_sql_str(u['codigo'])}, {_sql_str(u['zona'])}, "
        f"{_sql_str(u['tipo'])}, {_sql_str(bool(u['permite_alergenos']))})"
        for u in ubicaciones.CATALOGO.values()
    ]
    if valores_ubi:
        lineas.append(
            "INSERT INTO public.ubicaciones (codigo, zona, tipo, permite_alergenos)\nVALUES\n"
            + ",\n".join(valores_ubi)
            + "\nON CONFLICT (codigo) DO UPDATE SET\n"
            "  zona = EXCLUDED.zona,\n"
            "  tipo = EXCLUDED.tipo,\n"
            "  permite_alergenos = EXCLUDED.permite_alergenos;"
        )

    insumos = _cargar_catalogo_json(carpeta)
    lineas.extend([
        "",
        "-- -----------------------------------------------------------------------------",
        f"-- 2. Catálogo de códigos de producto del ERP ({len(insumos)} registros)",
        "-- -----------------------------------------------------------------------------",
    ])
    if insumos:
        valores_cat = [
            f"  ({_sql_str(i['sku'])}, {_sql_str(i['descripcion'])}, {_sql_str(i['tipo'])}, "
            f"{_sql_str(i['inventario'])}, {_sql_str(i['unidad'])}, {_sql_str(i['fabricante'])}, "
            f"{_sql_str(i['barra'])}, {_sql_str(i['maximo'])}, {_sql_str(i['minimo'])})"
            for i in insumos
        ]
        lineas.append(
            "INSERT INTO public.catalogo_insumos (\n"
            "  sku, descripcion, tipo, inventario, unidad, fabricante, barra, maximo, minimo\n"
            ")\nVALUES\n"
            + ",\n".join(valores_cat)
            + "\nON CONFLICT (sku) DO UPDATE SET\n"
            "  descripcion = EXCLUDED.descripcion,\n"
            "  tipo = EXCLUDED.tipo,\n"
            "  inventario = EXCLUDED.inventario,\n"
            "  unidad = EXCLUDED.unidad,\n"
            "  fabricante = EXCLUDED.fabricante,\n"
            "  barra = EXCLUDED.barra,\n"
            "  maximo = EXCLUDED.maximo,\n"
            "  minimo = EXCLUDED.minimo,\n"
            "  actualizado = NOW();"
        )

    datos_sqlite = _cargar_sqlite(carpeta)
    usuarios_db = datos_sqlite["usuario"]
    lineas.extend([
        "",
        "-- -----------------------------------------------------------------------------",
        f"-- 3. Usuarios del sistema ({len(usuarios_db)} cuentas)",
        "-- -----------------------------------------------------------------------------",
    ])
    if usuarios_db:
        valores_usr = [
            f"  ({_sql_str(u['id'])}, {_sql_str(u['nombre'])}, {_sql_str(u['rol'])}, "
            f"{_sql_str(u['hash'])}, {_sql_str(u['salt'])}, {_sql_str(u['iteraciones'])}, "
            f"{_sql_str(u['activo'])}, {_sql_str(u.get('debe_cambiar', 0))}, "
            f"{_sql_str(u['creado'])}, {_sql_str(u.get('ultimo_acceso'))})"
            for u in usuarios_db
        ]
        lineas.append(
            "INSERT INTO public.usuario (\n"
            "  id, nombre, rol, hash, salt, iteraciones, activo, debe_cambiar, creado, ultimo_acceso\n"
            ")\nVALUES\n"
            + ",\n".join(valores_usr)
            + "\nON CONFLICT (nombre) DO UPDATE SET\n"
            "  rol = EXCLUDED.rol,\n"
            "  hash = EXCLUDED.hash,\n"
            "  salt = EXCLUDED.salt,\n"
            "  iteraciones = EXCLUDED.iteraciones,\n"
            "  activo = EXCLUDED.activo,\n"
            "  debe_cambiar = EXCLUDED.debe_cambiar,\n"
            "  ultimo_acceso = EXCLUDED.ultimo_acceso;"
        )
        lineas.append(
            "SELECT setval(pg_get_serial_sequence('public.usuario', 'id'), "
            "COALESCE((SELECT MAX(id) FROM public.usuario), 1), true);"
        )

    if incluir_inventario:
        lotes = _cargar_inventarios_excel(carpeta)
        lineas.extend([
            "",
            "-- -----------------------------------------------------------------------------",
            f"-- 4. Inventario de lotes ({len(lotes)} renglones)",
            "-- -----------------------------------------------------------------------------",
        ])
        if lotes:
            valores_lotes = [
                f"  ({_sql_str(l['inventario'])}, {_sql_str(l['fila'])}, {_sql_str(l['descripcion'])}, "
                f"{_sql_str(l['categoria'])}, {_sql_str(l['proveedor'])}, {_sql_str(l['unidad'])}, "
                f"{_sql_str(l['cantidad'])}, {_sql_str(l['cantidad_texto'])}, "
                f"{_sql_str(l['cantidad_legible'])}, {_sql_str(l['cantidad_formula'])}, "
                f"{_sql_str(l['caducidad'])}, {_sql_str(l['caducidad_cruda'])}, "
                f"{_sql_str(l['motivo_caducidad'])}, {_sql_str(l['lote'])}, {_sql_str(l['erp'])}, "
                f"{_sql_str(l['qr_id'])}, {_sql_str(l['observaciones'])}, {_sql_str(l['ubicacion'])}, "
                f"{_sql_str(l['ubicacion_cruda'])}, {_sql_str(l['ubicacion_invalida'])}, "
                f"{_sql_str(l['presentacion'])}, {_sql_str(l['piezas'])})"
                for l in lotes
            ]
            lineas.append(
                "INSERT INTO public.inventario_lote (\n"
                "  inventario, fila, descripcion, categoria, proveedor, unidad,\n"
                "  cantidad, cantidad_texto, cantidad_legible, cantidad_formula,\n"
                "  caducidad, caducidad_cruda, motivo_caducidad, lote, erp, qr_id,\n"
                "  observaciones, ubicacion, ubicacion_cruda, ubicacion_invalida,\n"
                "  presentacion, piezas\n"
                ")\nVALUES\n"
                + ",\n".join(valores_lotes)
                + "\nON CONFLICT (inventario, fila) DO UPDATE SET\n"
                "  descripcion = EXCLUDED.descripcion,\n"
                "  categoria = EXCLUDED.categoria,\n"
                "  proveedor = EXCLUDED.proveedor,\n"
                "  unidad = EXCLUDED.unidad,\n"
                "  cantidad = EXCLUDED.cantidad,\n"
                "  cantidad_texto = EXCLUDED.cantidad_texto,\n"
                "  cantidad_legible = EXCLUDED.cantidad_legible,\n"
                "  cantidad_formula = EXCLUDED.cantidad_formula,\n"
                "  caducidad = EXCLUDED.caducidad,\n"
                "  caducidad_cruda = EXCLUDED.caducidad_cruda,\n"
                "  motivo_caducidad = EXCLUDED.motivo_caducidad,\n"
                "  lote = EXCLUDED.lote,\n"
                "  erp = EXCLUDED.erp,\n"
                "  qr_id = EXCLUDED.qr_id,\n"
                "  observaciones = EXCLUDED.observaciones,\n"
                "  ubicacion = EXCLUDED.ubicacion,\n"
                "  ubicacion_cruda = EXCLUDED.ubicacion_cruda,\n"
                "  ubicacion_invalida = EXCLUDED.ubicacion_invalida,\n"
                "  presentacion = EXCLUDED.presentacion,\n"
                "  piezas = EXCLUDED.piezas,\n"
                "  actualizado = NOW();"
            )

    lineas.extend(["", "COMMIT;", ""])
    return "\n".join(lineas)


def aplicar_en_supabase(dsn: str, sql: str) -> None:
    import psycopg2  # type: ignore

    con = psycopg2.connect(dsn)
    con.autocommit = True
    try:
        with con.cursor() as cur:
            cur.execute(sql)
    finally:
        con.close()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genera el SQL de Supabase (lienzo limpio por defecto: solo esquema + ubicaciones + catálogo)."
    )
    parser.add_argument(
        "--salida",
        default=None,
        help="Ruta del archivo SQL a generar (por defecto actualiza supabase/clean_canvas.sql y supabase/seed.sql)",
    )
    parser.add_argument(
        "--con-inventario",
        action="store_true",
        help="Incluye también los renglones actuales de data/Inventario_*.xlsx.",
    )
    parser.add_argument(
        "--aplicar",
        action="store_true",
        help="Conecta a SUPABASE_DB_URL / DATABASE_URL y ejecuta el SQL en Supabase.",
    )
    parser.add_argument(
        "--dsn",
        default=None,
        help="URL de conexión PostgreSQL (opcional, sobreescribe SUPABASE_DB_URL).",
    )
    args = parser.parse_args()

    sql = construir_sql_seed(incluir_inventario=args.con_inventario)
    if args.salida:
        rutas = [Path(args.salida)]
    else:
        rutas = [
            RAIZ / "supabase" / "clean_canvas.sql",
            RAIZ / "supabase" / "seed.sql",
        ]
    for ruta_salida in rutas:
        ruta_salida.parent.mkdir(parents=True, exist_ok=True)
        ruta_salida.write_text(sql, encoding="utf-8")
        print(f"[OK] Archivo SQL generado en: {ruta_salida} ({len(sql):,} bytes)")

    if args.aplicar:
        dsn = args.dsn or db.url_postgres()
        if not dsn:
            raise SystemExit(
                "Error: Define SUPABASE_DB_URL o DATABASE_URL, o pasa --dsn 'postgresql://...'"
            )
        print("Aplicando esquema y catálogos en Supabase PostgreSQL...")
        aplicar_en_supabase(dsn, sql)
        print("[OK] Lienzo limpio aplicado exitosamente en Supabase.")


if __name__ == "__main__":
    main()
