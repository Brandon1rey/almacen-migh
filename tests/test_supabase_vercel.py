"""Pruebas unitarias para la capa Supabase (PostgreSQL), modo en_db y bundle Vercel."""
from __future__ import annotations

import json
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path

from app import catalogo, db, reportes
from app.inventario import Inventario


class FakePgConnection:
    """Simula PostgresConnection sobre SQLite en memoria devolviendo PgRow."""

    es_postgres = True

    def __init__(self):
        self._con = sqlite3.connect(":memory:")
        self._con.row_factory = sqlite3.Row
        self._con.executescript(
            """
            CREATE TABLE catalogo_insumos (
                sku TEXT PRIMARY KEY,
                descripcion TEXT NOT NULL,
                tipo TEXT,
                inventario TEXT,
                unidad TEXT,
                fabricante TEXT,
                barra TEXT,
                maximo REAL,
                minimo REAL,
                actualizado TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE inventario_lote (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inventario TEXT NOT NULL,
                fila INTEGER NOT NULL,
                descripcion TEXT NOT NULL,
                categoria TEXT,
                proveedor TEXT,
                unidad TEXT NOT NULL DEFAULT 'kg',
                cantidad REAL,
                cantidad_texto TEXT,
                cantidad_legible INTEGER NOT NULL DEFAULT 1,
                cantidad_formula INTEGER NOT NULL DEFAULT 0,
                caducidad TEXT,
                caducidad_cruda TEXT,
                motivo_caducidad TEXT,
                lote TEXT,
                erp TEXT,
                qr_id TEXT,
                observaciones TEXT,
                ubicacion TEXT,
                ubicacion_cruda TEXT,
                ubicacion_invalida TEXT,
                presentacion TEXT,
                piezas TEXT,
                actualizado TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (inventario, fila)
            );
            """
        )

    def execute(self, sql: str, params: tuple = ()):
        sql_sqlite = sql.replace("NOW()", "CURRENT_TIMESTAMP")
        cur = self._con.execute(sql_sqlite, params)
        filas = [db.PgRow(dict(r)) for r in cur.fetchall()] if cur.description else []
        return db.PgCursor(filas=filas, lastrowid=cur.lastrowid, rowcount=cur.rowcount)

    def commit(self):
        self._con.commit()

    def close(self):
        self._con.close()


class TestSupabaseVercel(unittest.TestCase):
    def test_adaptar_sql_y_pgrow(self):
        sql, pide_id = db._adaptar_sql(
            "INSERT INTO movimiento (ts, accion, nota) VALUES (?, ?, '¿ok?') COLLATE NOCASE"
        )
        self.assertTrue(pide_id)
        self.assertIn("%s, %s, '¿ok?'", sql)
        self.assertNotIn("COLLATE NOCASE", sql)
        self.assertTrue(sql.endswith("RETURNING id"))

        fila = db.PgRow({"id": 7, "nombre": "Admin"})
        self.assertEqual(fila["id"], 7)
        self.assertEqual(fila[0], 7)
        self.assertEqual(fila[1], "Admin")

    def test_inventario_modo_db_ciclo_completo(self):
        con = FakePgConnection()
        try:
            con.execute(
                """INSERT INTO inventario_lote
                   (inventario, fila, descripcion, categoria, unidad, cantidad,
                    cantidad_legible, caducidad, lote, erp, ubicacion)
                   VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?, ?, ?)""",
                ("mp", 7, "ÁCIDO CÍTRICO ANHIDRO", "EXCIPIENTE", "kg", 100.0,
                 "2027-05-01", "L-100", "MP-001", "MPNA-A-041"),
            )
            con.commit()

            inv = Inventario(tipo="mp", con=con)
            self.assertTrue(inv.en_db)
            self.assertEqual(len(inv.filas), 1)
            self.assertEqual(inv.por_fila(7)["ubicacion"], "MPNA-A-041")

            # Asignar nueva ubicación
            inv.asignar(7, "MPNA-A-042")
            self.assertEqual(inv.por_fila(7)["ubicacion"], "MPNA-A-042")

            # Asignar QR
            inv.asignar_qr(7, "QR-TEST-01")
            self.assertEqual(inv.por_fila(7)["qr_id"], "QR-TEST-01")

            # Parcialidad
            nuevo_reg = inv.agregar_parcialidad(7, 25.0, "MPNA-A-043")
            nueva_fila = nuevo_reg["fila"]
            self.assertEqual(inv.por_fila(7)["cantidad"], 75.0)
            self.assertEqual(nuevo_reg["cantidad"], 25.0)
            self.assertEqual(nuevo_reg["ubicacion"], "MPNA-A-043")
            self.assertIsNone(nuevo_reg["qr_id"])

            # Alta
            alta = inv.agregar_alta({
                "descripcion": "LACTOSA MONOHIDRATADA",
                "proveedor": "HILMAR",
                "cantidad": 50.0,
                "unidad": "kg",
                "lote": "L-200",
                "caducidad": "2027-10-10",
                "ubicacion": "MPA-A-001",
                "categoria": "ALÉRGENO",
                "qr_id": "QR-TEST-02",
            })
            self.assertEqual(alta["ubicacion"], "MPA-A-001")
            self.assertEqual(len(inv.filas), 3)

            # Exportar workbook en memoria desde DB
            wb = inv.construir_workbook()
            self.assertTrue(len(wb.sheetnames) >= 1)
            with tempfile.TemporaryDirectory() as tmp:
                os.environ["ALMACEN_SALIDAS"] = tmp
                try:
                    ruta_xlsx = reportes.exportar_inventario(inv)
                    self.assertTrue(Path(ruta_xlsx).exists())
                finally:
                    os.environ.pop("ALMACEN_SALIDAS", None)

            # Borrar renglón
            borrados = inv.borrar_renglones([nueva_fila])
            self.assertEqual(len(borrados), 1)
            self.assertEqual(len(inv.filas), 2)
        finally:
            con.close()

    def test_catalogo_modo_db(self):
        con = FakePgConnection()
        try:
            con.execute(
                """INSERT INTO catalogo_insumos
                   (sku, descripcion, tipo, inventario, unidad, fabricante)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                ("MP-999", "SORBITOL POLVO", "Excipiente", "mp", "kg", "Roquette"),
            )
            con.commit()
            catalogo._CACHE.update({"ruta": None, "mtime": None, "datos": None})
            datos = catalogo.cargar(con=con)
            self.assertTrue(datos["existe"])
            self.assertEqual(datos["total"], 1)
            self.assertEqual(datos["insumos"][0]["sku"], "MP-999")
        finally:
            catalogo._CACHE.update({"ruta": None, "mtime": None, "datos": None})
            con.close()

    def test_bundle_vercel_y_archivos_supabase(self):
        raiz = Path(__file__).resolve().parent.parent
        self.assertTrue((raiz / "vercel.json").exists())
        self.assertTrue((raiz / "api" / "index.py").exists())
        self.assertTrue((raiz / "supabase" / "schema.sql").exists())
        self.assertTrue((raiz / "supabase" / "seed.sql").exists())
        cfg = json.loads((raiz / "vercel.json").read_text(encoding="utf-8"))
        self.assertIn("rewrites", cfg)
        self.assertIn("functions", cfg)


if __name__ == "__main__":
    unittest.main()
