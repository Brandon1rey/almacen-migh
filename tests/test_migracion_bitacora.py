"""La base de una versión anterior debe admitir las acciones nuevas (SALIDA) sin perder movimientos."""
import pathlib
import sqlite3
import tempfile
import unittest

from app import bitacora

ESQUEMA_VIEJO = """
CREATE TABLE movimiento (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    ts                TEXT NOT NULL,
    accion            TEXT NOT NULL CHECK (accion IN
                      ('UBICAR','MOVER','CONFIRMAR','NO_ENCONTRADO','LIMPIAR')),
    fila              INTEGER,
    descripcion       TEXT NOT NULL,
    lote              TEXT,
    ubicacion_origen  TEXT,
    ubicacion_destino TEXT,
    usuario           TEXT NOT NULL,
    nota              TEXT,
    revertido         INTEGER NOT NULL DEFAULT 0
);
INSERT INTO movimiento (ts, accion, fila, descripcion, lote, ubicacion_origen,
                        ubicacion_destino, usuario, nota)
VALUES ('2026-10-01 08:00:00', 'UBICAR', 4, 'WPC 80%', 'L-1', NULL, 'MPA-A-005',
        'Maria Elena', 'primer acomodo');
"""


class TestMigracionAcciones(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.ruta = pathlib.Path(self._tmp.name) / "vieja.db"
        con = sqlite3.connect(self.ruta)
        con.executescript(ESQUEMA_VIEJO)
        con.commit()
        con.close()

    def tearDown(self):
        self._tmp.cleanup()

    def test_conserva_los_movimientos_previos(self):
        con = bitacora.conectar(self.ruta)
        self.assertEqual(bitacora.contar(con), 1)
        previo = bitacora.ultimos(con, 1)[0]
        self.assertEqual(previo["accion"], "UBICAR")
        self.assertEqual(previo["ubicacion_destino"], "MPA-A-005")
        con.close()

    def test_admite_salida_a_produccion(self):
        con = bitacora.conectar(self.ruta)
        ts, mov_id = bitacora.registrar(
            con, "SALIDA", {"fila": 4, "descripcion": "WPC 80%", "lote": "L-1"},
            "Esteban", destino="PRODUCCION", origen="MPA-A-005", motivo="surtido",
            cantidad=25, ref_doc="OP-1234", usuario_id=1, inventario="mp")
        guardado = bitacora.por_id(con, mov_id)
        self.assertEqual(guardado["accion"], "SALIDA")
        self.assertEqual(guardado["ubicacion_destino"], "PRODUCCION")
        self.assertAlmostEqual(guardado["cantidad"], 25)
        con.close()

    def test_sigue_admitiendo_las_acciones_viejas(self):
        con = bitacora.conectar(self.ruta)
        bitacora.registrar(con, "MOVER", {"fila": 4, "descripcion": "WPC 80%", "lote": "L-1"},
                           "Esteban", destino="MPA-A-006", origen="MPA-A-005", usuario_id=1)
        self.assertEqual(bitacora.contar(con), 2)
        con.close()


if __name__ == "__main__":
    unittest.main()
