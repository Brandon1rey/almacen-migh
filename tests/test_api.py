import pathlib
import shutil
import sqlite3
import sys
import tempfile
import unittest

RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from app import api, bitacora, inventario  # noqa: E402

ORIGEN = pathlib.Path(r"C:\Users\asegm\Desktop\Copia de 06 OCT 26 INVENTARIO MP.xlsx")


class EstadoFalso:
    """Lo mínimo que la API necesita del servidor."""

    def __init__(self, ruta_xlsx, ruta_db):
        self._inv = inventario.Inventario(ruta_xlsx)
        self.con = bitacora.conectar(ruta_db)
        self.archivo = pathlib.Path(ruta_xlsx).name
        self.demo = True

    def inventario(self):
        return self._inv


@unittest.skipUnless(ORIGEN.exists(), "no está el Excel del área")
class TestMover(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        carpeta = pathlib.Path(self.tmp.name)
        self.ruta = carpeta / "Inventario.xlsx"
        shutil.copy2(ORIGEN, self.ruta)
        self.estado = EstadoFalso(self.ruta, carpeta / "bitacora.db")
        self.inv = self.estado.inventario()
        self.fila = self.inv.filas[0]["fila"]
        self.inv.asignar(self.fila, "MPA-A-005")
        self.inv.guardar()

    def tearDown(self):
        self.estado.con.close()
        self.inv.wb.close()
        self.tmp.cleanup()

    def mover(self, **cambios):
        cuerpo = {"fila": self.fila, "destino": "MPA-A-020", "motivo": "reacomodo",
                  "usuario": "Esteban Soto", "nota": "prueba"}
        cuerpo.update(cambios)
        return api.manejar(self.estado, "POST", "/api/mover", {}, cuerpo)

    def test_mover_registra_y_persiste(self):
        codigo, cuerpo = self.mover()
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["accion"], "MOVER")
        self.assertEqual(cuerpo["anterior"], "MPA-A-005")
        self.assertEqual(cuerpo["nueva"], "MPA-A-020")
        self.assertEqual(cuerpo["motivo"], "reacomodo")
        # el Excel en disco quedó con el destino
        otra = inventario.Inventario(self.ruta)
        self.assertEqual(otra.por_fila(self.fila)["ubicacion"], "MPA-A-020")
        otra.wb.close()
        # la bitácora tiene el movimiento con motivo
        movs = bitacora.por_fila(self.estado.con, self.fila)
        self.assertEqual(movs[0]["accion"], "MOVER")
        self.assertEqual(movs[0]["motivo"], "reacomodo")
        self.assertEqual(movs[0]["ubicacion_origen"], "MPA-A-005")

    def test_destino_igual_es_rechazado(self):
        codigo, cuerpo = self.mover(destino="MPA-A-005")
        self.assertEqual(codigo, 400)
        self.assertIn("ubicación actual", cuerpo["error"])

    def test_motivo_invalido(self):
        codigo, cuerpo = self.mover(motivo="porque si")
        self.assertEqual(codigo, 400)
        self.assertIn("motivo", cuerpo["error"])

    def test_sin_usuario(self):
        codigo, cuerpo = self.mover(usuario="")
        self.assertEqual(codigo, 400)
        self.assertIn("nombre", cuerpo["error"])

    def test_destino_inexistente(self):
        codigo, cuerpo = self.mover(destino="MPA-A-999")
        self.assertEqual(codigo, 400)
        self.assertIn("inexistente", cuerpo["error"])

    def test_lote_sin_ubicacion_pide_asignar(self):
        self.inv.asignar(self.fila, None)
        self.inv.guardar()
        codigo, cuerpo = self.mover()
        self.assertEqual(codigo, 400)
        self.assertIn("no tiene ubicación", cuerpo["error"])

    def test_historial_por_lote(self):
        self.mover()
        codigo, cuerpo = api.manejar(self.estado, "GET", "/api/historial", {"fila": self.fila}, None)
        self.assertEqual(codigo, 200)
        self.assertEqual(len(cuerpo["movimientos"]), 1)
        self.assertEqual(cuerpo["movimientos"][0]["motivo"], "reacomodo")

    def test_alta_con_ubicar_sigue_funcionando(self):
        otra_fila = self.inv.filas[1]["fila"]
        codigo, cuerpo = api.manejar(self.estado, "POST", "/api/ubicar", {},
                                     {"fila": otra_fila, "ubicacion": "MPNA-A-041",
                                      "usuario": "Galilea Medina"})
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["accion"], "UBICAR")


class TestMigracion(unittest.TestCase):
    def test_agrega_motivo_y_ref_a_bitacora_vieja(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = pathlib.Path(tmp) / "vieja.db"
            con = sqlite3.connect(ruta)
            con.executescript("""
                CREATE TABLE movimiento (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL,
                    accion TEXT NOT NULL, fila INTEGER, descripcion TEXT NOT NULL, lote TEXT,
                    ubicacion_origen TEXT, ubicacion_destino TEXT, usuario TEXT NOT NULL,
                    nota TEXT);
                INSERT INTO movimiento (ts, accion, descripcion, usuario)
                VALUES ('2026-10-01 09:00:00', 'UBICAR', 'X', 'Esteban Soto');
            """)
            con.commit()
            con.close()

            con = bitacora.conectar(ruta)
            columnas = {fila["name"] for fila in con.execute("PRAGMA table_info(movimiento)")}
            self.assertTrue({"motivo", "ref_doc"} <= columnas)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0], 1)
            con.close()


if __name__ == "__main__":
    unittest.main()
