import pathlib
import shutil
import sqlite3
import sys
import tempfile
import unittest

RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from app import api, bitacora, inventario, usuarios  # noqa: E402

ORIGEN_AREA = pathlib.Path(r"C:\Users\asegm\Desktop\Copia de 06 OCT 26 INVENTARIO MP.xlsx")
# Si la PC no es la del área, se usa el archivo real versionado en el repositorio: sin este
# respaldo, 34 pruebas se saltaban en silencio (y un salto no es una prueba que pasa).
ORIGEN = ORIGEN_AREA if ORIGEN_AREA.exists() else RAIZ / "data" / "Inventario_MP.xlsx"


class EstadoFalso:
    def __init__(self, ruta_xlsx, ruta_db):
        self._inv = inventario.Inventario(ruta_xlsx)
        self.con = bitacora.conectar(ruta_db)
        self.archivo = pathlib.Path(ruta_xlsx).name
        self.carpeta_datos = pathlib.Path(ruta_xlsx).parent

    def inventario(self, tipo=None):
        if tipo not in (None, "mp"):
            raise FileNotFoundError(f"el inventario {tipo} no está cargado")
        return self._inv

    def cargados(self):
        return [("mp", self._inv)]


def sesion_de(con, nombre):
    fila = usuarios.por_nombre(con, nombre)
    return {"usuario_id": fila["id"], "nombre": fila["nombre"], "rol": fila["rol"],
            "debe_cambiar": fila["debe_cambiar"]}


@unittest.skipUnless(ORIGEN.exists(), "no está el Excel del área")
class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        carpeta = pathlib.Path(self.tmp.name)
        self.ruta = carpeta / "Inventario.xlsx"
        shutil.copy2(ORIGEN, self.ruta)
        self.estado = EstadoFalso(self.ruta, carpeta / "bitacora.db")
        self.inv = self.estado.inventario()
        self.con = self.estado.con
        usuarios.crear(self.con, "Aseguramiento", "clave-segura-1", "admin")
        usuarios.crear(self.con, "Esteban", "clave-segura-2", "almacen")
        usuarios.crear(self.con, "Galilea", "clave-segura-3", "almacen")
        usuarios.crear(self.con, "Auditor", "clave-segura-4", "consulta")
        self.admin = sesion_de(self.con, "Aseguramiento")
        self.almacen = sesion_de(self.con, "Esteban")
        self.otro = sesion_de(self.con, "Galilea")
        self.consulta = sesion_de(self.con, "Auditor")
        # un lote grande para probar parcialidades
        self.fila = next(f["fila"] for f in self.inv.filas if float(f["cantidad"] or 0) >= 100)
        self.inv.asignar(self.fila, "MPA-A-005")
        self.inv.guardar()

    def tearDown(self):
        self.con.close()
        self.inv.wb.close()
        self.tmp.cleanup()

    # helpers
    def api(self, metodo, ruta, params=None, cuerpo=None, sesion=None):
        return api.manejar(self.estado, metodo, ruta, params or {}, cuerpo or {}, sesion)

    def mover(self, **cambios):
        cuerpo = {"fila": self.fila, "destino": "MPA-A-020", "motivo": "reacomodo",
                  "usuario": "Esteban", "nota": "prueba"}
        cuerpo.update(cambios)
        return self.api("POST", "/api/mover", cuerpo=cuerpo, sesion=self.almacen)

    def cantidad(self, fila=None):
        return float(self.inv.por_fila(fila or self.fila)["cantidad"])


class TestSesion(Base):
    def test_login_correcto(self):
        codigo, cuerpo, extra = self.api("POST", "/api/login",
                                         cuerpo={"usuario": "Esteban", "password": "clave-segura-2"})
        self.assertEqual(codigo, 200, cuerpo)
        galleta = extra["cookies"][0]
        token = galleta.split(";")[0].split("=", 1)[1]
        sesion = bitacora.sesion_valida(self.con, token)
        self.assertIsNotNone(sesion)
        self.assertEqual(sesion["rol"], "almacen")

    def test_login_incorrecto(self):
        codigo, cuerpo, _ = self.api("POST", "/api/login",
                                     cuerpo={"usuario": "Esteban", "password": "mala"})
        self.assertEqual(codigo, 401)
        self.assertIn("incorrectos", cuerpo["error"])

    def test_sin_sesion_no_lee_inventario(self):
        codigo, _, _ = self.api("GET", "/api/inventario")
        self.assertEqual(codigo, 401)

    def test_consulta_no_escribe(self):
        codigo, cuerpo, _ = self.api("POST", "/api/mover",
                                     cuerpo={"fila": self.fila, "destino": "MPA-A-020",
                                             "motivo": "reacomodo"}, sesion=self.consulta)
        self.assertEqual(codigo, 403)
        self.assertIn("consulta", cuerpo["error"])

    def test_cambio_de_password(self):
        codigo, _, _ = self.api("POST", "/api/password",
                                cuerpo={"actual": "clave-segura-2", "nueva": "otra-clave-99"},
                                sesion=self.almacen)
        self.assertEqual(codigo, 200)
        self.assertIsNotNone(usuarios.verificar(self.con, "Esteban", "otra-clave-99"))
        self.assertIsNone(usuarios.verificar(self.con, "Esteban", "clave-segura-2"))


class TestMovimientos(Base):
    def test_mover_completo(self):
        codigo, cuerpo, _ = self.mover()
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["accion"], "MOVER")
        self.assertEqual(cuerpo["anterior"], "MPA-A-005")
        self.assertEqual(cuerpo["nueva"], "MPA-A-020")
        self.assertAlmostEqual(cuerpo["cantidad"], self.cantidad())
        self.assertEqual(self.inv.por_fila(self.fila)["ubicacion"], "MPA-A-020")

    def test_mover_parcial_crea_parcialidad(self):
        antes = self.cantidad()
        codigo, cuerpo, _ = self.mover(cantidad=5)
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["accion"], "PARCIAL")
        self.assertAlmostEqual(self.cantidad(), antes - 5)
        nueva = cuerpo["fila_destino"]
        self.assertIsNotNone(nueva)
        parcial = self.inv.por_fila(nueva)
        self.assertAlmostEqual(parcial["cantidad"], 5)
        self.assertEqual(parcial["ubicacion"], "MPA-A-020")
        self.assertEqual(parcial["lote"], self.inv.por_fila(self.fila)["lote"])
        mov = bitacora.ultimo_vigente(self.con, self.fila)
        self.assertEqual(mov["accion"], "PARCIAL")
        self.assertEqual(mov["fila_destino"], nueva)
        self.assertAlmostEqual(mov["cantidad"], 5)

    def test_parcial_no_puede_ser_mayor(self):
        codigo, cuerpo, _ = self.mover(cantidad=99999)
        self.assertEqual(codigo, 400)
        self.assertIn("no puedes mover", cuerpo["error"])

    def test_parcial_con_todo_es_mover_completo(self):
        codigo, cuerpo, _ = self.mover(cantidad=self.cantidad())
        self.assertEqual(codigo, 200)
        self.assertEqual(cuerpo["accion"], "MOVER")

    def test_parcial_guarda_en_excel(self):
        codigo, cuerpo, _ = self.mover(cantidad=7.5)
        self.assertEqual(codigo, 200, cuerpo)
        otra = inventario.Inventario(self.ruta)
        try:
            self.assertAlmostEqual(float(otra.por_fila(cuerpo["fila_destino"])["cantidad"]), 7.5)
            self.assertEqual(otra.por_fila(cuerpo["fila_destino"])["ubicacion"], "MPA-A-020")
        finally:
            otra.wb.close()


class TestDeshacer(Base):
    def test_deshacer_mover_devuelve_ubicacion(self):
        _, cuerpo, _ = self.mover()
        codigo, res, _ = self.api("POST", "/api/deshacer",
                                  cuerpo={"movimiento": cuerpo["movimiento"]}, sesion=self.almacen)
        self.assertEqual(codigo, 200, res)
        self.assertEqual(self.inv.por_fila(self.fila)["ubicacion"], "MPA-A-005")
        self.assertEqual(bitacora.por_id(self.con, cuerpo["movimiento"])["revertido"], 1)

    def test_deshacer_parcial_devuelve_kilos(self):
        antes = self.cantidad()
        _, cuerpo, _ = self.mover(cantidad=12)
        self.assertAlmostEqual(self.cantidad(), antes - 12)
        codigo, res, _ = self.api("POST", "/api/deshacer",
                                  cuerpo={"movimiento": cuerpo["movimiento"]}, sesion=self.almacen)
        self.assertEqual(codigo, 200, res)
        self.assertAlmostEqual(self.cantidad(), antes)
        parcial = self.inv.por_fila(cuerpo["fila_destino"])
        self.assertAlmostEqual(float(parcial["cantidad"] or 0), 0)
        self.assertIsNone(parcial["ubicacion"])

    def test_no_puedo_deshacer_lo_de_otro(self):
        _, cuerpo, _ = self.mover()
        codigo, res, _ = self.api("POST", "/api/deshacer",
                                  cuerpo={"movimiento": cuerpo["movimiento"]}, sesion=self.otro)
        self.assertEqual(codigo, 403)

    def test_admin_deshace_lo_de_otro(self):
        _, cuerpo, _ = self.mover()
        codigo, res, _ = self.api("POST", "/api/deshacer",
                                  cuerpo={"movimiento": cuerpo["movimiento"]}, sesion=self.admin)
        self.assertEqual(codigo, 200, res)

    def test_no_se_puede_deshacer_si_hay_uno_posterior(self):
        _, primero, _ = self.mover(cantidad=5)
        _, segundo, _ = self.mover(cantidad=3)
        codigo, res, _ = self.api("POST", "/api/deshacer",
                                  cuerpo={"movimiento": primero["movimiento"]}, sesion=self.almacen)
        self.assertEqual(codigo, 409)
        self.assertIn("posteriores", res["error"])

    def test_deshacer_dos_veces(self):
        _, cuerpo, _ = self.mover()
        self.api("POST", "/api/deshacer", cuerpo={"movimiento": cuerpo["movimiento"]},
                 sesion=self.almacen)
        codigo, res, _ = self.api("POST", "/api/deshacer",
                                  cuerpo={"movimiento": cuerpo["movimiento"]}, sesion=self.almacen)
        self.assertEqual(codigo, 400)
        self.assertIn("ya fue deshecho", res["error"])


class TestPurgar(Base):
    def test_admin_purga_y_queda_auditoria(self):
        _, cuerpo, _ = self.mover()
        codigo, res, _ = self.api("POST", "/api/purgar",
                                  cuerpo={"movimiento": cuerpo["movimiento"], "motivo": "prueba"},
                                  sesion=self.admin)
        self.assertEqual(codigo, 200, res)
        self.assertIsNone(bitacora.por_id(self.con, cuerpo["movimiento"]))
        self.assertEqual(self.inv.por_fila(self.fila)["ubicacion"], "MPA-A-005")
        auditoria = self.con.execute("SELECT * FROM auditoria WHERE accion='PURGA_MOVIMIENTO'").fetchall()
        self.assertEqual(len(auditoria), 1)

    def test_almacen_no_purga(self):
        _, cuerpo, _ = self.mover()
        codigo, _, _ = self.api("POST", "/api/purgar",
                                cuerpo={"movimiento": cuerpo["movimiento"], "motivo": "prueba"},
                                sesion=self.almacen)
        self.assertEqual(codigo, 403)

    def test_purga_exige_motivo(self):
        _, cuerpo, _ = self.mover()
        codigo, res, _ = self.api("POST", "/api/purgar",
                                  cuerpo={"movimiento": cuerpo["movimiento"], "motivo": "x"},
                                  sesion=self.admin)
        self.assertEqual(codigo, 400)
        self.assertIn("motivo", res["error"])


class TestUsuarios(Base):
    def test_admin_crea_usuario(self):
        codigo, res, _ = self.api("POST", "/api/usuarios",
                                  cuerpo={"nombre": "Maria Elena", "password": "clave-larga-1",
                                          "rol": "almacen"}, sesion=self.admin)
        self.assertEqual(codigo, 200, res)
        self.assertIsNotNone(usuarios.verificar(self.con, "Maria Elena", "clave-larga-1"))

    def test_almacen_no_gestiona_usuarios(self):
        codigo, _, _ = self.api("GET", "/api/usuarios", sesion=self.almacen)
        self.assertEqual(codigo, 403)

    def test_password_corta(self):
        codigo, res, _ = self.api("POST", "/api/usuarios",
                                  cuerpo={"nombre": "Corto", "password": "123", "rol": "almacen"},
                                  sesion=self.admin)
        self.assertEqual(codigo, 400)
        self.assertIn("8 caracteres", res["error"])

    def test_no_quitar_ultimo_admin(self):
        uid = self.admin["usuario_id"]
        codigo, res, _ = self.api("PATCH", f"/api/usuarios/{uid}", cuerpo={"rol": "almacen"},
                                  sesion=self.admin)
        self.assertEqual(codigo, 400)
        self.assertIn("último administrador", res["error"])

    def test_desactivar_usuario_cierra_sesion(self):
        uid = self.almacen["usuario_id"]
        token, _ = bitacora.abrir_sesion(self.con, uid)
        codigo, _, _ = self.api("PATCH", f"/api/usuarios/{uid}", cuerpo={"activo": False},
                                sesion=self.admin)
        self.assertEqual(codigo, 200)
        self.assertIsNone(bitacora.sesion_valida(self.con, token))
        self.assertIsNone(usuarios.verificar(self.con, "Esteban", "clave-segura-2"))


class TestMigracion(unittest.TestCase):
    def test_bitacora_vieja_se_migra_sin_perder_movimientos(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = pathlib.Path(tmp) / "vieja.db"
            con = sqlite3.connect(ruta)
            con.executescript("""
                CREATE TABLE movimiento (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL,
                    accion TEXT NOT NULL CHECK (accion IN
                      ('UBICAR','MOVER','CONFIRMAR','NO_ENCONTRADO','LIMPIAR','IMPORTAR')),
                    fila INTEGER, descripcion TEXT NOT NULL, lote TEXT,
                    ubicacion_origen TEXT, ubicacion_destino TEXT,
                    usuario TEXT NOT NULL, nota TEXT);
                INSERT INTO movimiento (ts, accion, descripcion, usuario)
                VALUES ('2026-10-01 09:00:00', 'UBICAR', 'X', 'Esteban');
            """)
            con.commit()
            con.close()

            con = bitacora.conectar(ruta)
            columnas = {f["name"] for f in con.execute("PRAGMA table_info(movimiento)")}
            self.assertTrue({"motivo", "ref_doc", "revertido", "cantidad", "fila_destino"} <= columnas)
            self.assertEqual(con.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0], 1)
            # la tabla reconstruida acepta las acciones nuevas
            con.execute("INSERT INTO movimiento (ts, accion, descripcion, usuario) "
                        "VALUES ('2026-10-02 09:00:00', 'PARCIAL', 'X', 'Esteban')")
            con.commit()
            con.close()


if __name__ == "__main__":
    unittest.main()
