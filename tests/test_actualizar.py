"""Actualizador desde la USB: versiones, respaldo, reemplazo de código y reversión."""
import json
import pathlib
import sqlite3
import tempfile
import unittest
import zipfile

from scripts import actualizar

DIGEST = "0" * 64


def crear_instalacion(base, version="1.4.0"):
    """Una instalación de mentira: código, datos, manifiesto y un archivo viejo (demo.cmd)."""
    base = pathlib.Path(base)
    (base / "app").mkdir(parents=True)
    (base / "data").mkdir(parents=True)
    (base / "app" / "server.py").write_text("# servidor viejo\n", encoding="utf-8")
    (base / "app" / "api.py").write_text("# api vieja\n", encoding="utf-8")
    (base / "app" / "inventario.py").write_text("# inventario viejo\n", encoding="utf-8")
    (base / "app" / "version.py").write_text(f'VERSION = "{version}"\n', encoding="utf-8")
    (base / "demo.cmd").write_text("@echo off\n", encoding="utf-8")
    (base / "NOTAS.txt").write_text("notas del usuario\n", encoding="utf-8")
    (base / "data" / "Inventario_MP.xlsx").write_bytes(b"excel-viejo")
    con = sqlite3.connect(base / "data" / "bitacora.db")
    con.execute("CREATE TABLE movimiento (id INTEGER PRIMARY KEY, nota TEXT)")
    con.execute("INSERT INTO movimiento (nota) VALUES ('movimiento previo')")
    con.commit()
    con.close()
    lineas = [f"{DIGEST} {10:>10}  app/server.py",
              f"{DIGEST} {10:>10}  app/api.py",
              f"{DIGEST} {10:>10}  app/inventario.py",
              f"{DIGEST} {10:>10}  app/version.py",
              f"{DIGEST} {10:>10}  demo.cmd"]
    (base / "MANIFIESTO.txt").write_text("\n".join(lineas) + "\n", encoding="utf-8")
    return base


def crear_paquete(ruta, version="1.5.0", extra=None):
    extra = extra or {}
    ruta = pathlib.Path(ruta)
    with zipfile.ZipFile(ruta, "w") as z:
        z.writestr("VERSION.txt", f"{version}\n")
        z.writestr("app/server.py", "# servidor nuevo\n")
        z.writestr("app/api.py", "# api nueva\n")
        z.writestr("app/inventario.py", "# inventario nuevo\n")
        z.writestr("app/version.py", f'VERSION = "{version}"\n')
        z.writestr("app/nuevo.py", "# módulo nuevo\n")
        z.writestr("MANIFIESTO.txt", f"{DIGEST} {10:>10}  app/server.py\n")
        for nombre, contenido in extra.items():
            z.writestr(nombre, contenido)
    return ruta


class Base(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.raiz = pathlib.Path(self._tmp.name)
        self.instalada = crear_instalacion(self.raiz / "instalada")
        self.paquete = crear_paquete(self.raiz / "v1.5.0.zip")

    def tearDown(self):
        self._tmp.cleanup()

    def correr(self, **kwargs):
        datos = {"instalada": self.instalada, "sin_servidor": True, "log": lambda *_: None}
        datos.update(kwargs)
        return actualizar.actualizar(self.paquete, **datos)


class TestVersiones(unittest.TestCase):
    def test_compara_numeros_no_texto(self):
        self.assertGreater(actualizar.version_de_texto("1.10.0"),
                           actualizar.version_de_texto("1.9.0"))
        self.assertEqual(actualizar.version_de_texto("1.5"), (1, 5, 0))
        self.assertEqual(actualizar.version_de_texto("1.5.0-rc1"), (1, 5, 0))

    def test_lee_la_version_instalada(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = crear_instalacion(pathlib.Path(tmp) / "x", version="2.3.4")
            self.assertEqual(actualizar.leer_version(base), "2.3.4")

    def test_lee_la_version_del_paquete(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = crear_paquete(pathlib.Path(tmp) / "p.zip", version="3.1.0")
            self.assertEqual(actualizar.version_del_paquete(ruta), "3.1.0")


class TestActualizar(Base):
    def test_reemplaza_el_codigo(self):
        resultado = self.correr()
        self.assertTrue(resultado["ok"], resultado)
        self.assertEqual(resultado["version_nueva"], "1.5.0")
        self.assertIn("servidor nuevo",
                      (self.instalada / "app" / "server.py").read_text(encoding="utf-8"))
        self.assertTrue((self.instalada / "app" / "nuevo.py").exists())
        self.assertEqual(actualizar.leer_version(self.instalada), "1.5.0")

    def test_no_toca_los_datos(self):
        self.correr()
        con = sqlite3.connect(self.instalada / "data" / "bitacora.db")
        filas = con.execute("SELECT nota FROM movimiento").fetchall()
        con.close()
        self.assertEqual(filas, [("movimiento previo",)])
        self.assertEqual((self.instalada / "data" / "Inventario_MP.xlsx").read_bytes(),
                         b"excel-viejo")

    def test_retira_solo_lo_que_instalo_un_paquete(self):
        resultado = self.correr()
        self.assertFalse((self.instalada / "demo.cmd").exists())
        self.assertIn("demo.cmd", resultado["detalle"]["obsoletos"])
        # un archivo del usuario no se toca
        self.assertTrue((self.instalada / "NOTAS.txt").exists())

    def test_respalda_codigo_y_datos(self):
        resultado = self.correr()
        respaldo = pathlib.Path(resultado["respaldo"])
        self.assertTrue(respaldo.exists())
        with zipfile.ZipFile(respaldo) as z:
            nombres = z.namelist()
            self.assertIn("codigo/app/server.py", nombres)
            self.assertIn("datos/bitacora.db", nombres)
            self.assertIn("datos/Inventario_MP.xlsx", nombres)
            self.assertIn(b"servidor viejo", z.read("codigo/app/server.py"))
            info = json.loads(z.read("info.json"))
            self.assertEqual(info["version"], "1.4.0")

    def test_no_baja_de_version(self):
        paquete_viejo = crear_paquete(self.raiz / "v1.0.0.zip", version="1.0.0")
        resultado = actualizar.actualizar(paquete_viejo, self.instalada, sin_servidor=True,
                                          log=lambda *_: None)
        self.assertFalse(resultado["ok"])
        self.assertIn("más viejo", resultado["motivo"])
        self.assertIn("servidor viejo",
                      (self.instalada / "app" / "server.py").read_text(encoding="utf-8"))

    def test_forzar_permite_volver_atras(self):
        paquete_viejo = crear_paquete(self.raiz / "v1.0.0.zip", version="1.0.0")
        resultado = actualizar.actualizar(paquete_viejo, self.instalada, sin_servidor=True,
                                          forzar=True, log=lambda *_: None)
        self.assertTrue(resultado["ok"], resultado)
        self.assertEqual(actualizar.leer_version(self.instalada), "1.0.0")

    def test_simular_no_escribe(self):
        resultado = self.correr(simular=True)
        self.assertTrue(resultado["simulado"])
        self.assertIn("servidor viejo",
                      (self.instalada / "app" / "server.py").read_text(encoding="utf-8"))
        self.assertFalse((self.instalada / "app" / "nuevo.py").exists())
        self.assertFalse((self.instalada / "data" / "respaldos").exists())

    def test_revierte_si_la_version_nueva_no_arranca(self):
        llamadas = {"n": 0}
        original = actualizar.verificar

        def verificar_falso(carpeta, python=None):
            llamadas["n"] += 1
            if llamadas["n"] == 1:
                return False, "ImportError: módulo roto"
            return original(carpeta, python)

        actualizar.verificar = verificar_falso
        try:
            resultado = self.correr()
        finally:
            actualizar.verificar = original
        self.assertFalse(resultado["ok"])
        self.assertIn("no arranca", resultado["motivo"])
        self.assertIn("servidor viejo",
                      (self.instalada / "app" / "server.py").read_text(encoding="utf-8"))

    def test_paquete_sin_version_se_rechaza(self):
        sin_version = self.raiz / "sin.zip"
        with zipfile.ZipFile(sin_version, "w") as z:
            z.writestr("app/server.py", "# x\n")
        resultado = actualizar.actualizar(sin_version, self.instalada, sin_servidor=True,
                                          log=lambda *_: None)
        self.assertFalse(resultado["ok"])
        self.assertIn("no declara versión", resultado["motivo"])

    def test_instalacion_inexistente(self):
        resultado = actualizar.actualizar(self.paquete, self.raiz / "no-existe", sin_servidor=True,
                                          log=lambda *_: None)
        self.assertFalse(resultado["ok"])
        self.assertEqual(resultado["codigo"], 2)


class TestBuscar(unittest.TestCase):
    def test_encuentra_la_carpeta_por_pista(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = crear_instalacion(pathlib.Path(tmp) / "app-instalada")
            self.assertEqual(actualizar.buscar_instalacion([base]), base)

    def test_devuelve_none_si_no_hay(self):
        original = actualizar.CANDIDATOS
        actualizar.CANDIDATOS = []
        try:
            with tempfile.TemporaryDirectory() as tmp:
                self.assertIsNone(
                    actualizar.buscar_instalacion([pathlib.Path(tmp)], perfiles=False))
        finally:
            actualizar.CANDIDATOS = original


class TestRegistro(unittest.TestCase):
    def test_puerto_libre_no_devuelve_pids(self):
        # 65500 casi nunca está ocupado; lo que importa es que no truene con la salida de netstat
        self.assertEqual(actualizar.pid_en_puerto(65500), [])

    def test_deja_constancia_en_la_usb(self):
        with tempfile.TemporaryDirectory() as tmp:
            archivo = actualizar.registrar_resultado("EXITO: 1.5.0 -> 1.6.0", carpeta=tmp)
            actualizar.registrar_resultado("FALLO: disco lleno", carpeta=tmp)
            texto = archivo.read_text(encoding="utf-8")
            self.assertIn("EXITO: 1.5.0 -> 1.6.0", texto)
            self.assertIn("FALLO: disco lleno", texto)
            self.assertEqual(len(texto.strip().splitlines()), 2)

    def test_no_falla_si_no_se_puede_escribir(self):
        with tempfile.TemporaryDirectory() as tmp:
            sin_permiso = pathlib.Path(tmp) / "no-existe" / "sub"
            archivo = actualizar.registrar_resultado("EXITO", carpeta=sin_permiso)
            self.assertFalse(archivo.exists())      # no hay carpeta: no truena, solo no escribe


if __name__ == "__main__":
    unittest.main()
