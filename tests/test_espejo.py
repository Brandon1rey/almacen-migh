"""Espejo del proyecto en D:\\: copia el código, nunca los datos ni las notas personales."""
import pathlib
import tempfile
import unittest

from scripts import espejo


class Base(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        raiz = pathlib.Path(self._tmp.name)
        self.origen = raiz / "proyecto"
        self.destino = raiz / "espejo"
        for carpeta, archivos in {
            "app": {"server.py": "print('hola')", "static/index.html": "<html>"},
            "scripts": {"instalar.py": "# instalador"},
            "data": {"bitacora.db": "datos del servidor"},
            "datos_iniciales": {"LEEME.txt": "aqui va el excel"},
            "__pycache__": {"basura.pyc": "x"},
            ".venv/Scripts": {"python.exe": "no copiar"},
        }.items():
            for nombre, contenido in archivos.items():
                ruta = self.origen / carpeta / nombre
                ruta.parent.mkdir(parents=True, exist_ok=True)
                ruta.write_text(contenido, encoding="utf-8")
        (self.origen / "arrancar.cmd").write_text("@echo off", encoding="utf-8")
        (self.origen / "NOTASBRANDON.txt").write_text("mis notas", encoding="utf-8")
        (self.origen / ".py_path").write_text("C:/python.exe", encoding="utf-8")
        self.destino.mkdir()

    def tearDown(self):
        self._tmp.cleanup()


class TestComparar(Base):
    def test_espejo_vacio_todo_pendiente(self):
        estado = espejo.comparar(self.origen, self.destino)
        self.assertTrue(estado["pendientes"])
        self.assertEqual(estado["distintos"], [])
        self.assertIn("app/server.py", estado["faltan"])

    def test_ignora_datos_venv_y_notas(self):
        estado = espejo.comparar(self.origen, self.destino)
        todo = set(estado["faltan"]).union(estado["distintos"])
        self.assertNotIn("data/bitacora.db", todo)
        self.assertNotIn(".venv/Scripts/python.exe", todo)
        self.assertNotIn("NOTASBRANDON.txt", todo)
        self.assertNotIn(".py_path", todo)
        self.assertNotIn("__pycache__/basura.pyc", todo)
        self.assertIn("datos_iniciales/LEEME.txt", todo)


class TestEspejar(Base):
    def test_copia_el_codigo(self):
        espejo.espejar(self.origen, self.destino)
        self.assertTrue((self.destino / "app" / "server.py").exists())
        self.assertTrue((self.destino / "arrancar.cmd").exists())
        self.assertTrue((self.destino / "datos_iniciales" / "LEEME.txt").exists())

    def test_no_copia_datos_ni_notas(self):
        espejo.espejar(self.origen, self.destino)
        self.assertFalse((self.destino / "data" / "bitacora.db").exists())
        self.assertFalse((self.destino / "NOTASBRANDON.txt").exists())
        self.assertFalse((self.destino / ".venv").exists())

    def test_no_borra_los_datos_del_destino(self):
        (self.destino / "data").mkdir()
        (self.destino / "data" / "bitacora.db").write_text("datos del espejo", encoding="utf-8")
        espejo.espejar(self.origen, self.destino)
        self.assertEqual((self.destino / "data" / "bitacora.db").read_text(encoding="utf-8"),
                         "datos del espejo")

    def test_conserva_las_notas_del_destino(self):
        (self.destino / "NOTASBRANDON.txt").write_text("notas del espejo", encoding="utf-8")
        espejo.espejar(self.origen, self.destino)
        self.assertEqual((self.destino / "NOTASBRANDON.txt").read_text(encoding="utf-8"),
                         "notas del espejo")

    def test_actualiza_un_archivo_cambiado(self):
        espejo.espejar(self.origen, self.destino)
        (self.origen / "app" / "server.py").write_text("print('v2')", encoding="utf-8")
        resultado = espejo.espejar(self.origen, self.destino)
        self.assertEqual(resultado["copiados"], ["app/server.py"])
        self.assertEqual((self.destino / "app" / "server.py").read_text(encoding="utf-8"),
                         "print('v2')")

    def test_verifica_al_final(self):
        resultado = espejo.espejar(self.origen, self.destino)
        self.assertTrue(resultado["al_dia"])
        self.assertEqual(resultado["distintos"], [])

    def test_simular_no_escribe(self):
        resultado = espejo.espejar(self.origen, self.destino, simular=True)
        self.assertTrue(resultado["pendientes"])
        self.assertFalse((self.destino / "app").exists())


if __name__ == "__main__":
    unittest.main()
