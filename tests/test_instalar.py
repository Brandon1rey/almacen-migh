"""Instalación: arranque automático en la carpeta de Inicio (sin permisos de administrador)."""
import pathlib
import tempfile
import unittest

from scripts import instalar


class TestArranque(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.carpeta = pathlib.Path(self._tmp.name) / "Startup"

    def tearDown(self):
        self._tmp.cleanup()

    def test_crea_el_archivo_en_la_carpeta_de_inicio(self):
        ruta = instalar.crear_arranque_automatico(carpeta=self.carpeta)
        self.assertTrue(ruta.exists())
        self.assertEqual(ruta.name, "Almacen MIGH.cmd")

    def test_usa_el_interprete_del_entorno(self):
        # pythonw: el arranque no deja ventana; el servidor escribe su propio log
        texto = instalar.contenido_arranque("C:/Python312/python.exe")
        self.assertIn(str(pathlib.Path("C:/Python312/pythonw.exe")), texto)
        self.assertIn("app\\server.py", texto)
        self.assertNotIn(">>", texto)

    def test_respeta_un_pythonw_explicito(self):
        texto = instalar.contenido_arranque("C:/x/pythonw.exe")
        self.assertIn(str(pathlib.Path("C:/x/pythonw.exe")), texto)
        self.assertNotIn("pythonw.exe.exe", texto)

    def test_arranca_desde_la_carpeta_del_proyecto(self):
        texto = instalar.contenido_arranque(r"C:\x\python.exe")
        self.assertIn(f'cd /d "{instalar.RAIZ}"', texto)
        self.assertIn("app\\server.py", texto)

    def test_lleva_el_puerto(self):
        texto = instalar.contenido_arranque(r"C:\x\python.exe", puerto=8399)
        self.assertIn('set "ALMACEN_PUERTO=8399"', texto)
        self.assertIn('set "ALMACEN_HOST=0.0.0.0"', texto)

    def test_usa_fin_de_linea_de_windows(self):
        # cmd.exe necesita CRLF; con solo CR el archivo se lee como una sola línea
        texto = instalar.contenido_arranque(r"C:\x\python.exe")
        self.assertIn(chr(13) + chr(10), texto)
        self.assertTrue(texto.startswith("@echo off\r\n"))

    def test_dos_instalaciones_no_duplican(self):
        instalar.crear_arranque_automatico(carpeta=self.carpeta)
        instalar.crear_arranque_automatico(carpeta=self.carpeta)
        self.assertEqual(len(list(self.carpeta.glob("*.cmd"))), 1)

    def test_quitar_arranque(self):
        instalar.crear_arranque_automatico(carpeta=self.carpeta)
        self.assertTrue(instalar.quitar_arranque(carpeta=self.carpeta))
        self.assertFalse((self.carpeta / "Almacen MIGH.cmd").exists())
        self.assertFalse(instalar.quitar_arranque(carpeta=self.carpeta))


if __name__ == "__main__":
    unittest.main()
