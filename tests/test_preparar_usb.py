"""La carpeta de la USB debe quedar lista para doble clic y conservar el paquete anterior."""
import pathlib
import tempfile
import unittest
import zipfile

from scripts import preparar_usb


def paquete(ruta, version):
    ruta = pathlib.Path(ruta)
    with zipfile.ZipFile(ruta, "w") as z:
        z.writestr("VERSION.txt", f"{version}\n")
        z.writestr("app/server.py", "# ok\n")
    return ruta


class TestPrepararUSB(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.base = pathlib.Path(self._tmp.name)
        self.destino = self.base / "usb" / "ALMACEN_MIGH_ACTUALIZACION"

    def tearDown(self):
        self._tmp.cleanup()

    def test_deja_todo_lo_necesario(self):
        nuevo = paquete(self.base / "almacen-ubicaciones_v1.5.0_20261023-1200.zip", "1.5.0")
        resultado = preparar_usb.preparar(self.destino, nuevo)
        self.assertEqual(resultado["version"], "1.5.0")
        for nombre in ("ACTUALIZAR.cmd", "INSTALAR_EN_PC_NUEVA.cmd", "actualizar.py",
                       "version.txt", "LEEME_ACTUALIZACION.md"):
            self.assertTrue((self.destino / nombre).exists(), nombre)
        self.assertTrue((self.destino / "instalacion" / nuevo.name).exists())
        self.assertIn("1.5.0", (self.destino / "version.txt").read_text(encoding="utf-8"))

    def test_los_cmd_usan_fin_de_linea_de_windows(self):
        nuevo = paquete(self.base / "almacen-ubicaciones_v1.5.0_a.zip", "1.5.0")
        preparar_usb.preparar(self.destino, nuevo)
        datos = (self.destino / "ACTUALIZAR.cmd").read_bytes()
        self.assertIn(b"\r\n", datos)
        # todos los saltos son CRLF: no queda ningún LF suelto
        self.assertEqual(datos.count(b"\n"), datos.count(b"\r\n"))

    def test_el_paquete_anterior_pasa_al_historico(self):
        viejo = paquete(self.base / "almacen-ubicaciones_v1.4.0_a.zip", "1.4.0")
        nuevo = paquete(self.base / "almacen-ubicaciones_v1.5.0_b.zip", "1.5.0")
        preparar_usb.preparar(self.destino, viejo)
        resultado = preparar_usb.preparar(self.destino, nuevo)
        self.assertIn(viejo.name, resultado["anteriores"])
        self.assertFalse((self.destino / "instalacion" / viejo.name).exists())
        self.assertTrue((self.destino / "versiones" / viejo.name).exists())
        self.assertTrue((self.destino / "instalacion" / nuevo.name).exists())

    def test_el_lanzador_busca_el_paquete_mas_reciente(self):
        nuevo = paquete(self.base / "almacen-ubicaciones_v1.5.0_c.zip", "1.5.0")
        preparar_usb.preparar(self.destino, nuevo)
        texto = (self.destino / "ACTUALIZAR.cmd").read_text(encoding="utf-8", errors="ignore")
        self.assertIn("dir /b /o-d", texto)
        self.assertIn("actualizar.py", texto)
        self.assertIn("--paquete", texto)

    def test_detecta_la_instalacion_del_escritorio(self):
        nuevo = paquete(self.base / "almacen-ubicaciones_v1.5.0_d.zip", "1.5.0")
        preparar_usb.preparar(self.destino, nuevo)
        texto = (self.destino / "ACTUALIZAR.cmd").read_text(encoding="utf-8", errors="ignore")
        self.assertIn(r"%USERPROFILE%\Desktop\almacen-ubicaciones", texto)
        self.assertIn(r"C:\Users\Almacen Migh\Desktop\almacen-ubicaciones", texto)
        self.assertIn("--instalada", texto)
        self.assertIn("app\\server.py", texto)

    def test_avisa_el_resultado(self):
        nuevo = paquete(self.base / "almacen-ubicaciones_v1.5.0_e.zip", "1.5.0")
        preparar_usb.preparar(self.destino, nuevo)
        texto = (self.destino / "ACTUALIZAR.cmd").read_text(encoding="utf-8", errors="ignore")
        self.assertIn("ACTUALIZACION CORRECTA", texto)
        self.assertIn("NO SE ACTUALIZO", texto)
        self.assertIn("ultima_actualizacion.txt", texto)

    def test_paquete_sin_version(self):
        sin = self.base / "sin.zip"
        with zipfile.ZipFile(sin, "w") as z:
            z.writestr("app/server.py", "# x\n")
        with self.assertRaises(ValueError):
            preparar_usb.preparar(self.destino, sin)


if __name__ == "__main__":
    unittest.main()
