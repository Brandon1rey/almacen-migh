"""Registro de depuración y bus de eventos: lo que hace falta para diagnosticar sin adivinar."""
import pathlib
import tempfile
import unittest

from app import api, eventos, logger  # noqa: E402  (app primero: habilita vendor\)

from tests.test_regresiones import BaseInstalacion


class TestLogger(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.dir = pathlib.Path(self._tmp.name)
        logger.reiniciar()
        self.addCleanup(logger.reiniciar)
        self.ruta = logger.configurar(self.dir, "DEBUG", consola=False)

    def leer(self):
        return pathlib.Path(self.ruta).read_text(encoding="utf-8")

    def test_escribe_en_la_carpeta_de_datos(self):
        self.assertEqual(pathlib.Path(self.ruta).parent, self.dir / "logs")
        logger.evento("prueba.hola", usuario="Ana", fila=4)
        self.assertIn("evento=prueba.hola", self.leer())

    def test_no_registra_contrasenas_ni_tokens(self):
        logger.evento("sesion.iniciada", usuario="Ana", password="Super.Secreta.9", token="t0k3n")
        texto = self.leer()
        self.assertNotIn("Super.Secreta.9", texto)
        self.assertNotIn("t0k3n", texto)
        self.assertIn("password=***", texto)
        self.assertIn("token=***", texto)

    def test_peticion_guarda_metodo_ruta_estado_y_usuario(self):
        logger.peticion("POST", "/api/mover", 200, 12.34, "Esteban")
        linea = [l for l in self.leer().splitlines() if "peticion POST" in l][-1]
        self.assertIn("-> 200", linea)
        self.assertIn("12.3ms", linea)
        self.assertIn("usuario=Esteban", linea)

    def test_un_error_del_servidor_queda_como_error(self):
        logger.peticion("GET", "/api/lo_que_sea", 500, 1, None)
        self.assertIn("ERROR", self.leer())

    def test_configurar_dos_veces_no_duplica(self):
        logger.configurar(self.dir, "DEBUG", consola=False)
        logger.evento("una.sola")
        self.assertEqual(self.leer().count("evento=una.sola"), 1)

    def test_los_valores_gigantes_se_recortan(self):
        logger.evento("prueba", detalle="x" * 5000)
        linea = [l for l in self.leer().splitlines() if "evento=prueba" in l][-1]
        self.assertLess(len(linea), 1200)


class TestEventos(unittest.TestCase):
    def setUp(self):
        eventos.reiniciar()
        self.addCleanup(eventos.reiniciar)

    def test_publicar_y_escuchar(self):
        vistos = []
        eventos.escuchar("movimiento.registrado", vistos.append)
        eventos.publicar("movimiento.registrado", fila=4, accion="MOVER")
        eventos.publicar("otra.cosa")
        self.assertEqual(len(vistos), 1)
        self.assertEqual(vistos[0]["fila"], 4)
        self.assertEqual(vistos[0]["accion"], "MOVER")
        self.assertIn("id", vistos[0])
        self.assertIn("ts", vistos[0])

    def test_escucha_por_prefijo(self):
        vistos = []
        eventos.escuchar("movimiento.*", vistos.append)
        eventos.publicar("movimiento.registrado")
        eventos.publicar("sesion.iniciada")
        self.assertEqual(len(vistos), 1)

    def test_un_escucha_roto_no_tumba_la_accion(self):
        eventos.escuchar("*", lambda registro: 1 / 0)
        registro = eventos.publicar("prueba", valor=1)
        self.assertEqual(registro["evento"], "prueba")
        self.assertEqual(len(eventos.ultimos()), 1)

    def test_historial_y_desde(self):
        primero = eventos.publicar("uno")
        eventos.publicar("dos")
        self.assertEqual([r["evento"] for r in eventos.desde(primero["id"])], ["dos"])
        self.assertEqual(len(eventos.ultimos(1)), 1)

    def test_no_guarda_objetos_vivos(self):
        class Raro:
            def __repr__(self):
                return "raro"

        registro = eventos.publicar("prueba", objeto=Raro(), lista=[1, "dos", object()])
        self.assertIsInstance(registro["objeto"], str)
        self.assertEqual(registro["lista"], [1, "dos"])

    def test_el_escucha_del_registro_escribe_cada_evento(self):
        import shutil
        carpeta = pathlib.Path(tempfile.mkdtemp(prefix="registro-"))
        self.addCleanup(shutil.rmtree, carpeta, True)
        self.addCleanup(logger.reiniciar)
        ruta = logger.configurar(carpeta, "DEBUG", consola=False)
        eventos.conectar_registro()
        eventos.publicar("sesion.iniciada", usuario="Ana")
        texto = pathlib.Path(ruta).read_text(encoding="utf-8")
        logger.reiniciar()          # cierra el archivo antes de que se borre la carpeta
        self.assertIn("evento=sesion.iniciada", texto)
        self.assertIn("usuario=Ana", texto)


class TestApiDepuracion(BaseInstalacion):
    def setUp(self):
        super().setUp()
        eventos.reiniciar()
        self.addCleanup(eventos.reiniciar)
        logger.reiniciar()
        self.addCleanup(logger.reiniciar)
        self.ruta_log = logger.configurar(self.dir, "DEBUG", consola=False)
        eventos.conectar_registro()

    def nombres(self, limite=30):
        return [e["evento"] for e in eventos.ultimos(limite)]

    def test_un_movimiento_publica_evento_y_queda_en_el_log(self):
        codigo, cuerpo, _ = self.mover()
        self.assertEqual(codigo, 200, cuerpo)
        self.assertIn("movimiento.registrado", self.nombres())
        texto = pathlib.Path(self.ruta_log).read_text(encoding="utf-8")
        self.assertIn("evento=movimiento.registrado", texto)
        self.assertIn("accion=MOVER", texto)

    def test_login_y_logout_publican_eventos(self):
        codigo, _, _ = api.manejar(self.estado, "POST", "/api/login", {},
                                   {"usuario": "Esteban", "password": "clave-segura-2"})
        self.assertEqual(codigo, 200)
        self.assertIn("sesion.iniciada", self.nombres())

    def test_diagnostico_solo_para_admin(self):
        codigo, cuerpo, _ = self.api("GET", "/api/diagnostico", sesion=self.admin)
        self.assertEqual(codigo, 200, cuerpo)
        for clave in ("version", "python", "plataforma", "carpeta_datos", "inventarios",
                      "catalogo", "bitacora", "eventos", "log", "nivel_log", "disco"):
            self.assertIn(clave, cuerpo)
        self.assertTrue(cuerpo["log"].endswith("depuracion.log"), cuerpo["log"])
        codigo2, _, _ = self.api("GET", "/api/diagnostico", sesion=self.almacen)
        self.assertEqual(codigo2, 403)

    def test_eventos_recientes_solo_para_admin(self):
        codigo, cuerpo, _ = self.api("GET", "/api/eventos", sesion=self.admin)
        self.assertEqual(codigo, 200, cuerpo)
        self.assertIsInstance(cuerpo["eventos"], list)
        codigo2, _, _ = self.api("GET", "/api/eventos", sesion=self.almacen)
        self.assertEqual(codigo2, 403)

    def test_el_navegador_reporta_un_error(self):
        codigo, cuerpo, _ = self.api("POST", "/api/log", sesion=self.almacen, cuerpo={
            "nivel": "error", "mensaje": "TypeError: x is undefined",
            "contexto": {"linea": 42}, "url": "/"})
        self.assertEqual(codigo, 200, cuerpo)
        self.assertIn("navegador.error", self.nombres())
        texto = pathlib.Path(self.ruta_log).read_text(encoding="utf-8")
        self.assertIn("navegador", texto)
        self.assertIn("TypeError", texto)

    def test_reporte_vacio_se_rechaza(self):
        codigo, _, _ = self.api("POST", "/api/log", sesion=self.almacen, cuerpo={"mensaje": " "})
        self.assertEqual(codigo, 400)

    def test_sin_sesion_no_se_puede_reportar(self):
        codigo, _, _ = api.manejar(self.estado, "POST", "/api/log", {}, {"mensaje": "x"}, None)
        self.assertEqual(codigo, 401)


if __name__ == "__main__":
    unittest.main()
