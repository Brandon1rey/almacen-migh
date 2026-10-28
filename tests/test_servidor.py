"""Pruebas de verdad por HTTP: el mismo camino que recorre el navegador.

Aquí se encontró uno de los peores fallos: el servidor no atendía PATCH, así que la
pantalla de Usuarios (cambiar rol, desactivar, reponer contraseña) devolvía 501 y no
hacía nada. Estas pruebas levantan el servidor real en un puerto libre.
"""
import http.cookiejar
import json
import pathlib
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer

from app import reportes, usuarios
from app import server as servidor
from tests.test_regresiones import _carpeta


class TestHttp(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        cls.dir = pathlib.Path(cls._tmp.name)
        _carpeta(cls.dir)
        reportes.SALIDAS = cls.dir / "salidas"
        cls.estado = servidor.Estado(carpeta=cls.dir)
        usuarios.crear(cls.estado.con, "Aseguramiento", "clave-segura-1", "admin")
        usuarios.crear(cls.estado.con, "Esteban", "clave-segura-2", "almacen")
        cls._estado_previo = servidor.ESTADO
        cls._log_previo = servidor.Handler.log_message
        servidor.ESTADO = cls.estado
        servidor.Handler.log_message = lambda *a, **k: None
        cls.servidor = ThreadingHTTPServer(("127.0.0.1", 0), servidor.Handler)
        cls.base = f"http://127.0.0.1:{cls.servidor.server_address[1]}"
        cls.hilo = threading.Thread(target=cls.servidor.serve_forever, daemon=True)
        cls.hilo.start()

    @classmethod
    def tearDownClass(cls):
        cls.servidor.shutdown()
        cls.servidor.server_close()
        cls.hilo.join(timeout=5)
        servidor.Handler.log_message = cls._log_previo
        servidor.ESTADO = cls._estado_previo
        for inv in list(cls.estado._invs.values()):
            inv.wb.close()
        cls.estado.con.close()
        cls._tmp.cleanup()

    # ---------- utilidades ----------
    def opener(self):
        return urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))

    def pedir(self, abridor, ruta, metodo="GET", cuerpo=None, crudo=None):
        datos = crudo
        if cuerpo is not None:
            datos = json.dumps(cuerpo).encode("utf-8")
        peticion = urllib.request.Request(self.base + ruta, data=datos, method=metodo,
                                          headers={"Content-Type": "application/json"})
        try:
            with abridor.open(peticion, timeout=20) as respuesta:
                return respuesta.status, dict(respuesta.headers), respuesta.read()
        except urllib.error.HTTPError as e:
            try:
                return e.code, dict(e.headers), e.read()
            finally:
                e.close()          # sin esto queda el archivo de la respuesta abierto

    def entrar(self, usuario, password):
        abridor = self.opener()
        codigo, _, cuerpo = self.pedir(abridor, "/api/login", "POST",
                                       {"usuario": usuario, "password": password})
        self.assertEqual(codigo, 200, cuerpo)
        return abridor

    def json_de(self, cuerpo):
        return json.loads(cuerpo.decode("utf-8"))

    # ---------- pruebas ----------
    def test_la_pagina_se_sirve_sin_cache(self):
        abridor = self.opener()
        codigo, cabeceras, cuerpo = self.pedir(abridor, "/")
        self.assertEqual(codigo, 200)
        self.assertIn("text/html", cabeceras.get("Content-Type", ""))
        self.assertIn("no-cache", cabeceras.get("Cache-Control", ""))
        self.assertIn("Control de ubicación", cuerpo.decode("utf-8"))

    def test_login_y_lecturas(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        codigo, _, cuerpo = self.pedir(abridor, "/api/inventario?tipo=mp")
        self.assertEqual(codigo, 200)
        datos = self.json_de(cuerpo)
        self.assertEqual(datos["tipo"], "mp")
        self.assertGreater(len(datos["filas"]), 0)

    def test_patch_de_usuario_ya_funciona(self):
        abridor = self.entrar("Aseguramiento", "clave-segura-1")
        _, _, cuerpo = self.pedir(abridor, "/api/usuarios")
        gente = {u["nombre"]: u for u in self.json_de(cuerpo)["usuarios"]}
        uid = gente["Esteban"]["id"]
        codigo, _, cuerpo = self.pedir(abridor, f"/api/usuarios/{uid}", "PATCH",
                                       {"rol": "consulta"})
        self.assertEqual(codigo, 200, cuerpo)
        _, _, cuerpo = self.pedir(abridor, "/api/usuarios")
        gente = {u["nombre"]: u for u in self.json_de(cuerpo)["usuarios"]}
        self.assertEqual(gente["Esteban"]["rol"], "consulta")
        # se regresa para no afectar a las demás pruebas
        self.pedir(abridor, f"/api/usuarios/{uid}", "PATCH", {"rol": "almacen"})

    def test_almacen_no_puede_tocar_usuarios(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        codigo, _, _ = self.pedir(abridor, "/api/usuarios/1", "PATCH", {"rol": "admin"})
        self.assertEqual(codigo, 403)

    def test_tipo_invalido_es_400(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        codigo, _, cuerpo = self.pedir(abridor, "/api/inventario?tipo=xyz")
        self.assertEqual(codigo, 400, cuerpo)
        self.assertIn("inventario desconocido", self.json_de(cuerpo)["error"])

    def test_parametros_basura_no_rompen(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        codigo, _, cuerpo = self.pedir(abridor, "/api/bitacora?limite=abc")
        self.assertEqual(codigo, 200, cuerpo)
        codigo, _, cuerpo = self.pedir(abridor, "/api/historial?fila=abc")
        self.assertEqual(codigo, 400, cuerpo)

    def test_nan_del_navegador_no_llega_al_excel(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        fila = self.estado.inventario("mp").filas[0]["fila"]
        crudo = json.dumps({"fila": fila, "destino": "MPA-A-020", "motivo": "reacomodo",
                            "tipo": "mp", "cantidad": float("nan")}).encode("utf-8")
        codigo, _, cuerpo = self.pedir(abridor, "/api/mover", "POST", crudo=crudo)
        self.assertEqual(codigo, 400, cuerpo)
        self.assertIn("cantidad", self.json_de(cuerpo)["error"])

    def test_json_invalido_es_400(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        codigo, _, cuerpo = self.pedir(abridor, "/api/mover", "POST", crudo=b"{no es json")
        self.assertEqual(codigo, 400, cuerpo)

    def test_descarga_f01(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        codigo, cabeceras, cuerpo = self.pedir(abridor, "/descargar/f01")
        self.assertEqual(codigo, 200, cuerpo[:200])
        self.assertIn("spreadsheetml", cabeceras.get("Content-Type", ""))
        self.assertEqual(cuerpo[:2], b"PK")

    def test_descarga_sin_sesion_es_401(self):
        abridor = self.opener()
        codigo, _, _ = self.pedir(abridor, "/descargar/f01")
        self.assertEqual(codigo, 401)

    def test_descarga_de_inventario_no_cargado(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        codigo, _, cuerpo = self.pedir(abridor, "/descargar/inventario?tipo=mep")
        self.assertEqual(codigo, 200, cuerpo[:120])

    def test_ruta_fuera_de_static(self):
        abridor = self.opener()
        codigo, _, _ = self.pedir(abridor, "/%2e%2e/app/api.py")
        self.assertIn(codigo, (400, 404))

    def test_logout_cierra_la_sesion(self):
        abridor = self.entrar("Esteban", "clave-segura-2")
        codigo, _, cuerpo = self.pedir(abridor, "/api/yo")
        self.assertEqual(codigo, 200)
        self.assertTrue(self.json_de(cuerpo)["autenticado"])
        self.assertEqual(self.pedir(abridor, "/api/logout", "POST", {})[0], 200)
        codigo, _, cuerpo = self.pedir(abridor, "/api/yo")
        self.assertEqual(codigo, 200)
        self.assertFalse(self.json_de(cuerpo)["autenticado"])

    def test_favicon_no_ensucia_la_consola(self):
        abridor = self.opener()
        codigo, _, cuerpo = self.pedir(abridor, "/favicon.ico")
        self.assertEqual(codigo, 204)
        self.assertEqual(cuerpo, b"")


if __name__ == "__main__":
    unittest.main()
