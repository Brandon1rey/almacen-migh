"""Regresiones de los fallos encontrados en la cacería de bugs.

Cada prueba de este archivo existe porque el sistema hacía algo mal y alguien lo podía
sufrir en el almacén: un 500 por un parámetro raro, un borrado que dejaba kilos fantasma,
un primer arranque que se caía después de crear el administrador, etc.
"""
import contextlib
import io
import pathlib
import shutil
import tempfile
import unittest

# `app` agrega vendor\ al sys.path; por eso se importa antes que openpyxl, para que
# estas pruebas también corran solas (`python -m unittest tests.test_regresiones`).
from app import api, bitacora, inventario, inventarios, reportes, ubicaciones, usuarios
from app import server as servidor
from tests.test_multitipo import ENCABEZADOS, filas_de, libro

import openpyxl  # noqa: E402  (después de app, que habilita vendor\)


def _carpeta(dir_datos):
    """Escribe los tres inventarios con el nombre que espera el servidor."""
    for tipo, nombre in (("mp", "Inventario_MP.xlsx"), ("mep", "Inventario_MEP.xlsx"),
                         ("pt", "Inventario_PT.xlsx")):
        libro(pathlib.Path(dir_datos) / nombre, filas_de(tipo))


class BaseInstalacion(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        # La limpieza del directorio se registra primero para que corra al final:
        # si corriera antes de cerrar las conexiones, Windows no puede borrar el .db.
        self.addCleanup(self._tmp.cleanup)
        self.dir = pathlib.Path(self._tmp.name)
        _carpeta(self.dir)
        self.estado = servidor.Estado(carpeta=self.dir)
        self.addCleanup(self.con_cerrar)
        self.con = self.estado.con
        usuarios.crear(self.con, "Aseguramiento", "clave-segura-1", "admin")
        usuarios.crear(self.con, "Esteban", "clave-segura-2", "almacen")
        usuarios.crear(self.con, "Auditor", "clave-segura-3", "consulta")
        self.admin = self.sesion("Aseguramiento")
        self.almacen = self.sesion("Esteban")
        self.consulta = self.sesion("Auditor")
        self.fila = self.estado.inventario("mp").filas[0]["fila"]
        self.estado.inventario("mp").asignar(self.fila, "MPA-A-005")
        self.estado.inventario("mp").guardar()

    def con_cerrar(self):
        api._INTENTOS.clear()          # el freno de contraseñas no se filtra entre pruebas
        for inv in list(self.estado._invs.values()):
            inv.wb.close()
        self.estado.con.close()

    def sesion(self, nombre):
        fila = usuarios.por_nombre(self.con, nombre)
        return {"usuario_id": fila["id"], "nombre": fila["nombre"], "rol": fila["rol"],
                "debe_cambiar": fila["debe_cambiar"]}

    def api(self, metodo, ruta, params=None, cuerpo=None, sesion=None, tipo="mp"):
        return api.manejar(self.estado, metodo, ruta, params or {}, cuerpo or {},
                           sesion if sesion is not None else self.almacen)

    def mover(self, **cambios):
        cuerpo = {"fila": self.fila, "destino": "MPA-A-020", "motivo": "reacomodo",
                  "tipo": "mp"}
        cuerpo.update(cambios)
        return self.api("POST", "/api/mover", cuerpo=cuerpo)

    def cantidad(self, fila=None):
        return float(self.estado.inventario("mp").por_fila(fila or self.fila)["cantidad"])


class TestArranque(BaseInstalacion):
    def test_la_bitacora_vive_junto_a_los_datos(self):
        """Antes, abrir otra carpeta de datos seguía escribiendo en data\\bitacora.db."""
        self.assertEqual(self.estado.ruta_bitacora, self.dir / "bitacora.db")
        self.assertTrue((self.dir / "bitacora.db").is_file())

    def test_primer_arranque_crea_admin_y_deja_las_credenciales(self):
        """El servidor se caía con NameError después de crear el administrador."""
        otra = pathlib.Path(tempfile.mkdtemp(prefix="primer-arranque-"))
        self.addCleanup(shutil.rmtree, otra, True)
        self.addCleanup(servidor.cerrar_registro)
        _carpeta(otra)
        salida = io.StringIO()
        with contextlib.redirect_stdout(salida):
            estado = servidor.arrancar(otra, servir=False)
        self.addCleanup(estado.con.close)
        archivo = otra / "CREDENCIALES_INICIALES.txt"
        self.assertTrue(archivo.is_file(), "no se escribieron las credenciales")
        texto = archivo.read_text(encoding="utf-8")
        self.assertIn("usuario : admin", texto)
        clave = texto.split("contraseña: ", 1)[1].splitlines()[0].strip()
        self.assertTrue(clave)
        self.assertIsNotNone(usuarios.verificar(estado.con, "admin", clave),
                             "la contraseña impresa no sirve para entrar")
        self.assertIn("PRIMER ARRANQUE", salida.getvalue())

    def test_primer_arranque_no_duplica_al_segundo_intento(self):
        otra = pathlib.Path(tempfile.mkdtemp(prefix="segundo-arranque-"))
        self.addCleanup(shutil.rmtree, otra, True)
        self.addCleanup(servidor.cerrar_registro)
        _carpeta(otra)
        with contextlib.redirect_stdout(io.StringIO()):
            primero = servidor.arrancar(otra, servir=False)
            segundo = servidor.arrancar(otra, servir=False)
        self.addCleanup(primero.con.close)
        self.addCleanup(segundo.con.close)
        self.assertEqual(segundo.con.execute("SELECT COUNT(*) FROM usuario").fetchone()[0], 1)


class TestParametros(BaseInstalacion):
    def test_tipo_invalido_es_400_no_500(self):
        codigo, cuerpo, _ = self.api("GET", "/api/inventario", {"tipo": "xyz"})
        self.assertEqual(codigo, 400, cuerpo)
        self.assertIn("inventario desconocido", cuerpo["error"])

    def test_limite_y_fila_basura_no_truenan(self):
        codigo, _, _ = self.api("GET", "/api/bitacora", {"limite": "abc"})
        self.assertEqual(codigo, 200)
        codigo, cuerpo, _ = self.api("GET", "/api/historial", {"fila": "abc"})
        self.assertEqual(codigo, 400, cuerpo)
        codigo, cuerpo, _ = self.api("GET", "/api/historial", {"fila": self.fila,
                                                               "limite": "muchos"})
        self.assertEqual(codigo, 200)

    def test_movimiento_no_numerico_es_404(self):
        for ruta in ("/api/deshacer", "/api/purgar"):
            codigo, cuerpo, _ = self.api("POST", ruta, cuerpo={"movimiento": "abc"}, sesion=self.admin)
            self.assertEqual(codigo, 404, (ruta, cuerpo))

    def test_cantidad_nan_o_infinito_se_rechaza(self):
        for valor in (float("nan"), float("inf"), "mucho", "-3"):
            codigo, cuerpo, _ = self.mover(cantidad=valor)
            self.assertEqual(codigo, 400, (valor, cuerpo))
        self.assertAlmostEqual(self.cantidad(), 1458)

    def test_cantidad_con_coma_decimal_se_acepta(self):
        codigo, cuerpo, _ = self.mover(cantidad="5,5")
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["accion"], "PARCIAL")
        self.assertAlmostEqual(cuerpo["cantidad"], 5.5)

    def test_ubicar_vacio_ya_no_borra_la_ubicacion(self):
        codigo, cuerpo, _ = self.api("POST", "/api/ubicar", cuerpo={"fila": self.fila,
                                                                    "ubicacion": ""})
        self.assertEqual(codigo, 400, cuerpo)
        self.assertEqual(self.estado.inventario("mp").por_fila(self.fila)["ubicacion"],
                         "MPA-A-005")

    def test_limpiar_si_quita_la_ubicacion(self):
        codigo, cuerpo, _ = self.api("POST", "/api/limpiar", cuerpo={"fila": self.fila})
        self.assertEqual(codigo, 200, cuerpo)
        self.assertIsNone(self.estado.inventario("mp").por_fila(self.fila)["ubicacion"])

    def test_no_encontrado_exige_nota(self):
        codigo, cuerpo, _ = self.api("POST", "/api/no_encontrado", cuerpo={"fila": self.fila})
        self.assertEqual(codigo, 400, cuerpo)
        codigo, _, _ = self.api("POST", "/api/no_encontrado",
                                cuerpo={"fila": self.fila, "nota": "no estaba en el rack"})
        self.assertEqual(codigo, 200)

    def test_destino_fuera_del_catalogo_es_400(self):
        codigo, cuerpo, _ = self.mover(destino="ZZZ-999")
        self.assertEqual(codigo, 400, cuerpo)
        self.assertIn("no existe en el catálogo", cuerpo["error"])

    def test_motivo_otro_exige_nota(self):
        codigo, cuerpo, _ = self.mover(motivo="otro")
        self.assertEqual(codigo, 400, cuerpo)
        self.assertIn("nota", cuerpo["error"])
        codigo, _, _ = self.mover(motivo="otro", nota="se juntó con otro lote")
        self.assertEqual(codigo, 200)


class TestParcialidades(BaseInstalacion):
    def test_purgar_parcial_devuelve_los_kilos(self):
        """Borrar el movimiento de una parcialidad dejaba los kilos en el limbo."""
        antes = self.cantidad()
        _, cuerpo, _ = self.mover(cantidad=7.5)
        self.assertEqual(cuerpo["accion"], "PARCIAL")
        nueva = cuerpo["fila_destino"]
        codigo, res, _ = self.api("POST", "/api/purgar", sesion=self.admin,
                                  cuerpo={"movimiento": cuerpo["movimiento"],
                                          "motivo": "captura equivocada"})
        self.assertEqual(codigo, 200, res)
        self.assertAlmostEqual(self.cantidad(), antes)
        parcial = self.estado.inventario("mp").por_fila(nueva)
        self.assertAlmostEqual(float(parcial["cantidad"] or 0), 0)
        self.assertIsNone(parcial["ubicacion"])
        self.assertIsNone(bitacora.por_id(self.con, cuerpo["movimiento"]))

    def test_purgar_parcial_ya_movida_no_rompe_los_datos(self):
        _, cuerpo, _ = self.mover(cantidad=7.5)
        nueva = cuerpo["fila_destino"]
        # la parcialidad se vuelve a mover a otro rack: ya no es reversible
        inv = self.estado.inventario("mp")
        inv.asignar(nueva, "MPA-A-030")
        inv.guardar()
        codigo, res, _ = self.api("POST", "/api/purgar", sesion=self.admin,
                                  cuerpo={"movimiento": cuerpo["movimiento"],
                                          "motivo": "captura equivocada"})
        self.assertEqual(codigo, 409, res)
        self.assertAlmostEqual(self.cantidad(), 1450.5)
        self.assertIsNotNone(bitacora.por_id(self.con, cuerpo["movimiento"]))

    def test_deshacer_parcial_dos_veces_no_duplica_kilos(self):
        antes = self.cantidad()
        _, cuerpo, _ = self.mover(cantidad=10)
        self.api("POST", "/api/deshacer", cuerpo={"movimiento": cuerpo["movimiento"]})
        codigo, _, _ = self.api("POST", "/api/deshacer", cuerpo={"movimiento": cuerpo["movimiento"]})
        self.assertEqual(codigo, 400)
        self.assertAlmostEqual(self.cantidad(), antes)


class TestBitacora(BaseInstalacion):
    def test_guarda_la_unidad_del_movimiento(self):
        """Sin unidad, la pantalla mostraba «1,450» sin decir de qué."""
        _, cuerpo, _ = self.mover()
        mov = bitacora.por_id(self.con, cuerpo["movimiento"])
        self.assertEqual(mov["unidad"], "kg")
        _, res, _ = self.api("GET", "/api/bitacora", {"inventario": "mp"})
        self.assertEqual(res["movimientos"][0]["unidad"], "kg")

    def test_confirmar_queda_registrado(self):
        codigo, cuerpo, _ = self.api("POST", "/api/ubicar",
                                     cuerpo={"fila": self.fila, "ubicacion": "MPA-A-005"})
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["accion"], "CONFIRMAR")
        self.assertIsNotNone(bitacora.por_id(self.con, cuerpo["movimiento"]))

    def test_bitacora_de_todos_los_inventarios(self):
        _, res, _ = self.api("GET", "/api/bitacora", {"inventario": "todos"})
        self.assertEqual(res["tipo"], None)

    def test_comentario_de_tipo_invalido_es_400(self):
        codigo, cuerpo, _ = self.api("GET", "/api/bitacora", {"inventario": "xyz"})
        self.assertEqual(codigo, 400, cuerpo)


class TestCantidadesIlegibles(BaseInstalacion):
    def _libro_con_texto(self, texto):
        ruta = self.dir / "Inventario_MP.xlsx"
        libro(ruta, [["WPC 80%", "VILHER", "Proteína", "kg", texto, "2028-04-23",
                      "22260425002", "MIGH-MP114", "", None]], encabezados=ENCABEZADOS)
        return inventario.Inventario(ruta, tipo="mp")

    def test_texto_numerico_inequivoco_se_lee(self):
        inv = self._libro_con_texto("1,458.50")
        self.assertEqual(inv.filas[0]["cantidad"], 1458.5)
        self.assertTrue(inv.filas[0]["cantidad_legible"])
        inv.wb.close()

    def test_cantidad_ambigua_se_reporta_y_no_se_adivina(self):
        inv = self._libro_con_texto("1.458")
        registro = inv.filas[0]
        self.assertFalse(registro["cantidad_legible"])
        self.assertEqual(registro["cantidad"], "1.458")
        self.assertEqual(inv.resumen()["sin_cantidad_legible"], 1)
        self.assertEqual(inv.resumen()["filas_sin_cantidad"], [4])
        self.assertEqual(len(inv.buscar(solo_ilegibles=True)), 1)
        # el renglón ilegible no se cuela en los totales
        self.assertEqual(inv.resumen()["total"], 0)
        inv.wb.close()

    def test_mover_un_renglon_ilegible_da_mensaje_claro(self):
        self._libro_con_texto("1.458").wb.close()
        estado = servidor.Estado(carpeta=self.dir)
        self.addCleanup(estado.con.close)
        estado.inventario("mp").asignar(4, "MPA-A-001")
        fila = estado.inventario("mp").filas[0]["fila"]
        codigo, cuerpo, _ = api.manejar(estado, "POST", "/api/mover", {},
                                        {"fila": fila, "destino": "MPA-A-020",
                                         "motivo": "reacomodo", "tipo": "mp"}, self.almacen)
        self.assertEqual(codigo, 400, cuerpo)
        self.assertIn("CANTIDAD", cuerpo["error"])
        estado.inventario("mp").wb.close()


class TestReportes(BaseInstalacion):
    def test_f01_no_incluye_renglones_anulados(self):
        inv = self.estado.inventario("mp")
        inv.anular(self.fila, "anulado para la prueba")
        inv.guardar()
        reportes.SALIDAS = self.dir / "salidas"
        ruta = reportes.generar_f01([("mp", self.estado.inventario("mp"))], version="2",
                                    destino=self.dir / "f01.xlsx")
        wb = openpyxl.load_workbook(ruta)
        try:
            self.assertIn("MP", wb.sheetnames)
            texto = " ".join(str(c.value) for fila in wb["MP"].iter_rows()
                             for c in fila if c.value is not None)
            self.assertNotIn("WPC 80%", texto)
            self.assertIn("Total de lotes: 1", texto)
        finally:
            wb.close()

    def test_bitacora_exportada_trae_la_unidad(self):
        self.mover()
        reportes.SALIDAS = self.dir / "salidas"
        ruta = reportes.exportar_bitacora(self.con, destino=self.dir / "bitacora.xlsx")
        wb = openpyxl.load_workbook(ruta)
        try:
            hoja = wb["Bitácora"]
            encabezados = [hoja.cell(row=4, column=c).value for c in range(1, 15)]
            self.assertIn("UNIDAD", encabezados)
            self.assertEqual(hoja.cell(row=5, column=10).value, "kg")
        finally:
            wb.close()


class TestSesionYUsuarios(BaseInstalacion):
    def test_bloqueo_temporal_tras_ocho_fallos(self):
        codigo = 0
        for _ in range(8):
            codigo, _, _ = self.api("POST", "/api/login",
                                    cuerpo={"usuario": "Esteban", "password": "mala"})
            self.assertEqual(codigo, 401)
        codigo, cuerpo, _ = self.api("POST", "/api/login",
                                     cuerpo={"usuario": "Esteban", "password": "mala"})
        self.assertEqual(codigo, 429, cuerpo)
        # el usuario correcto puede entrar cuando pasa la ventana; aquí comprobamos
        # que el mensaje dice cuánto esperar
        self.assertIn("espera", cuerpo["error"])

    def test_login_sin_datos_es_400(self):
        codigo, cuerpo, _ = self.api("POST", "/api/login", cuerpo={})
        self.assertEqual(codigo, 400)
        self.assertIn("usuario", cuerpo["error"])

    def test_no_puedo_purgar_un_usuario_inexistente(self):
        codigo, cuerpo, _ = self.api("PATCH", "/api/usuarios/99999",
                                     cuerpo={"rol": "almacen"}, sesion=self.admin)
        self.assertEqual(codigo, 404, cuerpo)

    def test_cambio_de_clave_igual_es_400(self):
        codigo, cuerpo, _ = self.api("POST", "/api/password",
                                     cuerpo={"actual": "clave-segura-2",
                                             "nueva": "clave-segura-2"})
        self.assertEqual(codigo, 400, cuerpo)


class TestEntradaTolerante(BaseInstalacion):
    """Capturar datos a mano: códigos escritos de cualquier forma y cantidades con coma."""

    def test_codigos_escritos_de_cualquier_forma(self):
        for escrito in ("mpa a 5", "MPAA005", "mpa-a-5", " MPA-A-005 ", "mpaa-005", "MPA A 005"):
            self.assertEqual(ubicaciones.normalizar(escrito), "MPA-A-005", escrito)
        self.assertEqual(ubicaciones.normalizar("jaula mp"), "JAULA-MP")
        self.assertEqual(ubicaciones.normalizar("PASILLO A"), "PASILLO-A")
        self.assertEqual(ubicaciones.normalizar("mep e 220"), "MEP-E-220")
        self.assertFalse(ubicaciones.existe("MPA-A-999"))
        self.assertFalse(ubicaciones.existe("pasillo b"))

    def test_mover_acepta_codigo_sin_guiones(self):
        codigo, cuerpo, _ = self.mover(destino="mpaa20")
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["nueva"], "MPA-A-020")

    def test_cantidades_escritas_a_mano(self):
        self.assertEqual(inventario.interpretar_cantidad("5 kg", total=150), 5)
        self.assertEqual(inventario.interpretar_cantidad("12,5"), 12.5)
        self.assertEqual(inventario.interpretar_cantidad("1 458", total=1500), 1458)
        self.assertEqual(inventario.interpretar_cantidad("1,458.5"), 1458.5)
        self.assertEqual(inventario.interpretar_cantidad("1.458,5"), 1458.5)
        self.assertEqual(inventario.interpretar_cantidad("1,458", total=2000), 1458)
        self.assertEqual(inventario.interpretar_cantidad("1.458", total=2000), 1458)
        self.assertEqual(inventario.interpretar_cantidad("1.4585"), 1.4585)
        for basura in ("mucho", "", None, True, "1.2.3"):
            self.assertIsNone(inventario.interpretar_cantidad(basura), basura)

    def test_lectura_ambigua_no_cabe_en_el_total(self):
        # en un renglón de 2 kg, «1,458» no puede ser 1458: se lee como 1.458
        self.assertEqual(inventario.interpretar_cantidad("1,458", total=2), 1.458)
        # sin total manda la lectura mexicana de miles
        self.assertEqual(inventario.interpretar_cantidad("1,458"), 1458)

    def test_mover_con_cantidad_con_unidad(self):
        codigo, cuerpo, _ = self.mover(cantidad="5 kg")
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["accion"], "PARCIAL")
        self.assertAlmostEqual(cuerpo["cantidad"], 5)

    def test_unidad_escrita_tiene_que_coincidir(self):
        self.assertEqual(inventario.interpretar_cantidad("5 kg", total=150, unidad="kg"), 5)
        self.assertEqual(inventario.interpretar_cantidad("5 pzas", total=50, unidad="pza"), 5)
        # 5 g en un lote de kg movería 5 kg en vez de 5 gramos: se rechaza
        self.assertIsNone(inventario.interpretar_cantidad("5 g", total=150, unidad="kg"))
        self.assertIsNone(inventario.interpretar_cantidad("5 kg", total=50, unidad="pza"))
        # unidades desconocidas y formatos rotos tampoco se adivinan
        self.assertIsNone(inventario.interpretar_cantidad("5 m", total=10, unidad="kg"))
        self.assertIsNone(inventario.interpretar_cantidad("1.2.3", total=10, unidad="kg"))

    def test_mover_con_unidad_equivocada_es_400(self):
        codigo, cuerpo, _ = self.mover(cantidad="5 g")
        self.assertEqual(codigo, 400, cuerpo)
        self.assertIn("unidad", cuerpo["error"])

    def test_busqueda_tolera_errores_de_dedo(self):
        inv = self.estado.inventario("mp")
        self.assertEqual(inv.buscar("protiena"), [])          # exacta: no hay
        aproximados = inv.buscar("protiena", difuso=True)
        self.assertTrue(aproximados)
        self.assertIn("WPC", aproximados[0]["descripcion"].upper())
        self.assertTrue(inv.sugerencias("protiena"))
        self.assertEqual(inv.sugerencias("zzzzzz"), [])

    def test_la_busqueda_normal_no_es_difusa(self):
        codigo, cuerpo, _ = self.api("GET", "/api/inventario", {"q": "wpc"})
        self.assertEqual(codigo, 200, cuerpo)
        self.assertFalse(cuerpo["difuso"])
        self.assertTrue(cuerpo["filas"])

    def test_api_ofrece_aproximados_y_sugerencias(self):
        codigo, cuerpo, _ = self.api("GET", "/api/inventario", {"q": "protiena"})
        self.assertEqual(codigo, 200, cuerpo)
        self.assertTrue(cuerpo["difuso"])
        self.assertTrue(cuerpo["filas"])
        self.assertTrue(cuerpo["sugerencias"])
        self.assertEqual(cuerpo["consulta"], "protiena")

    def test_api_sin_parecido_no_inventa_sugerencias(self):
        codigo, cuerpo, _ = self.api("GET", "/api/inventario", {"q": "zzzzzz"})
        self.assertEqual(codigo, 200, cuerpo)
        self.assertEqual(cuerpo["filas"], [])
        self.assertFalse(cuerpo["difuso"])
        self.assertEqual(cuerpo["sugerencias"], [])


if __name__ == "__main__":
    unittest.main()
