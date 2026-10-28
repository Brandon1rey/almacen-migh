"""Sugerencias de ubicación al asignar: orden por utilidad, no alfabético."""
import unittest

from app import ubicaciones


def codigos(sugerencias):
    return [s["codigo"] for s in sugerencias]


class TestBusqueda(unittest.TestCase):
    def test_codigo_exacto_primero(self):
        self.assertEqual(codigos(ubicaciones.sugerir("MPNA-A-105", "mp"))[0], "MPNA-A-105")

    def test_prefijo_gana_a_coincidencia_interna(self):
        resultado = codigos(ubicaciones.sugerir("MPA-A-00", "mp"))
        self.assertTrue(resultado[0].startswith("MPA-A-00"))
        self.assertNotIn("MPNA-A-041", resultado)

    def test_numero_suelto_encuentra_el_rack(self):
        self.assertIn("MPNA-A-105", codigos(ubicaciones.sugerir("105", "mp")))

    def test_forma_corta_a105(self):
        self.assertEqual(codigos(ubicaciones.sugerir("a105", "mp"))[0], "MPNA-A-105")

    def test_nombre_de_zona(self):
        resultado = codigos(ubicaciones.sugerir("alergeno", "mp"))
        self.assertTrue(resultado)
        self.assertTrue(resultado[0].startswith("MPA-A"))
        # la zona «no alérgenos» también aparece, pero siempre después
        if "MPNA-A-041" in resultado:
            self.assertGreater(resultado.index("MPNA-A-041"), resultado.index("MPA-A-001"))

    def test_sin_acentos_ni_minusculas(self):
        self.assertIn("CUARENTENA-MP", codigos(ubicaciones.sugerir("cuarenténa", "mp")))

    def test_no_sugiere_produccion(self):
        self.assertNotIn("PRODUCCION", codigos(ubicaciones.sugerir("PRODUC", "mp")))

    def test_respeta_el_limite(self):
        self.assertLessEqual(len(ubicaciones.sugerir("MP", "mp", limite=5)), 5)

    def test_consulta_sin_coincidencias(self):
        self.assertEqual(ubicaciones.sugerir("ZZZZ", "mp"), [])


class TestInventario(unittest.TestCase):
    def test_mp_no_sugiere_racks_de_empaque(self):
        resultado = codigos(ubicaciones.sugerir("", "mp", limite=400))
        self.assertTrue(resultado)
        self.assertFalse([c for c in resultado if c.startswith("MEP-")])
        self.assertFalse([c for c in resultado if c.startswith("MPT-")])

    def test_empaque_solo_empaque(self):
        resultado = codigos(ubicaciones.sugerir("", "mep", limite=400))
        self.assertTrue(resultado)
        self.assertFalse([c for c in resultado if c.startswith("MPA-") or c.startswith("MPNA-")])

    def test_pt_solo_producto_terminado(self):
        resultado = codigos(ubicaciones.sugerir("", "pt", limite=400))
        self.assertTrue(resultado)
        self.assertFalse([c for c in resultado if c.startswith("MEP-") or c.startswith("MPNA-")])

    def test_jaula_pt_solo_para_producto_terminado(self):
        self.assertNotIn("JAULA-PT", codigos(ubicaciones.sugerir("jaula", "mp")))
        self.assertIn("JAULA-PT", codigos(ubicaciones.sugerir("jaula", "pt")))


class TestOcupacion(unittest.TestCase):
    def setUp(self):
        self.ocupacion = {
            "MPA-A-001": {"lotes": 2, "cantidad": 500, "unidad": "kg", "material": "LECHE EN POLVO"},
            "MPA-A-002": {"lotes": 1, "cantidad": 200, "unidad": "kg", "material": "LECHE EN POLVO"},
            "MPA-A-003": {"lotes": 1, "cantidad": 80, "unidad": "kg", "material": "AZUCAR"},
        }

    def test_libre_gana_a_ocupada(self):
        resultado = codigos(ubicaciones.sugerir("MPA-A-00", "mp", self.ocupacion, limite=12))
        self.assertEqual(resultado[0], "MPA-A-004")
        self.assertIn("MPA-A-001", resultado)

    def test_marca_libre_y_ocupada(self):
        resultado = {s["codigo"]: s for s in
                     ubicaciones.sugerir("MPA-A-00", "mp", self.ocupacion, limite=12)}
        self.assertTrue(resultado["MPA-A-004"]["libre"])
        self.assertFalse(resultado["MPA-A-001"]["libre"])
        self.assertEqual(resultado["MPA-A-001"]["lotes"], 2)

    def test_prefiere_consolidar_el_mismo_material(self):
        resultado = codigos(ubicaciones.sugerir("MPA-A-00", "mp", self.ocupacion,
                                                material="LECHE EN POLVO", limite=12))
        self.assertLess(resultado.index("MPA-A-002"), resultado.index("MPA-A-003"))


class TestAlergenos(unittest.TestCase):
    def test_material_con_alergeno_prefiere_zona_de_alergenos(self):
        resultado = codigos(ubicaciones.sugerir("", "mp", limite=400, material="SUERO DE LECHE"))
        self.assertTrue(resultado[0].startswith("MPA-A"))

    def test_avisa_si_la_zona_no_admite_alergenos(self):
        sugerencias = {s["codigo"]: s for s in ubicaciones.sugerir("", "mp", limite=400,
                                                                    material="LECHE EN POLVO")}
        self.assertTrue(sugerencias["MPNA-A-041"]["aviso"])
        self.assertFalse(sugerencias["MPA-A-001"]["aviso"])

    def test_material_sin_alergeno_no_se_avisa(self):
        sugerencias = {s["codigo"]: s for s in ubicaciones.sugerir("", "mp", limite=400,
                                                                    material="AZUCAR")}
        self.assertFalse(sugerencias["MPNA-A-041"]["aviso"])

    def test_reconoce_los_alergenos_de_la_lista_maestra(self):
        for material in ("LECHE EN POLVO", "LACTOSA", "WPC 80%", "SOYA", "TRIGO", "HUEVO",
                         "CASEINATO DE SODIO", "ALBUMINA"):
            self.assertTrue(ubicaciones.es_alergeno(material), material)
        for material in ("AZUCAR", "MALTODEXTRINA", "ALMIDON DE MAIZ"):
            self.assertFalse(ubicaciones.es_alergeno(material), material)


class TestMotivo(unittest.TestCase):
    def test_cada_sugerencia_explica_por_que(self):
        resultado = ubicaciones.sugerir("MPA-A-00", "mp")
        self.assertTrue(resultado[0]["motivo"])
        self.assertIn("libre", resultado[0]["motivo"].lower())


if __name__ == "__main__":
    unittest.main()
