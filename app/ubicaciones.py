"""Catálogo de ubicaciones físicas, derivado de ALM-PNO-002 §V.2.1.4."""

# (prefijo, desde, hasta, zona, tipo, permite_alergenos)
BLOQUES = [
    ("MPA-A",   1,  40, "MP-ALERGENO",     "RACK",  1),
    ("MPNA-A", 41, 130, "MP-NO-ALERGENO",  "RACK",  0),
    ("MPNA-B", 131, 136, "MP-NO-ALERGENO", "RACK",  0),
    ("MPT-B",   1,  20, "PT",              "RACK",  0),
    ("MEP-B",   1,  20, "EM",              "RACK",  0),
    ("MEP-C",  21, 100, "EM",              "RACK",  0),
    ("MEP-D", 101, 180, "EM",              "RACK",  0),
    ("MEP-E", 181, 220, "EM",              "RACK",  0),
]

ESPECIALES = [
    # codigo,            zona,        tipo,     permite_alergenos
    ("PASILLO-A",        "ESPECIAL", "PISO",   1),
    ("JAULA-MP",         "ESPECIAL", "JAULA",  1),
    ("JAULA-PT",         "ESPECIAL", "JAULA",  0),
    ("LAB-RETENCION",    "ESPECIAL", "CUARTO", 0),
    ("CUARENTENA-MP",    "ESPECIAL", "PISO",   1),
    ("ADUANA-MP",        "ESPECIAL", "PISO",   0),
    ("ANDEN",            "ESPECIAL", "ANDEN",  1),
]


def _construir():
    catalogo = {}
    for prefijo, desde, hasta, zona, tipo, alergenos in BLOQUES:
        for n in range(desde, hasta + 1):
            codigo = f"{prefijo}-{n:03d}"
            catalogo[codigo] = {"codigo": codigo, "zona": zona, "tipo": tipo,
                                "permite_alergenos": alergenos}
    for codigo, zona, tipo, alergenos in ESPECIALES:
        catalogo[codigo] = {"codigo": codigo, "zona": zona, "tipo": tipo,
                            "permite_alergenos": alergenos}
    return catalogo


CATALOGO = _construir()
CODIGOS = sorted(CATALOGO)
RACKS = [c for c in CODIGOS if CATALOGO[c]["tipo"] == "RACK"]


def existe(codigo):
    return (codigo or "").strip().upper() in CATALOGO


def normalizar(codigo):
    return (codigo or "").strip().upper().replace(" ", "")


def zona_de(codigo):
    u = CATALOGO.get(normalizar(codigo))
    return u["zona"] if u else "SIN UBICAR"


def por_zona():
    """{zona: [codigos]} en el orden en que el formato los imprime."""
    orden = ["MP-ALERGENO", "MP-NO-ALERGENO", "PT", "EM", "ESPECIAL"]
    salida = {z: [] for z in orden}
    for codigo in CODIGOS:
        salida[CATALOGO[codigo]["zona"]].append(codigo)
    return salida
