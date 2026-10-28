"""Catálogo de ubicaciones físicas (fuente única) y sugerencias para asignar.

Los códigos válidos salen de ALM-PNO-002 §V.2.1.4 y SOLO de aquí: `canonico()` es la
única puerta de entrada para escribir una ubicación. La captura tolerante («mpa a 5»,
«MPAA005») se convierte al código del catálogo; lo que no está en el catálogo se rechaza
y nunca llega al Excel.
"""
import re
import unicodedata

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


def clave(codigo):
    """Forma compacta de un código: MPA-A-005, mpa a 5 y MPAA005 -> 'MPAA5'.

    Las zonas especiales conservan su nombre sin guiones: JAULA-MP -> 'JAULAMP'.
    """
    texto = str(codigo or "").strip().upper()
    texto = re.sub(r"[\s_.]+", "-", texto).strip("-")
    if not texto:
        return ""
    partes = [p for p in texto.split("-") if p]
    if len(partes) >= 2 and partes[-1].isdigit() and partes[-2].isalpha():
        prefijo = "".join(partes[:-1])
        return f"{prefijo}{int(partes[-1])}"
    compacto = "".join(partes)
    # 'MPAA005' (todo pegado) también debe caer en MPA-A-005
    coincidencia = re.fullmatch(r"([A-Z]+)(\d+)", compacto)
    if coincidencia:
        return f"{coincidencia.group(1)}{int(coincidencia.group(2))}"
    return compacto


# clave compacta -> código canónico (para aceptar 'mpa-a-5' y 'MPAA005')
CLAVES = {clave(codigo): codigo for codigo in CODIGOS}
# 'PASILLO A', 'pasilloa', 'PASILLO-A' -> PASILLO-A
CLAVES.update({codigo.replace("-", ""): codigo for codigo in CODIGOS})


def canonico(codigo):
    """Código del catálogo o None.

    Es la ÚNICA forma aceptable de escribir una ubicación: la captura tolerante
    («mpa a 5», «MPAA005») se convierte al código canónico, y lo que no está en el
    catálogo no se guarda jamás. Así no se pueden crear ubicaciones fantasma.
    """
    if codigo is None:
        return None
    texto = str(codigo).strip().upper()
    if not texto:
        return None
    if texto in CATALOGO:
        return texto
    return CLAVES.get(clave(texto))


def existe(codigo):
    return canonico(codigo) is not None


def normalizar(codigo):
    """Para mostrar y para los mensajes: el canónico si existe; si no, el texto limpio.

    No usar el resultado para escribir en el Excel: para eso está `canonico()`.
    """
    if codigo is None:
        return ""
    texto = str(codigo).strip().upper()
    if not texto:
        return ""
    return canonico(texto) or texto


def zona_de(codigo):
    u = CATALOGO.get(canonico(codigo))
    return u["zona"] if u else "SIN UBICAR"


def por_zona():
    """{zona: [codigos]} en el orden en que el formato los imprime."""
    orden = ["MP-ALERGENO", "MP-NO-ALERGENO", "PT", "EM", "ESPECIAL"]
    salida = {z: [] for z in orden}
    for codigo in CODIGOS:
        salida[CATALOGO[codigo]["zona"]].append(codigo)
    return salida


# ============================ sugerencias al asignar ============================
# El acomodo no es alfabético: primero lo que sirve. Reglas, en orden de peso:
#   1. qué tan bien coincide lo que escribe el usuario (código exacto, empieza, termina,
#      contiene, o el nombre de la zona),
#   2. que la zona sea del inventario correcto (una bolsa no va en el rack de materias primas),
#   3. que el rack esté libre, o que ya tenga el mismo material (consolidar),
#   4. que la zona admita alérgenos cuando el material lo es (ALM-PNO-002 §V.2.1.4).
# Las sugerencias SOLO ordenan códigos del catálogo; nunca crean uno nuevo.

ZONAS_INVENTARIO = {"mp": ("MP-ALERGENO", "MP-NO-ALERGENO"), "mep": ("EM",), "pt": ("PT",)}

ESPECIALES_POR_INVENTARIO = {
    "mp": {"PASILLO-A", "JAULA-MP", "CUARENTENA-MP", "ADUANA-MP", "ANDEN", "LAB-RETENCION"},
    "mep": {"PASILLO-A", "ANDEN", "LAB-RETENCION", "CUARENTENA-MP"},
    "pt": {"JAULA-PT", "ANDEN", "LAB-RETENCION"},
}

ZONA_TEXTO = {"MP-ALERGENO": "zona alergenos", "MP-NO-ALERGENO": "zona no alergenos",
              "PT": "producto terminado", "EM": "material de empaque", "ESPECIAL": "zona especial"}

# Lista maestra de alérgenos (F02-ALM-PNO-001): se busca por palabra, no por SKU.
PALABRAS_ALERGENO = (
    "LECHE", "LACTOSA", "LACTEO", "LACTEOS", "SUERO", "WPC", "WPI", "WPH", "MPI",
    "CASEINA", "CASEINATO", "MANTEQUILLA", "CREMA", "QUESO",
    "SOYA", "SOJA", "TRIGO", "GLUTEN", "CENTENO", "CEBADA", "AVENA",
    "PESCADO", "HUEVO", "ALBUMINA", "ALMENDRA", "AVELLANA", "NUEZ", "CACAHUATE", "MANI",
    "AJONJOLI", "SESAMO", "APIO", "MOSTAZA", "SULFITO",
)


def normalizar_texto(texto):
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return " ".join("".join(c if c.isalnum() else " " for c in s.upper()).split())


def es_alergeno(material):
    """¿El material declara alérgeno? Se compara por palabra completa, sin acentos."""
    texto = normalizar_texto(material)
    return any(palabra in texto.split() for palabra in PALABRAS_ALERGENO)


def _zonas_para(tipo_inventario):
    tipo = (tipo_inventario or "mp").lower()
    return ZONAS_INVENTARIO.get(tipo, ZONAS_INVENTARIO["mp"]), \
        ESPECIALES_POR_INVENTARIO.get(tipo, ESPECIALES_POR_INVENTARIO["mp"])


def _puntaje_texto(codigo, consulta):
    """Cuánto se parece el código a lo escrito, y con qué motivo."""
    q = consulta.replace(" ", "")
    if not q:
        return 0, None
    compacto = codigo.replace("-", "")
    if compacto == q:
        return 1000, "código exacto"
    if compacto.startswith(q):
        return 400, None
    if compacto.endswith(q):
        return 250, None
    if any(parte.startswith(q) for parte in codigo.split("-")):
        return 300, None
    if q in compacto:
        return 150, None
    return 0, None


def _puntaje_zona(zona, consulta):
    """Coincidencia contra el nombre de la zona (permite escribir «alergeno» o «cuarentena»)."""
    if not consulta:
        return 0, None
    palabras = set(normalizar_texto(ZONA_TEXTO.get(zona, zona)).split())
    for token in consulta.split():
        # «alergeno» debe encontrar «alergenos»: se compara por prefijo en ambos sentidos
        if not any(p.startswith(token) or (len(token) >= 4 and token.startswith(p))
                   for p in palabras if len(p) >= 3):
            return 0, None
    if "NO" in palabras:          # «zona no alergenos» no es la zona de alérgenos
        return 40, "zona que no admite alérgenos"
    return 120, f"zona {ZONA_TEXTO.get(zona, zona).replace('zona ', '')}"


def sugerir(consulta="", tipo_inventario="mp", ocupacion=None, material=None, limite=8):
    """Ubicaciones ordenadas para asignar, cada una con su motivo.

    ocupacion: {codigo: {"lotes": n, "cantidad": x, "unidad": "kg", "material": "..."}}
    """
    ocupacion = ocupacion or {}
    zonas, especiales = _zonas_para(tipo_inventario)
    texto_q = normalizar_texto(consulta)
    q = texto_q.replace(" ", "")
    alergeno = es_alergeno(material)

    filas = []
    for codigo in CODIGOS:
        datos = CATALOGO[codigo]
        zona = datos["zona"]
        if zona == "ESPECIAL":
            if codigo not in especiales:
                continue
        elif zona not in zonas:
            continue

        puntos, motivo = _puntaje_texto(codigo, q)
        pz, mz = _puntaje_zona(zona, texto_q)
        if q or texto_q:
            if not puntos and not pz:
                continue
        puntos += pz
        if mz and not motivo:
            motivo = mz

        puntos += 60 if zona in zonas else 10
        puntos += 10 if datos["tipo"] == "RACK" else 0

        uso = ocupacion.get(codigo)
        libre = not uso or not (uso.get("lotes") or 0)
        if libre:
            puntos += 80
            motivo = motivo or "libre"
        else:
            mismo = bool(material) and normalizar_texto(uso.get("material")) == normalizar_texto(material)
            puntos += 20 if mismo else -60
            if (uso.get("lotes") or 0) >= 2:
                puntos -= 20

        aviso = None
        if alergeno:
            if datos["permite_alergenos"]:
                puntos += 40
            else:
                puntos -= 100
                aviso = "esta zona no admite alérgenos"

        filas.append({"codigo": codigo, "zona": zona, "tipo": datos["tipo"], "libre": libre,
                      "lotes": (uso or {}).get("lotes", 0),
                      "cantidad": (uso or {}).get("cantidad"),
                      "unidad": (uso or {}).get("material") and (uso or {}).get("unidad"),
                      "aviso": aviso, "motivo": motivo, "puntaje": puntos})

    filas.sort(key=lambda f: (-f["puntaje"], not f["libre"], f["codigo"]))
    return filas[:max(1, int(limite))]
