"""Los tres inventarios del almacén y sus reglas, según ALM-PNO-002.

§V.2.1.8: reporte de inventario F01-ALM-PNO-002 **mensual** de todos los materiales almacenados.
§V.2.3.8: reporte de inventario F01-ALM-PNO-002 **semanal** para producto terminado.
§V.2.1.4: zonas del almacén (MPA/MPNA, MPT, MEP, jaulas y pasillo A).
"""
import pathlib

RAIZ = pathlib.Path(__file__).resolve().parents[1]
CARPETA_DATOS = RAIZ / "data"

TIPOS = {
    "mp": {
        "clave": "mp",
        "nombre": "Materias primas",
        "corto": "MP",
        "archivo": "Inventario_MP.xlsx",
        "unidad": "kg",
        "reporte": "mensual",
        "orden": 1,
        "base": "ALM-PNO-002 §V.2.1.8 (reporte mensual)",
        "zonas": ("MP-ALERGENO", "MP-NO-ALERGENO", "ESPECIAL"),
        "exige_caducidad": True,
        "unidades": ("kg", "lt", "m", "pza"),
        "unidad_preferida": "kg",
    },
    "mep": {
        "clave": "mep",
        "nombre": "Material de empaque",
        "corto": "MEP",
        "archivo": "Inventario_MEP.xlsx",
        "unidad": "pza",
        "reporte": "mensual",
        "orden": 2,
        "base": "ALM-PNO-002 §V.2.1.8 (reporte mensual)",
        "zonas": ("EM", "ESPECIAL"),
        "exige_caducidad": False,
        "unidades": ("pza", "kg", "m"),
        "unidad_preferida": "pza",
    },
    "pt": {
        "clave": "pt",
        "nombre": "Producto terminado",
        "corto": "PT",
        "archivo": "Inventario_PT.xlsx",
        "unidad": "pza",
        "reporte": "semanal",
        "orden": 3,
        "base": "ALM-PNO-002 §V.2.3.8 (reporte semanal)",
        "zonas": ("PT", "ESPECIAL"),
        "exige_caducidad": True,
        "unidades": ("pza", "kg"),
        "unidad_preferida": "pza",
    },
}

ALIAS = {
    "mp": "mp", "materia prima": "mp", "materias primas": "mp", "materias-primas": "mp",
    "mep": "mep", "empaque": "mep", "material de empaque": "mep", "em": "mep",
    "pt": "pt", "producto terminado": "pt", "terminado": "pt",
}


def normalizar_tipo(valor):
    """'MEP', 'empaque', 'Material de empaque' -> 'mep'. Falla claro si no existe."""
    clave = str(valor or "mp").strip().lower()
    tipo = ALIAS.get(clave)
    if tipo is None:
        raise ValueError(f"inventario desconocido: {valor!r} (usa mp, mep o pt)")
    return tipo


def ruta(tipo, carpeta=None):
    """Ruta del Excel de ese inventario (no exige que exista)."""
    return pathlib.Path(carpeta or CARPETA_DATOS) / TIPOS[normalizar_tipo(tipo)]["archivo"]


def ficha(tipo):
    """Copia de la configuración del tipo, con su ruta."""
    clave = normalizar_tipo(tipo)
    return {**TIPOS[clave], "ruta": str(ruta(clave))}


def lista(carpeta=None, en_db=False):
    """Los tres tipos en orden, con 'existe' para que la interfaz sepa cuáles hay."""
    salida = []
    for clave in sorted(TIPOS, key=lambda t: TIPOS[t]["orden"]):
        f = ficha(clave)
        f["ruta"] = str(ruta(clave, carpeta))
        f["existe"] = True if en_db else pathlib.Path(f["ruta"]).exists()
        salida.append(f)
    return salida


def zona_esperada(tipo, zona):
    """¿La zona corresponde a ese inventario? Solo para avisar, no para bloquear."""
    return zona in TIPOS[normalizar_tipo(tipo)]["zonas"]
