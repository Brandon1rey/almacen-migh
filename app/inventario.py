"""El inventario vive en el Excel que ya usa el área; aquí solo se le agrega la columna UBICACIÓN."""
import datetime as dt
import ast
import operator
import pathlib
import re
import unicodedata
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter

from app import fechas, inventarios, ubicaciones

RAIZ = pathlib.Path(__file__).resolve().parents[1]
RUTA = RAIZ / "data" / "Inventario_MP.xlsx"
FILA_ENCABEZADO = 3
FILA_DATOS = 4
ENCABEZADO_UBICACION = "UBICACIÓN"

# encabezado del Excel del área  ->  nombre interno
MAPA_COLUMNAS = {
    "descripcion": "descripcion",
    "producto": "descripcion",
    "material": "descripcion",
    "nombre del material": "descripcion",
    "presentacion": "presentacion",
    "piezas por caja": "piezas",
    "piezas": "piezas",
    "sku": "erp",
    "clave erp": "erp",
    "categoria": "categoria",
    "proveedor": "proveedor",
    "unidad de medida": "unidad",
    "unidad": "unidad",
    "conteo": "cantidad",
    "cantidad": "cantidad",
    "stock": "cantidad",
    "existencia": "cantidad",
    "um": "unidad",
    "caducidad vigente": "caducidad",
    "caducidad": "caducidad",
    "lote actual": "lote",
    "lote": "lote",
    "entrada erp": "erp",
    "observaciones": "observaciones",
    "ubicacion": "ubicacion",
}


UNIDADES = {"pz": "pza", "pzas": "pza", "pieza": "pza", "piezas": "pza", "pza": "pza",
            "kg": "kg", "kgs": "kg", "kilogramo": "kg", "kilogramos": "kg",
            "lt": "lt", "l": "lt", "litro": "lt", "litros": "lt",
            "m": "m", "ml": "ml", "g": "g", "caja": "caja", "millar": "millar"}


def normalizar_unidad(valor, defecto="kg"):
    """'PZ', 'Piezas', 'KG ' -> unidad canónica, para no sumar peras con manzanas."""
    clave = str(valor or "").strip().lower()
    return UNIDADES.get(clave, clave or defecto)


def normalizar(texto):
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9]+", " ", s)).strip().upper()


_OPS = {ast.Add: operator.add, ast.Sub: operator.sub,
        ast.Mult: operator.mul, ast.Div: operator.truediv}


def evaluar_formula(texto):
    """Evalúa una suma/resta aritmética simple. Excel la muestra, pero el archivo
    puede no traer el valor calculado (openpyxl lo guarda sin caché)."""
    nodo = ast.parse(str(texto).lstrip("="), mode="eval").body

    def ev(n):
        if isinstance(n, ast.BinOp) and type(n.op) in _OPS:
            return _OPS[type(n.op)](ev(n.left), ev(n.right))
        if isinstance(n, ast.UnaryOp) and isinstance(n.op, ast.USub):
            return -ev(n.operand)
        if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)):
            return n.value
        raise ValueError("fórmula no soportada")

    return round(float(ev(nodo)), 3)


class Inventario:
    """Carga el libro del área, expone sus filas y escribe la columna UBICACIÓN."""

    def __init__(self, ruta=None, tipo="mp"):
        self.tipo = inventarios.normalizar_tipo(tipo)
        self.ruta = pathlib.Path(ruta or inventarios.ruta(self.tipo))
        self.fila_encabezado = FILA_ENCABEZADO
        self.nombre = inventarios.TIPOS[self.tipo]["nombre"]
        if not self.ruta.exists():
            raise FileNotFoundError(f"no existe el inventario: {self.ruta}")
        self.wb = openpyxl.load_workbook(self.ruta)
        self._fijar_imagenes()
        self.ws = self._hoja_inventario()
        self.columnas = self._detectar_columnas()
        self.filas = self._leer()

    # ---------- estructura ----------
    def _fijar_imagenes(self):
        """openpyxl reutiliza el buffer de las imágenes del libro y la segunda escritura
        falla con 'I/O operation on closed file'. Se congelan los bytes al cargar."""
        for hoja in self.wb.worksheets:
            for imagen in getattr(hoja, "_images", []):
                datos = imagen._data()

                def congelar(d):
                    return lambda: d

                imagen._data = congelar(datos)

    def _hoja_inventario(self):
        """La hoja del inventario es la que trae el encabezado del material, no siempre la primera.

        Acepta 'Descripción', 'Producto' o 'Material' y busca en las primeras filas, porque las
        tablas de empaque y producto terminado no siempre repiten el formato del archivo de MP.
        """
        validos = {"DESCRIPCION", "PRODUCTO", "MATERIAL", "NOMBRE DEL MATERIAL"}
        for hoja in self.wb.worksheets:
            for fila in range(1, min(hoja.max_row, 8) + 1):
                for col in range(1, min(hoja.max_column, 15) + 1):
                    if normalizar(hoja.cell(row=fila, column=col).value) in validos:
                        self.fila_encabezado = fila
                        return hoja
        return self.wb.worksheets[0]

    def _detectar_columnas(self):
        columnas = {}
        for col in range(1, self.ws.max_column + 1):
            valor = self.ws.cell(row=self.fila_encabezado, column=col).value
            if not valor:
                continue
            clave = MAPA_COLUMNAS.get(normalizar(valor).lower())
            if clave and clave not in columnas:
                columnas[clave] = col
        if "descripcion" not in columnas:
            raise ValueError(f"no se encontró la columna de descripción en la fila {self.fila_encabezado}")
        if "ubicacion" not in columnas:
            col = max(columnas.values()) + 1
            self._crear_columna_ubicacion(col)
            columnas["ubicacion"] = col
        return columnas

    def _crear_columna_ubicacion(self, col):
        origen = self.ws.cell(row=self.fila_encabezado, column=col - 1)
        celda = self.ws.cell(row=self.fila_encabezado, column=col, value=ENCABEZADO_UBICACION)
        if origen.has_style:
            celda._style = copy(origen._style)
        celda.value = ENCABEZADO_UBICACION
        self.ws.column_dimensions[get_column_letter(col)].width = 16
        for fila in range(self.fila_encabezado + 1, self.ws.max_row + 1):
            modelo = self.ws.cell(row=fila, column=col - 1)
            nueva = self.ws.cell(row=fila, column=col)
            if modelo.has_style:
                nueva._style = copy(modelo._style)

    def _leer(self):
        filas, ocurrencias = [], {}
        for fila in range(self.fila_encabezado + 1, self.ws.max_row + 1):
            descripcion = self.ws.cell(row=fila, column=self.columnas["descripcion"]).value
            if descripcion in (None, ""):
                continue
            crudo = self._valor(fila, "caducidad")
            iso, motivo = fechas.parsear(crudo)
            lote = self._valor(fila, "lote")
            registro = {
                "fila": fila,
                "descripcion": str(descripcion).strip(),
                "categoria": self._valor(fila, "categoria"),
                "proveedor": self._valor(fila, "proveedor"),
                "unidad": normalizar_unidad(self._valor(fila, "unidad"),
                                            inventarios.TIPOS[self.tipo]["unidad"]),
                "cantidad": self._valor(fila, "cantidad"),
                "caducidad": iso,
                "caducidad_cruda": (crudo.date().isoformat() if isinstance(crudo, dt.datetime)
                                    else str(crudo) if crudo not in (None, "") else None),
                "motivo_caducidad": motivo,
                "lote": str(lote).strip() if lote not in (None, "") else None,
                "erp": self._valor(fila, "erp"),
                "observaciones": self._valor(fila, "observaciones"),
                "ubicacion": self._ubicacion_actual(fila),
                "cantidad_formula": self._es_formula(fila, "cantidad"),
                "presentacion": self._valor(fila, "presentacion"),
                "piezas": self._valor(fila, "piezas"),
                "tipo": self.tipo,
            }
            dias, sem = fechas.semaforo(iso)
            registro["dias_restantes"] = dias
            registro["semaforo"] = sem
            clave = (normalizar(registro["descripcion"]), normalizar(registro["lote"]))
            ocurrencias[clave] = ocurrencias.get(clave, 0) + 1
            registro["clave"] = f"{clave[0]}|{clave[1]}|{ocurrencias[clave]}"
            filas.append(registro)
        return filas

    def _valor(self, fila, campo):
        col = self.columnas.get(campo)
        if not col:
            return None
        valor = self.ws.cell(row=fila, column=col).value
        if isinstance(valor, str):
            valor = valor.strip()
            if valor.startswith("=") and campo == "cantidad":
                try:
                    return evaluar_formula(valor)
                except (ValueError, SyntaxError, ZeroDivisionError, RecursionError):
                    return None
        return valor if valor != "" else None

    def _es_formula(self, fila, campo):
        col = self.columnas.get(campo)
        if not col:
            return False
        valor = self.ws.cell(row=fila, column=col).value
        return isinstance(valor, str) and valor.startswith("=")

    def _ubicacion_actual(self, fila):
        col = self.columnas["ubicacion"]
        valor = self.ws.cell(row=fila, column=col).value
        if valor in (None, ""):
            return None
        return ubicaciones.normalizar(valor)

    # ---------- consultas ----------
    def por_fila(self, fila):
        for registro in self.filas:
            if registro["fila"] == fila:
                return registro
        raise KeyError(f"fila inexistente: {fila}")

    def buscar(self, texto="", solo_sin_ubicar=False, solo_vencidos=False, incluir_cero=False):
        termino = normalizar(texto)
        salida = []
        for f in self.filas:
            if not incluir_cero and not self._tiene_existencia(f):
                continue
            if solo_sin_ubicar and f["ubicacion"]:
                continue
            if solo_vencidos and f["semaforo"] != "VENCIDO":
                continue
            if termino:
                pajar = normalizar(" ".join(str(f[c] or "") for c in
                                            ("descripcion", "lote", "proveedor", "categoria", "erp")))
                if termino not in pajar:
                    continue
            salida.append(f)
        return salida

    @staticmethod
    def _tiene_existencia(registro):
        cantidad = registro.get("cantidad")
        return isinstance(cantidad, (int, float)) and cantidad > 0

    def resumen(self):
        activas = [f for f in self.filas if self._tiene_existencia(f)]
        total = len(activas)
        con = [f for f in activas if f["ubicacion"]]
        semaforo = {"VENCIDO": 0, "CRITICO": 0, "AVISO": 0, "OK": 0, "SIN DATO": 0}
        for f in activas:
            semaforo[f["semaforo"]] += 1
        por_unidad = {}
        for f in activas:
            if isinstance(f["cantidad"], (int, float)):
                unidad = (f["unidad"] or "kg").strip().lower() or "kg"
                por_unidad[unidad] = round(por_unidad.get(unidad, 0) + f["cantidad"], 3)
        return {
            "tipo": self.tipo,
            "nombre": inventarios.TIPOS[self.tipo]["nombre"],
            "total": total,
            "registros": len(self.filas),
            "anulados": len(self.filas) - total,
            "con_ubicacion": len(con),
            "sin_ubicacion": total - len(con),
            "cobertura": round(100 * len(con) / total, 1) if total else 0.0,
            "kg": round(por_unidad.get("kg", 0.0), 1),
            "por_unidad": {k: round(v, 3) for k, v in por_unidad.items()},
            "semaforo": semaforo,
            "sin_lote": sum(1 for f in activas if not f["lote"]),
            "sin_caducidad": sum(1 for f in activas if not f["caducidad"]),
            "con_formula": sum(1 for f in self.filas if f["cantidad_formula"]),
            "por_zona": self._por_zona(),
            "actualizado": dt.datetime.now().isoformat(sep=" ", timespec="seconds"),
        }

    def _por_zona(self):
        conteo = {}
        for f in self.filas:
            if not f["ubicacion"] or not self._tiene_existencia(f):
                continue
            zona = ubicaciones.zona_de(f["ubicacion"])
            conteo[zona] = conteo.get(zona, 0) + 1
        return conteo

    # ---------- escritura ----------
    def asignar(self, fila, codigo):
        """Escribe la ubicación en el libro. Devuelve (anterior, nueva)."""
        registro = self.por_fila(fila)
        if codigo in (None, "", "-"):
            nueva = None
        else:
            nueva = ubicaciones.normalizar(codigo)
            if not ubicaciones.existe(nueva):
                raise ValueError(f"ubicación inexistente en el catálogo: {codigo}")
        anterior = registro["ubicacion"]
        celda = self.ws.cell(row=fila, column=self.columnas["ubicacion"], value=nueva)
        if self.ws.cell(row=fila, column=self.columnas["ubicacion"] - 1).has_style:
            celda._style = copy(
                self.ws.cell(row=fila, column=self.columnas["ubicacion"] - 1)._style)
        registro["ubicacion"] = nueva
        return anterior, nueva

    def guardar(self, destino=None):
        destino = pathlib.Path(destino or self.ruta)
        destino.parent.mkdir(parents=True, exist_ok=True)
        temporal = destino.with_suffix(".tmp.xlsx")
        self.wb.save(temporal)
        temporal.replace(destino)
        return destino

    # ---------- cantidades y parcialidades ----------
    def ajustar_cantidad(self, fila, nueva):
        col = self.columnas.get("cantidad")
        if not col:
            raise ValueError("el libro no tiene columna de cantidad")
        modelo = self.ws.cell(row=fila, column=col)
        celda = self.ws.cell(row=fila, column=col, value=float(nueva))
        if modelo.has_style:
            celda._style = copy(modelo._style)
        registro = self.por_fila(fila)
        registro["cantidad"] = float(nueva)
        return registro

    def agregar_parcialidad(self, fila_origen, cantidad, destino, nota=None):
        """Crea un renglón nuevo con la parte que se mueve y deja el resto en el original."""
        if cantidad is None or cantidad <= 0:
            raise ValueError("la cantidad a mover debe ser mayor que cero")
        origen = self.por_fila(fila_origen)
        disponible = float(origen["cantidad"] or 0)
        if cantidad >= disponible:
            raise ValueError(f"la cantidad ({cantidad}) debe ser menor que la existencia "
                             f"({disponible}); para mover todo usa el movimiento completo")
        if not ubicaciones.existe(destino):
            raise ValueError(f"ubicación inexistente en el catálogo: {destino}")

        fila = self.ws.max_row + 1
        col_desc = self.columnas["descripcion"]
        while fila > self.fila_encabezado + 1 and not self.ws.cell(row=fila - 1, column=col_desc).value:
            fila -= 1

        for campo, col in self.columnas.items():
            valor = self.ws.cell(row=fila_origen, column=col).value
            if campo == "cantidad":
                valor = round(float(cantidad), 3)
            elif campo == "ubicacion":
                valor = destino
            elif campo in ("presentacion", "piezas"):
                valor = self.ws.cell(row=fila_origen, column=col).value
            elif campo == "observaciones":
                previo = str(valor).strip() if valor not in (None, "") else ""
                marca = nota or f"parcialidad de la fila {fila_origen}"
                valor = f"{previo} · {marca}" if previo else marca
            celda = self.ws.cell(row=fila, column=col, value=valor)
            modelo = self.ws.cell(row=fila_origen, column=col)
            if modelo.has_style:
                celda._style = copy(modelo._style)

        self.ajustar_cantidad(fila_origen, round(disponible - float(cantidad), 3))
        self.filas = self._leer()
        return self.por_fila(fila)

    def anular(self, fila, nota):
        """Deja el renglón en cero y sin ubicación, con la marca en observaciones."""
        self.ajustar_cantidad(fila, 0)
        self.asignar(fila, None)
        col = self.columnas.get("observaciones")
        if col:
            previo = self.ws.cell(row=fila, column=col).value
            texto = f"{str(previo).strip() + ' · ' if previo else ''}{nota}"
            self.ws.cell(row=fila, column=col, value=texto)
        return self.por_fila(fila)

    def copia(self, destino):
        destino = pathlib.Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(destino)
        return destino
