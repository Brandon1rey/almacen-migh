"""Catálogo de insumos del ERP (MP y ME): clave, descripción, tipo, unidad y fabricante.

Es la **fuente única de los nombres de insumo**: se importa una vez desde el archivo del
sistema (`.xls` o `.xlsx`) y queda en `data\\catalogo_insumos.json`, para que la app no
dependa de Excel ni de COM.

    python scripts\\importar_catalogo.py --origen "C:\\ruta\\insumo_CATALOGO.xls"

Regla: la app NUNCA inventa un insumo. Si un renglón trae clave del ERP, la descripción,
unidad, categoría y fabricante salen de aquí (no de lo que se teclee), y una clave que no
existe se rechaza al dar de alta.
"""
import datetime as dt
import difflib
import json
import os
import pathlib
import re
import unicodedata

RAIZ = pathlib.Path(__file__).resolve().parents[1]
RUTA = pathlib.Path(os.environ.get("ALMACEN_CATALOGO")
                    or RAIZ / "data" / "catalogo_insumos.json")


def ruta(carpeta_datos=None):
    """Dónde vive el catálogo de ESTA instalación.

    Prioridad: la variable de entorno ALMACEN_CATALOGO (pruebas y casos raros), la carpeta
    de datos del servidor que está corriendo y, por último, data\\ del código. Importa
    porque una instalación puede arrancar con otra carpeta de datos (demo.cmd, una copia
    en la USB): el catálogo tiene que ser el de esa instalación, no el del repositorio.
    """
    desde_entorno = os.environ.get("ALMACEN_CATALOGO")
    if desde_entorno:
        return pathlib.Path(desde_entorno)
    if carpeta_datos:
        return pathlib.Path(carpeta_datos) / "catalogo_insumos.json"
    return RUTA

# Tipo del ERP -> inventario del sistema (ALM-PNO-002)
TIPO_A_INVENTARIO = {"empaque": "mep", "excipiente": "mp", "principio activo": "mp",
                     "materia prima": "mp", "producto terminado": "pt", "producto": "pt"}
INVENTARIOS = ("mp", "mep", "pt")
UNIDADES = {"kg": "kg", "kgs": "kg", "kilogramo": "kg", "l": "lt", "lt": "lt", "litro": "lt",
            "ml": "ml", "pieza": "pza", "piezas": "pza", "pza": "pza", "m": "m"}
COLUMNAS = {"clave de insumo": "sku", "clave": "sku", "sku": "sku", "codigo": "sku",
            "descripcion": "descripcion", "producto": "descripcion",
            "tipo": "tipo", "unidad de medida": "unidad", "unidad": "unidad", "um": "unidad",
            "codigo de barra": "barra", "maximo": "maximo", "minimo": "minimo",
            "fabricante": "fabricante", "proveedor": "fabricante"}


def normalizar(texto):
    s = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9]+", " ", s)).strip().upper()


def normalizar_unidad(valor):
    clave = str(valor or "").strip().lower()
    return UNIDADES.get(clave, clave or "")


def _numero(valor):
    try:
        return float(str(valor).replace(",", "")) if valor not in (None, "") else None
    except (TypeError, ValueError):
        return None


def _leer_xlsx(ruta):
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hoja = wb.worksheets[0]
    filas = list(hoja.iter_rows(values_only=True))
    wb.close()
    return filas


def _leer_xls(ruta):
    """Catálogo del ERP en .xls (formato viejo, BIFF): se abre con Excel y se pasa a .xlsx.

    Primero por pywin32 (rápido y sin ventanas); si no está instalado, por el COM de Windows
    desde PowerShell. La PC del almacén no tiene internet ni pip, y el catálogo del ERP llega
    justo en .xls: sin este respaldo habría que pedirle a alguien que lo convierta a mano.
    """
    ruta = pathlib.Path(ruta).resolve()
    try:
        return _leer_xls_com(ruta)
    except RuntimeError:
        return _leer_xls_powershell(ruta)


def _leer_xls_com(ruta):
    try:
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError("pywin32 no está instalado") from e
    excel = win32.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(ruta), ReadOnly=True)
        hoja = wb.Sheets(1)
        filas = [list(fila) for fila in hoja.UsedRange.Value]
        wb.Close(False)
        return filas
    except Exception as e:                      # noqa: BLE001 - COM falla de mil formas
        raise RuntimeError(f"Excel no pudo abrir el archivo: {e}") from e
    finally:
        excel.Quit()


def _leer_xls_powershell(ruta):
    """Convierte el .xls a .xlsx con el Excel de la PC (sin instalar nada)."""
    if os.name != "nt":
        raise RuntimeError("para leer .xls se necesita Excel instalado; guarda el archivo "
                           "como .xlsx y vuelve a intentar")
    import subprocess
    import tempfile
    carpeta = pathlib.Path(tempfile.mkdtemp(prefix="catalogo-xls-"))
    destino = carpeta / "catalogo_convertido.xlsx"
    script = (
        "$ErrorActionPreference='Stop';"
        "$excel = New-Object -ComObject Excel.Application;"
        "$excel.Visible = $false; $excel.DisplayAlerts = $false;"
        "try {"
        "  $wb = $excel.Workbooks.Open($env:ARCHIVO_XLS, 0, $true);"
        "  $wb.SaveAs($env:ARCHIVO_XLSX, 51);"
        "  $wb.Close($false);"
        "} finally {"
        "  $excel.Quit();"
        "  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)"
        "}"
    )
    entorno = dict(os.environ, ARCHIVO_XLS=str(ruta), ARCHIVO_XLSX=str(destino))
    try:
        proceso = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, text=True, timeout=300, env=entorno)
        if proceso.returncode != 0 or not destino.exists():
            detalle = (proceso.stderr or proceso.stdout or "").strip()[-400:]
            raise RuntimeError(
                "no se pudo leer el .xls: hace falta Excel instalado (o guarda el archivo "
                f"como .xlsx). Detalle: {detalle}")
        return _leer_xlsx(destino)
    finally:
        try:
            destino.unlink(missing_ok=True)
            carpeta.rmdir()
        except OSError:
            pass


def leer_filas(ruta):
    ruta = pathlib.Path(ruta)
    if ruta.suffix.lower() == ".xls":
        return _leer_xls(ruta)
    return _leer_xlsx(ruta)


def _mapear(filas):
    """De las filas crudas del ERP a registros del catálogo, por nombre de columna."""
    if not filas:
        return []
    encabezado = [normalizar(c) for c in filas[0]]
    indices = {}
    for i, texto in enumerate(encabezado):
        campo = COLUMNAS.get(texto.lower())
        if campo and campo not in indices:
            indices[campo] = i
    if "sku" not in indices or "descripcion" not in indices:
        raise ValueError("el archivo no trae 'Clave de Insumo' y 'Descripción' en la primera fila")

    insumos, vistos = [], set()
    for fila in filas[1:]:
        def valor(campo):
            i = indices.get(campo)
            if i is None or i >= len(fila):
                return None
            return fila[i]

        sku = str(valor("sku") or "").strip()
        descripcion = str(valor("descripcion") or "").strip()
        if not sku or not descripcion or sku.upper() in vistos:
            continue
        vistos.add(sku.upper())
        tipo = str(valor("tipo") or "").strip()
        insumos.append({
            "sku": sku,
            "descripcion": descripcion,
            "tipo": tipo,
            "unidad": normalizar_unidad(valor("unidad")),
            "fabricante": (str(valor("fabricante") or "").strip() or None),
            "barra": (str(valor("barra") or "").strip() or None),
            "maximo": _numero(valor("maximo")),
            "minimo": _numero(valor("minimo")),
            "inventario": TIPO_A_INVENTARIO.get(tipo.strip().lower()),
        })
    return insumos


def importar(origen, destino=None):
    """Importa el catálogo y lo deja en JSON (escritura atómica: nunca a medias)."""
    destino = pathlib.Path(destino or RUTA)
    insumos = _mapear(leer_filas(origen))
    if not insumos:
        raise ValueError(f"el catálogo {origen} no tiene renglones con clave y descripción")
    por_tipo = {}
    for i in insumos:
        por_tipo[i["tipo"] or "(sin tipo)"] = por_tipo.get(i["tipo"] or "(sin tipo)", 0) + 1
    datos = {"generado": dt.datetime.now().isoformat(sep=" ", timespec="seconds"),
             "fuente": str(pathlib.Path(origen).resolve()),
             "total": len(insumos), "por_tipo": por_tipo,
             "por_inventario": {t: sum(1 for i in insumos if i["inventario"] == t)
                                for t in INVENTARIOS},
             "insumos": insumos}
    destino.parent.mkdir(parents=True, exist_ok=True)
    temporal = destino.with_suffix(".tmp.json")
    try:
        temporal.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
        temporal.replace(destino)
    except Exception:
        try:
            temporal.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    _CACHE.update({"ruta": None, "mtime": None, "datos": None})
    return datos


_CACHE = {"ruta": None, "mtime": None, "datos": None}


def _vacio(existe=False, error=None):
    datos = {"generado": None, "fuente": None, "total": 0, "por_tipo": {},
             "por_inventario": {}, "insumos": [], "existe": existe}
    if error:
        datos["error"] = error
    return datos


def cargar(ruta=None, con=None):
    """Catálogo en memoria; lee desde PostgreSQL (Supabase) si hay conexión, o desde JSON."""
    if con is not None and getattr(con, "es_postgres", False):
        import time
        ahora = time.monotonic()
        if _CACHE.get("ruta") == "postgres" and (ahora - (_CACHE.get("mtime") or 0)) < 60:
            return _CACHE["datos"]
        try:
            filas = con.execute(
                "SELECT sku, descripcion, tipo, unidad, fabricante, barra, maximo, minimo, "
                "inventario FROM catalogo_insumos ORDER BY sku").fetchall()
            if filas:
                insumos = [dict(f) for f in filas]
                por_tipo = {}
                for i in insumos:
                    clave_t = i.get("tipo") or "(sin tipo)"
                    por_tipo[clave_t] = por_tipo.get(clave_t, 0) + 1
                datos = {
                    "generado": "supabase",
                    "fuente": "supabase:catalogo_insumos",
                    "total": len(insumos),
                    "por_tipo": por_tipo,
                    "por_inventario": {t: sum(1 for i in insumos if i.get("inventario") == t)
                                       for t in INVENTARIOS},
                    "insumos": insumos,
                    "existe": True,
                }
                _CACHE.update({"ruta": "postgres", "mtime": ahora, "datos": datos})
                return datos
        except Exception:  # noqa: BLE001 - cae al respaldo JSON si la tabla aún no está poblada
            pass
    ruta = pathlib.Path(ruta or RUTA)
    if not ruta.exists():
        return _vacio()
    try:
        mtime = ruta.stat().st_mtime
    except OSError as e:
        return _vacio(error=f"no se pudo leer el catálogo: {e}")
    if _CACHE["ruta"] == str(ruta) and _CACHE["mtime"] == mtime:
        return _CACHE["datos"]
    try:
        datos = json.loads(ruta.read_text(encoding="utf-8"))
        if not isinstance(datos, dict) or not isinstance(datos.get("insumos"), list):
            raise ValueError("el archivo no tiene la forma de un catálogo")
    except (OSError, ValueError, UnicodeDecodeError) as e:
        datos = _vacio(error=f"el catálogo no se pudo leer ({e}); vuelve a importarlo")
    datos["existe"] = True
    _CACHE.update({"ruta": str(ruta), "mtime": mtime, "datos": datos})
    return datos


def por_sku(datos, sku):
    buscado = normalizar(sku).replace(" ", "")
    if not buscado:
        return None
    for insumo in datos.get("insumos", []):
        if normalizar(insumo.get("sku")).replace(" ", "") == buscado:
            return insumo
    return None


def buscar(datos, texto="", limite=30, tipo_inventario=None):
    """Busca por clave, descripción o fabricante; si no hay texto, devuelve los primeros."""
    termino = normalizar(texto)
    salida = []
    for insumo in datos.get("insumos", []):
        if tipo_inventario and insumo.get("inventario") != tipo_inventario:
            continue
        if termino:
            pajar = normalizar(" ".join(str(insumo.get(c) or "") for c in
                                        ("sku", "descripcion", "fabricante", "tipo")))
            if not all(palabra in pajar for palabra in termino.split()):
                continue
        salida.append(insumo)
        if len(salida) >= limite:
            break
    return salida


def sugerencias(datos, texto, limite=5):
    """Insumos con nombre/clave parecidos (solo del catálogo; nunca inventa uno)."""
    termino = normalizar(texto)
    if not termino:
        return []
    candidatos = {}
    for insumo in datos.get("insumos", []):
        candidatos.setdefault(normalizar(insumo.get("descripcion")), insumo)
        candidatos.setdefault(normalizar(insumo.get("sku")), insumo)
    parecidos = difflib.get_close_matches(termino, list(candidatos), n=limite, cutoff=0.55)
    return [candidatos[p] for p in parecidos]


def para_inventario(datos, tipo):
    return [i for i in datos.get("insumos", []) if i.get("inventario") == tipo]


def resumen(datos):
    return {"existe": datos.get("existe", False), "total": datos.get("total", 0),
            "generado": datos.get("generado"), "fuente": datos.get("fuente"),
            "por_tipo": datos.get("por_tipo", {}),
            "por_inventario": datos.get("por_inventario", {}),
            "error": datos.get("error")}
