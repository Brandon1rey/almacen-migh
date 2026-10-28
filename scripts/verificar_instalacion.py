"""Verifica que la instalación en esta PC esté completa. Sale con código 0 si todo pasa."""
import json
import os
import pathlib
import socket
import sqlite3
import sys

# La consola de Windows no siempre es UTF-8: sin esto, imprimir «→» o «·» lanza
# UnicodeEncodeError y el script parece fallar aunque ya haya hecho su trabajo.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass


RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
_VENDOR = RAIZ / "vendor"
if _VENDOR.exists():
    sys.path.insert(0, str(_VENDOR))

PUERTO = int(os.environ.get("ALMACEN_PUERTO", 8300))
resultados = []


def revisar(nombre, funcion):
    try:
        detalle = funcion()
        resultados.append((True, nombre, detalle or ""))
    except Exception as e:  # noqa: BLE001 - cualquier falla es un hallazgo
        resultados.append((False, nombre, f"{type(e).__name__}: {e}"))


def python_ok():
    if sys.version_info < (3, 10):
        raise RuntimeError(f"se requiere Python 3.10 o superior; esta PC tiene {sys.version.split()[0]}")
    return sys.version.split()[0]


def openpyxl_ok():
    import openpyxl
    return f"openpyxl {openpyxl.__version__}"


def estructura_ok():
    faltan = [n for n in ("app/server.py", "app/static/index.html", "app/inventario.py")
              if not (RAIZ / n).exists()]
    if faltan:
        raise FileNotFoundError(f"faltan archivos: {', '.join(faltan)}")
    return "app completa"


def inventario_ok():
    ruta = RAIZ / "data" / "Inventario_MP.xlsx"
    if not ruta.exists():
        raise FileNotFoundError("no existe data/Inventario_MP.xlsx (corre preparar_inventario.py)")
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True)
    hoja = wb.worksheets[0]
    filas = sum(1 for fila in hoja.iter_rows(min_row=4, max_col=1, values_only=True) if fila[0])
    wb.close()
    if filas == 0:
        raise ValueError("el inventario no tiene filas de material")
    return f"{filas} lotes · {ruta.stat().st_size // 1024} KB"


def escritura_ok():
    prueba = RAIZ / "data" / ".prueba_escritura"
    prueba.write_text("ok", encoding="utf-8")
    prueba.unlink()
    return "data/ escribible"


def bitacora_ok():
    from app import bitacora
    con = bitacora.conectar()
    n = bitacora.contar(con)
    columnas = {fila["name"] for fila in con.execute("PRAGMA table_info(movimiento)")}
    con.close()
    if not {"motivo", "ref_doc"} <= columnas:
        raise RuntimeError("la bitácora no tiene las columnas motivo/ref_doc")
    return f"{n} movimientos"


def puerto_ok():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        ocupado = s.connect_ex(("127.0.0.1", PUERTO)) == 0
    if ocupado:
        return f"puerto {PUERTO} ya en uso (¿el sistema ya está corriendo?)"
    return f"puerto {PUERTO} libre"


def main():
    print("VERIFICACIÓN DE INSTALACIÓN — CONTROL DE UBICACIÓN DE ALMACÉN")
    print("=" * 64)
    revisar("Python", python_ok)
    revisar("openpyxl", openpyxl_ok)
    revisar("Estructura de la aplicación", estructura_ok)
    revisar("Inventario (Excel del área)", inventario_ok)
    revisar("Permisos de escritura", escritura_ok)
    revisar("Bitácora SQLite", bitacora_ok)
    revisar("Puerto del servidor", puerto_ok)

    for ok, nombre, detalle in resultados:
        print(f"  [{'OK ' if ok else 'FALLA'}] {nombre:<32} {detalle}")
    fallas = [r for r in resultados if not r[0]]
    print("=" * 64)
    if fallas:
        print(f"{len(fallas)} problema(s). Revisa el detalle de arriba.")
        return 1
    print("Todo listo. Arranca con arrancar.cmd y abre http://127.0.0.1:%d" % PUERTO)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
