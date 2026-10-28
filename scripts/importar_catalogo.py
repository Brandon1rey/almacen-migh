"""Importa el catálogo de insumos del ERP (MP y ME) al formato que usa la app.

    python scripts\\importar_catalogo.py
    python scripts\\importar_catalogo.py --origen "C:\\Users\\asegm\\Downloads\\insumo_CATALOGO.xls"

Deja data\\catalogo_insumos.json con clave, descripción, tipo, unidad y fabricante de cada insumo.
El sistema usa ese archivo para sugerir y validar los ingresos de material; no lee Excel al operar.
"""
import argparse
import pathlib
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

from app import catalogo  # noqa: E402

ORIGEN_DEFECTO = pathlib.Path(r"C:\Users\asegm\Downloads\insumo_CATALOGO.xls")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Importa el catálogo de insumos")
    ap.add_argument("--origen", default=str(ORIGEN_DEFECTO))
    ap.add_argument("--destino", default=str(catalogo.RUTA))
    args = ap.parse_args(argv)

    origen = pathlib.Path(args.origen)
    if not origen.exists():
        print(f"ERROR: no existe {origen}")
        return 2
    try:
        datos = catalogo.importar(origen, args.destino)
    except (ValueError, RuntimeError) as e:
        print(f"ERROR: {e}")
        return 3

    print(f"Catálogo : {args.destino}")
    print(f"Fuente   : {datos['fuente']}")
    print(f"Insumos  : {datos['total']}")
    for tipo, cantidad in sorted(datos["por_tipo"].items(), key=lambda par: -par[1]):
        print(f"  {tipo:<20} {cantidad:>4}")
    print(f"  -> materias primas   {datos['por_inventario'].get('mp', 0):>4}")
    print(f"  -> material empaque  {datos['por_inventario'].get('mep', 0):>4}")
    if datos['por_inventario'].get('pt'):
        print(f"  -> producto terminado {datos['por_inventario']['pt']:>3}")
    print("\nYa se puede usar en la pantalla de Ingresar material (sugiere clave, descripción y unidad).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
