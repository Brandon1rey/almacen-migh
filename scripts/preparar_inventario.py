"""Prepara (o actualiza) el Excel del área agregándole la columna UBICACIÓN.

Uso:
  python scripts/preparar_inventario.py                          # materias primas
  python scripts/preparar_inventario.py --tipo mep               # material de empaque
  python scripts/preparar_inventario.py --tipo pt --origen "C:/.../INVENTARIO PT.xlsx"
  python scripts/preparar_inventario.py --tipo mep --reimportar  # toma el archivo nuevo del área
                                                                 # y conserva las ubicaciones ya capturadas

Sin --origen, si el destino no existe, crea la plantilla vacía con los encabezados correctos.
"""
import argparse
import datetime as dt
import pathlib
import shutil
import sys

RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from app import inventario, inventarios  # noqa: E402
from scripts import plantilla_inventario  # noqa: E402

ORIGEN_DEFECTO = {  # de dónde suele venir cada archivo del área
    "mp": pathlib.Path(r"C:\Users\asegm\Desktop\Copia de 06 OCT 26 INVENTARIO MP.xlsx"),
}
RESPALDO = RAIZ / "data" / "respaldos"


def preparar(origen, destino, reimportar=False, tipo="mp"):
    tipo = inventarios.normalizar_tipo(tipo)
    destino = pathlib.Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    mapa_ubicaciones = {}

    if destino.exists() and reimportar:
        anterior = inventario.Inventario(destino, tipo=tipo)
        mapa_ubicaciones = {f["clave"]: f["ubicacion"] for f in anterior.filas if f["ubicacion"]}
        RESPALDO.mkdir(parents=True, exist_ok=True)
        copia = RESPALDO / (f"Inventario_{tipo.upper()}_antes_"
                            f"{dt.datetime.now():%Y%m%d-%H%M}.xlsx")
        shutil.copy2(destino, copia)
        print(f"Respaldo del inventario anterior: {copia}")
    elif destino.exists() and origen is None:
        actual = inventario.Inventario(destino, tipo=tipo)
        print(f"El inventario {tipo.upper()} ya existe ({len(actual.filas)} filas). "
              f"Usa --reimportar para volver a tomarlo del archivo del área.")
        return destino, actual.resumen(), 0

    if origen is not None:
        origen = pathlib.Path(origen)
        if not origen.exists():
            raise FileNotFoundError(f"no existe el archivo de origen: {origen}")
        shutil.copy2(origen, destino)
    elif not destino.exists():
        plantilla_inventario.crear(tipo, destino)
        print(f"No había archivo: se creó la plantilla vacía de {tipo.upper()} en {destino}")

    nuevo = inventario.Inventario(destino, tipo=tipo)      # agrega la columna si falta
    recuperadas = 0
    if mapa_ubicaciones:
        for fila in nuevo.filas:
            ubicacion_previa = mapa_ubicaciones.get(fila["clave"])
            if ubicacion_previa:
                nuevo.asignar(fila["fila"], ubicacion_previa)
                recuperadas += 1
    nuevo.guardar()
    return destino, nuevo.resumen(), recuperadas


def main(argv=None):
    ap = argparse.ArgumentParser(description="Prepara un inventario del almacén")
    ap.add_argument("--tipo", default="mp", choices=list(inventarios.TIPOS))
    ap.add_argument("--origen", default=None,
                    help="archivo del área; si se omite se usa el de MP o la plantilla")
    ap.add_argument("--destino", default=None)
    ap.add_argument("--reimportar", action="store_true")
    args = ap.parse_args(argv)

    tipo = inventarios.normalizar_tipo(args.tipo)
    origen = args.origen
    destino = pathlib.Path(args.destino) if args.destino else inventarios.ruta(tipo)
    if origen is None and tipo in ORIGEN_DEFECTO and not destino.exists():
        origen = ORIGEN_DEFECTO[tipo]

    destino, resumen, recuperadas = preparar(origen, destino, args.reimportar, tipo)
    print(f"Inventario: {tipo.upper()} — {resumen['nombre']}")
    print(f"Archivo   : {destino}")
    print(f"Lotes     : {resumen['total']}")
    print(f"Con ubic. : {resumen['con_ubicacion']} ({resumen['cobertura']} %)")
    print(f"Sin ubicar: {resumen['sin_ubicacion']}")
    print(f"Vencidos  : {resumen['semaforo']['VENCIDO']}")
    print(f"Vencen 30d: {resumen['semaforo']['CRITICO']}")
    print(f"Sin lote  : {resumen['sin_lote']}")
    print(f"Sin caduc.: {resumen['sin_caducidad']}")
    for unidad, cantidad in sorted(resumen["por_unidad"].items()):
        print(f"Cantidad  : {cantidad:,.1f} {unidad}")
    if recuperadas:
        print(f"Ubicaciones conservadas al reimportar: {recuperadas}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
