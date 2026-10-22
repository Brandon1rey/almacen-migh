"""Prepara (o actualiza) el Excel del área agregándole la columna UBICACIÓN.

Uso:
  python scripts/preparar_inventario.py
  python scripts/preparar_inventario.py --origen "C:/.../Copia de 06 OCT 26 INVENTARIO MP.xlsx"
  python scripts/preparar_inventario.py --reimportar      # toma un archivo nuevo del área
                                                          # y conserva las ubicaciones ya capturadas
"""
import argparse
import pathlib
import shutil
import sys

RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from app import inventario  # noqa: E402

ORIGEN_DEFECTO = pathlib.Path(r"C:\Users\asegm\Desktop\Copia de 06 OCT 26 INVENTARIO MP.xlsx")
DESTINO = RAIZ / "data" / "Inventario_MP.xlsx"
RESPALDO = RAIZ / "data" / "respaldos"


def preparar(origen, destino, reimportar=False):
    origen, destino = pathlib.Path(origen), pathlib.Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    mapa_ubicaciones = {}

    if destino.exists():
        anterior = inventario.Inventario(destino)
        mapa_ubicaciones = {f["clave"]: f["ubicacion"] for f in anterior.filas if f["ubicacion"]}
        if not reimportar:
            print(f"El inventario ya existe ({len(anterior.filas)} filas). "
                  f"Usa --reimportar para volver a tomarlo del archivo del área.")
            return destino, anterior.resumen(), 0
        RESPALDO.mkdir(parents=True, exist_ok=True)
        copia = RESPALDO / f"Inventario_MP_antes_{__import__('datetime').datetime.now():%Y%m%d-%H%M}.xlsx"
        shutil.copy2(destino, copia)
        print(f"Respaldo del inventario anterior: {copia}")

    shutil.copy2(origen, destino)
    nuevo = inventario.Inventario(destino)          # agrega la columna si falta
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
    ap = argparse.ArgumentParser()
    ap.add_argument("--origen", default=str(ORIGEN_DEFECTO))
    ap.add_argument("--destino", default=str(DESTINO))
    ap.add_argument("--reimportar", action="store_true")
    args = ap.parse_args(argv)

    destino, resumen, recuperadas = preparar(args.origen, args.destino, args.reimportar)
    print(f"Archivo   : {destino}")
    print(f"Lotes     : {resumen['total']}")
    print(f"Con ubic. : {resumen['con_ubicacion']} ({resumen['cobertura']} %)")
    print(f"Sin ubicar: {resumen['sin_ubicacion']}")
    print(f"Vencidos  : {resumen['semaforo']['VENCIDO']}")
    print(f"Vencen 30d: {resumen['semaforo']['CRITICO']}")
    print(f"Sin lote  : {resumen['sin_lote']}")
    print(f"Sin caduc.: {resumen['sin_caducidad']}")
    if recuperadas:
        print(f"Ubicaciones conservadas al reimportar: {recuperadas}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
