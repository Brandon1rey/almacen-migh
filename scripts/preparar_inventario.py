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

# La consola de Windows no siempre es UTF-8: sin esto, imprimir «→» o «·» lanza
# UnicodeEncodeError y el script parece fallar aunque ya haya hecho su trabajo.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass


RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from app import inventario, inventarios  # noqa: E402
from scripts import plantilla_inventario  # noqa: E402

ORIGEN_DEFECTO = {  # de dónde suele venir cada archivo del área
    "mp": pathlib.Path(r"C:\Users\asegm\Desktop\Copia de 06 OCT 26 INVENTARIO MP.xlsx"),
}


def preparar(origen, destino, reimportar=False, tipo="mp"):
    tipo = inventarios.normalizar_tipo(tipo)
    destino = pathlib.Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    mapa_ubicaciones = {}
    mapa_qr = {}
    claves_con_qr = set()
    repetidas = set()

    if destino.exists() and reimportar:
        anterior = inventario.Inventario(destino, tipo=tipo)
        mapa_ubicaciones = {f["clave"]: f["ubicacion"] for f in anterior.filas if f["ubicacion"]}
        # La etiqueta física no viene en el archivo del área: se conserva por clave, y solo
        # cuando la clave es inequívoca (una parcialidad comparte descripción+lote).
        todas_anterior = [f["clave"].rsplit("|", 1)[0] for f in anterior.filas]
        for f in anterior.filas:
            if f.get("qr_id"):
                base = f["clave"].rsplit("|", 1)[0]
                claves_con_qr.add(base)
                mapa_qr[base] = f["qr_id"]
        repetidas = {c for c in claves_con_qr if todas_anterior.count(c) > 1}
        # El respaldo se deja JUNTO AL ARCHIVO que se va a reemplazar: con --destino a otra
        # carpeta, el respaldo debe quedar con esos datos y no en el data\ del repositorio.
        respaldo = destino.parent / "respaldos"
        respaldo.mkdir(parents=True, exist_ok=True)
        copia = respaldo / (f"Inventario_{tipo.upper()}_antes_"
                            f"{dt.datetime.now():%Y%m%d-%H%M}.xlsx")
        shutil.copy2(destino, copia)
        print(f"Respaldo del inventario anterior: {copia}")
    elif destino.exists() and origen is None:
        actual = inventario.Inventario(destino, tipo=tipo)
        print(f"El inventario {tipo.upper()} ya existe ({len(actual.filas)} filas). "
              f"Usa --reimportar para volver a tomarlo del archivo del área.")
        return destino, actual.resumen(), 0

    descartadas = []
    if origen is not None:
        origen = pathlib.Path(origen)
        if not origen.exists():
            raise FileNotFoundError(
                f"no existe el archivo de origen: {origen}\n"
                f"Pásalo con --origen \"ruta\\al\\INVENTARIO {tipo.upper()}.xlsx\" "
                f"(o corre el script sin --origen para crear la plantilla vacía).")
        if origen.suffix.lower() != ".xlsx":
            raise ValueError(f"el archivo del área debe ser .xlsx, no {origen.suffix}: {origen}")
        if origen.resolve() != destino.resolve():      # copiarse a sí mismo truena en Windows
            shutil.copy2(origen, destino)
    elif not destino.exists():
        plantilla_inventario.crear(tipo, destino)
        print(f"No había archivo: se creó la plantilla vacía de {tipo.upper()} en {destino}")

    nuevo = inventario.Inventario(destino, tipo=tipo)      # agrega la columna si falta
    if claves_con_qr:
        todas_nuevo = [f["clave"].rsplit("|", 1)[0] for f in nuevo.filas]
        repetidas |= {c for c in claves_con_qr if todas_nuevo.count(c) > 1}
        for clave in repetidas:
            mapa_qr.pop(clave, None)
    etiquetas_ambiguas = len(repetidas)
    recuperadas = 0
    if mapa_ubicaciones:
        for fila in nuevo.filas:
            ubicacion_previa = mapa_ubicaciones.get(fila["clave"])
            if not ubicacion_previa:
                continue
            try:
                nuevo.asignar(fila["fila"], ubicacion_previa)
                recuperadas += 1
            except ValueError:
                # Código capturado con una versión anterior (fuera del catálogo): no se
                # reescribe para no crear ubicaciones fantasma; se avisa para corregirlo.
                descartadas.append((fila["fila"], ubicacion_previa))
    etiquetas = 0
    for fila in nuevo.filas:
        codigo = mapa_qr.get(fila["clave"].rsplit("|", 1)[0])
        if codigo:
            nuevo.asignar_qr(fila["fila"], codigo)
            etiquetas += 1
    nuevo.guardar()
    resumen = nuevo.resumen()
    resumen["ubicaciones_descartadas"] = descartadas
    resumen["etiquetas"] = etiquetas
    resumen["etiquetas_ambiguas"] = etiquetas_ambiguas
    return destino, resumen, recuperadas


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
    if resumen.get("etiquetas"):
        print(f"Códigos QR/Barcode conservados al reimportar: {resumen['etiquetas']}")
    if resumen.get("etiquetas_ambiguas"):
        print(f"AVISO: {resumen['etiquetas_ambiguas']} clave(s) con varias filas tenían código; "
              f"no se reasignó para no duplicar etiquetas")
    descartadas = resumen.get("ubicaciones_descartadas") or []
    if descartadas:
        print(f"\nAVISO: {len(descartadas)} ubicación(es) del archivo anterior no existen en el")
        print("catálogo ALM-PNO-002 y NO se volvieron a escribir (no se crean racks fantasma):")
        for fila, codigo in descartadas[:15]:
            print(f"  fila {fila}: «{codigo}»")
        if len(descartadas) > 15:
            print(f"  … y {len(descartadas) - 15} más")
        print("Corrige esos códigos en el Excel del área o asígnales un rack válido desde la app.")
    if resumen.get("ubicaciones_invalidas"):
        print(f"\nAVISO: {resumen['ubicaciones_invalidas']} renglón(es) traen en UBICACIÓN un código")
        print("que no está en el catálogo; la app los trata como «sin ubicar».")
    if resumen.get("sin_cantidad_legible"):
        print(f"AVISO: {resumen['sin_cantidad_legible']} renglón(es) tienen la CANTIDAD como texto")
        print("raro y no suman al total. Corrige esas celdas en el Excel del área.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
