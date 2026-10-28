"""Respaldo de los tres inventarios, el catálogo del ERP y la bitácora, con copia
opcional a otra unidad (USB).

    python scripts\\respaldo.py
    python scripts\\respaldo.py --copia-a "E:\\respaldo_almacen"
    python scripts\\respaldo.py --dias 90

Por qué respalda los tres inventarios y el catálogo: la versión anterior solo copiaba
`Inventario_MP.xlsx` y la bitácora. En una instalación con material de empaque (MEP) y
producto terminado (PT), esos dos libros y el catálogo de insumos del ERP —que es la
fuente única de claves y nombres— se quedaban sin copia.
"""
import argparse
import datetime as dt
import hashlib
import json
import pathlib
import shutil
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
CARPETA_DATOS = RAIZ / "data"
DB = CARPETA_DATOS / "bitacora.db"
CATALOGO = CARPETA_DATOS / "catalogo_insumos.json"
DESTINO = CARPETA_DATOS / "respaldos"

# Los tres inventarios del almacén (ALM-PNO-002). Si agregas otro tipo, agrégalo aquí.
INVENTARIOS = ("mp", "mep", "pt")

# Solo se rota lo que este script crea: nunca archivos que el usuario dejó en respaldos\.
PATRONES_RESPALDO = ("Inventario_*.xlsx", "bitacora_*.db", "catalogo_*.json",
                     "respaldo_*.json")


def sha256(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(65536), b""):
            h.update(bloque)
    return h.hexdigest()


def copiar_db(origen, destino):
    """Copia consistente aunque el servidor esté escribiendo."""
    src = sqlite3.connect(str(origen))
    dst = sqlite3.connect(str(destino))
    try:
        with dst:
            src.backup(dst)
        try:
            movimientos = dst.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0]
        except sqlite3.DatabaseError:
            movimientos = None          # bitácora de una versión vieja o archivo raro
    finally:
        src.close()
        dst.close()
    return movimientos


def _copiar_archivo(origen, destino, marca, tipo):
    salida = destino / f"{origen.stem}_{marca}{origen.suffix}"
    shutil.copy2(origen, salida)
    hecho = {"archivo": salida.name, "bytes": salida.stat().st_size,
             "sha256": sha256(salida), "tipo": tipo, "origen": origen.name}
    return hecho, salida


def _misma_ruta(a, b):
    try:
        return pathlib.Path(a).resolve() == pathlib.Path(b).resolve()
    except OSError:
        return False


def respaldar(destino=DESTINO, dias=60, copia_a=None):
    marca = dt.datetime.now().strftime("%Y%m%d-%H%M")
    destino = pathlib.Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    hechos, archivos = [], []

    for tipo in INVENTARIOS:
        origen = CARPETA_DATOS / f"Inventario_{tipo.upper()}.xlsx"
        if not origen.exists():
            continue
        hecho, salida = _copiar_archivo(origen, destino, marca, f"inventario_{tipo}")
        hechos.append(hecho)
        archivos.append(salida)

    if CATALOGO.exists():
        hecho, salida = _copiar_archivo(CATALOGO, destino, marca, "catalogo")
        hechos.append(hecho)
        archivos.append(salida)

    if DB.exists():
        salida = destino / f"bitacora_{marca}.db"
        movimientos = copiar_db(DB, salida)
        hechos.append({"archivo": salida.name, "bytes": salida.stat().st_size,
                       "sha256": sha256(salida), "tipo": "bitacora",
                       "movimientos": movimientos})
        archivos.append(salida)

    if not hechos:
        raise SystemExit("no hay nada que respaldar: no existe data\\Inventario_MP.xlsx "
                         "ni data\\bitacora.db")

    manifiesto = destino / f"respaldo_{marca}.json"
    manifiesto.write_text(json.dumps({"fecha": marca, "equipo": sys.platform,
                                      "archivos": hechos}, indent=2, ensure_ascii=False),
                          encoding="utf-8")
    archivos.append(manifiesto)

    if copia_a:
        externo = pathlib.Path(copia_a)
        if _misma_ruta(externo, destino):
            print("copia externa: se omite (es la misma carpeta del respaldo)")
        else:
            externo.mkdir(parents=True, exist_ok=True)
            for archivo in archivos:
                shutil.copy2(archivo, externo / archivo.name)
            print(f"copia externa: {externo} ({len(archivos)} archivo(s))")

    limite = dt.datetime.now() - dt.timedelta(days=dias)
    borrados = 0
    for patron in PATRONES_RESPALDO:
        for archivo in destino.glob(patron):
            if archivo.is_file() and dt.datetime.fromtimestamp(archivo.stat().st_mtime) < limite:
                archivo.unlink()
                borrados += 1

    print(f"destino: {destino}  (rotación {dias} días, {borrados} archivo(s) viejo(s) eliminado(s))")
    for hecho in hechos:
        print(f"  {hecho['archivo']:<44} {hecho['bytes']:>10,} bytes")
    faltan = [t.upper() for t in INVENTARIOS
              if not (CARPETA_DATOS / f"Inventario_{t.upper()}.xlsx").exists()]
    if faltan:
        print(f"  (sin archivo todavía: {', '.join(faltan)})")
    return hechos


def main(argv=None):
    ap = argparse.ArgumentParser(description="Respalda inventarios, catálogo y bitácora")
    ap.add_argument("--destino", default=str(DESTINO))
    ap.add_argument("--copia-a", default=None, help="carpeta externa (USB) para una segunda copia")
    ap.add_argument("--dias", type=int, default=60, help="días que se conserva cada respaldo")
    args = ap.parse_args(argv)
    respaldar(args.destino, args.dias, args.copia_a)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
