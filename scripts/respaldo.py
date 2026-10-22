"""Respaldo del inventario y la bitácora, con copia opcional a otra unidad (USB).

    python scripts\\respaldo.py
    python scripts\\respaldo.py --copia-a "E:\\respaldo_almacen"
    python scripts\\respaldo.py --dias 90
"""
import argparse
import datetime as dt
import hashlib
import json
import pathlib
import shutil
import sqlite3
import sys

RAIZ = pathlib.Path(__file__).resolve().parents[1]
INVENTARIO = RAIZ / "data" / "Inventario_MP.xlsx"
DB = RAIZ / "data" / "bitacora.db"
DESTINO = RAIZ / "data" / "respaldos"


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
    with dst:
        src.backup(dst)
    movimientos = dst.execute("SELECT COUNT(*) FROM movimiento").fetchone()[0]
    src.close()
    dst.close()
    return movimientos


def respaldar(destino=DESTINO, dias=60, copia_a=None):
    marca = dt.datetime.now().strftime("%Y%m%d-%H%M")
    destino = pathlib.Path(destino)
    destino.mkdir(parents=True, exist_ok=True)
    hechos = []

    if INVENTARIO.exists():
        salida = destino / f"Inventario_MP_{marca}.xlsx"
        shutil.copy2(INVENTARIO, salida)
        hechos.append({"archivo": salida.name, "bytes": salida.stat().st_size,
                       "sha256": sha256(salida), "tipo": "inventario"})

    if DB.exists():
        salida = destino / f"bitacora_{marca}.db"
        movimientos = copiar_db(DB, salida)
        hechos.append({"archivo": salida.name, "bytes": salida.stat().st_size,
                       "sha256": sha256(salida), "tipo": "bitacora",
                       "movimientos": movimientos})

    if not hechos:
        raise SystemExit("no hay nada que respaldar (falta data/Inventario_MP.xlsx)")

    manifiesto = destino / f"respaldo_{marca}.json"
    manifiesto.write_text(json.dumps({"fecha": marca, "equipo": sys.platform,
                                      "archivos": hechos}, indent=2, ensure_ascii=False),
                          encoding="utf-8")

    if copia_a:
        externo = pathlib.Path(copia_a)
        externo.mkdir(parents=True, exist_ok=True)
        for hecho in hechos + [{"archivo": manifiesto.name}]:
            shutil.copy2(destino / hecho["archivo"], externo / hecho["archivo"])
        print(f"copia externa: {externo}")

    limite = dt.datetime.now() - dt.timedelta(days=dias)
    borrados = 0
    for archivo in destino.glob("*"):
        if archivo.is_file() and dt.datetime.fromtimestamp(archivo.stat().st_mtime) < limite:
            archivo.unlink()
            borrados += 1

    print(f"destino: {destino}  (rotación {dias} días, {borrados} archivo(s) viejo(s) eliminado(s))")
    for hecho in hechos:
        print(f"  {hecho['archivo']:<40} {hecho['bytes']:>10,} bytes")
    return hechos


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--destino", default=str(DESTINO))
    ap.add_argument("--copia-a", default=None)
    ap.add_argument("--dias", type=int, default=60)
    args = ap.parse_args(argv)
    respaldar(args.destino, args.dias, args.copia_a)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
