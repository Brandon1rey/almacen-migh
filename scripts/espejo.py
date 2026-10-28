"""Espeja el proyecto en D:\\ (el disco espejo) sin tocar los datos ni las notas personales.

    python scripts\\espejo.py --simular     # solo dice qué falta
    python scripts\\espejo.py               # copia y verifica con SHA-256
    python scripts\\espejo.py --destino "E:\\almacen-ubicaciones"

Qué copia: código, scripts, pruebas, documentación, .cmd y datos_iniciales.
Qué NO toca: data\\ (cada instalación tiene los suyos), .venv, .git, dist, __pycache__,
NOTASBRANDON.txt y .py_path. En el destino nunca borra nada.
"""
import argparse
import hashlib
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
DESTINO = pathlib.Path(r"D:\almacen-ubicaciones")
IGNORAR_DIRS = {"data", ".git", ".venv", "__pycache__", "dist", ".pytest_cache"}
IGNORAR_ARCHIVOS = {"NOTASBRANDON.txt", ".py_path"}
IGNORAR_SUFIJOS = {".pyc", ".pyo", ".tmp", ".db-journal"}


def _sha256(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _archivos(raiz):
    """{ruta relativa: ruta absoluta} de lo que hay que espejar."""
    salida = {}
    raiz = pathlib.Path(raiz)
    if not raiz.exists():
        return salida
    for p in raiz.rglob("*"):
        if p.is_dir():
            continue
        rel = p.relative_to(raiz)
        if any(parte in IGNORAR_DIRS for parte in rel.parts):
            continue
        if p.name in IGNORAR_ARCHIVOS or p.suffix in IGNORAR_SUFIJOS:
            continue
        salida[rel.as_posix()] = p
    return salida


def comparar(origen=None, destino=None):
    """Qué falta, qué cambió y qué está igual entre el proyecto y el espejo."""
    origen = pathlib.Path(origen or RAIZ)
    destino = pathlib.Path(destino or DESTINO)
    a, b = _archivos(origen), _archivos(destino)
    faltan = sorted(set(a).difference(b))
    sobran = sorted(set(b).difference(a))
    distintos = sorted(r for r in set(a).intersection(b) if _sha256(a[r]) != _sha256(b[r]))
    return {"origen": str(origen), "destino": str(destino), "faltan": faltan,
            "distintos": distintos, "sobran": sobran,
            "pendientes": sorted(faltan + distintos),
            "al_dia": not faltan and not distintos}


def espejar(origen=None, destino=None, simular=False):
    """Copia lo faltante o cambiado y vuelve a comparar para confirmar."""
    origen = pathlib.Path(origen or RAIZ)
    destino = pathlib.Path(destino or DESTINO)
    estado = comparar(origen, destino)

    if simular:
        estado["copiados"] = []
        return estado

    if not destino.parent.exists():
        raise FileNotFoundError(f"no existe {destino.parent} (¿el disco está conectado?)")

    copiados = []
    for relativa in estado["pendientes"]:
        origen_archivo = origen / relativa
        destino_archivo = destino / relativa
        destino_archivo.parent.mkdir(parents=True, exist_ok=True)
        if destino_archivo.exists():
            shutil.copy2(origen_archivo, destino_archivo)     # conserva la fecha
        else:
            shutil.copy2(origen_archivo, destino_archivo)
        copiados.append(relativa)

    verificado = comparar(origen, destino)
    verificado["copiados"] = copiados
    return verificado


def main(argv=None):
    ap = argparse.ArgumentParser(description="Espejo del proyecto en D:\\")
    ap.add_argument("--origen", default=None)
    ap.add_argument("--destino", default=None, help=f"por omisión {DESTINO}")
    ap.add_argument("--simular", action="store_true")
    args = ap.parse_args(argv)

    try:
        resultado = espejar(args.origen, args.destino, simular=args.simular)
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
        return 2
    except OSError as e:
        print(f"ERROR al copiar: {e}")
        return 3

    print(f"Origen   : {resultado['origen']}")
    print(f"Espejo   : {resultado['destino']}")
    if args.simular:
        print(f"Faltan   : {len(resultado['pendientes'])} archivo(s)")
        for relativa in resultado["pendientes"][:30]:
            print(f"   {relativa}")
        if len(resultado["pendientes"]) > 30:
            print(f"   … y {len(resultado['pendientes']) - 30} más")
        return 0
    print(f"Copiados : {len(resultado['copiados'])}")
    print(f"Distintos: {len(resultado['distintos'])}")
    if resultado["al_dia"]:
        print("Resultado: espejo al día (verificado por SHA-256).")
        return 0
    print("Resultado: QUEDARON DIFERENCIAS. Revisar:")
    for relativa in resultado["pendientes"][:20]:
        print(f"   {relativa}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
