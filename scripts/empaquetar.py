"""Empaqueta el sistema en un ZIP listo para copiar a otra PC, con manifiesto SHA-256.

    python scripts\\empaquetar.py                  → dist\\almacen-ubicaciones_<fecha>.zip
    python scripts\\empaquetar.py --vendor          → además incluye vendor\\ (PC sin internet/pip)
    python scripts\\empaquetar.py --con-datos       → además incluye data\\Inventario_MP.xlsx y bitacora.db
"""
import sys

# La consola de Windows no siempre es UTF-8: sin esto, imprimir «→» o «·» lanza
# UnicodeEncodeError y el script parece fallar aunque ya haya hecho su trabajo.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

import argparse
import datetime as dt
import hashlib
import importlib.util
import pathlib
import re
import shutil
import zipfile

RAIZ = pathlib.Path(__file__).resolve().parents[1]
DIST = RAIZ / "dist"

INCLUIR = [
    "app", "scripts", "tests", "docs",
    "requirements.txt", "requirements-docs.txt",
    "README.md", "LEEME_INSTALACION.md",
]
# todos los .cmd del proyecto entran solos: agregar uno nuevo no debe obligar a tocar esta lista
INCLUIR.extend(sorted(p.name for p in RAIZ.glob("*.cmd")))
EXCLUIR_DIRS = {"__pycache__", ".venv", ".git", "__MACOSX"}
EXCLUIR_SUFIJOS = {".pyc", ".pyo", ".tmp"}
PAQUETES_VENDOR = ("openpyxl", "et_xmlfile")


def version_actual():
    """Versión declarada en app\\version.py (la que compara el actualizador)."""
    texto = (RAIZ / "app" / "version.py").read_text(encoding="utf-8")
    encontrado = re.search(r"VERSION\s*=\s*[\"']([^\"']+)[\"']", texto)
    return encontrado.group(1) if encontrado else "0.0.0"


def sha256(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(65536), b""):
            h.update(bloque)
    return h.hexdigest()


def archivos_de(item):
    ruta = RAIZ / item
    if ruta.is_file():
        yield ruta
        return
    for hija in sorted(ruta.rglob("*")):
        if hija.is_dir() or hija.suffix in EXCLUIR_SUFIJOS:
            continue
        if any(parte in EXCLUIR_DIRS for parte in hija.parts):
            continue
        yield hija


def copiar_vendor():
    """Copia openpyxl y et_xmlfile desde el Python actual a vendor\\ (sin pip en destino)."""
    destino = RAIZ / "vendor"
    if destino.exists():
        shutil.rmtree(destino)
    destino.mkdir()
    copiados = []
    for nombre in PAQUETES_VENDOR:
        especificacion = importlib.util.find_spec(nombre)
        if not especificacion or not especificacion.submodule_search_locations:
            print(f"   aviso: {nombre} no está instalado en este Python; se omite")
            continue
        origen = pathlib.Path(list(especificacion.submodule_search_locations)[0])
        shutil.copytree(origen, destino / nombre,
                        ignore=shutil.ignore_patterns("__pycache__", "*.dist-info"))
        copiados.append(nombre)
    return copiados


def empaquetar(salida=None, con_vendor=False, con_datos=False):
    DIST.mkdir(parents=True, exist_ok=True)
    marca = dt.datetime.now().strftime("%Y%m%d-%H%M")
    version = version_actual()
    salida = pathlib.Path(salida) if salida else DIST / f"almacen-ubicaciones_v{version}_{marca}.zip"

    if con_vendor:
        print("copiando dependencias a vendor\\ ...")
        print("   ", ", ".join(copiar_vendor()) or "nada")

    items = list(INCLUIR)
    if con_vendor:
        items.append("vendor")
    if con_datos:
        items += ["data/Inventario_MP.xlsx", "data/bitacora.db"]

    registro, faltantes = [], []
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
        for item in items:
            if not (RAIZ / item).exists():
                faltantes.append(item)
                continue
            for archivo in archivos_de(item):
                relativa = archivo.relative_to(RAIZ).as_posix()
                z.write(archivo, relativa)
                registro.append((relativa, archivo.stat().st_size, sha256(archivo)))

        z.writestr("VERSION.txt", f"{version}\n"
                   f"generado: {dt.datetime.now():%Y-%m-%d %H:%M}\n"
                   f"dependencias: {'vendor' if con_vendor else 'pip'}\n")
        manifiesto = ["MANIFIESTO DE EMPAQUE — CONTROL DE UBICACIÓN DE ALMACÉN",
                      f"version: {version}",
                      f"generado: {dt.datetime.now():%Y-%m-%d %H:%M}",
                      f"archivos: {len(registro)}",
                      f"dependencias incluidas: {'sí (vendor/)' if con_vendor else 'no (requiere pip)'}",
                      f"datos incluidos: {'sí' if con_datos else 'no'}",
                      "",
                      f"{'SHA-256':<64} {'BYTES':>10}  ARCHIVO"]
        for relativa, tamano, digest in sorted(registro):
            manifiesto.append(f"{digest} {tamano:>10}  {relativa}")
        z.writestr("MANIFIESTO.txt", "\n".join(manifiesto) + "\n")

    print(f"\nZIP        : {salida}")
    print(f"archivos   : {len(registro)}")
    print(f"tamaño     : {salida.stat().st_size/1024/1024:.2f} MB")
    print(f"sha256 zip : {sha256(salida)}")
    if faltantes:
        print(f"omitidos   : {', '.join(faltantes)}")
    print("\nPara instalar en otra PC: descomprimir y doble clic en instalar_todo.cmd")
    return salida


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--salida", default=None)
    ap.add_argument("--vendor", action="store_true", help="incluir dependencias (PC sin internet)")
    ap.add_argument("--con-datos", action="store_true", help="incluir inventario y bitácora actuales")
    args = ap.parse_args(argv)
    empaquetar(args.salida, args.vendor, args.con_datos)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
