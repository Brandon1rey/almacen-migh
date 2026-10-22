"""Instalador para otra PC: crea el entorno, instala dependencias y deja los datos listos.

Uso normal (doble clic en instalar.cmd):

    python scripts\\instalar.py

Con el Excel del área en otro lado:

    python scripts\\instalar.py --origen "C:\\ruta\\INVENTARIO MP.xlsx"
"""
import argparse
import pathlib
import shutil
import subprocess
import sys

RAIZ = pathlib.Path(__file__).resolve().parents[1]
VENV = RAIZ / ".venv"
REQUISITOS = RAIZ / "requirements.txt"
VENDOR = RAIZ / "vendor"
DATOS_INICIALES = RAIZ / "datos_iniciales"
INVENTARIO = RAIZ / "data" / "Inventario_MP.xlsx"


def paso(texto):
    print(f"\n== {texto}")


def python_del_entorno():
    """Devuelve el intérprete a usar: el del .venv si se pudo crear."""
    if VENV.exists():
        exe = VENV / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")
        if exe.exists():
            return exe
    return pathlib.Path(sys.executable)


def crear_entorno():
    paso("Creando entorno aislado (.venv)")
    if VENV.exists():
        print("   ya existe, se reutiliza")
        return True
    try:
        subprocess.run([sys.executable, "-m", "venv", str(VENV)], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        print("   creado")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"   no se pudo crear el entorno ({e}); se usará el Python del sistema")
        return False


def instalar_dependencias(py):
    paso("Instalando dependencias (openpyxl)")
    try:
        subprocess.run([str(py), "-m", "pip", "install", "--disable-pip-version-check",
                        "-r", str(REQUISITOS)], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        print("   instaladas desde requirements.txt")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        if (VENDOR / "openpyxl").exists():
            print("   sin internet/pip: se usará la carpeta vendor\\ (ya viene incluida)")
            return True
        print("   FALLA: no hay internet, no hay pip y no existe la carpeta vendor\\")
        print("   Solución: copiar la carpeta vendor\\ de una PC donde ya funcione,")
        print("   o instalar Python desde python.org con 'Add python.exe to PATH'.")
        return False


def preparar_datos(origen=None):
    paso("Preparando los datos")
    if INVENTARIO.exists():
        print("   ya existe data\\Inventario_MP.xlsx")
        return True

    if origen:
        comando = [str(python_del_entorno()), str(RAIZ / "scripts" / "preparar_inventario.py"),
                   "--origen", str(origen)]
    elif (DATOS_INICIALES / "Inventario_MP.xlsx").exists():
        comando = [str(python_del_entorno()), str(RAIZ / "scripts" / "preparar_inventario.py"),
                   "--origen", str(DATOS_INICIALES / "Inventario_MP.xlsx")]
    else:
        print("   ATENCIÓN: no hay inventario todavía.")
        print("   Opción A: copia el Excel del área a datos_iniciales\\Inventario_MP.xlsx y vuelve a correr")
        print("   Opción B: python scripts\\preparar_inventario.py --origen \"C:\\ruta\\INVENTARIO MP.xlsx\"")
        return False

    resultado = subprocess.run(comando, cwd=str(RAIZ))
    return resultado.returncode == 0


def respaldar_datos_previos():
    """Si se reinstala, no pisar la bitácora existente."""
    db = RAIZ / "data" / "bitacora.db"
    if not db.exists():
        return
    destino = RAIZ / "data" / "respaldos"
    destino.mkdir(parents=True, exist_ok=True)
    import datetime as dt
    copia = destino / f"bitacora_antes_de_instalar_{dt.datetime.now():%Y%m%d-%H%M}.db"
    shutil.copy2(db, copia)
    print(f"   respaldo de la bitácora anterior: {copia.name}")


def verificar(py):
    paso("Verificando la instalación")
    resultado = subprocess.run([str(py), str(RAIZ / "scripts" / "verificar_instalacion.py")])
    return resultado.returncode == 0


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--origen", help="ruta del Excel de inventario del área")
    ap.add_argument("--sin-venv", action="store_true", help="no crear .venv (usar el Python actual)")
    args = ap.parse_args(argv)

    print("=" * 64)
    print("  INSTALACIÓN — CONTROL DE UBICACIÓN DE ALMACÉN (MIGH)")
    print("=" * 64)
    print(f"  Carpeta: {RAIZ}")
    print(f"  Python : {sys.version.split()[0]} ({sys.executable})")

    respaldar_datos_previos()
    if not args.sin_venv:
        crear_entorno()
    py = python_del_entorno()
    ok_deps = instalar_dependencias(py)
    ok_datos = preparar_datos(args.origen)
    ok_verif = verificar(py)

    print("\n" + "=" * 64)
    if ok_deps and ok_datos and ok_verif:
        print("  LISTO. Abre arrancar.cmd y entra a http://127.0.0.1:8300")
        return 0
    print("  INSTALACIÓN INCOMPLETA. Revisa los mensajes de arriba.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
