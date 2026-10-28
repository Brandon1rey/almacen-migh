"""Instalación completa en una PC nueva: entorno, dependencias, datos y arranque automático.

Uso normal (doble clic en instalar.cmd o instalar_todo.cmd):

    python scripts\\instalar.py

Con el Excel del área en otro lado y arranque automático:

    python scripts\\instalar.py --origen "C:\\ruta\\INVENTARIO MP.xlsx" --arranque --firewall
"""
import argparse
import datetime as dt
import os
import pathlib
import shutil
import subprocess
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

from app import inventarios  # noqa: E402

VENV = RAIZ / ".venv"
REQUISITOS = RAIZ / "requirements.txt"
VENDOR = RAIZ / "vendor"
DATOS_INICIALES = RAIZ / "datos_iniciales"
INVENTARIO = RAIZ / "data" / "Inventario_MP.xlsx"
NOMBRE_ARRANQUE = "Almacen MIGH.cmd"
NOMBRE_TAREA = "Almacen MIGH"


def paso(texto):
    print(f"\n== {texto}")


def carpeta_inicio():
    """Carpeta de Inicio del usuario (arranca al iniciar sesión, sin permisos de administrador)."""
    appdata = os.environ.get("APPDATA")
    if appdata:
        return (pathlib.Path(appdata) / "Microsoft" / "Windows" / "Start Menu"
                / "Programs" / "Startup")
    return (pathlib.Path.home() / "AppData" / "Roaming" / "Microsoft" / "Windows"
            / "Start Menu" / "Programs" / "Startup")


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
    except (subprocess.CalledProcessError, FileNotFoundError):
        if (VENDOR / "openpyxl").exists():
            print("   sin internet/pip: se usará la carpeta vendor\\ (ya viene incluida)")
        else:
            print("   FALLA: no hay internet, no hay pip y no existe la carpeta vendor\\")
            print("   Solución: copiar la carpeta vendor\\ de una PC donde ya funcione,")
            print("   o instalar Python desde python.org con 'Add python.exe to PATH'.")
            return False

    prueba = subprocess.run([str(py), "-c", "import openpyxl; print(openpyxl.__version__)"],
                            capture_output=True, text=True)
    if prueba.returncode != 0:
        if (VENDOR / "openpyxl").exists():
            print("   se usará vendor\\ al arrancar (lo agrega el servidor)")
            return True
        print(f"   FALLA: openpyxl no importa: {prueba.stderr.strip()[:200]}")
        return False
    print(f"   openpyxl {prueba.stdout.strip()} OK")
    return True


def preparar_datos(origen=None):
    """Deja los tres inventarios listos: MP del archivo del área, MEP/PT con plantilla si faltan."""
    paso("Preparando los datos")
    py = python_del_entorno()
    ok = True

    if INVENTARIO.exists():
        print("   data\\Inventario_MP.xlsx ya existe")
    else:
        candidatos = [origen] if origen else []
        if (DATOS_INICIALES / "Inventario_MP.xlsx").exists():
            candidatos.append(DATOS_INICIALES / "Inventario_MP.xlsx")
        candidato = next((c for c in candidatos if c and pathlib.Path(c).exists()), None)
        if candidato:
            resultado = subprocess.run([str(py), str(RAIZ / "scripts" / "preparar_inventario.py"),
                                        "--origen", str(candidato)], cwd=str(RAIZ))
            ok = resultado.returncode == 0
        else:
            print("   ATENCIÓN: no hay inventario de materias primas todavía.")
            print("   Copia el Excel del área a datos_iniciales\\Inventario_MP.xlsx y vuelve a correr,")
            print("   o usa: python scripts\\preparar_inventario.py --origen \"C:\\ruta\\INVENTARIO MP.xlsx\"")
            ok = False

    for tipo in ("mep", "pt"):
        ruta = inventarios.ruta(tipo)
        if ruta.exists():
            print(f"   {ruta.name} ya existe")
            continue
        resultado = subprocess.run([str(py), str(RAIZ / "scripts" / "plantilla_inventario.py"),
                                    "--tipo", tipo, "--vacia"], cwd=str(RAIZ))
        print(f"   plantilla vacía de {tipo.upper()} creada" if resultado.returncode == 0
              else f"   no se pudo crear la plantilla de {tipo.upper()}")

    # Catálogo de insumos del ERP: si el paquete lo trae ya importado, se copia tal cual y la
    # PC no necesita Excel para leer el .xls. Si ya existe uno, no se pisa (puede ser más nuevo).
    catalogo_datos = RAIZ / "data" / "catalogo_insumos.json"
    catalogo_inicial = DATOS_INICIALES / "catalogo_insumos.json"
    if catalogo_datos.exists():
        print("   data\\catalogo_insumos.json ya existe")
    elif catalogo_inicial.exists():
        shutil.copy2(catalogo_inicial, catalogo_datos)
        print("   catálogo de insumos copiado de datos_iniciales")
    else:
        print("   ATENCIÓN: no hay catálogo de insumos todavía.")
        print("   Corre: python scripts\\importar_catalogo.py --origen \"C:\\ruta\\insumo.xls\"")
    return ok


def respaldar_datos_previos():
    """Si se reinstala, no pisar la bitácora existente."""
    db = RAIZ / "data" / "bitacora.db"
    if not db.exists():
        return
    destino = RAIZ / "data" / "respaldos"
    destino.mkdir(parents=True, exist_ok=True)
    copia = destino / f"bitacora_antes_de_instalar_{dt.datetime.now():%Y%m%d-%H%M}.db"
    shutil.copy2(db, copia)
    print(f"   respaldo de la bitácora anterior: {copia.name}")


def contenido_arranque(py, puerto=8300):
    """El .cmd que queda en Inicio: arranca el servidor sin ventana; el servidor escribe su log."""
    interprete = pathlib.Path(py)
    if interprete.name.lower() == "python.exe":
        interprete = interprete.with_name("pythonw.exe")   # sin consola
    partes = [
        "@echo off",
        "rem Arranque automatico del control de ubicacion (MIGH). Quitarlo con quitar_arranque.cmd",
        'cd /d "' + str(RAIZ) + '"',
        'set "ALMACEN_PUERTO=' + str(int(puerto)) + '"',
        'set "ALMACEN_HOST=0.0.0.0"',
        'start "" "' + str(interprete) + '" app' + chr(92) + 'server.py',
        "",
    ]
    return (chr(13) + chr(10)).join(partes) + (chr(13) + chr(10))


def crear_arranque_automatico(py=None, puerto=8300, carpeta=None):
    """Deja el arranque en la carpeta de Inicio del usuario. No requiere administrador."""
    carpeta = pathlib.Path(carpeta or carpeta_inicio())
    carpeta.mkdir(parents=True, exist_ok=True)
    destino = carpeta / NOMBRE_ARRANQUE
    destino.write_text(contenido_arranque(py or python_del_entorno(), puerto),
                       encoding="utf-8", newline="")
    return destino


def quitar_arranque(carpeta=None):
    carpeta = pathlib.Path(carpeta or carpeta_inicio())
    destino = carpeta / NOMBRE_ARRANQUE
    if destino.exists():
        destino.unlink()
        return True
    return False


def crear_tarea(puerto=8300, nombre=NOMBRE_TAREA, py=None):
    """Tarea de Windows al iniciar sesión (requiere consola elevada). Alternativa al Inicio."""
    py = pathlib.Path(py or python_del_entorno())
    pyw = py.with_name("pythonw.exe") if py.name.lower() == "python.exe" else py
    log = RAIZ / "data" / "logs" / "servidor.log"
    (RAIZ / "data" / "logs").mkdir(parents=True, exist_ok=True)
    comando = f'cmd /c cd /d "{RAIZ}" && "{pyw}" app\\server.py >> "{log}" 2>&1'
    resultado = subprocess.run(["schtasks", "/Create", "/F", "/TN", nombre, "/TR", comando,
                                "/SC", "ONLOGON", "/RL", "LIMITED"], capture_output=True, text=True)
    return resultado.returncode == 0, (resultado.stdout + resultado.stderr).strip()


def abrir_firewall(puerto=8300, nombre="Almacen MIGH 8300"):
    """Regla de entrada para que entren teléfonos y otras PC (requiere consola elevada)."""
    subprocess.run(["netsh", "advfirewall", "firewall", "delete", "rule", f"name={nombre}"],
                   capture_output=True)
    resultado = subprocess.run(["netsh", "advfirewall", "firewall", "add", "rule", f"name={nombre}",
                                "dir=in", "action=allow", "protocol=TCP",
                                f"localport={puerto}", "profile=any",
                                "remoteip=LocalSubnet,192.168.0.0/16,10.0.0.0/8,172.16.0.0/12"],
                               capture_output=True, text=True)
    return resultado.returncode == 0, (resultado.stdout + resultado.stderr).strip()


def verificar(py):
    paso("Verificando la instalación")
    resultado = subprocess.run([str(py), str(RAIZ / "scripts" / "verificar_instalacion.py")])
    return resultado.returncode == 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="Instalación completa del control de ubicación")
    ap.add_argument("--origen", help="ruta del Excel de inventario del área")
    ap.add_argument("--sin-venv", action="store_true", help="no crear .venv (usar el Python actual)")
    ap.add_argument("--arranque", action="store_true", help="dejar el arranque automático en Inicio")
    ap.add_argument("--sin-arranque", action="store_true", help="no tocar el arranque automático")
    ap.add_argument("--tarea", action="store_true", help="además, tarea de Windows (requiere admin)")
    ap.add_argument("--firewall", action="store_true", help="abrir el puerto 8300 (requiere admin)")
    ap.add_argument("--puerto", type=int, default=8300)
    ap.add_argument("--quitar-arranque", action="store_true",
                    help="solo quitar el arranque automático y salir")
    args = ap.parse_args(argv)

    if args.quitar_arranque:
        quitado = quitar_arranque()
        print("Arranque automático quitado." if quitado
              else "No había arranque automático que quitar.")
        return 0

    print("=" * 64)
    print("  INSTALACIÓN — CONTROL DE UBICACIÓN DE ALMACÉN (MIGH)")
    print("=" * 64)
    print(f"  Carpeta: {RAIZ}")
    print(f"  Python : {sys.version.split()[0]} ({sys.executable})")

    respaldar_datos_previos()
    DATOS_INICIALES.mkdir(parents=True, exist_ok=True)   # deja lista la carpeta de entrada
    # Aquí se puede dejar Inventario_MP.xlsx y catalogo_insumos.json (el catálogo ya importado
    # desde insumo.xls con scripts\\importar_catalogo.py) para que la instalación los use.
    if not args.sin_venv:
        crear_entorno()
    py = python_del_entorno()
    ok_deps = instalar_dependencias(py)
    ok_datos = preparar_datos(args.origen)

    arranque = None
    if not args.sin_arranque:
        paso("Dejando el arranque automático")
        try:
            arranque = crear_arranque_automatico(py, args.puerto)
            print(f"   {arranque}")
            print("   arranca al iniciar sesión de este usuario (sin permisos de administrador)")
        except OSError as e:
            print(f"   no se pudo crear el arranque: {e}")

    if args.tarea or args.firewall:
        paso("Tareas que necesitan administrador")
        if args.tarea:
            ok, salida = crear_tarea(args.puerto, py=py)
            print(("   tarea creada: " + NOMBRE_TAREA) if ok
                  else f"   tarea NO creada: {salida[:160]}")
        if args.firewall:
            ok, salida = abrir_firewall(args.puerto)
            print(("   puerto abierto en el firewall") if ok
                  else f"   firewall NO modificado: {salida[:160]}")

    ok_verif = verificar(py)

    print("\n" + "=" * 64)
    if ok_deps and ok_datos and ok_verif:
        print("  LISTO.")
        print(f"  Abre: http://127.0.0.1:{args.puerto}")
        print("  Usuarios: python scripts\\crear_usuario.py --listar")
        if arranque:
            print(f"  Arranque automático: {arranque.name} (carpeta Inicio)")
        return 0
    print("  INSTALACIÓN INCOMPLETA. Revisa los mensajes de arriba.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
