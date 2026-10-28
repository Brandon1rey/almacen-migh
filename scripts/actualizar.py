"""Actualiza una instalación de Almacén MIGH desde el paquete de la USB.

Solo biblioteca estándar: este archivo también se copia a la USB y corre en la PC de destino,
antes o después de instalar, sin depender del código que está actualizando.

    python actualizar.py --paquete "instalacion\\almacen-ubicaciones_v1.5.0_20261023.zip"
    python actualizar.py --paquete <zip> --instalada "D:\\almacen-ubicaciones" --puerto 8300
    python actualizar.py --paquete <zip> --simular

Qué hace, en orden:
  1. encuentra la instalación y compara versiones (no baja de versión salvo --forzar),
  2. respalda el código actual y los datos (bitácora con copia consistente + Excel + catálogo),
  3. detiene el servidor si está corriendo,
  4. copia el código nuevo, retira los archivos que ya no existen en la versión nueva
     y **nunca toca data\\** (inventarios, bitácora, salidas, respaldos),
  5. verifica que la app arranque (importa app.server/api/inventario; openpyxl entra por
     inventario, así no falla en una PC sin pip) y que el servidor responda,
  6. si algo falla, restaura el código del respaldo y avisa.

Códigos de salida: 0 todo bien · 1 error controlado · 2 no encontró la instalación o el paquete.
"""
import argparse
import datetime as dt
import json
import os
import pathlib
import re
import shutil
import socket
import sqlite3
import subprocess
import sys

# La consola de Windows no siempre es UTF-8: sin esto, imprimir «→» o «·» lanza
# UnicodeEncodeError y el script parece fallar aunque ya haya hecho su trabajo.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass

import tempfile
import time
import urllib.error
import urllib.request
import zipfile

NOMBRE_APP = "almacen-ubicaciones"
# carpetas que nunca se pisan: son datos de la instalación, no código
NO_TOCAR = {"data", "dist", ".git", ".venv", "__pycache__", ".pytest_cache"}
CANDIDATOS = [
    # la instalación normal de esta PC (el repo vive en el Escritorio del usuario del almacén)
    pathlib.Path(os.path.expanduser(r"~\Desktop\almacen-ubicaciones")),
    pathlib.Path(r"C:\Users\Almacen Migh\Desktop\almacen-ubicaciones"),
    pathlib.Path(r"C:\almacen-ubicaciones"),
    pathlib.Path(r"D:\almacen-ubicaciones"),
    pathlib.Path(os.path.expanduser(r"~\almacen-ubicaciones")),
    pathlib.Path(os.path.expanduser(r"~\Desktop\almacen-ubicaciones-datos")),
]


# --------------------------------------------------------------------------- versiones

def version_de_texto(texto):
    """'1.5.0' -> (1, 5, 0). Tolera sufijos como '1.5.0-rc1'."""
    numeros = re.findall(r"\d+", str(texto or "").split("-")[0])
    while len(numeros) < 3:
        numeros.append("0")
    return tuple(int(n) for n in numeros[:3])


def leer_version(carpeta):
    """Lee la versión de una instalación (app\\version.py). Devuelve None si no hay."""
    archivo = pathlib.Path(carpeta) / "app" / "version.py"
    if not archivo.exists():
        return None
    texto = archivo.read_text(encoding="utf-8", errors="ignore")
    encontrado = re.search(r'VERSION\s*=\s*["\']([^"\']+)["\']', texto)
    return encontrado.group(1) if encontrado else None


def version_del_paquete(ruta_zip):
    """Lee VERSION.txt del paquete; si no está, cae a app/version.py dentro del zip."""
    with zipfile.ZipFile(ruta_zip) as z:
        if "VERSION.txt" in z.namelist():
            return z.read("VERSION.txt").decode("utf-8", "ignore").strip().splitlines()[0].strip()
        if "app/version.py" in z.namelist():
            texto = z.read("app/version.py").decode("utf-8", "ignore")
            encontrado = re.search(r'VERSION\s*=\s*["\']([^"\']+)["\']', texto)
            return encontrado.group(1) if encontrado else None
    return None


# --------------------------------------------------------------------------- instalación

def buscar_instalacion(pistas=(), perfiles=True):
    """Carpeta instalada: la que tenga app\\server.py y data\\."""
    for pista in list(pistas) + list(CANDIDATOS):
        if not pista:
            continue
        carpeta = pathlib.Path(pista)
        if (carpeta / "app" / "server.py").exists():
            return carpeta
    if not perfiles:
        return None
    # última pasada: cualquier perfil de usuario
    for base in (pathlib.Path(r"C:\Users"), pathlib.Path("D:\\")):
        if not base.exists():
            continue
        for patron in ("*/almacen-ubicaciones", "*/Desktop/almacen-ubicaciones",
                       "*/Desktop/almacen-ubicaciones*"):
            for hija in base.glob(patron):
                if (hija / "app" / "server.py").exists():
                    return hija
    return None


def _codigo_instalado(carpeta):
    """Archivos de código (todo menos datos, venv, git)."""
    salida = []
    for ruta in pathlib.Path(carpeta).rglob("*"):
        if ruta.is_dir():
            continue
        relativa = ruta.relative_to(carpeta)
        if any(parte in NO_TOCAR for parte in relativa.parts):
            continue
        if ruta.suffix in (".pyc", ".pyo", ".tmp"):
            continue
        salida.append(relativa.as_posix())
    return sorted(salida)


# --------------------------------------------------------------------------- respaldo

def _copiar_db(origen, destino):
    """Copia consistente de SQLite, aunque el servidor esté corriendo."""
    con = sqlite3.connect(str(origen))
    try:
        salida = sqlite3.connect(str(destino))
        try:
            con.backup(salida)
        finally:
            salida.close()
    finally:
        con.close()


def respaldar(carpeta, version_actual=None, etiqueta="pre_actualizacion"):
    """ZIP con el código y los datos actuales, en data\\respaldos\\<etiqueta>.zip."""
    carpeta = pathlib.Path(carpeta)
    carpeta_respaldos = carpeta / "data" / "respaldos"
    carpeta_respaldos.mkdir(parents=True, exist_ok=True)
    marca = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    destino = carpeta_respaldos / f"{etiqueta}_{marca}.zip"
    codigo = _codigo_instalado(carpeta)
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for relativa in codigo:
            z.write(carpeta / relativa, f"codigo/{relativa}")
        datos = carpeta / "data"
        for nombre in ("Inventario_MP.xlsx", "Inventario_MEP.xlsx", "Inventario_PT.xlsx",
                       "catalogo_insumos.json"):
            if (datos / nombre).exists():
                z.write(datos / nombre, f"datos/{nombre}")
        if (datos / "bitacora.db").exists():
            temporal = carpeta_respaldos / f"_{marca}_bitacora.db"
            _copiar_db(datos / "bitacora.db", temporal)
            z.write(temporal, "datos/bitacora.db")
            temporal.unlink(missing_ok=True)
        z.writestr("info.json", json.dumps({
            "fecha": dt.datetime.now().isoformat(sep=" ", timespec="seconds"),
            "version": version_actual, "carpeta": str(carpeta), "archivos_codigo": len(codigo),
        }, ensure_ascii=False, indent=1))
    return destino


def restaurar(respaldo, carpeta):
    """Devuelve el código del respaldo (los datos nunca se tocaron)."""
    carpeta = pathlib.Path(carpeta)
    with zipfile.ZipFile(respaldo) as z:
        for nombre in z.namelist():
            if nombre.startswith("codigo/"):
                relativa = nombre[len("codigo/"):]
                destino = carpeta / relativa
                destino.parent.mkdir(parents=True, exist_ok=True)
                with z.open(nombre) as origen, open(destino, "wb") as salida:
                    shutil.copyfileobj(origen, salida)
    return carpeta


# --------------------------------------------------------------------------- servidor

def puerto_ocupado(puerto, host="127.0.0.1"):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.4)
        return s.connect_ex((host, int(puerto))) == 0


def pid_en_puerto(puerto):
    """PIDs escuchando en el puerto (solo Windows; en otros sistemas devuelve [])."""
    if os.name != "nt":
        return []
    try:
        # netstat escribe en la página de códigos de Windows, no en UTF-8: se lee en bytes
        proceso = subprocess.run(["netstat", "-ano"], capture_output=True, timeout=20)
    except (OSError, subprocess.SubprocessError):
        return []
    salida = (proceso.stdout or b"").decode("latin-1", "replace")
    pids = []
    for linea in salida.splitlines():
        partes = linea.split()
        if len(partes) >= 5 and partes[0].upper().startswith("TCP") \
                and partes[1].endswith(f":{puerto}") and partes[3].upper() == "LISTENING":
            if partes[4].isdigit() and int(partes[4]) not in pids:
                pids.append(int(partes[4]))
    return pids


def detener_servidor(puerto, limite=15):
    """Mata el proceso que escucha en el puerto. Devuelve los PID detenidos."""
    detenidos = []
    for pid in pid_en_puerto(puerto):
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True, timeout=20)
        detenidos.append(pid)
    fin = time.time() + limite
    while puerto_ocupado(puerto) and time.time() < fin:
        time.sleep(0.5)
    return detenidos


def _python_de(carpeta):
    """Intérprete de la instalación: su .venv si existe, si no el que está corriendo."""
    for relativa in (r".venv\Scripts\python.exe", r".venv\Scripts\pythonw.exe"):
        ruta = pathlib.Path(carpeta) / relativa
        if ruta.exists():
            return str(ruta)
    return sys.executable


def verificar(carpeta, python=None):
    """La app nueva debe importarse. Devuelve (ok, salida).

    No se importa openpyxl aparte: ya entra por app.inventario, que es lo que la app
    hace realmente al arrancar. Así la comprobación también sirve en una PC sin pip
    (donde openpyxl solo existe dentro de vendor\\) y no da un falso "no arranca".
    """
    python = python or _python_de(carpeta)
    carpeta = pathlib.Path(carpeta)
    codigo = ("import app.server, app.api, app.inventario, app.version;"
              "print(app.version.VERSION)")
    # vendor\ va en el PYTHONPATH por si el paquete es tan viejo que app\__init__.py
    # todavía no lo enganchaba solo.
    entorno = dict(os.environ, PYTHONPATH=os.pathsep.join(
        [str(carpeta), str(carpeta / "vendor")]))
    try:
        proceso = subprocess.run([python, "-c", codigo], cwd=str(carpeta), capture_output=True,
                                 text=True, timeout=120, env=entorno)
    except (OSError, subprocess.SubprocessError) as e:
        return False, str(e)
    if proceso.returncode != 0:
        return False, (proceso.stderr or proceso.stdout or "").strip()[-600:]
    return True, proceso.stdout.strip()


def arrancar(carpeta, puerto, limite=25):
    """Levanta el servidor en segundo plano y espera a que responda."""
    carpeta = pathlib.Path(carpeta)
    python = _python_de(carpeta)
    pythonw = str(pathlib.Path(python).with_name("pythonw.exe"))
    ejecutable = pythonw if pathlib.Path(pythonw).exists() else python
    banderas = 0
    if os.name == "nt":
        banderas = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | \
            getattr(subprocess, "DETACHED_PROCESS", 0)
    entorno = dict(os.environ)
    entorno.setdefault("ALMACEN_PUERTO", str(puerto))
    subprocess.Popen([ejecutable, str(carpeta / "app" / "server.py")], cwd=str(carpeta),
                     env=entorno, creationflags=banderas, stdout=subprocess.DEVNULL,
                     stderr=subprocess.DEVNULL)
    fin = time.time() + limite
    while time.time() < fin:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{puerto}/api/yo", timeout=2) as r:
                if r.status == 200:
                    return True
        except (urllib.error.URLError, OSError):
            time.sleep(0.7)
    return False


# --------------------------------------------------------------------------- instalar

def _manifiesto(carpeta):
    """Nombres de archivo según MANIFIESTO.txt de la instalación (si existe)."""
    archivo = pathlib.Path(carpeta) / "MANIFIESTO.txt"
    if not archivo.exists():
        return set()
    nombres = set()
    for linea in archivo.read_text(encoding="utf-8", errors="ignore").splitlines():
        partes = linea.split()
        if len(partes) >= 3 and len(partes[0]) == 64:
            nombres.add(partes[-1].strip())
    return nombres


def instalar(paquete, carpeta, respaldo=None):
    """Copia el código del paquete sobre la instalación. Devuelve el detalle."""
    carpeta = pathlib.Path(carpeta)
    antes = set(_codigo_instalado(carpeta))
    # solo se retira lo que instaló un paquete anterior (nunca archivos del usuario)
    anteriores = _manifiesto(carpeta)
    copiados, obsoletos = [], []
    with zipfile.ZipFile(paquete) as z:
        nombres = [n for n in z.namelist() if not n.endswith("/")]
        nuevos = {n for n in nombres if n.split("/")[0] not in NO_TOCAR}
        for nombre in nombres:
            if nombre.split("/")[0] in NO_TOCAR:
                continue
            destino = carpeta / nombre
            destino.parent.mkdir(parents=True, exist_ok=True)
            with z.open(nombre) as origen, open(destino, "wb") as salida:
                shutil.copyfileobj(origen, salida)
            copiados.append(nombre)
    # archivos que el paquete anterior instaló y el nuevo ya no trae (ej. demo.cmd)
    for relativa in sorted(anteriores.difference(nuevos)):
        if relativa in ("MANIFIESTO.txt",) or relativa.startswith("data/"):
            continue
        if relativa.split("/")[0] in NO_TOCAR:
            continue
        ruta = carpeta / relativa
        if ruta.exists():
            ruta.unlink()
            obsoletos.append(relativa)
    for relativa in obsoletos:      # dejar vacías las carpetas que quedaron sin archivos
        padre = (carpeta / relativa).parent
        while padre != carpeta and padre.exists() and not any(padre.iterdir()):
            padre.rmdir()
            padre = padre.parent
    return {"copiados": copiados, "obsoletos": obsoletos,
            "nuevos": sorted(nuevos.difference(antes))}


# --------------------------------------------------------------------------- flujo completo

def actualizar(paquete, instalada=None, puerto=8300, forzar=False, simular=False,
               sin_servidor=False, log=print):
    paquete = pathlib.Path(paquete)
    if not paquete.exists():
        return {"ok": False, "motivo": f"no existe el paquete {paquete}", "codigo": 2}
    carpeta = pathlib.Path(instalada) if instalada else buscar_instalacion()
    if not carpeta or not (carpeta / "app" / "server.py").exists():
        return {"ok": False, "motivo": "no se encontró una instalación de almacén-ubicaciones", "codigo": 2}

    v_nueva = version_del_paquete(paquete)
    v_actual = leer_version(carpeta)
    if not v_nueva:
        return {"ok": False, "motivo": "el paquete no declara versión (falta VERSION.txt)", "codigo": 2}
    if v_actual and version_de_texto(v_nueva) < version_de_texto(v_actual) and not forzar:
        return {"ok": False, "codigo": 1,
                "motivo": f"el paquete es más viejo ({v_nueva}) que lo instalado ({v_actual}); "
                          f"usa --forzar si de verdad quieres volver atrás",
                "version_actual": v_actual, "version_nueva": v_nueva}

    log(f"Paquete     : {paquete}")
    log(f"Instalación : {carpeta}")
    log(f"Versión     : {v_actual or '(sin versión)'} → {v_nueva}")
    if simular:
        with zipfile.ZipFile(paquete) as z:
            cantidad = len([n for n in z.namelist() if n.split('/')[0] not in NO_TOCAR])
        return {"ok": True, "simulado": True, "instalada": str(carpeta),
                "version_actual": v_actual, "version_nueva": v_nueva,
                "archivos_paquete": cantidad, "codigo": 0}

    respaldo = respaldar(carpeta, v_actual)
    log(f"Respaldo    : {respaldo}")
    estaba = False if sin_servidor else puerto_ocupado(puerto)
    detenidos = [] if sin_servidor else detener_servidor(puerto)
    if detenidos:
        log(f"Servidor    : detenido ({', '.join(map(str, detenidos))})")

    try:
        detalle = instalar(paquete, carpeta, respaldo)
        ok, salida = verificar(carpeta)
        if not ok:
            raise RuntimeError(f"la versión nueva no arranca: {salida}")
        log(f"Código      : {len(detalle['copiados'])} archivos copiados, "
            f"{len(detalle['obsoletos'])} retirados")
        if detalle["obsoletos"]:
            log(f"Retirados   : {', '.join(detalle['obsoletos'][:8])}")
        arrancado = False
        if not sin_servidor:
            arrancado = arrancar(carpeta, puerto)
            log(f"Servidor    : {'en línea' if arrancado else 'NO respondió'}")
            if estaba and not arrancado:
                raise RuntimeError("el servidor no respondió después de actualizar")
        return {"ok": True, "instalada": str(carpeta), "respaldo": str(respaldo),
                "version_actual": v_actual, "version_nueva": v_nueva, "detalle": detalle,
                "servidor": arrancado, "codigo": 0}
    except Exception as e:  # noqa: BLE001 - cualquier falla se revierte
        log(f"ERROR       : {e}")
        log("Restaurando el código anterior desde el respaldo ...")
        restaurar(respaldo, carpeta)
        if not sin_servidor and estaba:
            arrancar(carpeta, puerto)
        return {"ok": False, "motivo": str(e), "respaldo": str(respaldo), "codigo": 1}


LINEA = "=" * 62


def registrar_resultado(texto, carpeta=None):
    """Deja constancia en ultima_actualizacion.txt, junto al actualizador (la USB)."""
    carpeta = pathlib.Path(carpeta or pathlib.Path(__file__).resolve().parent)
    archivo = carpeta / "ultima_actualizacion.txt"
    marca = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(archivo, "a", encoding="utf-8") as f:
            f.write("[" + marca + "] " + texto + chr(10))
    except OSError:
        pass          # si la USB es de solo lectura, no se interrumpe la actualización
    return archivo


def main(argv=None):
    ap = argparse.ArgumentParser(description="Actualiza Almacén MIGH desde un paquete ZIP")
    ap.add_argument("--paquete", required=True, help="ZIP de la versión nueva")
    ap.add_argument("--instalada", default=None, help="carpeta instalada (si no, se busca sola)")
    ap.add_argument("--puerto", type=int, default=int(os.environ.get("ALMACEN_PUERTO", 8300)))
    ap.add_argument("--forzar", action="store_true", help="permite instalar una versión menor")
    ap.add_argument("--simular", action="store_true", help="solo compara, no toca nada")
    ap.add_argument("--sin-servidor", action="store_true", help="no detiene ni arranca el servidor")
    args = ap.parse_args(argv)

    print(LINEA)
    print("  ACTUALIZACION DEL SISTEMA DE ALMACEN — MIGH")
    print(LINEA)
    resultado = actualizar(args.paquete, args.instalada, args.puerto, args.forzar,
                           args.simular, args.sin_servidor)
    codigo = resultado.get("codigo", 1)

    if resultado["ok"] and resultado.get("simulado"):
        texto = (f"SIMULACION OK: se instalaria {resultado['version_nueva']} "
                 f"({resultado['archivos_paquete']} archivos) en {resultado['instalada']}. "
                 f"Nada se modifico.")
        print(LINEA)
        print("  RESULTADO: SIMULACION CORRECTA (no se cambio nada)")
        print(f"  Version    : {resultado['version_actual']} -> {resultado['version_nueva']}")
        print(f"  Instalacion: {resultado['instalada']}")
        print(f"  Paquete    : {resultado['archivos_paquete']} archivos")
        print(LINEA)
    elif resultado["ok"]:
        texto = (f"EXITO: {resultado['version_actual']} -> {resultado['version_nueva']} en "
                 f"{resultado['instalada']}. Respaldo: {resultado['respaldo']}. "
                 f"Servidor: {'en linea' if resultado.get('servidor') else 'no verificado'}.")
        print(LINEA)
        print("  RESULTADO: ACTUALIZACION CORRECTA")
        print(f"  Version    : {resultado['version_actual']} -> {resultado['version_nueva']}")
        print(f"  Instalacion: {resultado['instalada']}")
        print(f"  Respaldo   : {resultado['respaldo']}")
        print(f"  Servidor   : {'en linea' if resultado.get('servidor') else 'no se detuvo (--sin-servidor)'}")
        print("  Datos      : inventarios y bitacora intactos")
        print(LINEA)
    else:
        texto = (f"FALLO: {resultado['motivo']} | instalacion: {resultado.get('instalada', '?')}"
                 + (f" | respaldo: {resultado['respaldo']}" if resultado.get("respaldo") else ""))
        print(LINEA)
        print("  RESULTADO: NO SE ACTUALIZO")
        print(f"  Motivo     : {resultado['motivo']}")
        if resultado.get("respaldo"):
            print(f"  Respaldo   : {resultado['respaldo']}")
            print("  El codigo anterior quedo restaurado; los datos no se tocaron.")
        print("  Pide apoyo a calidad antes de reintentar.")
        print(LINEA)

    registro = registrar_resultado(texto)
    print(f"  Registro   : {registro}")
    return codigo


if __name__ == "__main__":
    raise SystemExit(main())
