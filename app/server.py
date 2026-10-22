"""Servidor local del almacén. Solo biblioteca estándar: nada que instalar."""
import json
import mimetypes
import os
import pathlib
import socket
import sys
import threading
import traceback
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
_VENDOR = RAIZ / "vendor"          # dependencias copiadas para PC sin internet ni pip
if _VENDOR.exists():
    sys.path.insert(0, str(_VENDOR))

from app import api, bitacora, inventario, usuarios  # noqa: E402

PUERTO = int(os.environ.get("ALMACEN_PUERTO", 8300))
HOST = os.environ.get("ALMACEN_HOST", "127.0.0.1")
STATIC = pathlib.Path(__file__).parent / "static"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def ip_local():
    """IP de esta PC en la red local (para el modo servidor de LAN)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return "127.0.0.1"


class Estado:
    """Una sola instancia del Excel en memoria, recargada si el archivo cambia en disco."""

    def __init__(self, ruta=None):
        self._lock = threading.RLock()
        self._ruta = pathlib.Path(ruta or inventario.RUTA)
        self.demo = "demo" in str(self._ruta).lower()
        self.archivo = self._ruta.name
        self._inv = inventario.Inventario(self._ruta)
        self._mtime = self._ruta.stat().st_mtime
        ruta_bitacora = (self._ruta.parent / "bitacora_demo.db") if self.demo else None
        self.con = bitacora.conectar(ruta_bitacora)

    def inventario(self):
        with self._lock:
            mtime = self._ruta.stat().st_mtime
            if mtime != self._mtime:
                self._inv = inventario.Inventario(self._ruta)
                self._mtime = mtime
            return self._inv

    def recargar(self):
        with self._lock:
            self._inv = inventario.Inventario(self._ruta)
            self._mtime = self._ruta.stat().st_mtime
            return self._inv


ESTADO = None


class Handler(BaseHTTPRequestHandler):
    server_version = "AlmacenMIGH/1.0"

    # ---------- utilidades ----------
    def _json(self, codigo, cuerpo, cookies=None):
        datos = json.dumps(cuerpo, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(datos)))
        self.send_header("Cache-Control", "no-store")
        for galleta in cookies or []:
            self.send_header("Set-Cookie", galleta)
        self.end_headers()
        self.wfile.write(datos)

    def _galleta(self, nombre):
        crudo = self.headers.get("Cookie", "") or ""
        for parte in crudo.split(";"):
            if "=" in parte:
                clave, valor = parte.split("=", 1)
                if clave.strip() == nombre:
                    return valor.strip()
        return None

    def _sesion(self):
        token = self._galleta(api.COOKIE)
        return bitacora.sesion_valida(ESTADO.con, token), token

    def _bytes(self, datos, tipo, nombre=None):
        self.send_response(200)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(datos)))
        if nombre:
            self.send_header("Content-Disposition", f'attachment; filename="{nombre}"')
        self.end_headers()
        self.wfile.write(datos)

    def _archivo(self, ruta, nombre=None):
        tipo = mimetypes.guess_type(ruta.name)[0] or "application/octet-stream"
        self._bytes(ruta.read_bytes(), tipo, nombre)

    # ---------- rutas ----------
    def do_GET(self):
        partes = urllib.parse.urlparse(self.path)
        ruta = partes.path
        params = dict(urllib.parse.parse_qsl(partes.query))

        if ruta.startswith("/api/"):
            sesion, token = self._sesion()
            try:
                codigo, cuerpo, extra = api.manejar(ESTADO, "GET", ruta, params, None,
                                                    sesion, token)
            except Exception:
                traceback.print_exc()
                return self._json(500, {"error": "error interno, revisa la consola"})
            return self._json(codigo, cuerpo, extra.get("cookies"))

        if ruta.startswith("/descargar/"):
            sesion, _ = self._sesion()
            if not sesion:
                return self._json(401, {"error": "inicia sesión para descargar"})
            tipo = ruta.rsplit("/", 1)[-1]
            try:
                destino = api.descargar(ESTADO, tipo)
            except KeyError:
                return self._json(404, {"error": f"reporte desconocido: {tipo}"})
            except PermissionError:
                return self._json(409, {"error": "cierra el Excel antes de descargar"})
            bitacora.auditar(ESTADO.con, sesion["nombre"], "DESCARGA", tipo)
            return self._archivo(destino, destino.name)

        if ruta in ("/", ""):
            return self._archivo(STATIC / "index.html")
        archivo = STATIC / ruta.lstrip("/")
        if archivo.exists() and archivo.resolve().is_relative_to(STATIC.resolve()):
            return self._archivo(archivo)
        return self._json(404, {"error": "no encontrado"})

    def do_POST(self):
        partes = urllib.parse.urlparse(self.path)
        largo = int(self.headers.get("Content-Length", 0))
        try:
            cuerpo = json.loads(self.rfile.read(largo) or b"{}")
        except json.JSONDecodeError:
            return self._json(400, {"error": "JSON inválido"})
        try:
            sesion, token = self._sesion()
            codigo, respuesta, extra = api.manejar(ESTADO, "POST", partes.path, {}, cuerpo,
                                                   sesion, token)
        except Exception:
            traceback.print_exc()
            return self._json(500, {"error": "error interno, revisa la consola"})
        return self._json(codigo, respuesta, extra.get("cookies"))

    def log_message(self, formato, *args):
        if "/api/" in (args[0] if args else ""):
            return
        print(f"  {self.address_string()} · {formato % args}")


def main():
    global ESTADO
    ruta = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else inventario.RUTA
    ESTADO = Estado(ruta)
    resumen = ESTADO.inventario().resumen()
    servidor = ThreadingHTTPServer((HOST, PUERTO), Handler)

    print("=" * 66)
    print("  CONTROL DE UBICACIÓN — ALMACÉN DE MATERIAS PRIMAS MIGH")
    if ESTADO.demo:
        print("  *** MODO DEMO — DATOS DE DEMOSTRACIÓN, NO ES UN REGISTRO ***")
    print("=" * 66)
    print(f"  Archivo      : {ruta}")
    print(f"  Lotes        : {resumen['total']}")
    print(f"  Con ubicación: {resumen['con_ubicacion']} ({resumen['cobertura']} %)")
    print(f"  Sin ubicar   : {resumen['sin_ubicacion']}")
    print(f"  Vencidos     : {resumen['semaforo']['VENCIDO']}")

    nombre, clave, creada = usuarios.bootstrap(ESTADO.con)
    if creada:
        archivo = ruta.parent / "CREDENCIALES_INICIALES.txt"
        archivo.write_text(
            "CREDENCIALES INICIALES — CONTROL DE UBICACIÓN DE ALMACÉN\n"
            f"usuario : {nombre}\ncontraseña: {clave}\n\n"
            "Entra con estos datos y cambia la contraseña de inmediato.\n"
            "Después borra este archivo.\n", encoding="utf-8")
        bitacora.auditar(ESTADO.con, "sistema", "USUARIO_INICIAL", nombre)
        print("-" * 66)
        print("  PRIMER ARRANQUE: se creó el administrador")
        print(f"    usuario    : {nombre}")
        print(f"    contraseña : {clave}")
        print(f"  (también quedó en {archivo})")
        print("  Cámbiala al entrar. Luego borra ese archivo.")

    print("-" * 66)
    if HOST in ("127.0.0.1", "localhost"):
        print(f"  Abre en el navegador:  http://127.0.0.1:{PUERTO}")
    else:
        print(f"  Esta PC     : http://127.0.0.1:{PUERTO}")
        print(f"  Otras PC    : http://{ip_local()}:{PUERTO}")
        print("  AVISO: modo red. Cada usuario entra con su nombre y contraseña.")
        print("         No exponer fuera de la red de la planta.")
    print("  Ctrl+C para detener.")
    print("=" * 66)
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nCerrando…")
    finally:
        servidor.server_close()


if __name__ == "__main__":
    main()
