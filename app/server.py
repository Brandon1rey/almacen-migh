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

from app import api, bitacora, inventario  # noqa: E402

PUERTO = int(os.environ.get("ALMACEN_PUERTO", 8300))
STATIC = pathlib.Path(__file__).parent / "static"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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
    def _json(self, codigo, cuerpo):
        datos = json.dumps(cuerpo, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(datos)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(datos)

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
            try:
                codigo, cuerpo = api.manejar(ESTADO, "GET", ruta, params, None)
            except Exception:
                traceback.print_exc()
                return self._json(500, {"error": "error interno, revisa la consola"})
            return self._json(codigo, cuerpo)

        if ruta.startswith("/descargar/"):
            tipo = ruta.rsplit("/", 1)[-1]
            try:
                destino = api.descargar(ESTADO, tipo)
            except KeyError:
                return self._json(404, {"error": f"reporte desconocido: {tipo}"})
            except PermissionError:
                return self._json(409, {"error": "cierra el Excel antes de descargar"})
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
            codigo, respuesta = api.manejar(ESTADO, "POST", partes.path, {}, cuerpo)
        except Exception:
            traceback.print_exc()
            return self._json(500, {"error": "error interno, revisa la consola"})
        return self._json(codigo, respuesta)

    def log_message(self, formato, *args):
        if "/api/" in (args[0] if args else ""):
            return
        print(f"  {self.address_string()} · {formato % args}")


def main():
    global ESTADO
    ruta = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else inventario.RUTA
    ESTADO = Estado(ruta)
    resumen = ESTADO.inventario().resumen()
    servidor = ThreadingHTTPServer(("127.0.0.1", PUERTO), Handler)

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
    print("-" * 66)
    print(f"  Abre en el navegador:  http://127.0.0.1:{PUERTO}")
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
