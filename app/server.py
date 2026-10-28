"""Servidor local del almacén. Solo biblioteca estándar: nada que instalar."""
import json
import mimetypes
import os
import pathlib
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

from app import (api, bitacora, db, eventos, inventario, inventarios, logger,  # noqa: E402
                 usuarios, version)

import time  # noqa: E402  (para medir las peticiones)


def _entero_env(nombre, defecto):
    try:
        return int(os.environ.get(nombre, defecto))
    except (TypeError, ValueError):
        return defecto


PUERTO = _entero_env("ALMACEN_PUERTO", 8300)
HOST = os.environ.get("ALMACEN_HOST", "0.0.0.0")
STATIC = pathlib.Path(__file__).parent / "static"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class _Tee:
    """Escribe en la consola y en el archivo de registro a la vez (o solo en el archivo)."""

    def __init__(self, *destinos):
        self.destinos = [d for d in destinos if d is not None]

    def write(self, texto):
        for destino in self.destinos:
            try:
                destino.write(texto)
                destino.flush()
            except (OSError, ValueError):
                pass

    def flush(self):
        for destino in self.destinos:
            try:
                destino.flush()
            except (OSError, ValueError):
                pass


_REGISTRO = {}      # {"archivo": <log abierto>} mientras el proceso tenga el registro activo


def _abrir_registro(carpeta):
    """Deja el banner y los errores en data\\logs\\servidor.log.

    Es lo que hace usable el arranque automático: con pythonw no hay consola, así que sin esto
    no quedaría rastro de por qué no arrancó. El archivo vive JUNTO A LOS DATOS de esta
    instalación (no en la carpeta del código), para que dos instalaciones no se mezclen.
    """
    global _REGISTRO
    try:
        destino = pathlib.Path(carpeta) / "logs"
        destino.mkdir(parents=True, exist_ok=True)
        archivo = open(destino / "servidor.log", "a", encoding="utf-8", buffering=1)
    except OSError:
        return
    # Se envuelve la salida ACTUAL del proceso, no sys.__stdout__: así se respeta una
    # redirección (`python app\\server.py > salida.txt`) y las pruebas que capturan la
    # consola. Si ya había un _Tee, se toma su primer destino para no anidar capas.
    def base(actual):
        return actual.destinos[0] if isinstance(actual, _Tee) and actual.destinos else actual

    anterior = _REGISTRO.get("archivo")
    _REGISTRO = {"archivo": archivo}
    sys.stdout = _Tee(base(sys.stdout), archivo)
    sys.stderr = _Tee(base(sys.stderr), archivo)
    if anterior is not None and anterior is not archivo:
        try:                      # no dejar archivos abiertos al rearmar el registro
            anterior.close()
        except OSError:
            pass


def cerrar_registro():
    """Cierra data\\logs\\servidor.log (lo usan las pruebas y el cierre ordenado de main())."""
    global _REGISTRO
    archivo = _REGISTRO.get("archivo") if isinstance(_REGISTRO, dict) else None
    _REGISTRO = {}
    if archivo is not None:
        try:
            archivo.close()
        except OSError:
            pass


class Estado:
    """Los tres inventarios del almacén (desde Supabase PostgreSQL o archivos locales)."""

    def __init__(self, ruta=None, carpeta=None, dsn=None):
        self._lock = threading.RLock()
        url_pg = dsn or (db.url_postgres() if (ruta is None and carpeta is None) else None)
        self.en_db = bool(url_pg)
        if carpeta is None:
            carpeta = pathlib.Path(ruta).parent if ruta else inventarios.CARPETA_DATOS
            if not self.en_db and ruta is None and db.es_vercel():
                import shutil
                import tempfile
                tmp_datos = pathlib.Path(tempfile.gettempdir()) / "almacen-datos"
                if not tmp_datos.exists():
                    shutil.copytree(carpeta, tmp_datos, dirs_exist_ok=True)
                carpeta = tmp_datos
        self.carpeta_datos = pathlib.Path(carpeta)
        self._rutas = {t: inventarios.ruta(t, self.carpeta_datos) for t in inventarios.TIPOS}
        if ruta and pathlib.Path(ruta).is_file():
            self._rutas["mp"] = pathlib.Path(ruta)
        self.demo = False if self.en_db else ("demo" in str(self.carpeta_datos).lower())
        self._invs, self._mtimes = {}, {}
        if self.en_db:
            self.ruta_bitacora = None
            self.con = bitacora.conectar(dsn=url_pg)
            for tipo in inventarios.TIPOS:
                self._cargar(tipo)
            self.archivo = "Supabase PostgreSQL (MP · MEP · PT)"
            return
        for tipo, ruta_tipo in self._rutas.items():
            if ruta_tipo.exists():
                self._cargar(tipo)
        if not self._invs:
            raise FileNotFoundError(f"no hay ningún inventario en {self.carpeta_datos}")
        self.archivo = self._rutas["mp"].name if "mp" in self._invs else \
            ", ".join(r.name for t, r in self._rutas.items() if t in self._invs)
        # La bitácora SIEMPRE vive junto a los inventarios de esta instancia.
        # Antes, al abrir una carpeta de datos distinta, las sesiones y los movimientos
        # se guardaban en data\\bitacora.db del proyecto: dos instalaciones mezcladas.
        nombre_db = "bitacora_demo.db" if self.demo else "bitacora.db"
        self.ruta_bitacora = self.carpeta_datos / nombre_db
        self.con = bitacora.conectar(self.ruta_bitacora)

    def _marca_db(self, tipo):
        fila = self.con.execute(
            "SELECT COUNT(*), COALESCE(MAX(actualizado)::text, '') "
            "FROM inventario_lote WHERE inventario=?", (tipo,)).fetchone()
        return (fila[0], fila[1]) if fila else (0, "")

    def _cargar(self, tipo):
        ruta = self._rutas[tipo]
        if self.en_db:
            self._invs[tipo] = inventario.Inventario(ruta, tipo=tipo, con=self.con)
            self._mtimes[tipo] = self._marca_db(tipo)
        else:
            self._invs[tipo] = inventario.Inventario(ruta, tipo=tipo)
            self._mtimes[tipo] = ruta.stat().st_mtime

    def inventario(self, tipo=None):
        clave = inventarios.normalizar_tipo(tipo or "mp")
        if self.en_db:
            with self._lock:
                marca = self._marca_db(clave)
                if self._invs.get(clave) is None or marca != self._mtimes.get(clave):
                    self._cargar(clave)
                return self._invs[clave]
        ruta = self._rutas.get(clave)
        if ruta is None or not ruta.exists():
            raise FileNotFoundError(f"el inventario {clave.upper()} no está cargado")
        with self._lock:
            mtime = ruta.stat().st_mtime
            if self._invs.get(clave) is None or mtime != self._mtimes.get(clave):
                self._cargar(clave)
            return self._invs[clave]

    def cargados(self):
        """[(tipo, inventario)] en orden, solo de los archivos que existen."""
        salida = []
        for tipo in sorted(inventarios.TIPOS, key=lambda t: inventarios.TIPOS[t]["orden"]):
            try:
                salida.append((tipo, self.inventario(tipo)))
            except FileNotFoundError:
                continue
        return salida

    def recargar(self):
        with self._lock:
            for tipo in list(self._invs):
                self._cargar(tipo)
            return self._invs


ESTADO = None
_ESTADO_LOCK = threading.Lock()


def asegurar_estado():
    """Inicializa ESTADO bajo demanda (usado en funciones serverless de Vercel)."""
    global ESTADO
    if ESTADO is not None:
        return ESTADO
    with _ESTADO_LOCK:
        if ESTADO is None:
            ESTADO = Estado()
            logger.configurar(ESTADO.carpeta_datos, consola=True)
            eventos.conectar_registro()
            primer_arranque(ESTADO)
    return ESTADO


class Handler(BaseHTTPRequestHandler):
    server_version = "AlmacenMIGH/1.8"

    # ---------- utilidades ----------
    def _json(self, codigo, cuerpo, cookies=None):
        datos = json.dumps(cuerpo, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(codigo)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(datos)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for galleta in cookies or []:
            self.send_header("Set-Cookie", galleta)
        self.end_headers()
        self.wfile.write(datos)
        self._registrar_peticion(codigo)

    def _registrar_peticion(self, codigo):
        """Una línea por petición a la API, con duración y usuario (para depurar)."""
        datos = getattr(self, "_debug", None)
        if not datos:
            return
        self._debug = None
        logger.peticion(datos["metodo"], datos["ruta"], codigo,
                        (time.monotonic() - datos["inicio"]) * 1000, datos.get("usuario"))

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

    def _bytes(self, datos, tipo, nombre=None, cache="no-cache"):
        self.send_response(200)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(datos)))
        self.send_header("Cache-Control", cache)
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        if nombre:
            self.send_header("Content-Disposition", f'attachment; filename="{nombre}"')
        self.end_headers()
        self.wfile.write(datos)

    def _archivo(self, ruta, nombre=None):
        tipo = mimetypes.guess_type(ruta.name)[0] or "application/octet-stream"
        cache = "no-cache" if tipo.startswith("text/html") else "max-age=3600"
        self._bytes(ruta.read_bytes(), tipo, nombre, cache=cache)

    # ---------- despacho ----------
    def _despachar(self, metodo):
        partes = urllib.parse.urlparse(self.path)
        ruta = partes.path
        params = dict(urllib.parse.parse_qsl(partes.query))

        if metodo == "GET" and not ruta.startswith(("/api/", "/descargar/")):
            if ruta == "/favicon.ico":
                self.send_response(204)
                self.send_header("Content-Length", "0")
                self.end_headers()
                return
            if ruta in ("/", ""):
                return self._archivo(STATIC / "index.html")
            archivo = STATIC / ruta.lstrip("/")
            try:
                dentro = archivo.exists() and archivo.resolve().is_relative_to(STATIC.resolve())
            except OSError:
                dentro = False
            if dentro:
                return self._archivo(archivo)
            return self._json(404, {"error": "no encontrado"})

        try:
            asegurar_estado()
        except Exception as e:  # noqa: BLE001
            return self._json(503, {"error": f"Error al conectar con la base de datos: {e}"})

        if ruta.startswith("/api/"):
            return self._api(metodo, ruta, params)

        if metodo != "GET":
            return self._json(405, {"error": f"método no permitido: {metodo} {ruta}"})

        if ruta.startswith("/descargar/"):
            return self._descargar(ruta, params)

        return self._json(404, {"error": "no encontrado"})

    def _api(self, metodo, ruta, params):
        self._debug = {"metodo": metodo, "ruta": ruta, "inicio": time.monotonic(),
                       "usuario": None}
        sesion, token = self._sesion()
        self._debug["usuario"] = sesion["nombre"] if sesion else None
        cuerpo = {}
        if metodo in ("POST", "PATCH"):
            try:
                largo = int(self.headers.get("Content-Length", 0) or 0)
            except ValueError:
                largo = 0
            crudo = self.rfile.read(largo) if largo else b"{}"
            try:
                cuerpo = json.loads(crudo or b"{}")
            except (json.JSONDecodeError, UnicodeDecodeError):
                return self._json(400, {"error": "el cuerpo de la petición no es JSON válido"})
            if not isinstance(cuerpo, dict):
                return self._json(400, {"error": "el cuerpo de la petición debe ser un objeto JSON"})
        try:
            codigo, respuesta, extra = api.manejar(ESTADO, metodo, ruta, params, cuerpo,
                                                   sesion, token)
        except ValueError as e:
            logger.depurar("valor invalido", metodo=metodo, ruta=ruta, detalle=str(e))
            return self._json(400, {"error": str(e)})
        except Exception as e:                     # noqa: BLE001 - se registra y se responde
            logger.excepcion(f"error interno en {metodo} {ruta}")
            eventos.publicar("servidor.error", metodo=metodo, ruta=ruta,
                             detalle=f"{type(e).__name__}: {e}",
                             usuario=self._debug.get("usuario") if self._debug else None)
            return self._json(500, {"error": "error interno; revisa la ventana negra del servidor"})
        return self._json(codigo, respuesta, extra.get("cookies"))

    def _descargar(self, ruta, params):
        sesion, _ = self._sesion()
        if not sesion:
            return self._json(401, {"error": "inicia sesión para descargar"})
        tipo = ruta.rsplit("/", 1)[-1]
        try:
            destino = api.descargar(ESTADO, tipo, params.get("tipo"))
        except KeyError:
            return self._json(404, {"error": f"reporte desconocido: {tipo}"})
        except ValueError as e:
            return self._json(400, {"error": str(e)})
        except PermissionError:
            return self._json(409, {"error": "cierra el Excel en esta PC antes de descargar"})
        except FileNotFoundError as e:
            return self._json(409, {"error": str(e)})
        except Exception:
            traceback.print_exc()
            return self._json(500, {"error": "no se pudo generar el reporte; revisa la consola"})
        bitacora.auditar(ESTADO.con, sesion["nombre"], "DESCARGA", tipo)
        return self._archivo(destino, destino.name)

    def do_GET(self):
        self._despachar("GET")

    def do_POST(self):
        self._despachar("POST")

    def do_PATCH(self):
        self._despachar("PATCH")

    def log_message(self, formato, *args):
        """El servidor de la librería solo registra lo que NO es API (el detalle va en _json)."""
        if "/api/" in (args[0] if args else ""):
            return
        logger.depurar("estatico", cliente=self.address_string(), detalle=formato % args)


def _anunciar(estado, puerto, host):
    print("=" * 66, flush=True)
    print("  CONTROL DE UBICACIÓN — ALMACÉN MIGH", flush=True)
    print(f"  versión {version.VERSION} · {version.FECHA}", flush=True)
    if estado.demo:
        print("  *** MODO DEMO — DATOS DE DEMOSTRACIÓN, NO ES UN REGISTRO ***", flush=True)
    print("=" * 66, flush=True)
    print(f"  Datos        : {estado.carpeta_datos}", flush=True)
    for tipo, inv in estado.cargados():
        r = inv.resumen()
        print(f"  {tipo.upper():<3} {inv.nombre:<20} {r['total']:>5} lotes · "
              f"{r['con_ubicacion']:>5} con ubicación ({r['cobertura']:>5} %) · "
              f"{r['semaforo']['VENCIDO']:>3} vencidos", flush=True)
    faltan = [t for t in inventarios.TIPOS if t not in {t for t, _ in estado.cargados()}]
    if faltan:
        print(f"  Sin cargar   : {', '.join(t.upper() for t in faltan)} "
              f"(crea la plantilla con scripts\\plantilla_inventario.py --tipo {faltan[0]})",
              flush=True)
    print("-" * 66, flush=True)
    propia = "127.0.0.1" if host in ("0.0.0.0", "127.0.0.1", "localhost") else host
    print(f"  Abre en el navegador:  http://{propia}:{puerto}", flush=True)
    print("  Ctrl+C para detener.", flush=True)
    print("=" * 66, flush=True)


def primer_arranque(estado):
    """Crea el administrador inicial y deja sus credenciales a la vista.

    Devuelve (nombre, clave) si lo creó, o None. El archivo de credenciales se escribe
    junto a los datos de ESTA instalación: antes el servidor intentaba usar una variable
    inexistente y se caía con NameError justo después de crear el usuario, así que la
    contraseña se perdía y nadie podía entrar nunca.
    """
    nombre, clave, creada = usuarios.bootstrap(estado.con)
    if not creada:
        return None
    try:
        archivo = estado.carpeta_datos / "CREDENCIALES_INICIALES.txt"
        archivo.write_text(
            "CREDENCIALES INICIALES — CONTROL DE UBICACIÓN DE ALMACÉN\n"
            f"usuario : {nombre}\ncontraseña: {clave}\n\n"
            "Entra con estos datos y cambia la contraseña de inmediato.\n"
            "Después borra este archivo.\n", encoding="utf-8")
    except OSError as e:
        archivo = f"(no se pudo escribir el archivo de credenciales: {e})"
    bitacora.auditar(estado.con, "sistema", "USUARIO_INICIAL", nombre)
    print("-" * 66, flush=True)
    print("  PRIMER ARRANQUE: se creó el administrador", flush=True)
    print(f"    usuario    : {nombre}", flush=True)
    print(f"    contraseña : {clave}", flush=True)
    print(f"  (también quedó en {archivo})", flush=True)
    print("  Cámbiala al entrar. Luego borra ese archivo.", flush=True)
    return nombre, clave


def arrancar(argumento=None, host=None, puerto=None, servir=True):
    """Levanta el sistema. `servir=False` se usa en las pruebas (no abre el socket)."""
    global ESTADO
    if argumento is not None:
        argumento = pathlib.Path(argumento)
    if argumento is not None and argumento.is_dir():
        ESTADO = Estado(carpeta=argumento)
    else:
        ESTADO = Estado(ruta=argumento)
    _abrir_registro(ESTADO.carpeta_datos)
    logger.configurar(ESTADO.carpeta_datos)
    eventos.conectar_registro()
    eventos.publicar("servidor.arrancado", version=version.VERSION,
                     puerto=puerto or PUERTO, host=host or HOST,
                     demo=bool(getattr(ESTADO, "demo", False)),
                     carpeta=str(ESTADO.carpeta_datos))
    servidor = None
    if servir:
        # Se abre el puerto ANTES de crear al administrador: si el puerto está ocupado,
        # el operador no se queda con una contraseña que nunca vio en pantalla.
        servidor = ThreadingHTTPServer((host or HOST, puerto or PUERTO), Handler)
    _anunciar(ESTADO, puerto or PUERTO, host or HOST)
    primer_arranque(ESTADO)
    if servidor is None:
        return ESTADO
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nCerrando…", flush=True)
    finally:
        servidor.server_close()
        eventos.publicar("servidor.detenido", version=version.VERSION)
    return ESTADO


def main():
    argumento = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else None
    try:
        arrancar(argumento)
    except FileNotFoundError as e:
        print("\nNo se puede arrancar:", e, flush=True)
        print("Revisa que la carpeta data\\ tenga Inventario_MP.xlsx "
              "(o corre instalar.cmd).\n", flush=True)
        return 1
    except OSError as e:
        print("\nNo se puede arrancar: el puerto está ocupado o los datos no son accesibles.",
              flush=True)
        print(f"Detalle: {e}\n", flush=True)
        return 1
    finally:
        cerrar_registro()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
