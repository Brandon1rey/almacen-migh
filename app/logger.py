"""Registro para depurar: todo lo que pasa queda en `data\\logs\\depuracion.log`.

Tres piezas:

- `configurar(carpeta_datos)`: deja el registro listo (una vez por proceso). El nivel sale de
  `ALMACEN_LOG` (`DEBUG`, `INFO`, `AVISO`, `ERROR`); por omisión `INFO`.
- `evento(nombre, **datos)` / `peticion(...)` / `excepcion(...)`: lo que se escribe.
- `ruta()`: dónde está el archivo (la interfaz de diagnóstico lo muestra).

Reglas: se escribe en UTF-8, el archivo rota a 2 MB (5 copias) y **nunca** se registran
contraseñas ni tokens: las claves que suenan a secreto se tachan como `***`.
"""
import logging
import logging.handlers
import os
import pathlib
import sys
import threading

RAIZ = pathlib.Path(__file__).resolve().parents[1]
NOMBRE = "almacen"
NIVELES = {"DEBUG": logging.DEBUG, "INFO": logging.INFO, "AVISO": logging.INFO,
           "WARNING": logging.WARNING, "ERROR": logging.ERROR, "CRITICO": logging.CRITICAL}
SECRETOS = ("password", "contrasena", "contraseña", "clave", "token", "cookie", "secreto")
LARGO_MAX = 300

_LOG = logging.getLogger(NOMBRE)
_LOG.addHandler(logging.NullHandler())      # si nadie configuró el registro, no ensucia la consola
_ESTADO = {"ruta": None, "nivel": "INFO"}
_LOCK = threading.Lock()


def _tachar(valor):
    texto = str(valor)
    if len(texto) > LARGO_MAX:
        texto = texto[:LARGO_MAX] + "…"
    return texto.replace("\n", " ")


def _formato(datos):
    """`clave=valor` legible, sin secretos y sin valores gigantes."""
    partes = []
    for clave, valor in (datos or {}).items():
        if valor is None:
            continue
        if any(secreto in str(clave).lower() for secreto in SECRETOS):
            partes.append(f"{clave}=***")
        else:
            partes.append(f"{clave}={_tachar(valor)}")
    return " ".join(partes)


def configurar(carpeta_datos=None, nivel=None, consola=True):
    """Deja el registro escribiendo en `data\\logs\\depuracion.log` (idempotente)."""
    with _LOCK:
        carpeta = pathlib.Path(carpeta_datos) if carpeta_datos else RAIZ / "data"
        if _LOG.handlers and _ESTADO["ruta"] and _ESTADO.get("carpeta") == str(carpeta):
            return _ESTADO["ruta"]
        if _LOG.handlers and _ESTADO.get("carpeta"):
            # Otra carpeta de datos (demo, pruebas, otra instalación en el mismo proceso):
            # se cierra el registro anterior y se abre el nuevo.
            for manejador in list(_LOG.handlers):
                _LOG.removeHandler(manejador)
                try:
                    manejador.close()
                except Exception:     # noqa: BLE001
                    pass
            _ESTADO.update({"ruta": None})
        _ESTADO["carpeta"] = str(carpeta)
        texto = str(nivel or os.environ.get("ALMACEN_LOG") or "INFO").strip().upper()
        _ESTADO["nivel"] = texto
        _LOG.setLevel(NIVELES.get(texto, logging.INFO))
        _LOG.propagate = False
        for manejador in list(_LOG.handlers):
            if isinstance(manejador, logging.NullHandler):
                _LOG.removeHandler(manejador)
        formato = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s",
                                    "%Y-%m-%d %H:%M:%S")
        destino = None
        try:
            (carpeta / "logs").mkdir(parents=True, exist_ok=True)
            destino = carpeta / "logs" / "depuracion.log"
            archivo = logging.handlers.RotatingFileHandler(
                destino, maxBytes=2_000_000, backupCount=5, encoding="utf-8")
            archivo.setFormatter(formato)
            _LOG.addHandler(archivo)
            _ESTADO["ruta"] = destino
        except OSError:
            destino = None            # sin permiso para escribir: al menos queda la consola
        try:
            if consola:
                salida = logging.StreamHandler(sys.stdout)
                salida.setFormatter(formato)
                _LOG.addHandler(salida)
        except Exception:             # noqa: BLE001 - consola cerrada (pythonw)
            pass
        if destino:
            _LOG.info("registro iniciado nivel=%s archivo=%s", texto, destino)
        return destino


def ruta():
    return _ESTADO["ruta"]


def nivel():
    return _ESTADO["nivel"]


def evento(nombre, **datos):
    _LOG.info("evento=%s %s", nombre, _formato(datos))


def registrar_evento(registro):
    """Escucha del bus de eventos: cada evento del sistema queda en el archivo."""
    if not isinstance(registro, dict):
        return
    datos = {k: v for k, v in registro.items() if k not in ("evento", "id")}
    _LOG.info("evento=%s id=%s %s", registro.get("evento"), registro.get("id"),
              _formato(datos))


def peticion(metodo, ruta_pedida, codigo, milisegundos, usuario=None):
    """Una línea por petición a la API: método, ruta, resultado, duración y quién."""
    texto = (f"peticion {metodo} {ruta_pedida} -> {codigo} "
             f"{float(milisegundos):.1f}ms")
    if usuario:
        texto += f" usuario={_tachar(usuario)}"
    if codigo >= 500:
        _LOG.error(texto)
    elif codigo >= 400:
        _LOG.warning(texto)
    else:
        _LOG.info(texto)


def navegador(mensaje, contexto=None, url=None, nivel_js="error"):
    """Errores que reporta el navegador (window.onerror y promesas rechazadas)."""
    _LOG.error("navegador nivel=%s url=%s mensaje=%s %s", nivel_js, url,
               _tachar(mensaje), _formato(contexto))


def excepcion(mensaje):
    """Excepción no controlada: deja el traceback completo en el archivo."""
    _LOG.exception("%s", mensaje)


def depurar(mensaje, **datos):
    if _LOG.isEnabledFor(logging.DEBUG):
        _LOG.debug("%s %s", mensaje, _formato(datos))


def reiniciar():
    """Solo para pruebas: suelta los manejadores y olvida el estado."""
    with _LOCK:
        for manejador in list(_LOG.handlers):
            _LOG.removeHandler(manejador)
            try:
                manejador.close()
            except Exception:     # noqa: BLE001
                pass
        _ESTADO.update({"ruta": None, "nivel": "INFO", "carpeta": None})
        _LOG.addHandler(logging.NullHandler())   # sin esto, logging escupe a stderr
