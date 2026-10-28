"""Bus de eventos del sistema: quien hace algo lo publica, quien quiere enterarse escucha.

Para qué sirve en este proyecto:

- **Depurar**: un solo escucha escribe cada evento en `data\\logs\\depuracion.log`, así el
  archivo cuenta la historia completa (quién entró, qué movió, qué falló) sin repartir
  `print()` por todo el código.
- **Diagnóstico**: los últimos eventos quedan en memoria y `GET /api/diagnostico` los
  devuelve, útil cuando alguien dice «no me guardó» y no hay forma de reproducirlo.

Un escucha que falla **no** tumba la operación que publicó el evento: se registra el fallo y
se sigue. Los eventos son datos planos (texto, números, booleanos): nunca objetos vivos.
"""
import collections
import datetime as dt
import itertools
import threading

from app import logger

MAX_EVENTOS = 300
_PREFIJOS_VALIDOS = (str, int, float, bool, type(None))

_cerrojo = threading.Lock()
_secuencia = itertools.count(1)
_eventos = collections.deque(maxlen=MAX_EVENTOS)
_escuchas = []          # [(patron, funcion)]
_escucha_log = {"hecho": False}


def _limpio(valor):
    if isinstance(valor, _PREFIJOS_VALIDOS):
        return valor
    if isinstance(valor, (list, tuple)):
        return [v for v in valor if isinstance(v, _PREFIJOS_VALIDOS)][:20]
    return str(valor)[:200]


def escuchar(patron, funcion):
    """Registra un escucha. `patron` es el nombre exacto, `"*"` o un prefijo (`"movimiento.*"`)."""
    if not callable(funcion):
        raise TypeError("el escucha tiene que ser una función")
    with _cerrojo:
        if not any(p == patron and f is funcion for p, f in _escuchas):
            _escuchas.append((str(patron or "*"), funcion))
    return funcion


def _escuchas_para(nombre):
    for patron, funcion in _escuchas:
        if patron == "*" or patron == nombre:
            yield funcion
        elif patron.endswith(".*") and nombre.startswith(patron[:-1]):
            yield funcion


def publicar(nombre, **datos):
    """Publica un evento. Devuelve el registro (con id y hora) por si se quiere usar."""
    registro = {
        "id": next(_secuencia),
        "ts": dt.datetime.now().isoformat(sep=" ", timespec="seconds"),
        "evento": str(nombre),
    }
    for clave, valor in datos.items():
        limpio = _limpio(valor)
        if limpio is not None:
            registro[str(clave)] = limpio
    with _cerrojo:
        _eventos.append(registro)
        escuchas = list(_escuchas_para(registro["evento"]))
    for funcion in escuchas:
        try:
            funcion(registro)
        except Exception:      # noqa: BLE001 - un escucha roto no rompe la acción
            logger.excepcion(f"falló un escucha de eventos para {registro['evento']}")
    return registro


def ultimos(limite=100, solo_errores=False):
    with _cerrojo:
        registros = list(_eventos)
    if solo_errores:
        registros = [r for r in registros if "error" in r["evento"]
                     or r.get("nivel") in ("error", "ERROR")]
    return registros[-int(limite):] if limite else registros


def desde(numero_desde, limite=100):
    """Eventos posteriores a `numero_desde` (lo usa la pantalla para avisar de otros cambios)."""
    with _cerrojo:
        registros = [r for r in _eventos if r["id"] > int(numero_desde or 0)]
    return registros[-int(limite):] if limite else registros


def reiniciar():
    """Solo para pruebas: vacía el historial y los escuchas del registro."""
    global _secuencia
    with _cerrojo:
        _eventos.clear()
        _escuchas.clear()
        _escucha_log["hecho"] = False
        _secuencia = itertools.count(1)


def conectar_registro():
    """Escucha que escribe cada evento en el archivo de depuración (una sola vez)."""
    if _escucha_log["hecho"]:
        return
    escuchar("*", logger.registrar_evento)
    _escucha_log["hecho"] = True
