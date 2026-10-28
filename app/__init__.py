"""Inventario de almacén MIGH: ubicación por tarima sobre el Excel que ya usa el área."""
import pathlib as _pathlib
import sys as _sys


def _habilitar_vendor():
    """Deja importable la carpeta vendor\\ (PC sin internet ni pip).

    Se agrega al final del sys.path: si la máquina ya tiene openpyxl instalado con pip,
    gana esa versión y vendor solo es el respaldo. Sin esto, cualquier script o prueba
    que importe `app` fallaba con ModuleNotFoundError aunque vendor\\ estuviera completo.
    """
    _vendor = _pathlib.Path(__file__).resolve().parents[1] / "vendor"
    if _vendor.is_dir() and str(_vendor) not in _sys.path:
        _sys.path.append(str(_vendor))


_habilitar_vendor()
