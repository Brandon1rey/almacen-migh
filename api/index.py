"""Punto de entrada Serverless para Vercel (`@vercel/python`).

Vercel detecta la clase `handler(BaseHTTPRequestHandler)` y despacha cada
petición HTTP a través del `Handler` de `app.server`. La conexión a Supabase
se inicializa bajo demanda en la primera petición de API/descarga, de modo que
los archivos estáticos (`/`, `/html5-qrcode.min.js`) cargan de inmediato sin
bloquear ni fallar en el arranque en frío.
"""
from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from app.server import Handler  # noqa: E402


class handler(Handler):
    """Handler compatible con el runtime Python de Vercel."""
    pass
