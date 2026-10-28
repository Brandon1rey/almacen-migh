"""Punto de entrada Serverless para Vercel (`@vercel/python`).

Vercel detecta la clase `handler(BaseHTTPRequestHandler)` y despacha cada
petición HTTP a través del mismo `Handler` de `app.server`. Cuando está
definida la variable `SUPABASE_DB_URL` (o `DATABASE_URL` / `POSTGRES_URL`),
todo el estado (usuarios, sesiones, inventarios MP/MEP/PT, catálogo, bitácora y
auditoría) se lee y escribe directamente en Supabase PostgreSQL.
"""
from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from app.server import Handler, asegurar_estado  # noqa: E402

# Inicialización perezosa en el arranque en frío (cold start) del contenedor
asegurar_estado()


class handler(Handler):
    """Handler compatible con el runtime Python de Vercel."""
    pass
