"""Lectura de fechas del Excel del área: nunca se inventa una fecha."""
import datetime as dt
import re

MESES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
         "jul": 7, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12}


def parsear(valor):
    """Devuelve (iso | None, motivo). Acepta solo lo que no es ambiguo."""
    if valor in (None, "", "N/A", "n/a", "NA", "-"):
        return None, "vacía"
    if isinstance(valor, dt.datetime):
        return valor.date().isoformat(), None
    if isinstance(valor, dt.date):
        return valor.isoformat(), None
    texto = str(valor).strip()

    # 2027/agos/16  ·  MAY/08/2027
    m = re.fullmatch(r"(\d{4})/([A-Za-zÁÉÍÓÚáéíóú]{3,})/(\d{1,2})", texto)
    if m and m.group(2)[:3].lower() in MESES:
        return _armar(int(m.group(1)), MESES[m.group(2)[:3].lower()], int(m.group(3)))
    m = re.fullmatch(r"([A-Za-zÁÉÍÓÚáéíóú]{3,})/(\d{1,2})/(\d{4})", texto)
    if m and m.group(1)[:3].lower() in MESES:
        return _armar(int(m.group(3)), MESES[m.group(1)[:3].lower()], int(m.group(2)))

    m = re.fullmatch(r"(\d{4})/(\d{2})/(\d{4})", texto)
    if m:
        return None, f"año repetido: '{texto}'"

    m = re.fullmatch(r"(\d{1,4})[/-](\d{1,2})[/-](\d{1,4})", texto)
    if m:
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 1900:
            # año primero: es ISO (AAAA-MM-DD), no hay ambigüedad que resolver
            if b > 12 and c <= 12:
                return _armar(a, c, b)
            return _armar(a, b, c)
        if c > 1900:
            anio, x, y = c, a, b
        else:
            return None, f"año ilegible: '{texto}'"
        if x > 12 and y <= 12:
            return _armar(anio, y, x)
        if y > 12 and x <= 12:
            return _armar(anio, x, y)
        return None, f"fecha ambigua día/mes: '{texto}'"

    return None, f"formato no reconocido: '{texto}'"


def _armar(anio, mes, dia):
    try:
        return dt.date(anio, mes, dia).isoformat(), None
    except ValueError:
        return None, f"fecha imposible: {anio}-{mes}-{dia}"


def semaforo(iso, hoy=None):
    """(dias_restantes | None, 'VENCIDO'|'CRITICO'|'AVISO'|'OK'|'SIN DATO')."""
    if not iso:
        return None, "SIN DATO"
    if isinstance(hoy, str):
        hoy = dt.date.fromisoformat(hoy)
    hoy = hoy or dt.date.today()
    dias = (dt.date.fromisoformat(iso) - hoy).days
    if dias < 0:
        return dias, "VENCIDO"
    if dias <= 30:
        return dias, "CRITICO"
    if dias <= 90:
        return dias, "AVISO"
    return dias, "OK"
