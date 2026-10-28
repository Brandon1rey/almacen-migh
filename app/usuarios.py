"""Usuarios y contraseñas: PBKDF2-HMAC-SHA256 con sal por usuario (solo biblioteca estándar)."""
import datetime as dt
import hashlib
import hmac
import re
import secrets

from app.bitacora import ROLES

ITERACIONES = 260_000
MINIMO_PASSWORD = 8
PATRON_NOMBRE = re.compile(r"^[A-Za-zÁÉÍÓÚáéíóúÑñ][A-Za-zÁÉÍÓÚáéíóúÑñ .'-]{2,39}$")


def _ahora():
    return dt.datetime.now().isoformat(sep=" ", timespec="seconds")


def _derivar(password, salt_bytes, iteraciones):
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt_bytes,
                               iteraciones).hex()


def validar(nombre, password, rol):
    if not PATRON_NOMBRE.match((nombre or "").strip()):
        raise ValueError("nombre inválido: entre 3 y 40 caracteres, sin números ni símbolos raros")
    if len(password or "") < MINIMO_PASSWORD:
        raise ValueError(f"la contraseña debe tener al menos {MINIMO_PASSWORD} caracteres")
    if rol not in ROLES:
        raise ValueError(f"rol inválido: usa uno de {', '.join(ROLES)}")


def crear(con, nombre, password, rol="almacen", debe_cambiar=False):
    nombre = (nombre or "").strip()
    validar(nombre, password, rol)
    salt = secrets.token_bytes(16)
    try:
        con.execute(
            "INSERT INTO usuario (nombre, rol, hash, salt, iteraciones, creado, debe_cambiar) "
            "VALUES (?,?,?,?,?,?,?)",
            (nombre, rol, _derivar(password, salt, ITERACIONES), salt.hex(), ITERACIONES,
             _ahora(), 1 if debe_cambiar else 0))
        con.commit()
    except Exception as e:
        raise ValueError(f"no se pudo crear el usuario {nombre}: {e}") from e
    return con.execute("SELECT * FROM usuario WHERE LOWER(nombre)=LOWER(?)", (nombre,)).fetchone()


def verificar(con, nombre, password):
    """Devuelve la fila del usuario si la contraseña es correcta, si no None."""
    fila = con.execute("SELECT * FROM usuario WHERE LOWER(nombre)=LOWER(?) AND activo=1",
                       ((nombre or "").strip(),)).fetchone()
    if fila is None:
        # mismo costo de cálculo aunque el usuario no exista (evita adivinar nombres)
        _derivar(password or "", b"0" * 16, ITERACIONES)
        return None
    candidato = _derivar(password or "", bytes.fromhex(fila["salt"]), fila["iteraciones"])
    if hmac.compare_digest(candidato, fila["hash"]):
        return fila
    return None


def cambiar_password(con, usuario_id, nueva, debe_cambiar=False):
    if len(nueva or "") < MINIMO_PASSWORD:
        raise ValueError(f"la contraseña debe tener al menos {MINIMO_PASSWORD} caracteres")
    salt = secrets.token_bytes(16)
    con.execute(
        "UPDATE usuario SET hash=?, salt=?, iteraciones=?, debe_cambiar=? WHERE id=?",
        (_derivar(nueva, salt, ITERACIONES), salt.hex(), ITERACIONES,
         1 if debe_cambiar else 0, int(usuario_id)))
    con.commit()
    # al cambiar la contraseña se cierran todas las sesiones de ese usuario
    con.execute("DELETE FROM sesion WHERE usuario_id=?", (int(usuario_id),))
    con.commit()


def listar(con):
    return [dict(r) for r in con.execute(
        "SELECT id, nombre, rol, activo, debe_cambiar, creado, ultimo_acceso "
        "FROM usuario ORDER BY LOWER(nombre)").fetchall()]


def por_nombre(con, nombre):
    return con.execute("SELECT * FROM usuario WHERE LOWER(nombre)=LOWER(?)",
                       ((nombre or "").strip(),)).fetchone()


def por_id(con, usuario_id):
    """Fila del usuario o None. La API la usa para responder 404 en vez de fingir éxito."""
    try:
        ident = int(usuario_id)
    except (TypeError, ValueError):
        return None
    return con.execute("SELECT * FROM usuario WHERE id=?", (ident,)).fetchone()


def cambiar_rol(con, usuario_id, rol):
    if rol not in ROLES:
        raise ValueError(f"rol inválido: {rol}")
    total_admin = con.execute(
        "SELECT COUNT(*) FROM usuario WHERE rol='admin' AND activo=1").fetchone()[0]
    actual = con.execute("SELECT rol FROM usuario WHERE id=?", (usuario_id,)).fetchone()
    if actual and actual["rol"] == "admin" and rol != "admin" and total_admin <= 1:
        raise ValueError("no puedes quitar el último administrador activo")
    con.execute("UPDATE usuario SET rol=? WHERE id=?", (rol, int(usuario_id)))
    con.commit()


def activar(con, usuario_id, activo):
    if not activo:
        total_admin = con.execute(
            "SELECT COUNT(*) FROM usuario WHERE rol='admin' AND activo=1").fetchone()[0]
        actual = con.execute("SELECT rol FROM usuario WHERE id=?", (usuario_id,)).fetchone()
        if actual and actual["rol"] == "admin" and total_admin <= 1:
            raise ValueError("no puedes desactivar el último administrador activo")
    con.execute("UPDATE usuario SET activo=? WHERE id=?", (1 if activo else 0, int(usuario_id)))
    con.execute("DELETE FROM sesion WHERE usuario_id=? AND ?=0", (int(usuario_id), 1 if activo else 0))
    con.commit()


def hay_usuarios(con):
    return con.execute("SELECT COUNT(*) FROM usuario").fetchone()[0] > 0


def bootstrap(con, password=None):
    """Si la base está vacía, crea el administrador. Devuelve (usuario, password, creada)."""
    if hay_usuarios(con):
        return None, None, False
    clave = password or secrets.token_urlsafe(9)
    crear(con, "admin", clave, rol="admin", debe_cambiar=True)
    return "admin", clave, True
