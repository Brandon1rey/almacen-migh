"""Alta y mantenimiento de usuarios desde la consola (sin abrir el navegador).

    python scripts\\crear_usuario.py --listar
    python scripts\\crear_usuario.py --nombre "Esteban Soto" --rol almacen
    python scripts\\crear_usuario.py --nombre "Galilea Medina" --rol almacen --password "Clave.2026"
    python scripts\\crear_usuario.py --reset "Esteban Soto"
"""
import argparse
import pathlib
import secrets
import sys

# La consola de Windows no siempre es UTF-8: sin esto, imprimir «→» o «·» lanza
# UnicodeEncodeError y el script parece fallar aunque ya haya hecho su trabajo.
for _flujo in (sys.stdout, sys.stderr):
    try:
        _flujo.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass


RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from app import bitacora, usuarios  # noqa: E402


def main(argv=None):
    ap = argparse.ArgumentParser(description="Usuarios del control de ubicación")
    ap.add_argument("--db", default=None, help="ruta de la base (por defecto data/bitacora.db)")
    ap.add_argument("--listar", action="store_true")
    ap.add_argument("--nombre")
    ap.add_argument("--rol", default="almacen", choices=list(bitacora.ROLES))
    ap.add_argument("--password", help="si se omite, se genera una")
    ap.add_argument("--reset", metavar="NOMBRE", help="genera contraseña nueva y la marca provisional")
    args = ap.parse_args(argv)

    con = bitacora.conectar(args.db)

    if args.listar:
        filas = usuarios.listar(con)
        if not filas:
            print("No hay usuarios todavía. Se crea el administrador al arrancar el servidor,")
            print("o con:  python scripts\\crear_usuario.py --nombre admin --rol admin")
            return 0
        print(f"{'USUARIO':<26}{'ROL':<10}{'ACTIVO':<8}{'ÚLTIMO ACCESO'}")
        for u in filas:
            print(f"{u['nombre']:<26}{u['rol']:<10}{'sí' if u['activo'] else 'no':<8}"
                  f"{u['ultimo_acceso'] or '—'}")
        return 0

    if args.reset:
        fila = usuarios.por_nombre(con, args.reset)
        if fila is None:
            print(f"No existe el usuario {args.reset}")
            return 1
        clave = args.password or secrets.token_urlsafe(9)
        usuarios.cambiar_password(con, fila["id"], clave, debe_cambiar=True)
        bitacora.auditar(con, "consola", "RESET_PASSWORD", args.reset)
        print(f"Contraseña nueva para {args.reset}: {clave}")
        print("Entra con ella y cámbiala de inmediato.")
        return 0

    if not args.nombre:
        ap.print_help()
        return 1

    clave = args.password or secrets.token_urlsafe(9)
    try:
        usuarios.crear(con, args.nombre, clave, args.rol)
    except ValueError as e:
        print(f"ERROR: {e}")
        return 1
    bitacora.auditar(con, "consola", "USUARIO_CREADO", {"nombre": args.nombre, "rol": args.rol})
    print(f"Usuario creado: {args.nombre}  ({args.rol})")
    print(f"Contraseña     : {clave}")
    print("Anótala y entrégala en mano; al entrar puede cambiarla.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
