"""Puebla el demo con material de empaque y producto terminado (determinista).

    python scripts\\sembrar_demo_mep_pt.py

Escribe data\\demo\\Inventario_MEP.xlsx, data\\demo\\Inventario_PT.xlsx y los movimientos
correspondientes en data\\demo\\bitacora_demo.db. Es ficticio: no usar como registro.
"""
import datetime as dt
import pathlib
import random
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

RAIZ = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from app import bitacora, inventario, inventarios  # noqa: E402

DEMO = RAIZ / "data" / "demo"
HOY = dt.date(2026, 10, 22)
SEMILLA = 7

ENCABEZADOS = ["Descripción", "Proveedor", "Categoría", "Unidad de medida", "Cantidad",
               "Caducidad vigente", "Lote actual", "Entrada ERP", "Observaciones", "UBICACIÓN"]
ANCHOS = [46, 24, 14, 10, 11, 16, 18, 14, 30, 16]
AZUL = PatternFill("solid", fgColor="1F3B73")
BORDE = Border(*[Side(style="thin", color="B7C0CC")] * 4)

# (descripción, proveedor, categoría, unidad, cantidad, caducidad, lote, ERP, nota)
MEP = [
    ("BOLSA LAMINADA 400 G", "IMPRESOS DEL NORTE", "Bolsa", "pza", 120000, "2029-03-31", "BL400-2601", "MIGH-ME001", ""),
    ("BOLSA LAMINADA 400 G", "IMPRESOS DEL NORTE", "Bolsa", "pza", 48000, "2029-03-31", "BL400-2602", "MIGH-ME001", ""),
    ("BOLSA LAMINADA 900 G", "IMPRESOS DEL NORTE", "Bolsa", "pza", 96000, "2029-05-30", "BL900-2601", "MIGH-ME002", ""),
    ("BOLSA LAMINADA 2 KG", "IMPRESOS DEL NORTE", "Bolsa", "pza", 24000, "2029-01-31", "BL2K-2601", "MIGH-ME003", "caja de 2,000 pzas"),
    ("BOLSA DOY PACK 500 G", "FLEXIBLES BAJIO", "Bolsa", "pza", 30000, "2028-11-30", "BDP500-2601", "MIGH-ME004", ""),
    ("BOTE P/POLVO 900 G", "PLASTICOS SUR", "Bote", "pza", 36000, "2030-06-30", "BO900-2601", "MIGH-ME010", ""),
    ("BOTE P/POLVO 900 G", "PLASTICOS SUR", "Bote", "pza", 12000, "2030-06-30", "BO900-2602", "MIGH-ME010", ""),
    ("BOTE P/POLVO 400 G", "PLASTICOS SUR", "Bote", "pza", 18000, "2030-04-30", "BO400-2601", "MIGH-ME011", ""),
    ("TAPA ROSCA 43 MM BLANCA", "PLASTICOS SUR", "Tapa", "pza", 50000, "2030-08-31", "TR43-B-2601", "MIGH-ME020", ""),
    ("TAPA ROSCA 43 MM AZUL", "PLASTICOS SUR", "Tapa", "pza", 26000, "2030-08-31", "TR43-A-2601", "MIGH-ME021", ""),
    ("TAPA FLIP-TOP 38 MM", "TAPAS DEL CENTRO", "Tapa", "pza", 15000, "2029-12-31", "TF38-2601", "MIGH-ME022", ""),
    ("CUCHARA DOSIFICADORA 5 G", "INYECTADOS MEX", "Cuchara", "pza", 200000, "2031-01-31", "CD5-2601", "MIGH-ME030", ""),
    ("CUCHARA DOSIFICADORA 5 G", "INYECTADOS MEX", "Cuchara", "pza", 75000, "2031-01-31", "CD5-2602", "MIGH-ME030", ""),
    ("ETIQUETA FRONTAL NUTRILIFE VAINILLA", "ETIQUETAS FINAS", "Etiqueta", "pza", 80000, "2028-06-30", "EF-NV-2601", "MIGH-ME040", ""),
    ("ETIQUETA TRASERA NUTRILIFE VAINILLA", "ETIQUETAS FINAS", "Etiqueta", "pza", 78000, "2028-06-30", "ET-NV-2601", "MIGH-ME041", ""),
    ("ETIQUETA FRONTAL NUTRILIFE CHOCOLATE", "ETIQUETAS FINAS", "Etiqueta", "pza", 40000, "2028-04-30", "EF-NC-2601", "MIGH-ME042", ""),
    ("ETIQUETA FRONTAL ISOLATE WHEY", "ETIQUETAS FINAS", "Etiqueta", "pza", 22000, "2028-02-29", "EF-IW-2601", "MIGH-ME043", ""),
    ("CAJA COLECTIVA 12 PZAS 900 G", "CARTONES DEL BAJIO", "Caja", "pza", 9000, "2030-01-31", "CC12-2601", "MIGH-ME050", ""),
    ("CAJA COLECTIVA 24 PZAS 400 G", "CARTONES DEL BAJIO", "Caja", "pza", 6500, "2030-01-31", "CC24-2601", "MIGH-ME051", ""),
    ("SEPARADOR DE CARTON 900 G", "CARTONES DEL BAJIO", "Caja", "pza", 15000, "2030-03-31", "SEP900-2601", "MIGH-ME052", ""),
    ("FILM TERMOENCOGIBLE 50 CM", "ENVOLVENTES MX", "Film", "kg", 850.5, "2028-09-30", "FT50-2601", "MIGH-ME060", ""),
    ("CINTA CANELA IMPRESA", "ENVOLVENTES MX", "Cinta", "pza", 2400, "2029-07-31", "CIN-2601", "MIGH-ME070", ""),
    ("BOLSA LAMINADA 900 G", "IMPRESOS DEL NORTE", "Bolsa", "pza", 30000, "2027-09-30", "BL900-2507", "MIGH-ME002", "lote viejo, ya no se usa"),
    ("ETIQUETA FRONTAL NUTRILIFE FRESA", "ETIQUETAS FINAS", "Etiqueta", "pza", 0, "2027-12-31", "EF-NF-2509", "MIGH-ME044", "lote agotado"),
]

PT = [
    ("NUTRILIFE VAINILLA 900 G", "MIGH", "Polvo", "pza", 1200, "2028-02-15", "PT-2601-A", "MIGH-PT001", "caja colectiva 12 pzas"),
    ("NUTRILIFE VAINILLA 900 G", "MIGH", "Polvo", "pza", 480, "2028-03-10", "PT-2601-B", "MIGH-PT001", "caja colectiva 12 pzas"),
    ("NUTRILIFE CHOCOLATE 900 G", "MIGH", "Polvo", "pza", 960, "2028-01-31", "PT-2601-C", "MIGH-PT002", ""),
    ("NUTRILIFE FRESA 900 G", "MIGH", "Polvo", "pza", 600, "2027-12-20", "PT-2512-A", "MIGH-PT003", ""),
    ("NUTRILIFE VAINILLA 400 G", "MIGH", "Polvo", "pza", 2400, "2028-04-30", "PT-2602-A", "MIGH-PT004", ""),
    ("NUTRILIFE CHOCOLATE 400 G", "MIGH", "Polvo", "pza", 1800, "2028-05-31", "PT-2602-B", "MIGH-PT005", ""),
    ("NUTRILIFE PRO VAINILLA 900 G", "MIGH", "Polvo", "pza", 720, "2028-06-30", "PT-2603-A", "MIGH-PT006", ""),
    ("ISOLATE WHEY 2 KG", "MIGH", "Proteína", "pza", 300, "2028-07-31", "PT-2604-A", "MIGH-PT007", "bolsa con sello"),
    ("WHEY CONCENTRATE 2 KG", "MIGH", "Proteína", "pza", 240, "2028-07-31", "PT-2604-B", "MIGH-PT008", ""),
    ("COLÁGENO HIDROLIZADO 500 G", "MIGH", "Suplemento", "pza", 900, "2028-09-30", "PT-2605-A", "MIGH-PT009", ""),
    ("MULTIVITAMÍNICO 60 CAPS", "MIGH", "Suplemento", "pza", 1500, "2028-10-31", "PT-2606-A", "MIGH-PT010", ""),
    ("NUTRILIFE VAINILLA 900 G", "MIGH", "Polvo", "pza", 36, "2026-08-31", "PT-2506-X", "MIGH-PT001", "vencido, retenido"),
    ("NUTRILIFE CHOCOLATE 400 G", "MIGH", "Polvo", "pza", 12, "2026-11-30", "PT-2511-Y", "MIGH-PT005", "cuarentena"),
]


def _fecha(fecha, rng, dias):
    return (dt.date.fromisoformat(fecha) - dt.timedelta(days=rng.randint(0, dias))).isoformat()


def escribir(tipo, filas, semilla):
    rng = random.Random(semilla)
    destino = DEMO / f"Inventario_{tipo.upper()}.xlsx"
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inventario"
    ws["A1"] = f"INVENTARIO DE {inventarios.TIPOS[tipo]['nombre'].upper()} — DEMO (DATOS FICTICIOS)"
    ws["A1"].font = Font(bold=True, size=12, color="B01C1C")
    ws["A2"] = "Generado por scripts\\sembrar_demo_mep_pt.py — no es un registro de calidad"
    for i, texto in enumerate(ENCABEZADOS, start=1):
        c = ws.cell(row=3, column=i, value=texto)
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = AZUL
        c.border = BORDE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = ANCHOS[i - 1]
    ws.row_dimensions[3].height = 28

    registros = []
    for n, (desc, prov, cat, unidad, cantidad, caducidad, lote, erp, nota) in enumerate(filas, start=4):
        valores = [desc, prov, cat, unidad, cantidad, caducidad, lote, erp, nota, ""]
        for i, valor in enumerate(valores, start=1):
            c = ws.cell(row=n, column=i, value=valor)
            c.font = Font(size=9)
            c.border = BORDE
        registros.append({"fila": n, "descripcion": desc, "lote": lote,
                          "cantidad": cantidad, "unidad": unidad, "tipo": tipo})
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:J{max(ws.max_row, 4)}"
    wb.save(destino)
    return destino, registros


def ubicaciones_para(tipo, n, rng):
    """Asigna racks plausibles: empaque en MEP-B..E, producto terminado en MPT-B."""
    if tipo == "mep":
        bloque = rng.choice([("MEP-B", 1, 20), ("MEP-C", 21, 100), ("MEP-D", 101, 180),
                             ("MEP-E", 181, 220)])
    else:
        bloque = ("MPT-B", 1, 20)
    return f"{bloque[0]}-{rng.randint(bloque[1], bloque[2]):03d}"


def sembrar_movimientos(con, tipo, registros, rng):
    total = 0
    for r in registros:
        if r["cantidad"] in (0, None):
            continue
        destino = ubicaciones_para(tipo, r["fila"], rng)
        if rng.random() < 0.12:                     # algunos quedan pendientes de ubicar
            continue
        ts = dt.datetime(2026, 9, 1) + dt.timedelta(days=rng.randint(0, 50),
                                                    minutes=rng.randint(0, 600))
        _registrar_en(con, "UBICAR", r, destino, ts, tipo, rng, motivo="recepcion")
        total += 1
        if rng.random() < 0.22:                     # reacomodos
            nuevo = ubicaciones_para(tipo, r["fila"], rng)
            if nuevo != destino:
                ts2 = ts + dt.timedelta(days=rng.randint(1, 20), minutes=rng.randint(0, 300))
                _registrar_en(con, "MOVER", r, nuevo, ts2, tipo, rng, origen=destino,
                              motivo="reacomodo")
                destino = nuevo
                total += 1
        if rng.random() < 0.10 and isinstance(r["cantidad"], (int, float)) and r["cantidad"] > 100:
            parte = round(r["cantidad"] * rng.choice([0.05, 0.1, 0.2]), 0)
            ts3 = ts + dt.timedelta(days=rng.randint(5, 30), minutes=rng.randint(0, 300))
            _registrar_en(con, "PARCIAL", r, ubicaciones_para(tipo, r["fila"], rng), ts3, tipo,
                          rng, origen=destino, motivo="surtido", cantidad=parte)
            total += 1
    return total


def _registrar_en(con, accion, registro, destino, ts, tipo, rng, origen=None, motivo=None,
                  cantidad=None):
    """Inserta con fecha controlada (la bitácora no acepta ts en registrar: se ajusta después)."""
    _, mov_id = bitacora.registrar(
        con, accion, registro, rng.choice(["Esteban", "Galilea"]), destino=destino,
        origen=origen, motivo=motivo, cantidad=cantidad if cantidad is not None
        else registro.get("cantidad"), inventario=tipo)
    con.execute("UPDATE movimiento SET ts=? WHERE id=?", (ts.isoformat(sep=" ", timespec="seconds"), mov_id))
    con.commit()
    return mov_id


def aplicar_ubicaciones(tipo, registros, con):
    """Deja en el Excel la última ubicación de cada lote (como si se hubiera capturado en la app)."""
    inv = inventario.Inventario(DEMO / f"Inventario_{tipo.upper()}.xlsx", tipo=tipo)
    ultimas = {}
    for m in bitacora.ultimos(con, 10000, inventario=tipo):
        if m["descripcion"] not in ultimas and m["ubicacion_destino"]:
            ultimas[(m["descripcion"], m["lote"])] = m["ubicacion_destino"]
    for f in inv.filas:
        destino = ultimas.get((f["descripcion"], f["lote"]))
        if destino:
            inv.asignar(f["fila"], destino)
    inv.guardar()
    return inv.resumen()


def main():
    rng = random.Random(SEMILLA)
    con = bitacora.conectar(DEMO / "bitacora_demo.db")

    resumenes = {}
    for tipo, filas, semilla in (("mep", MEP, SEMILLA), ("pt", PT, SEMILLA + 1)):
        _, registros = escribir(tipo, filas, semilla)
        movimientos = sembrar_movimientos(con, tipo, registros, rng)
        r = aplicar_ubicaciones(tipo, registros, con)
        resumenes[tipo] = (r, movimientos)

    print("DEMO ampliado con empaque y producto terminado")
    for tipo, (r, movs) in resumenes.items():
        print(f"  {tipo.upper():<3} {r['nombre']:<20} {r['total']:>4} lotes · "
              f"{r['con_ubicacion']:>4} con ubicación ({r['cobertura']} %) · {movs} movimientos")
    print(f"  archivos: {DEMO}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
