"""Consolidado 2026: todo por propiedad, por mes y por tipo de gasto.

Lee las planillas canónicas del repo (CIUDAD, OHIGGINS, BONORINO, AVA, SOL)
y genera CONSOLIDADO/Consolidado_2026.xlsx con:
  - Consolidado: Propiedad · Tipo · Gasto × Ene..Dic + Total
  - Faltantes: servicios sin importe en un mes ya cerrado
  - Liquidación Chris: CIUDAD 50/50 por mes
Uso: python3 tools/consolidar_2026.py
"""
import math
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MES_CERRADO = 10  # último mes cerrado (octubre); se controla faltantes hasta acá
OUT = ROOT / "CONSOLIDADO/Consolidado_2026.xlsx"

NAVY = PatternFill("solid", fgColor="17365D")
GREY = PatternFill("solid", fgColor="D9E1F2")
YEL = PatternFill("solid", fgColor="FFF2CC")
WHITE = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)
NUM = "#,##0"


def mes_idx(v):
    if v is None:
        return None
    s = str(v).strip().capitalize()
    return MESES.index(s) if s in MESES else None


def matriz(path, sheet="2026", first_row=5, last_row=20, skip=()):
    """Filas de servicio de una matriz Tipo/Categoría/Descripción × meses."""
    ws = load_workbook(path)[sheet]
    out = []
    for r in range(first_row, last_row + 1):
        desc = ws.cell(r, 3).value
        if not desc or any(s in str(desc) for s in skip):
            continue
        vals = []
        for c in range(4, 16):
            v = ws.cell(r, c).value
            vals.append(math.ceil(v) if isinstance(v, (int, float)) else None)
        out.append((ws.cell(r, 2).value or ws.cell(r, 1).value, desc, vals))
    return out


def almacen_por_tipo(path):
    ws = load_workbook(path)["Almacen Detalle"]
    sol, chris = defaultdict(lambda: [0] * 12), defaultdict(lambda: [0] * 12)
    for r in ws.iter_rows(min_row=4, values_only=True):
        m = mes_idx(r[0])
        if m is None:
            continue
        tipo = str(r[2] or "OTROS").split("·")[0].split("—")[0].strip().upper() or "OTROS"
        if tipo.startswith("ALMAC"):
            tipo = "SIN DETALLE"
        if isinstance(r[3], (int, float)):
            sol[tipo][m] += r[3]
        if isinstance(r[4], (int, float)):
            chris[tipo][m] += r[4]
    return sol, chris


def gps_por_categoria(path):
    ws = load_workbook(path)["Gasto Personal Sol"]
    cat = defaultdict(lambda: [0] * 12)
    for r in ws.iter_rows(min_row=3, values_only=True):
        fecha, mes, categoria, monto = r[0], r[1], r[2], r[3]
        if not hasattr(fecha, "year") or fecha.year != 2026 or not isinstance(monto, (int, float)):
            continue
        if "PROYECTADO" in str(r[5] or "").upper():
            continue
        cat[str(categoria or "SIN CATEGORÍA").split("·")[0].strip()][fecha.month - 1] += monto
    return cat


def main():
    rows = []  # (propiedad, tipo, gasto, [12])
    ciudad = ROOT / "CIUDAD/Gastos_CIUDAD_1132_2026.xlsx"
    serv_ciudad = [0] * 12
    for cat, desc, vals in matriz(ciudad, skip=("ALMACEN",)):
        if any(vals):
            rows.append(("CIUDAD 1132", "SERVICIOS · " + str(cat), desc, vals))
            serv_ciudad = [a + (b or 0) for a, b in zip(serv_ciudad, vals)]
    sol, chris = almacen_por_tipo(ciudad)
    for tipo in sorted(sol):
        if any(sol[tipo]):
            rows.append(("CIUDAD 1132", "ALMACÉN SOL", tipo, sol[tipo]))
    for tipo in sorted(chris):
        if any(chris[tipo]):
            rows.append(("CIUDAD 1132", "ALMACÉN CHRIS", tipo, chris[tipo]))
    for prop, path in (("OHIGGINS", "OHIGGINS/Gastos_OHIGGINS_2026.xlsx"),
                       ("BONORINO", "BONORINO/Gastos_BONORINO_2026.xlsx"),
                       ("AVA", "AVA/Gastos_AVA_2026.xlsx")):
        for cat, desc, vals in matriz(ROOT / path, first_row=5 if prop != "AVA" else 5):
            if any(vals):
                rows.append((prop, "SERVICIOS · " + str(cat), desc, vals))
    for cat, vals in sorted(gps_por_categoria(ciudad).items()):
        if any(vals):
            rows.append(("SOL · personal", "GASTO PERSONAL", cat, vals))

    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado"
    ws["A1"] = "Consolidado 2026 — por propiedad, mes y tipo de gasto"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "Generado por tools/consolidar_2026.py desde las planillas canónicas. Montos ARS enteros."
    hdr = ["Propiedad", "Tipo", "Gasto"] + [m[:3] for m in MESES] + ["Total"]
    ws.append([])
    ws.append(hdr)
    for c in range(1, len(hdr) + 1):
        ws.cell(4, c).fill = NAVY
        ws.cell(4, c).font = WHITE
    prop_actual, inicio = None, 5

    def subtotal(prop, desde, hasta):
        r = ws.max_row + 1
        ws.cell(r, 1, f"TOTAL {prop}").font = BOLD
        for c in range(4, 17):
            L = get_column_letter(c)
            ws.cell(r, c, f"=SUM({L}{desde}:{L}{hasta})").font = BOLD
            ws.cell(r, c).number_format = NUM
        for c in range(1, 17):
            ws.cell(r, c).fill = GREY
        ws.append([])

    for prop, tipo, gasto, vals in rows:
        if prop != prop_actual:
            if prop_actual:
                subtotal(prop_actual, inicio, ws.max_row)
            prop_actual, inicio = prop, ws.max_row + 1
        r = ws.max_row + 1
        ws.append([prop, tipo, gasto] + [v or None for v in vals] + [f"=SUM(D{r}:O{r})"])
        for c in range(4, 17):
            ws.cell(r, c).number_format = NUM
    subtotal(prop_actual, inicio, ws.max_row)
    ws.freeze_panes = "D5"
    for col, w in zip("ABC", (16, 26, 46)):
        ws.column_dimensions[col].width = w
    for c in range(4, 17):
        ws.column_dimensions[get_column_letter(c)].width = 12

    # Faltantes: servicios sin importe en meses cerrados
    fal = wb.create_sheet("Faltantes")
    fal.append(["Servicios sin importe en un mes cerrado (hasta " + MESES[MES_CERRADO - 1] + ")"])
    fal.append(["Propiedad", "Gasto", "Mes", "Qué hacer"])
    for c in range(1, 5):
        fal.cell(2, c).fill = NAVY
        fal.cell(2, c).font = WHITE
    for prop, tipo, gasto, vals in rows:
        if not tipo.startswith("SERVICIOS"):
            continue
        primero = next((i for i, v in enumerate(vals) if v), None)
        if primero is None:
            continue
        for i in range(primero, MES_CERRADO):
            if not vals[i]:
                fal.append([prop, gasto, MESES[i], "Buscar pago (extracto MP / BBVA / Gmail) o confirmar que no corresponde"])
    for col, w in zip("ABCD", (16, 46, 12, 70)):
        fal.column_dimensions[col].width = w

    # Liquidación Chris (CIUDAD 50/50)
    lq = wb.create_sheet("Liquidación Chris")
    lq.append(["Liquidación CIUDAD 50/50 — AVA, OHIGGINS y BONORINO no entran. Ene–abr: se registra, no se cobra."])
    h = ["Mes", "Servicios CIUDAD", "Almacén Sol", "Almacén Chris", "50% servicios", "50% almacén",
         "A cargo Chris", "Pagó Chris (almacén)", "Chris debe a Sol", "Estado"]
    lq.append(h)
    for c in range(1, len(h) + 1):
        lq.cell(2, c).fill = NAVY
        lq.cell(2, c).font = WHITE
    for i, mes in enumerate(MESES[:11]):
        s = serv_ciudad[i]
        a_sol = sum(v[i] for v in sol.values())
        a_ch = sum(v[i] for v in chris.values())
        cs, ca = math.ceil(s / 2), math.ceil((a_sol + a_ch) / 2)
        debe = cs + ca - a_ch
        estado = "CALCULADO · NO COBRAR" if i < 4 else "A COBRAR"
        lq.append([mes, s, a_sol, a_ch, cs, ca, cs + ca, a_ch, debe, estado])
        for c in range(2, 10):
            lq.cell(lq.max_row, c).number_format = NUM
    for col in "ABCDEFGHIJ":
        lq.column_dimensions[col].width = 18

    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print("OK", OUT, "filas", len(rows), "faltantes", fal.max_row - 2)


if __name__ == "__main__":
    main()
