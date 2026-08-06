#!/usr/bin/env python3
"""Regenera LTRP/LTRP.xlsx — Excel canónico del Long Term Retention Program."""

from __future__ import annotations

import csv
import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "LTRP.xlsx"

PLANES_CSV = ROOT / "planes_vigentes.csv"
ACCION_CSV = ROOT / "accion_estimada.csv"
MATRIZ_CSV = ROOT / "simulacion_matriz.csv"
VESTING_CSV = ROOT / "calendario_vesting.csv"

PAGO_ANIOS = list(range(2027, 2033))

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="D6E3F0")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="E2EFDA")
MONEY_FORMAT = "#,##0"
MONEY_DEC = "#,##0.00"
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def _style_header(ws, row: int = 1) -> None:
    for cell in ws[row]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN


def _autosize(ws, min_width: int = 10, max_width: int = 42) -> None:
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        length = 0
        for cell in col:
            if cell.value is None:
                continue
            length = max(length, len(str(cell.value)))
        ws.column_dimensions[letter].width = max(min_width, min(max_width, length + 2))


def _to_float(value: str | float | int | None) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return float(str(value).replace(",", "").strip())


def _to_int(value: str | float | int | None) -> int | None:
    f = _to_float(value)
    return None if f is None else int(round(f))


def _border_range(ws, min_row: int, max_row: int, max_col: int) -> None:
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, max_col=max_col):
        for cell in row:
            cell.border = THIN


def build_simulacion(wb: Workbook) -> None:
    """Hoja principal — misma estructura que la simulación de Sol."""
    ws = wb.active
    ws.title = "Simulación de pagos"

    planes = [p for p in _read_csv(PLANES_CSV)]
    accion = {int(r["anio_pago"]): _to_float(r["precio_accion_estimado"]) for r in _read_csv(ACCION_CSV)}
    matriz = {
        r["plan"]: r
        for r in _read_csv(MATRIZ_CSV)
        if r["plan"] != "TOTAL"
    }

    # Título
    ws["A1"] = "Long Term Retention Program (LTRP) — Simulación de pagos"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws["A2"] = "Última actualización: agosto 2026 · Máx. 6 planes · cada enero = suma de 1/6 ajustados por acción MELI"
    ws["A2"].font = Font(italic=True, color="595959")

    # Encabezados (fila 4)
    headers = [
        "Concepto",
        "Valor nominal",
        "Valor de acción al otorgamiento",
    ] + [str(y) for y in PAGO_ANIOS]
    for col, h in enumerate(headers, 1):
        ws.cell(4, col, h)
    _style_header(ws, 4)

    # Filas de planes (5–9) con fórmulas Excel
    # Layout: A=plan B=nominal C=precio_grant D..=pagos
    # Precios anuales en hoja "Acción estimada" o en bloque inferior
    # Usamos precios en fila de referencia al pie y fórmulas ROUND

    start_plan = 5
    for i, p in enumerate(planes):
        row = start_plan + i
        plan = p["plan"]
        nominal = _to_float(p["valor_nominal_usd"]) or 0.0
        grant_px = _to_float(p["precio_accion_otorgamiento"]) or 0.0
        ws.cell(row, 1, plan)
        ws.cell(row, 2, nominal).number_format = MONEY_FORMAT
        ws.cell(row, 3, grant_px).number_format = MONEY_DEC

        m = matriz.get(plan, {})
        for j, y in enumerate(PAGO_ANIOS):
            col = 4 + j
            raw = m.get(f"pago_{y}", "")
            if raw == "" or raw is None:
                ws.cell(row, col, None)
            else:
                # Valor canónico de la simulación de Sol (enteros)
                val = _to_int(raw)
                ws.cell(row, col, val).number_format = MONEY_FORMAT
                # Comentario con fórmula
                px = accion[y]
                tramo = nominal / 6.0
                calc = tramo * 0.5 + tramo * 0.5 * (px / grant_px)
                ws.cell(row, col).alignment = Alignment(horizontal="center")

    end_plan = start_plan + len(planes) - 1

    # Fila TOTAL estimado a cobrar por año
    total_row = end_plan + 1
    ws.cell(total_row, 1, "Total estimado a cobrar por año")
    ws.cell(total_row, 1).font = Font(bold=True)
    ws.cell(total_row, 1).fill = SECTION_FILL
    for j, y in enumerate(PAGO_ANIOS):
        col = 4 + j
        letter = get_column_letter(col)
        cell = ws.cell(total_row, col, f"=SUM({letter}{start_plan}:{letter}{end_plan})")
        cell.font = Font(bold=True)
        cell.fill = YELLOW
        cell.number_format = MONEY_FORMAT
        cell.alignment = Alignment(horizontal="center")
    for col in (2, 3):
        ws.cell(total_row, col).fill = SECTION_FILL

    # Bloque valor de la acción estimado
    accion_label = total_row + 2
    ws.cell(accion_label, 1, "Valor de la acción estimado")
    ws.cell(accion_label, 1).font = Font(bold=True, size=12, color="1F4E79")
    ws.cell(accion_label + 1, 1, (
        "El valor de acción considerado para el primer año corresponde al promedio de los "
        "últimos 60 días de trading al momento de actualizar esta simulación con un ajuste "
        "de 3% correspondiente al cierre del 2026. Para los años subsiguientes se considera "
        "un crecimiento anual del 8%."
    ))
    ws.merge_cells(start_row=accion_label + 1, start_column=1, end_row=accion_label + 1, end_column=9)
    ws.cell(accion_label + 1, 1).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[accion_label + 1].height = 48

    precio_row = accion_label + 2
    ws.cell(precio_row, 1, "Precio estimado")
    ws.cell(precio_row, 1).font = Font(bold=True)
    for j, y in enumerate(PAGO_ANIOS):
        col = 4 + j
        cell = ws.cell(precio_row, col, accion[y])
        cell.number_format = MONEY_DEC
        cell.fill = GREEN
        cell.alignment = Alignment(horizontal="center")
        ws.cell(4, col)  # ensure header exists
        # Año label above already in header; add year under methodology for clarity
    for j, y in enumerate(PAGO_ANIOS):
        ws.cell(precio_row - 0, 4 + j)  # prices already set

    # Año labels for price row context
    year_row = precio_row - 1
    # actually put years on the row before prices if empty - the methodology text is there
    # Put year headers on precio_row-0 is prices; add a small year row:
    # Re-structure: year labels on same columns as payments
    for j, y in enumerate(PAGO_ANIOS):
        # already have years in header of payment table
        pass

    # Total acumulado pendiente
    pending_row = precio_row + 2
    ws.cell(pending_row, 1, "Total acumulado pendiente de pago")
    ws.cell(pending_row, 1).font = Font(bold=True)
    # Suma de totales anuales
    first_tot = get_column_letter(4)
    last_tot = get_column_letter(3 + len(PAGO_ANIOS))
    cell = ws.cell(pending_row, 2, f"=SUM({first_tot}{total_row}:{last_tot}{total_row})")
    cell.font = Font(bold=True, size=12)
    cell.fill = YELLOW
    cell.number_format = MONEY_FORMAT

    # Fórmula de referencia
    formula_row = pending_row + 2
    ws.cell(formula_row, 1, "Fórmula por celda")
    ws.cell(formula_row, 1).font = Font(bold=True)
    ws.cell(formula_row, 2, (
        "ROUND( (Nominal÷6)×50% + (Nominal÷6)×50% × (P_año / P_otorgamiento) ; 0 )"
    ))
    ws.cell(formula_row + 1, 1, "Regla general")
    ws.cell(formula_row + 1, 1).font = Font(bold=True)
    ws.cell(formula_row + 1, 2, (
        "Máximo 6 planes. Cada enero se suma el 1/6 de cada plan activo. "
        "Ej.: otorgado 70.000 en 2026 → ene-2027 = 1/6·2026 + 1/6·2025 + … + 1/6·2022, "
        "cada tramo ajustado 50/50 por la acción."
    ))
    ws.merge_cells(start_row=formula_row + 1, start_column=2, end_row=formula_row + 1, end_column=9)
    ws.cell(formula_row + 1, 2).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[formula_row + 1].height = 36

    _border_range(ws, 4, total_row, 3 + len(PAGO_ANIOS))
    for j in range(len(PAGO_ANIOS)):
        ws.cell(precio_row, 4 + j).border = THIN

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 28
    for i in range(4, 10):
        ws.column_dimensions[get_column_letter(i)].width = 11

    # Highlight 2027 column
    for row in range(4, total_row + 1):
        ws.cell(row, 4).fill = YELLOW if row == total_row else (
            HEADER_FILL if row == 4 else PatternFill("solid", fgColor="DEEBF7")
        )
        if row == 4:
            ws.cell(row, 4).font = HEADER_FONT


def build_planes(wb: Workbook) -> None:
    ws = wb.create_sheet("Planes vigentes")
    planes = _read_csv(PLANES_CSV)
    headers = [
        "Concepto",
        "Año grant",
        "Valor nominal",
        "Valor de acción al otorgamiento",
        "1/6 anual (base)",
        "Fijo 50% base",
        "Variable 50% base",
        "Estado",
        "Notas",
    ]
    ws.append(headers)
    _style_header(ws)
    for p in planes:
        nom = _to_float(p["valor_nominal_usd"]) or 0.0
        tramo = nom / 6.0
        ws.append(
            [
                p["plan"],
                int(p["anio_grant"]),
                nom,
                _to_float(p["precio_accion_otorgamiento"]),
                tramo,
                tramo * 0.5,
                tramo * 0.5,
                p["estado"],
                p.get("notas", ""),
            ]
        )
    total_row = ws.max_row + 1
    ws.cell(total_row, 1, "TOTAL")
    ws.cell(total_row, 1).font = Font(bold=True)
    for col in (3, 5, 6, 7):
        letter = get_column_letter(col)
        cell = ws.cell(total_row, col, f"=SUM({letter}2:{letter}{total_row - 1})")
        cell.font = Font(bold=True)
        cell.fill = SECTION_FILL
        cell.number_format = MONEY_DEC if col > 3 else MONEY_FORMAT
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=7):
        for cell in row:
            if cell.column == 3:
                cell.number_format = MONEY_FORMAT
            elif cell.column == 4:
                cell.number_format = MONEY_DEC
            else:
                cell.number_format = MONEY_DEC
    _border_range(ws, 1, ws.max_row, 9)
    _autosize(ws)


def build_accion(wb: Workbook) -> None:
    ws = wb.create_sheet("Acción estimada")
    ws.append(["Año pago", "Precio acción estimado", "Notas"])
    _style_header(ws)
    for r in _read_csv(ACCION_CSV):
        ws.append(
            [
                int(r["anio_pago"]),
                _to_float(r["precio_accion_estimado"]),
                r.get("notas", ""),
            ]
        )
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=2):
        for cell in row:
            cell.number_format = MONEY_DEC
            cell.fill = GREEN
    ws.append([])
    ws.append(["Metodología", ""])
    ws["A" + str(ws.max_row)].font = Font(bold=True)
    ws.append(
        [
            "Año 1 (2027)",
            "Promedio últimos 60 días de trading + ajuste 3% cierre 2026",
        ]
    )
    ws.append(["Años siguientes", "Crecimiento anual 8%"])
    ws.append(
        [
            "Fórmula pago",
            "ROUND((N/6)*50% + (N/6)*50%*(P_año/P_otorgamiento); 0)",
        ]
    )
    _border_range(ws, 1, 7, 3)
    _autosize(ws, max_width=70)


def build_calendario(wb: Workbook) -> None:
    ws = wb.create_sheet("Calendario vesting")
    vesting = _read_csv(VESTING_CSV)
    # Filter columns relevant + full range from CSV
    years = list(range(2023, 2033))
    headers = ["Plan"] + [f"31/01/{y}" for y in years]
    ws.append(headers)
    _style_header(ws)
    for row in vesting:
        ws.append([row["plan"]] + [row.get(f"pago_{y}", "") for y in years])
    # Count row
    counts = ["# planes"]
    for y in years:
        n = sum(1 for row in vesting if row.get(f"pago_{y}", "").strip())
        counts.append(n)
    ws.append(counts)
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
        cell.fill = SECTION_FILL
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(headers)):
        for cell in r:
            cell.border = THIN
            cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 14
    for i in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11


def build_detalle_formula(wb: Workbook) -> None:
    """Desglose fijo/variable 2027 para auditar vs portal."""
    ws = wb.create_sheet("Detalle fórmula 2027")
    planes = _read_csv(PLANES_CSV)
    px = 1765.02  # precio 2027
    ws.append(["Detalle ene-2027 — base 50/50 × precio acción"])
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([f"Precio acción estimado 2027: {px}"])
    ws.append([])
    ws.append(
        [
            "Plan",
            "Nominal",
            "P otorgamiento",
            "1/6",
            "Fijo 50%",
            "Variable 50% × (P2027/Pgrant)",
            "Total (sin round)",
            "Total simulación (round)",
        ]
    )
    _style_header(ws, 4)
    matriz = {r["plan"]: r for r in _read_csv(MATRIZ_CSV) if r["plan"] != "TOTAL"}
    start = 5
    for i, p in enumerate(planes):
        nom = _to_float(p["valor_nominal_usd"]) or 0.0
        grant = _to_float(p["precio_accion_otorgamiento"]) or 0.0
        tramo = nom / 6.0
        fijo = tramo * 0.5
        variable = tramo * 0.5 * (px / grant)
        total = fijo + variable
        rounded = _to_int(matriz[p["plan"]]["pago_2027"])
        ws.append([p["plan"], nom, grant, tramo, fijo, variable, total, rounded])
        r = start + i
        for col in range(2, 9):
            fmt = MONEY_FORMAT if col in (2, 8) else MONEY_DEC
            ws.cell(r, col).number_format = fmt
    end = start + len(planes) - 1
    ws.append(
        [
            "TOTAL",
            f"=SUM(B{start}:B{end})",
            "",
            f"=SUM(D{start}:D{end})",
            f"=SUM(E{start}:E{end})",
            f"=SUM(F{start}:F{end})",
            f"=SUM(G{start}:G{end})",
            f"=SUM(H{start}:H{end})",
        ]
    )
    for col in range(1, 9):
        cell = ws.cell(end + 1, col)
        cell.font = Font(bold=True)
        cell.fill = YELLOW
        if col in (2, 8):
            cell.number_format = MONEY_FORMAT
        elif col >= 4:
            cell.number_format = MONEY_DEC
    ws.append([])
    ws.append(["Base sin acción (factor 1,0)", 32803.67])
    ws.append(["Estimado con acción (simulación)", 35332])
    ws.append(["Diferencia = efecto valor de la acción", 35332 - 32803.67])
    for r in range(ws.max_row - 2, ws.max_row + 1):
        ws.cell(r, 2).number_format = MONEY_DEC
    _border_range(ws, 4, end + 1, 8)
    _autosize(ws, max_width=40)


def build_resumen(wb: Workbook) -> None:
    ws = wb.create_sheet("Resumen", 0)
    # Move simulation to stay as practical main; resumen first for norms
    rows = [
        ("Long Term Retention Program (LTRP)", ""),
        ("Excel canónico", "LTRP.xlsx"),
        ("Última actualización", "agosto 2026"),
        ("", ""),
        (
            "REGLA GENERAL",
            "Máximo 6 planes; cada enero se suma el 1/6 de cada uno",
        ),
        (
            "Ejemplo",
            "Otorgado 70.000 en 2026 → ene-2027 = 1/6·2026 + 1/6·2025 + 1/6·2024 + 1/6·2023 + 1/6·2022",
        ),
        (
            "Fórmula con acción",
            "(N/6)×50% fijo + (N/6)×50% × (P_año / P_otorgamiento)",
        ),
        ("", ""),
        ("Total estimado a cobrar 2027", 35332),
        ("Total estimado a cobrar 2028", 36847),
        ("Total estimado a cobrar 2029", 32284),
        ("Total estimado a cobrar 2030", 25001),
        ("Total estimado a cobrar 2031", 19379),
        ("Total estimado a cobrar 2032", 13056),
        ("Total acumulado pendiente de pago", 161899),
        ("", ""),
        ("Base 2027 sin efecto acción (suma 1/6)", 32803.67),
        ("Diferencia 2027 por valor de la acción", 2528.33),
        ("Precio acción estimado 2027", 1765.02),
        ("", ""),
        (
            "Metodología precio",
            "Año 1: promedio 60 días + 3% cierre 2026 · Luego +8% anual",
        ),
        (
            "Nota",
            "Ingreso por bono MELI; independiente de CIUDAD / SOL / propiedades.",
        ),
    ]
    ws.append(["Campo", "Valor"])
    _style_header(ws)
    for campo, valor in rows:
        ws.append([campo, valor])
    for row_idx in range(9, 16):
        ws.cell(row_idx, 2).number_format = MONEY_FORMAT
    for row_idx in (17, 18, 19):
        ws.cell(row_idx, 2).number_format = MONEY_DEC
    for addr in ("A5", "B5", "A6", "B6", "A9", "B9", "A15", "B15"):
        ws[addr].fill = YELLOW
        ws[addr].font = Font(bold=True)
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    _border_range(ws, 1, ws.max_row, 2)
    _autosize(ws, min_width=28, max_width=95)


def main() -> None:
    wb = Workbook()
    # Simulación first as active, then re-order: Resumen, Simulación, ...
    build_simulacion(wb)
    build_planes(wb)
    build_accion(wb)
    build_calendario(wb)
    build_detalle_formula(wb)
    build_resumen(wb)

    # Ensure Resumen is first sheet
    wb.move_sheet("Resumen", offset=-len(wb.sheetnames) + 1)

    wb.save(OUT)
    print(f"Wrote {OUT}")
    print(f"Sheets: {wb.sheetnames}")

    # Sanity checks vs Sol's simulation
    matriz = _read_csv(MATRIZ_CSV)
    total_row = next(r for r in matriz if r["plan"] == "TOTAL")
    assert _to_int(total_row["pago_2027"]) == 35332
    assert _to_int(total_row["pago_2028"]) == 36847
    pending = sum(_to_int(total_row[f"pago_{y}"]) or 0 for y in PAGO_ANIOS)
    assert pending == 161899, pending
    print(f"OK total 2027 = 35332 · pendiente acumulado = {pending}")

    # Formula check LTRP 2026 / 2027
    calc = 70000 / 6 * 0.5 + 70000 / 6 * 0.5 * (1765.02 / 2094.65)
    assert abs(calc - 10749) < 1.0, calc
    print(f"OK fórmula LTRP 2026 ene-2027 ≈ {calc:.2f} → 10749")


if __name__ == "__main__":
    main()
