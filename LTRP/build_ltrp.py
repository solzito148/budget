#!/usr/bin/env python3
"""Regenera LTRP/LTRP_Sol.xlsx desde los CSV canónicos del programa."""

from __future__ import annotations

import csv
import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "LTRP_Sol.xlsx"

PLANES_CSV = ROOT / "planes_vigentes.csv"
VESTING_CSV = ROOT / "calendario_vesting.csv"
SIM_CSV = ROOT / "simulacion_pagos.csv"

# Años de pago cubiertos en la matriz (31/01 de cada año)
PAGO_ANIOS = list(range(2023, 2033))

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="D6E3F0")
MONEY_FORMAT = '#,##0.00'
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


def _frac_to_float(frac: str) -> float | None:
    frac = (frac or "").strip()
    if not frac:
        return None
    if "/" in frac:
        num, den = frac.split("/", 1)
        return float(num) / float(den)
    return float(frac)


def build_resumen(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "Resumen"
    rows = [
        ("Long Term Retention Program (LTRP)", ""),
        ("Última actualización", "agosto 2026"),
        ("Moneda", "USD"),
        (
            "Fórmula del cobro 31/01",
            "SUMAR (Valor nominal ÷ 6) de cada plan vigente desde 2022",
        ),
        (
            "Ejemplo plan 2026",
            "70.000 ÷ 6 = 11.666,67  +  1/6 de LTRP 2022…2025",
        ),
        ("Duración por plan", "6 cobros anuales (1/6 por año) en 31/01"),
        ("Composición del pago", "50% fijo + 50% variable (acción MELI)"),
        ("", ""),
        ("Pago base 31/01/2027 = suma de cinco 1/6", 32803.67),
        ("  · LTRP 2022 ÷ 6", 5000),
        ("  · LTRP 2023 ÷ 6", 5000),
        ("  · LTRP 2024 ÷ 6", 5000),
        ("  · LTRP 2025 ÷ 6", 6137),
        ("  · LTRP 2026 ÷ 6", 11666.6666666667),
        ("Estimación portal MELI 2027", 35332),
        ("Factor MELI implícito (solo variable)", 1.1542),
        ("", ""),
        ("Total valor nominal planes vigentes", 196822),
        ("Planes vigentes", 5),
        ("", ""),
        (
            "Nota",
            "El valor nominal (ej. 70.000) es el plan asignado ese año, "
            "NO el cobro anual. El cobro = suma de los ÷6 activos. "
            "El 50% variable mueve el total vs la base. "
            "Ingreso por bono; no es gasto CIUDAD/SOL.",
        ),
    ]
    ws.append(["Campo", "Valor"])
    _style_header(ws)
    for campo, valor in rows:
        ws.append([campo, valor])
    for row_idx in (9, 10, 11, 12, 13, 14, 15, 18):
        ws.cell(row_idx, 2).number_format = MONEY_FORMAT
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=2):
        for cell in row:
            cell.border = THIN
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    # Destacar fila de suma 2027
    ws["A9"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["B9"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws["A9"].font = Font(bold=True)
    ws["B9"].font = Font(bold=True)
    _autosize(ws, min_width=28, max_width=80)


def build_planes(wb: Workbook) -> None:
    ws = wb.create_sheet("Planes vigentes")
    planes = _read_csv(PLANES_CSV)
    headers = [
        "Plan",
        "Año grant",
        "Valor nominal USD",
        "Tramo anual 1/6 USD",
        "Fijo 50% USD",
        "Variable 50% base USD",
        "Estado",
        "Notas",
    ]
    ws.append(headers)
    _style_header(ws)
    for p in planes:
        ws.append(
            [
                p["plan"],
                int(p["anio_grant"]),
                _to_float(p["valor_nominal_usd"]),
                _to_float(p["tramo_anual_usd"]),
                _to_float(p["fijo_50_usd"]),
                _to_float(p["variable_50_base_usd"]),
                p["estado"],
                p["notas"],
            ]
        )
    # Totales
    total_row = ws.max_row + 1
    ws.cell(total_row, 1, "TOTAL")
    ws.cell(total_row, 1).font = Font(bold=True)
    ws.cell(total_row, 1).fill = SECTION_FILL
    for col in range(3, 7):
        letter = get_column_letter(col)
        ws.cell(total_row, col, f"=SUM({letter}2:{letter}{total_row - 1})")
        ws.cell(total_row, col).font = Font(bold=True)
        ws.cell(total_row, col).fill = SECTION_FILL
        ws.cell(total_row, col).number_format = MONEY_FORMAT
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=6):
        for cell in row:
            if cell.row < total_row:
                cell.number_format = MONEY_FORMAT
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=8):
        for cell in row:
            cell.border = THIN
    _autosize(ws)


def build_calendario(wb: Workbook) -> None:
    ws = wb.create_sheet("Calendario vesting")
    vesting = _read_csv(VESTING_CSV)
    headers = ["Plan"] + [f"Pago {y} (31/01/{y})" for y in PAGO_ANIOS]
    ws.append(headers)
    _style_header(ws)
    for row in vesting:
        ws.append([row["plan"]] + [row.get(f"pago_{y}", "") for y in PAGO_ANIOS])
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(headers)):
        for cell in r:
            cell.border = THIN
            cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 14
    for i in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 14


def build_matriz_usd(wb: Workbook) -> None:
    """Matriz USD base (factor MELI = 1,0) por plan y año de pago."""
    ws = wb.create_sheet("Matriz USD base")
    planes = _read_csv(PLANES_CSV)
    vesting = {r["plan"]: r for r in _read_csv(VESTING_CSV)}
    headers = ["Plan", "Nominal USD"] + [f"{y}" for y in PAGO_ANIOS] + ["Total plan"]
    ws.append(headers)
    _style_header(ws)

    totals_by_year = {y: 0.0 for y in PAGO_ANIOS}
    for p in planes:
        plan = p["plan"]
        nominal = _to_float(p["valor_nominal_usd"]) or 0.0
        tramo = nominal / 6.0
        row_vals: list[object] = [plan, nominal]
        plan_total = 0.0
        vrow = vesting.get(plan, {})
        for y in PAGO_ANIOS:
            frac = _frac_to_float(vrow.get(f"pago_{y}", ""))
            if frac is None:
                row_vals.append(None)
            else:
                # Cada celda con fracción activa aporta el tramo completo 1/6
                # (la fracción N/6 indica el número de tramo, no un múltiplo).
                amount = tramo
                row_vals.append(amount)
                plan_total += amount
                totals_by_year[y] += amount
        row_vals.append(plan_total)
        ws.append(row_vals)

    total_row = ["TOTAL", f"=SUM(B2:B{ws.max_row})"]
    for y in PAGO_ANIOS:
        total_row.append(totals_by_year[y])
    total_row.append(sum(totals_by_year.values()))
    ws.append(total_row)
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
        cell.fill = SECTION_FILL

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=ws.max_column):
        for cell in row:
            if isinstance(cell.value, (int, float)) or (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ):
                cell.number_format = MONEY_FORMAT
            cell.border = THIN
            cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.border = THIN
    _autosize(ws, min_width=12)


def build_simulacion(wb: Workbook) -> None:
    ws = wb.create_sheet("Simulación pagos")
    sims = _read_csv(SIM_CSV)
    headers = [
        "Año pago",
        "Fecha pago",
        "Planes activos",
        "Tramo fijo USD",
        "Tramo variable base USD",
        "Total base (factor 1,0) USD",
        "Estimado portal USD",
        "Factor MELI implícito (variable)",
        "Última actualización",
        "Notas",
    ]
    ws.append(headers)
    _style_header(ws)
    for s in sims:
        ws.append(
            [
                int(s["anio_pago"]),
                s["fecha_pago"],
                int(s["planes_activos"]),
                _to_float(s["tramo_fijo_usd"]),
                _to_float(s["tramo_variable_base_usd"]),
                _to_float(s["total_base_factor_1_usd"]),
                _to_float(s["estimado_portal_usd"]),
                _to_float(s["factor_meli_implicito_variable"]),
                s.get("ultima_actualizacion", ""),
                s.get("notas", ""),
            ]
        )
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=4, max_col=7):
        for cell in row:
            if cell.value is not None:
                cell.number_format = MONEY_FORMAT
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=8, max_col=8):
        for cell in row:
            if cell.value is not None:
                cell.number_format = "0.0000"
    # Destacar fila 2027
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=10):
        if row[0].value == 2027:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="FFF2CC")
        for cell in row:
            cell.border = THIN
    _autosize(ws, max_width=55)


def build_detalle_2027(wb: Workbook) -> None:
    ws = wb.create_sheet("Detalle pago 2027")
    ws.append(["Pago 31/01/2027 = suma de (Nominal ÷ 6) de cada plan desde 2022"])
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(
        [
            "Plan",
            "Valor nominal USD",
            "Nominal ÷ 6",
            "Fijo 50% USD",
            "Variable base 50% USD",
            "Variable estimado*",
            "Total estimado plan*",
        ]
    )
    _style_header(ws, row=3)

    # Factor implícito desde estimado portal
    factor = 1.1542
    detalle = [
        ("LTRP 2022", 30000.0),
        ("LTRP 2023", 30000.0),
        ("LTRP 2024", 30000.0),
        ("LTRP 2025", 36822.0),
        ("LTRP 2026", 70000.0),
    ]
    start = 4
    for i, (plan, nominal) in enumerate(detalle):
        tramo = nominal / 6.0
        fijo = tramo * 0.5
        var_base = tramo * 0.5
        var_est = var_base * factor
        total_est = fijo + var_est
        ws.append([plan, nominal, tramo, fijo, var_base, var_est, total_est])
        r = start + i
        for col in range(2, 8):
            ws.cell(r, col).number_format = MONEY_FORMAT

    end = start + len(detalle) - 1
    ws.append(
        [
            "SUMA (pago base 31/01/2027)",
            f"=SUM(B{start}:B{end})",
            f"=SUM(C{start}:C{end})",
            f"=SUM(D{start}:D{end})",
            f"=SUM(E{start}:E{end})",
            f"=SUM(F{start}:F{end})",
            f"=SUM(G{start}:G{end})",
        ]
    )
    for col in range(1, 8):
        cell = ws.cell(end + 1, col)
        cell.font = Font(bold=True)
        cell.fill = SECTION_FILL
        if col >= 2:
            cell.number_format = MONEY_FORMAT

    ws.append([])
    ws.append(["Fórmula", "Pago base = 30000/6 + 30000/6 + 30000/6 + 36822/6 + 70000/6"])
    ws.append(["", "= 5.000 + 5.000 + 5.000 + 6.137 + 11.666,67 = 32.803,67"])
    ws.append(["Estimación portal MELI (agosto 2026)", 35332])
    ws["B" + str(ws.max_row)].number_format = MONEY_FORMAT
    ws.append(["Factor MELI usado en esta hoja (implícito)", factor])
    ws.append(
        [
            "*Variable estimado",
            "Variable base × factor MELI implícito (35.332 vs base 32.803,67)",
        ]
    )
    ws.append(
        [
            "Pago final",
            "(suma ÷6)×50% fijo + (suma ÷6)×50%×factor_MELI",
        ]
    )
    for row in ws.iter_rows(min_row=3, max_row=end + 1, max_col=7):
        for cell in row:
            cell.border = THIN
    _autosize(ws, min_width=14, max_width=70)


def main() -> None:
    wb = Workbook()
    build_resumen(wb)
    build_planes(wb)
    build_calendario(wb)
    build_matriz_usd(wb)
    build_simulacion(wb)
    build_detalle_2027(wb)
    wb.save(OUT)
    print(f"Wrote {OUT}")

    # Sanity: base 2027 ≈ 32803.67
    planes = _read_csv(PLANES_CSV)
    total_nominal = sum(_to_float(p["valor_nominal_usd"]) or 0.0 for p in planes)
    base_2027 = total_nominal / 6.0
    assert math.isclose(base_2027, 32803.6666666667, rel_tol=1e-9), base_2027
    print(f"OK base 2027 factor 1.0 = {base_2027:.2f} USD")
    print("OK estimación portal 2027 = 35332 USD")


if __name__ == "__main__":
    main()
