#!/usr/bin/env python3
"""Regenera LTRP/LTRP.xlsx — Excel canónico del Long Term Retention Program."""

from __future__ import annotations

import csv
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

PAGO_ANIOS = list(range(2027, 2038))  # 2027–2037

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="D6E3F0")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREEN = PatternFill("solid", fgColor="E2EFDA")
BLUE = PatternFill("solid", fgColor="DEEBF7")
GRAY = PatternFill("solid", fgColor="F2F2F2")
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
    """Hoja principal — simulación Sol con edad + proyección 2027–2036."""
    ws = wb.active
    ws.title = "Simulación de pagos"

    matriz = [r for r in _read_csv(MATRIZ_CSV)]
    accion = {
        int(r["anio_pago"]): _to_float(r["precio_accion_estimado"])
        for r in _read_csv(ACCION_CSV)
    }

    ws["A1"] = "Long Term Retention Program (LTRP) — Simulación de pagos"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws["A2"] = (
        "Máx. 6 planes por enero · 50% fijo + 50% variable (acción MELI) · "
        "Planes 2027+ proyectados a 70.000 con 1/6≈11.666 (sin precio de acción aún) · "
        "Plan 2026 se completa recién en 2032"
    )
    ws["A2"].font = Font(italic=True, color="595959")
    ws.merge_cells("A2:P2")

    # Fila 3: "Año de PAGO" spanning payment columns
    ws.cell(3, 6, "Año de PAGO →")
    ws.cell(3, 6).font = Font(bold=True, color="1F4E79")
    for j, y in enumerate(PAGO_ANIOS):
        cell = ws.cell(3, 6 + j, y)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    # Encabezados fila 4
    headers = [
        "Edad Sol",
        "Concepto",
        "Valor nominal",
        "Valor por año Aprox",
        "Valor de acción al otorgamiento",
    ] + [str(y) for y in PAGO_ANIOS]
    for col, h in enumerate(headers, 1):
        ws.cell(4, col, h)
    _style_header(ws, 4)

    start_plan = 5
    for i, m in enumerate(matriz):
        row = start_plan + i
        edad = _to_int(m["edad_sol"])
        nominal = _to_float(m["valor_nominal"]) or 0.0
        por_ano = _to_float(m["valor_por_ano"])
        grant_px = _to_float(m["precio_otorgamiento"])
        proyectado = grant_px is None

        ws.cell(row, 1, edad).alignment = Alignment(horizontal="center")
        ws.cell(row, 2, m["plan"])
        ws.cell(row, 3, nominal).number_format = MONEY_FORMAT
        ws.cell(row, 4, por_ano).number_format = MONEY_DEC
        if grant_px is not None:
            ws.cell(row, 5, grant_px).number_format = MONEY_DEC
        else:
            ws.cell(row, 5, None)

        if proyectado:
            for col in range(1, 6):
                if ws.cell(row, col).fill.fgColor is None or True:
                    ws.cell(row, col).fill = GRAY

        for j, y in enumerate(PAGO_ANIOS):
            col = 6 + j
            raw = (m.get(f"pago_{y}") or "").strip()
            if not raw:
                ws.cell(row, col, None)
            else:
                # 11666 may be float-looking; keep as number
                val = float(raw) if "." in raw else int(float(raw))
                cell = ws.cell(row, col, val)
                cell.number_format = MONEY_FORMAT
                cell.alignment = Alignment(horizontal="center")
                if proyectado:
                    cell.fill = GRAY
                elif y == 2027:
                    cell.fill = BLUE

    end_plan = start_plan + len(matriz) - 1

    # Total estimado a cobrar por año
    total_row = end_plan + 1
    ws.cell(total_row, 1, "")
    ws.cell(total_row, 2, "Total estimado a cobrar por año")
    ws.cell(total_row, 2).font = Font(bold=True)
    for col in range(1, 6):
        ws.cell(total_row, col).fill = SECTION_FILL
        ws.cell(total_row, col).font = Font(bold=True)
    for j, y in enumerate(PAGO_ANIOS):
        col = 6 + j
        letter = get_column_letter(col)
        cell = ws.cell(
            total_row, col, f"=SUM({letter}{start_plan}:{letter}{end_plan})"
        )
        cell.font = Font(bold=True)
        cell.fill = YELLOW
        cell.number_format = MONEY_FORMAT
        cell.alignment = Alignment(horizontal="center")

    # Acción estimada (solo años con metodología Sol)
    accion_label = total_row + 2
    ws.cell(accion_label, 2, "Valor de la acción estimado")
    ws.cell(accion_label, 2).font = Font(bold=True, size=12, color="1F4E79")
    ws.cell(
        accion_label + 1,
        2,
        (
            "Año 1 (2027): promedio últimos 60 días de trading + ajuste 3% cierre 2026. "
            "Años siguientes: crecimiento anual 8%. "
            "Planes 2027+ aún sin precio al otorgamiento → placeholder 11.666 (1/6 nominal)."
        ),
    )
    ws.merge_cells(
        start_row=accion_label + 1,
        start_column=2,
        end_row=accion_label + 1,
        end_column=6 + len(PAGO_ANIOS) - 1,
    )
    ws.cell(accion_label + 1, 2).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[accion_label + 1].height = 40

    precio_row = accion_label + 2
    ws.cell(precio_row, 2, "Precio estimado")
    ws.cell(precio_row, 2).font = Font(bold=True)
    for j, y in enumerate(PAGO_ANIOS):
        col = 6 + j
        if y in accion:
            cell = ws.cell(precio_row, col, accion[y])
            cell.number_format = MONEY_DEC
            cell.fill = GREEN
        else:
            cell = ws.cell(precio_row, col, "n/d")
            cell.fill = GRAY
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN

    # Pendiente: suma de totales (incluye proyección)
    pending_row = precio_row + 2
    ws.cell(pending_row, 2, "Total acumulado pendiente de pago (con proyección)")
    ws.cell(pending_row, 2).font = Font(bold=True)
    first_tot = get_column_letter(6)
    last_tot = get_column_letter(5 + len(PAGO_ANIOS))
    cell = ws.cell(
        pending_row, 3, f"=SUM({first_tot}{total_row}:{last_tot}{total_row})"
    )
    cell.font = Font(bold=True, size=12)
    cell.fill = YELLOW
    cell.number_format = MONEY_FORMAT

    # Pendiente solo planes vigentes con acción (2022–2026, años 2027–2032)
    pending_known = pending_row + 1
    ws.cell(pending_known, 2, "Pendiente planes vigentes 2022–2026 (simulación con acción)")
    ws.cell(pending_known, 3, 161899)
    ws.cell(pending_known, 3).number_format = MONEY_FORMAT
    ws.cell(pending_known, 3).fill = GREEN

    # Notas
    note_row = pending_known + 2
    ws.cell(note_row, 2, "Fórmula (planes con precio al otorgamiento)")
    ws.cell(note_row, 2).font = Font(bold=True)
    ws.cell(
        note_row,
        3,
        "ROUND( (N÷6)×50% + (N÷6)×50% × (P_año / P_otorgamiento) ; 0 )",
    )
    ws.cell(note_row + 1, 2, "Regla general")
    ws.cell(note_row + 1, 2).font = Font(bold=True)
    ws.cell(
        note_row + 1,
        3,
        (
            "Máximo 6 planes por cobro de enero. "
            "Ej.: otorgado 70.000 en 2026 → ene-2027 = 1/6·2026+…+1/6·2022. "
            "Recién en 2032 se completa la totalidad del plan 2026. "
            "Filas grises = proyección futura (nominal 70.000 asumido)."
        ),
    )
    ws.merge_cells(
        start_row=note_row + 1,
        start_column=3,
        end_row=note_row + 1,
        end_column=16,
    )
    ws.cell(note_row + 1, 3).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[note_row + 1].height = 48

    _border_range(ws, 4, total_row, 5 + len(PAGO_ANIOS))

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 28
    for i in range(6, 6 + len(PAGO_ANIOS)):
        ws.column_dimensions[get_column_letter(i)].width = 10


def build_planes(wb: Workbook) -> None:
    ws = wb.create_sheet("Planes")
    planes = _read_csv(PLANES_CSV)
    headers = [
        "Edad Sol",
        "Concepto",
        "Año grant",
        "Valor nominal",
        "Valor por año Aprox",
        "Valor de acción al otorgamiento",
        "Estado",
        "Notas",
    ]
    ws.append(headers)
    _style_header(ws)
    for p in planes:
        px = _to_float(p.get("precio_accion_otorgamiento"))
        row = [
            _to_int(p["edad_sol"]),
            p["plan"],
            int(p["anio_grant"]),
            _to_float(p["valor_nominal_usd"]),
            _to_float(p["valor_por_ano_aprox"]),
            px,
            p["estado"],
            p.get("notas", ""),
        ]
        ws.append(row)
        r = ws.max_row
        ws.cell(r, 4).number_format = MONEY_FORMAT
        ws.cell(r, 5).number_format = MONEY_DEC
        if px is not None:
            ws.cell(r, 6).number_format = MONEY_DEC
        if p["estado"] == "proyectado":
            for col in range(1, 9):
                ws.cell(r, col).fill = GRAY
    _border_range(ws, 1, ws.max_row, 8)
    _autosize(ws, max_width=55)


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
    ws[f"A{ws.max_row}"].font = Font(bold=True)
    ws.append(["Año 1 (2027)", "Promedio 60 días trading + ajuste 3% cierre 2026"])
    ws.append(["Años siguientes", "Crecimiento anual 8%"])
    ws.append(
        [
            "Planes 2027+",
            "Sin precio al otorgamiento aún → simulación usa 11.666 (1/6 de 70.000)",
        ]
    )
    _border_range(ws, 1, 7, 3)
    _autosize(ws, max_width=70)


def build_calendario(wb: Workbook) -> None:
    ws = wb.create_sheet("Calendario vesting")
    vesting = _read_csv(VESTING_CSV)
    years = list(range(2023, 2038))
    headers = ["Plan"] + [f"31/01/{y}" for y in years]
    ws.append(headers)
    _style_header(ws)
    for row in vesting:
        ws.append([row["plan"]] + [row.get(f"pago_{y}", "") for y in years])
    counts = ["# planes"]
    for y in years:
        n = sum(1 for row in vesting if (row.get(f"pago_{y}") or "").strip())
        counts.append(min(n, 6) if n else 0)  # display actual count
    # recount actual
    counts = ["# planes"]
    for y in years:
        n = sum(1 for row in vesting if (row.get(f"pago_{y}") or "").strip())
        counts.append(n)
    ws.append(counts)
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
        cell.fill = SECTION_FILL
        # highlight if > 6 (should not happen for payment years with full grants)
        if isinstance(cell.value, int) and cell.value > 6:
            cell.fill = PatternFill("solid", fgColor="FCE4D6")
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(headers)):
        for cell in r:
            cell.border = THIN
            cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 14
    for i in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 10


def build_detalle_2027(wb: Workbook) -> None:
    ws = wb.create_sheet("Detalle fórmula 2027")
    planes = [p for p in _read_csv(PLANES_CSV) if p["estado"] == "vigente"]
    px = 1765.02
    matriz = {
        r["plan"]: r for r in _read_csv(MATRIZ_CSV) if (r.get("pago_2027") or "").strip()
    }
    ws.append(["Detalle ene-2027 — solo planes vigentes con precio de acción"])
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([f"Precio acción estimado 2027: {px}"])
    ws.append([])
    ws.append(
        [
            "Edad",
            "Plan",
            "Nominal",
            "P otorgamiento",
            "1/6",
            "Fijo 50%",
            "Variable ajustada",
            "Total (sin round)",
            "Simulación (round)",
        ]
    )
    _style_header(ws, 4)
    start = 5
    for i, p in enumerate(planes):
        nom = _to_float(p["valor_nominal_usd"]) or 0.0
        grant = _to_float(p["precio_accion_otorgamiento"]) or 0.0
        tramo = nom / 6.0
        fijo = tramo * 0.5
        variable = tramo * 0.5 * (px / grant)
        total = fijo + variable
        rounded = _to_int(matriz[p["plan"]]["pago_2027"])
        ws.append(
            [
                _to_int(p["edad_sol"]),
                p["plan"],
                nom,
                grant,
                tramo,
                fijo,
                variable,
                total,
                rounded,
            ]
        )
        r = start + i
        for col in range(3, 10):
            ws.cell(r, col).number_format = MONEY_FORMAT if col in (3, 9) else MONEY_DEC
    end = start + len(planes) - 1
    ws.append(
        [
            "",
            "TOTAL",
            f"=SUM(C{start}:C{end})",
            "",
            f"=SUM(E{start}:E{end})",
            f"=SUM(F{start}:F{end})",
            f"=SUM(G{start}:G{end})",
            f"=SUM(H{start}:H{end})",
            f"=SUM(I{start}:I{end})",
        ]
    )
    for col in range(1, 10):
        cell = ws.cell(end + 1, col)
        cell.font = Font(bold=True)
        cell.fill = YELLOW
        if col in (3, 9):
            cell.number_format = MONEY_FORMAT
        elif col >= 5:
            cell.number_format = MONEY_DEC
    ws.append([])
    ws.append(["Base sin acción (suma 1/6)", 32803.67])
    ws.append(["Estimado con acción", 35332])
    ws.append(["Diferencia = efecto valor de la acción", 2528.33])
    for r in range(ws.max_row - 2, ws.max_row + 1):
        ws.cell(r, 2).number_format = MONEY_DEC
    _border_range(ws, 4, end + 1, 9)
    _autosize(ws, max_width=40)


def build_resumen(wb: Workbook) -> None:
    ws = wb.create_sheet("Resumen", 0)
    matriz = [r for r in _read_csv(MATRIZ_CSV)]
    totals = {}
    for y in PAGO_ANIOS:
        totals[y] = sum(
            float(r[f"pago_{y}"])
            for r in matriz
            if (r.get(f"pago_{y}") or "").strip()
        )

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
            "Otorgado 70.000 en 2026 (edad 45) → ene-2027 = 1/6·2026+…+1/6·2022",
        ),
        (
            "Plan 2026 completo",
            "Recién en 2032 se cobra el 6/6 y se completa la totalidad del plan",
        ),
        (
            "Proyección",
            "LTRP 2027–2036 asumidos a 70.000 · 1/6≈11.666 sin precio de acción aún",
        ),
        (
            "Fórmula con acción",
            "(N/6)×50% fijo + (N/6)×50% × (P_año / P_otorgamiento)",
        ),
        ("", ""),
        ("Total estimado 2027 (vigentes)", 35332),
        ("Total estimado 2028 (con proy. 2027)", round(totals[2028])),
        ("Total estimado 2029", round(totals[2029])),
        ("Total estimado 2030", round(totals[2030])),
        ("Total estimado 2031", round(totals[2031])),
        ("Total estimado 2032 (cierra plan 2026)", round(totals[2032])),
        ("Total estimado 2033–2037 (c/u, tope 6×11666)", round(totals[2033])),
        ("", ""),
        ("Pendiente vigentes 2022–2026 (con acción)", 161899),
        ("Pendiente total con proyección 2027–2037", round(sum(totals.values()))),
        ("", ""),
        ("Base 2027 sin efecto acción", 32803.67),
        ("Diferencia 2027 por valor de la acción", 2528.33),
        ("Precio acción estimado 2027", 1765.02),
        ("", ""),
        (
            "Edad Sol",
            "Columna por plan al año de otorgamiento (41 en LTRP 2022 … 51 en LTRP 2036)",
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
    for row_idx in range(11, 18):
        ws.cell(row_idx, 2).number_format = MONEY_FORMAT
    for row_idx in (19, 20, 22, 23, 24):
        ws.cell(row_idx, 2).number_format = (
            MONEY_DEC if row_idx >= 22 else MONEY_FORMAT
        )
    for addr in ("A5", "B5", "A6", "B6", "A7", "B7", "A11", "B11", "A16", "B16", "A19", "B19"):
        ws[addr].fill = YELLOW
        ws[addr].font = Font(bold=True)
    ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    _border_range(ws, 1, ws.max_row, 2)
    _autosize(ws, min_width=28, max_width=95)


def main() -> None:
    wb = Workbook()
    build_simulacion(wb)
    build_planes(wb)
    build_accion(wb)
    build_calendario(wb)
    build_detalle_2027(wb)
    build_resumen(wb)
    wb.move_sheet("Resumen", offset=-len(wb.sheetnames) + 1)
    wb.save(OUT)
    print(f"Wrote {OUT}")
    print(f"Sheets: {wb.sheetnames}")

    matriz = _read_csv(MATRIZ_CSV)
    t2027 = sum(
        float(r["pago_2027"])
        for r in matriz
        if (r.get("pago_2027") or "").strip()
    )
    assert abs(t2027 - 35332) < 0.1, t2027
    assert (matriz[5].get("pago_2027") or "") == ""  # LTRP 2027
    assert float(matriz[5]["pago_2028"]) == 11666
    assert int(matriz[4]["edad_sol"]) == 45  # LTRP 2026
    print("OK 2027=35332 · LTRP 2026 edad 45 · proyección 2027+ alineada")


if __name__ == "__main__":
    main()
