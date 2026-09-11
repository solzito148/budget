#!/usr/bin/env python3
"""Rendición mensual CIUDAD 1132 + Almacén Sol/Chris.

Lee Gastos_CIUDAD_1132_2026.xlsx (hojas `2026` y `Almacen Detalle`),
calcula totales, desglose de almacén y liquidación Sol/Chris según las
fórmulas de la matriz, y escribe:

  - CIUDAD/Rendicion_CIUDAD_2026.xlsx
  - CIUDAD/RENDICION_CIUDAD_2026.md
"""

from __future__ import annotations

import math
from collections import defaultdict
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "Gastos_CIUDAD_1132_2026.xlsx"
OUT_XLSX = ROOT / "Rendicion_CIUDAD_2026.xlsx"
OUT_MD = ROOT / "RENDICION_CIUDAD_2026.md"

MONTHS = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
]
MONTH_COL = {m: i + 4 for i, m in enumerate(MONTHS)}  # D=4 … O=15

NAVY = "1F4E79"
TEAL = "0F6B5C"
LIGHT = "E8F0F7"
LIGHT_TEAL = "E6F4F1"
AMBER = "FFF4E5"
GRAY = "F5F5F5"
RED = "C00000"
GREEN = "006600"

THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def money(n: float | int | None) -> int:
    """Gastos: enteros redondeados hacia arriba."""
    if n is None:
        return 0
    if isinstance(n, str):
        return 0
    return int(math.ceil(float(n)))


def settle(n: float | int | None) -> int:
    """Liquidación 50/50: mitad con redondeo hacia arriba (Sol absorbe como máximo $0,5 extra)."""
    if n is None:
        return 0
    return int(math.ceil(float(n)))


def fmt_ars(n: int | float) -> str:
    v = settle(n)
    sign = "-" if v < 0 else ""
    return f"{sign}${abs(v):,}".replace(",", ".")


def family(detalle: str | None) -> str:
    if not detalle:
        return "SIN DETALLE"
    u = str(detalle).strip().upper()
    if u.startswith("SUPERMERCADO") or "PEDIDOS YA · SUPERMERCADO" in u:
        return "SUPERMERCADO"
    if u.startswith("PEDIDOS YA · COMIDA") or u.startswith("COMIDA"):
        return "COMIDA"
    if u.startswith("RAMON") or u.startswith("RAMÓN"):
        return "RAMÓN"
    if "COMBUSTIBLE" in u:
        return "COMBUSTIBLE"
    if u.startswith("ENTRETENIMIENTO"):
        return "ENTRETENIMIENTO"
    if "STREAMING" in u or "DIRECTV" in u:
        return "STREAMING TV"
    if u.startswith("FARMACIA"):
        return "FARMACIA"
    if u.startswith("REPUESTOS"):
        return "REPUESTOS"
    return "OTROS"


def load_matrix_rows(wb) -> list[dict]:
    ws = wb["2026"]
    rows = []
    for r in range(5, 21):
        tipo = ws.cell(r, 1).value
        cat = ws.cell(r, 2).value
        desc = ws.cell(r, 3).value
        if not tipo and not cat and not desc:
            continue
        if tipo in ("TOTAL", "CADA UNO", "SOL", "CHRIS"):
            continue
        vals = {}
        for mes, col in MONTH_COL.items():
            vals[mes] = ws.cell(r, col).value
        rows.append(
            {
                "row": r,
                "tipo": tipo or "",
                "categoria": cat or "",
                "descripcion": desc or "",
                "valores": vals,
                "es_almacen": (cat or "").upper() == "ALMACEN",
            }
        )
    return rows


def load_almacen(wb) -> tuple[dict, dict, dict, dict]:
    """Return (sol_by_month, chris_by_month, family_by_month, items_by_month)."""
    alm = wb["Almacen Detalle"]
    sol = defaultdict(int)
    chris = defaultdict(int)
    fam = {
        m: defaultdict(lambda: {"sol": 0, "chris": 0, "total": 0}) for m in MONTHS
    }
    items = {m: [] for m in MONTHS}

    for r in range(4, alm.max_row + 1):
        mes = alm.cell(r, 1).value
        if mes not in MONTH_COL:
            continue
        detalle = alm.cell(r, 3).value
        if not detalle:
            continue
        sol_v = alm.cell(r, 4).value
        chris_v = alm.cell(r, 5).value
        if isinstance(sol_v, str) or isinstance(chris_v, str):
            continue
        s = money(sol_v)
        c = money(chris_v)
        if s == 0 and c == 0:
            continue
        sol[mes] += s
        chris[mes] += c
        fam_key = family(detalle)
        fam[mes][fam_key]["sol"] += s
        fam[mes][fam_key]["chris"] += c
        fam[mes][fam_key]["total"] += s + c
        fecha = alm.cell(r, 2).value
        comentario = alm.cell(r, 7).value
        items[mes].append(
            {
                "fecha": fecha,
                "detalle": str(detalle),
                "sol": s,
                "chris": c,
                "total": s + c,
                "familia": fam_key,
                "comentario": comentario or "",
            }
        )
    return sol, chris, fam, items


def resolve_month_values(matrix_rows, alm_sol, alm_chris) -> dict:
    """Compute full rendición per month using matrix formulas.

    Regla Sol: no cobrarle nada a Chris en los meses anteriores al primero
    en que Chris pasó gastos de almacén (ene–abr 2026 = sin cargo).
    """
    first_chris_month = next((m for m in MONTHS if alm_chris.get(m, 0) > 0), None)
    out = {}
    for mes in MONTHS:
        conceptos = []
        total = 0
        for row in matrix_rows:
            if row["es_almacen"]:
                if "SOL" in (row["descripcion"] or "").upper():
                    val = alm_sol.get(mes, 0)
                else:
                    val = alm_chris.get(mes, 0)
            else:
                raw = row["valores"].get(mes)
                val = money(raw) if raw is not None and not isinstance(raw, str) else 0
            if val:
                conceptos.append(
                    {
                        "tipo": row["tipo"],
                        "categoria": row["categoria"],
                        "descripcion": row["descripcion"],
                        "monto": val,
                    }
                )
                total += val

        alm_s = alm_sol.get(mes, 0)
        alm_c = alm_chris.get(mes, 0)
        # Ensure ALMACEN rows appear even if matrix row exists with formula
        has_sol = any("SOL" in c["descripcion"].upper() and c["categoria"] == "ALMACEN" for c in conceptos)
        has_chris = any("CHRIS" in c["descripcion"].upper() and c["categoria"] == "ALMACEN" for c in conceptos)
        if alm_s and not has_sol:
            conceptos.append(
                {
                    "tipo": "SERVICIOS",
                    "categoria": "ALMACEN",
                    "descripcion": "ALMACEN SOL",
                    "monto": alm_s,
                }
            )
            total += alm_s
        if alm_c and not has_chris:
            conceptos.append(
                {
                    "tipo": "SERVICIOS",
                    "categoria": "ALMACEN",
                    "descripcion": "ALMACEN CHRIS",
                    "monto": alm_c,
                }
            )
            total += alm_c

        # --- Reconciliación (como la explicó Sol) ---
        # 1) Servicios CIUDAD (sin almacén) ÷ 2 = cuota de cada uno
        # 2) Almacén también se parte 50/50
        # 3) A cada uno se le resta lo que ya pagó (Sol: todos los servicios + su almacén;
        #    Chris: solo su almacén) → queda quién le debe a quién
        servicios = total - alm_s - alm_c
        almacen_total = alm_s + alm_c
        cuota_servicios = settle(servicios / 2) if servicios else 0
        cuota_almacen = settle(almacen_total / 2) if almacen_total else 0
        cada_uno = settle((servicios + almacen_total) / 2) if (servicios + almacen_total) else 0
        # sanity: cada_uno ≈ cuota_servicios + cuota_almacen (puede diferir $1 por redondeo)
        pagado_sol = servicios + alm_s  # Sol paga todos los servicios + su almacén
        pagado_chris = alm_c
        pendiente_sol = cada_uno - pagado_sol  # + Sol le debe a Chris; − a favor de Sol
        pendiente_chris = cada_uno - pagado_chris  # + Chris le debe a Sol

        # Antes del primer mes con almacén Chris: no cobrarle nada a Chris
        sin_cargo_chris = bool(
            first_chris_month and MONTHS.index(mes) < MONTHS.index(first_chris_month)
        )

        if sin_cargo_chris:
            veredicto = "Sin cargo a Chris (aún no pasó almacén)"
            transferencia = 0
            direccion = "sin cargo"
            pendiente_sol = 0
            pendiente_chris = 0
        elif pendiente_chris > 0 and pendiente_sol <= 0:
            veredicto = f"Chris le debe a Sol {fmt_ars(pendiente_chris)}"
            transferencia = pendiente_chris
            direccion = "Chris→Sol"
        elif pendiente_sol > 0 and pendiente_chris <= 0:
            veredicto = f"Sol le debe a Chris {fmt_ars(pendiente_sol)}"
            transferencia = pendiente_sol
            direccion = "Sol→Chris"
        elif pendiente_chris == 0 and pendiente_sol == 0:
            veredicto = "Liquidado — sin saldo"
            transferencia = 0
            direccion = "—"
        else:
            # ambos mismos signo por redondeo raro: usar el de mayor magnitud hacia Sol/Chris
            if abs(pendiente_chris) >= abs(pendiente_sol) and pendiente_chris > 0:
                veredicto = f"Chris le debe a Sol {fmt_ars(pendiente_chris)}"
                transferencia = pendiente_chris
                direccion = "Chris→Sol"
            elif pendiente_sol > 0:
                veredicto = f"Sol le debe a Chris {fmt_ars(pendiente_sol)}"
                transferencia = pendiente_sol
                direccion = "Sol→Chris"
            else:
                veredicto = "Liquidado / a favor de Sol"
                transferencia = abs(pendiente_sol)
                direccion = "a favor Sol"

        out[mes] = {
            "conceptos": conceptos,
            "total": total,
            "servicios": servicios,
            "almacen_sol": alm_s,
            "almacen_chris": alm_c,
            "almacen_total": almacen_total,
            "cuota_servicios": cuota_servicios,
            "cuota_almacen": cuota_almacen,
            "cada_uno": cada_uno,
            "pagado_sol": pagado_sol,
            "pagado_chris": pagado_chris,
            "pendiente_sol": pendiente_sol,
            "pendiente_chris": pendiente_chris,
            "transferencia": transferencia,
            "direccion": direccion,
            "veredicto": veredicto,
            "sin_cargo_chris": sin_cargo_chris,
            "primer_mes_chris": first_chris_month,
            "tiene_datos": total > 0,
            "estado": (
                "completo"
                if alm_s or alm_c or any(
                    c["categoria"] not in ("TV CABLE", "ALMACEN") for c in conceptos
                )
                else ("solo_telecentro" if total > 0 else "vacio")
            ),
        }
    return out


def style_header(cell, fill_color=NAVY):
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=fill_color)
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = THIN


def write_money(ws, row, col, value, bold=False, color=None, round_mode="ceil"):
    amount = settle(value) if round_mode == "settle" else money(value)
    cell = ws.cell(row, col, amount)
    cell.number_format = '"$"#,##0'
    cell.border = THIN
    cell.alignment = Alignment(horizontal="right")
    if bold:
        cell.font = Font(bold=True, color=color) if color else Font(bold=True)
    elif color:
        cell.font = Font(color=color)
    return cell


def build_excel(rendicion, fam, items) -> None:
    wb = Workbook()

    # --- Resumen ---
    ws = wb.active
    ws.title = "Resumen"
    ws["A1"] = "CIUDAD 1132 — Rendición 2026 (casa + almacén)"
    ws["A1"].font = Font(size=16, bold=True, color=NAVY)
    ws["A2"] = (
        f"Generado {date.today().isoformat()} · Fuente: Gastos_CIUDAD_1132_2026.xlsx · "
        "Servicios÷2 + Almacén÷2 − lo que cada uno ya pagó → quién le debe a quién"
    )
    ws["A2"].font = Font(italic=True, color="666666")

    headers = [
        "Mes",
        "Estado",
        "Servicios",
        "Cuota serv. (÷2)",
        "Almacén Sol",
        "Almacén Chris",
        "Cuota alm. (÷2)",
        "A cargo c/u",
        "Pagó Sol",
        "Pagó Chris",
        "Reconciliación",
    ]
    for c, h in enumerate(headers, 1):
        style_header(ws.cell(4, c, h))

    estado_label = {
        "completo": "Completo",
        "solo_telecentro": "Parcial (solo Telecentro proyectado)",
        "vacio": "Sin datos",
    }

    row = 5
    tot_serv = tot_as = tot_ac = tot_t = tot_tr = 0
    for mes in MONTHS:
        d = rendicion[mes]
        if d["estado"] == "solo_telecentro":
            veredicto = "Mes incompleto — no liquidar aún"
        else:
            veredicto = d.get("veredicto", "—")

        ws.cell(row, 1, mes).border = THIN
        ws.cell(row, 2, estado_label[d["estado"]]).border = THIN
        write_money(ws, row, 3, d["servicios"])
        write_money(ws, row, 4, d["cuota_servicios"], round_mode="settle")
        write_money(ws, row, 5, d["almacen_sol"])
        write_money(ws, row, 6, d["almacen_chris"])
        write_money(ws, row, 7, d["cuota_almacen"], round_mode="settle")
        write_money(ws, row, 8, d["cada_uno"], bold=True, round_mode="settle")
        write_money(ws, row, 9, d["pagado_sol"])
        write_money(ws, row, 10, d["pagado_chris"])
        cell = ws.cell(row, 11, veredicto)
        cell.border = THIN
        if d["estado"] != "completo":
            for c in range(1, 12):
                ws.cell(row, c).fill = PatternFill("solid", fgColor=GRAY)

        tot_serv += d["servicios"]
        tot_as += d["almacen_sol"]
        tot_ac += d["almacen_chris"]
        tot_t += d["total"]
        if d["estado"] == "completo" and d["direccion"] == "Chris→Sol":
            tot_tr += d["transferencia"]
        row += 1

    ws.cell(row, 1, "TOTAL 2026").font = Font(bold=True)
    ws.cell(row, 1).border = THIN
    ws.cell(row, 2).border = THIN
    write_money(ws, row, 3, tot_serv, bold=True)
    write_money(ws, row, 4, tot_serv / 2, bold=True, round_mode="settle")
    write_money(ws, row, 5, tot_as, bold=True)
    write_money(ws, row, 6, tot_ac, bold=True)
    write_money(ws, row, 7, (tot_as + tot_ac) / 2, bold=True, round_mode="settle")
    write_money(ws, row, 8, tot_t / 2, bold=True, round_mode="settle")
    for c in range(9, 11):
        ws.cell(row, c).border = THIN
        ws.cell(row, c).fill = PatternFill("solid", fgColor=LIGHT)
    ws.cell(row, 11, f"Chris → Sol (desde 1er mes con almacén Chris): {fmt_ars(tot_tr)}").border = THIN
    ws.cell(row, 11).font = Font(bold=True, color=RED)

    row += 2
    ws.cell(row, 1, "Cómo se calcula").font = Font(bold=True, color=NAVY)
    notes = [
        "1) Sumar servicios CIUDAD del mes (expensas, luz, gas, limpieza, etc. — sin almacén) y dividir por 2 = cuota de cada uno.",
        "2) Sumar almacén Sol + almacén Chris y dividir por 2 = cuota de almacén de cada uno.",
        "3) A cargo de cada uno = cuota servicios + cuota almacén.",
        "4) Restar lo que ya pagó: Sol pagó todos los servicios + su almacén; Chris pagó solo su almacén.",
        "5) Reconciliación: si el saldo de Chris es positivo → Chris le debe a Sol; si el de Sol es positivo → Sol le debe a Chris.",
        "6) Enero–abril: NO se le cobra nada a Chris (todavía no había pasado gastos de almacén). La deuda arranca en mayo.",
        "Agosto–diciembre: almacén vacío / Telecentro proyectado — no liquidar hasta completar el mes.",
    ]
    for note in notes:
        row += 1
        ws.cell(row, 1, f"• {note}")

    widths = [12, 34, 12, 14, 12, 12, 14, 12, 12, 12, 36]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # --- Almacén Sol vs Chris (por mes) ---
    wsc = wb.create_sheet("Almacén Sol vs Chris")
    wsc["A1"] = "Almacén — Sol vs Chris por mes"
    wsc["A1"].font = Font(size=14, bold=True, color=TEAL)
    for c, h in enumerate(["Mes", "Almacén Sol", "Almacén Chris", "Total almacén"], 1):
        style_header(wsc.cell(3, c, h), TEAL)
    r = 4
    sum_s = sum_c = 0
    for mes in MONTHS:
        d = rendicion[mes]
        wsc.cell(r, 1, mes).border = THIN
        write_money(wsc, r, 2, d["almacen_sol"])
        write_money(wsc, r, 3, d["almacen_chris"])
        write_money(wsc, r, 4, d["almacen_sol"] + d["almacen_chris"], bold=True)
        sum_s += d["almacen_sol"]
        sum_c += d["almacen_chris"]
        r += 1
    wsc.cell(r, 1, "TOTAL").font = Font(bold=True)
    wsc.cell(r, 1).border = THIN
    write_money(wsc, r, 2, sum_s, bold=True)
    write_money(wsc, r, 3, sum_c, bold=True)
    write_money(wsc, r, 4, sum_s + sum_c, bold=True)
    for i, w in enumerate([14, 14, 14, 14], 1):
        wsc.column_dimensions[get_column_letter(i)].width = w

    # --- Almacén por familia (Sol | Chris | Total) ---
    wsf = wb.create_sheet("Almacén por familia")
    wsf["A1"] = "Almacén — por familia (Sol / Chris / Total)"
    wsf["A1"].font = Font(size=14, bold=True, color=TEAL)
    families = [
        "SUPERMERCADO",
        "COMIDA",
        "RAMÓN",
        "COMBUSTIBLE",
        "ENTRETENIMIENTO",
        "STREAMING TV",
        "FARMACIA",
        "REPUESTOS",
        "OTROS",
    ]
    # Header: Familia | then for each month Sol/Chris/Total — too wide.
    # Instead: Familia | Total Sol | Total Chris | Total | then monthly total
    style_header(wsf.cell(3, 1, "Familia"), TEAL)
    style_header(wsf.cell(3, 2, "Almacén Sol"), TEAL)
    style_header(wsf.cell(3, 3, "Almacén Chris"), TEAL)
    style_header(wsf.cell(3, 4, "Total"), TEAL)
    for i, mes in enumerate(MONTHS):
        style_header(wsf.cell(3, i + 5, mes), TEAL)

    r = 4
    col_month_totals = [0] * 12
    tot_sol = tot_chris = 0
    for fam_name in families:
        wsf.cell(r, 1, fam_name).border = THIN
        row_sol = sum(fam[m].get(fam_name, {}).get("sol", 0) for m in MONTHS)
        row_chris = sum(fam[m].get(fam_name, {}).get("chris", 0) for m in MONTHS)
        row_tot = row_sol + row_chris
        write_money(wsf, r, 2, row_sol)
        write_money(wsf, r, 3, row_chris)
        write_money(wsf, r, 4, row_tot, bold=True)
        tot_sol += row_sol
        tot_chris += row_chris
        for i, mes in enumerate(MONTHS):
            v = fam[mes].get(fam_name, {}).get("total", 0)
            write_money(wsf, r, i + 5, v)
            col_month_totals[i] += v
        r += 1
    wsf.cell(r, 1, "TOTAL").font = Font(bold=True)
    wsf.cell(r, 1).border = THIN
    write_money(wsf, r, 2, tot_sol, bold=True)
    write_money(wsf, r, 3, tot_chris, bold=True)
    write_money(wsf, r, 4, tot_sol + tot_chris, bold=True)
    for i, v in enumerate(col_month_totals):
        write_money(wsf, r, i + 5, v, bold=True)
    wsf.column_dimensions["A"].width = 18
    for i in range(2, 17):
        wsf.column_dimensions[get_column_letter(i)].width = 12

    # --- Una hoja por mes con datos ---
    for mes in MONTHS:
        d = rendicion[mes]
        if not d["tiene_datos"] and not items[mes]:
            continue
        wsm = wb.create_sheet(mes[:31])
        wsm["A1"] = f"Rendición {mes} 2026 — CIUDAD 1132"
        wsm["A1"].font = Font(size=14, bold=True, color=NAVY)
        wsm["A2"] = (
            f"TOTAL {fmt_ars(d['total'])} · Cada uno {fmt_ars(d['cada_uno'])} · "
            f"Pagó Sol {fmt_ars(d['pagado_sol'])} · Pagó Chris {fmt_ars(d['pagado_chris'])}"
        )

        # Liquidación box — pasos de reconciliación
        wsm["A4"] = "Reconciliación Sol / Chris"
        wsm["A4"].font = Font(bold=True, color=NAVY)
        wsm["A5"] = (
            "1) Servicios ÷ 2  →  2) + Almacén ÷ 2  →  3) − lo que cada uno ya pagó  →  quién le debe a quién"
        )
        wsm["A5"].font = Font(italic=True, color="666666")
        for c, h in enumerate(["Paso", "Sol", "Chris", "Total"], 1):
            style_header(wsm.cell(6, c, h))
        liq_rows = [
            ("1. Servicios CIUDAD (sin almacén)", d["servicios"], 0, d["servicios"], False),
            ("2. Cuota servicios (÷2)", d["cuota_servicios"], d["cuota_servicios"], d["servicios"], False),
            ("3. Almacén pagado", d["almacen_sol"], d["almacen_chris"], d["almacen_total"], False),
            ("4. Cuota almacén (÷2)", d["cuota_almacen"], d["cuota_almacen"], d["almacen_total"], False),
            (
                "5. A cargo de cada uno (servicios÷2 + almacén÷2)",
                d["cada_uno"],
                d["cada_uno"],
                d["total"],
                False,
            ),
            (
                "6. Ya pagó (Sol: servicios+su almacén / Chris: su almacén)",
                d["pagado_sol"],
                d["pagado_chris"],
                d["total"],
                False,
            ),
            (
                "7. Saldo (+ debe / − a favor)",
                d["pendiente_sol"],
                d["pendiente_chris"],
                None,
                True,
            ),
        ]
        rr = 7
        for label, sol_v, chris_v, tot_v, highlight in liq_rows:
            wsm.cell(rr, 1, label).border = THIN
            write_money(wsm, rr, 2, sol_v, bold=highlight, round_mode="settle")
            write_money(wsm, rr, 3, chris_v, bold=highlight, round_mode="settle")
            if tot_v is None:
                wsm.cell(rr, 4).border = THIN
            else:
                write_money(wsm, rr, 4, tot_v, bold=highlight, round_mode="settle")
            if highlight:
                for c in range(1, 5):
                    wsm.cell(rr, c).fill = PatternFill("solid", fgColor=AMBER)
            rr += 1

        wsm.cell(rr, 1, f"→ {d['veredicto']}").font = Font(bold=True, color=TEAL)

        # Conceptos casa
        rr += 2
        wsm.cell(rr, 1, "Conceptos de la casa (matriz 2026)").font = Font(bold=True, color=NAVY)
        rr += 1
        for c, h in enumerate(["Tipo", "Categoría", "Descripción", "Monto"], 1):
            style_header(wsm.cell(rr, c, h))
        rr += 1
        for cpto in d["conceptos"]:
            wsm.cell(rr, 1, cpto["tipo"]).border = THIN
            wsm.cell(rr, 2, cpto["categoria"]).border = THIN
            wsm.cell(rr, 3, cpto["descripcion"]).border = THIN
            write_money(wsm, rr, 4, cpto["monto"])
            if cpto["categoria"] == "ALMACEN":
                for c in range(1, 5):
                    wsm.cell(rr, c).fill = PatternFill("solid", fgColor=LIGHT_TEAL)
            rr += 1
        wsm.cell(rr, 1, "TOTAL").font = Font(bold=True)
        wsm.cell(rr, 1).border = THIN
        wsm.cell(rr, 2).border = THIN
        wsm.cell(rr, 3).border = THIN
        write_money(wsm, rr, 4, d["total"], bold=True)

        # Almacén familia — columnas Sol / Chris
        rr += 2
        wsm.cell(rr, 1, "Almacén — por familia").font = Font(bold=True, color=TEAL)
        rr += 1
        for c, h in enumerate(["Familia", "Almacén Sol", "Almacén Chris", "Total"], 1):
            style_header(wsm.cell(rr, c, h), TEAL)
        rr += 1
        for fam_name, vals in sorted(
            fam[mes].items(), key=lambda x: -x[1].get("total", 0)
        ):
            wsm.cell(rr, 1, fam_name).border = THIN
            write_money(wsm, rr, 2, vals.get("sol", 0))
            write_money(wsm, rr, 3, vals.get("chris", 0))
            write_money(wsm, rr, 4, vals.get("total", 0), bold=True)
            rr += 1
        if fam[mes]:
            wsm.cell(rr, 1, "TOTAL ALMACÉN").font = Font(bold=True)
            wsm.cell(rr, 1).border = THIN
            write_money(wsm, rr, 2, d["almacen_sol"], bold=True)
            write_money(wsm, rr, 3, d["almacen_chris"], bold=True)
            write_money(wsm, rr, 4, d["almacen_sol"] + d["almacen_chris"], bold=True)
            rr += 1

        # Detalle almacén
        if items[mes]:
            rr += 1
            wsm.cell(rr, 1, "Almacén — detalle día a día").font = Font(bold=True, color=TEAL)
            rr += 1
            for c, h in enumerate(
                ["Fecha", "Familia", "Detalle", "Importe Sol", "Importe Chris", "Total", "Comentario"], 1
            ):
                style_header(wsm.cell(rr, c, h), TEAL)
            rr += 1
            def _fecha_key(it):
                f = it["fecha"]
                if f is None:
                    return date.min
                if hasattr(f, "date"):
                    return f.date()
                return f

            for it in sorted(items[mes], key=lambda x: (_fecha_key(x), x["detalle"])):
                fecha = _fecha_key(it)
                if fecha == date.min:
                    fecha = None
                wsm.cell(rr, 1, fecha).border = THIN
                wsm.cell(rr, 1).number_format = "YYYY-MM-DD"
                wsm.cell(rr, 2, it["familia"]).border = THIN
                wsm.cell(rr, 3, it["detalle"]).border = THIN
                write_money(wsm, rr, 4, it["sol"])
                write_money(wsm, rr, 5, it["chris"])
                write_money(wsm, rr, 6, it["total"])
                wsm.cell(rr, 7, it["comentario"]).border = THIN
                rr += 1

        wsm.column_dimensions["A"].width = 14
        wsm.column_dimensions["B"].width = 16
        wsm.column_dimensions["C"].width = 42
        wsm.column_dimensions["D"].width = 14
        wsm.column_dimensions["E"].width = 14
        wsm.column_dimensions["F"].width = 12
        wsm.column_dimensions["G"].width = 50

    wb.save(OUT_XLSX)


def build_markdown(rendicion, fam) -> None:
    lines = [
        "# Rendición CIUDAD 1132 — 2026",
        "",
        f"Generado: **{date.today().isoformat()}** · Fuente: `Gastos_CIUDAD_1132_2026.xlsx`",
        "",
        "## Cómo se calcula la reconciliación",
        "",
        "1. **Servicios CIUDAD** (sin almacén) ÷ 2 = cuota de cada uno",
        "2. **Almacén Sol + Almacén Chris** ÷ 2 = cuota de almacén de cada uno",
        "3. **A cargo de cada uno** = cuota servicios + cuota almacén",
        "4. **Restar lo ya pagado**: Sol = todos los servicios + su almacén; Chris = solo su almacén",
        "5. **Reconciliación**: si Chris quedó debiendo → le paga a Sol; si Sol quedó debiendo → le paga a Chris",
        "6. **Enero–abril: sin cargo a Chris** (aún no había pasado almacén). La deuda arranca en **mayo**.",
        "",
        "## Resumen anual",
        "",
        "| Mes | Estado | Servicios | Cuota serv. | Alm. Sol | Alm. Chris | A cargo c/u | Pagó Sol | Pagó Chris | Reconciliación |",
        "|---|---|---:|---:|---:|---:|---:|---:|---:|---|",
    ]
    estado_label = {
        "completo": "Completo",
        "solo_telecentro": "Parcial",
        "vacio": "Sin datos",
    }
    tot_s = tot_as = tot_ac = tot_t = tot_tr = 0
    for mes in MONTHS:
        d = rendicion[mes]
        if d["estado"] == "solo_telecentro":
            verd = "no liquidar aún"
        else:
            verd = d.get("veredicto", "—")
        lines.append(
            f"| {mes} | {estado_label[d['estado']]} | {fmt_ars(d['servicios'])} | "
            f"{fmt_ars(d['cuota_servicios'])} | {fmt_ars(d['almacen_sol'])} | "
            f"{fmt_ars(d['almacen_chris'])} | {fmt_ars(d['cada_uno'])} | "
            f"{fmt_ars(d['pagado_sol'])} | {fmt_ars(d['pagado_chris'])} | {verd} |"
        )
        tot_s += d["servicios"]
        tot_as += d["almacen_sol"]
        tot_ac += d["almacen_chris"]
        tot_t += d["total"]
        if d["estado"] == "completo" and d.get("direccion") == "Chris→Sol":
            tot_tr += d["transferencia"]
    lines.append(
        f"| **TOTAL liquidable*** | | **{fmt_ars(sum(rendicion[m]['servicios'] for m in MONTHS if rendicion[m].get('direccion')=='Chris→Sol'))}** | "
        f"| **{fmt_ars(sum(rendicion[m]['almacen_sol'] for m in MONTHS if rendicion[m].get('direccion')=='Chris→Sol'))}** | "
        f"**{fmt_ars(sum(rendicion[m]['almacen_chris'] for m in MONTHS if rendicion[m].get('direccion')=='Chris→Sol'))}** | "
        f"| | | **Chris → Sol {fmt_ars(tot_tr)}** |"
    )
    lines += [
        "",
        "\\* Solo meses con cargo a Chris (mayo–julio). Enero–abril: sin cargo.",
        "",
        f"**Chris le debe a Sol (desde mayo):** {fmt_ars(tot_tr)}",
        "",
        "## Detalle por mes",
        "",
    ]

    for mes in MONTHS:
        d = rendicion[mes]
        if d["estado"] == "vacio":
            lines += [f"### {mes}", "", "_Sin gastos cargados._", ""]
            continue
        if d["estado"] == "solo_telecentro":
            lines += [
                f"### {mes}",
                "",
                f"_Mes incompleto — solo Telecentro proyectado {fmt_ars(d['total'])}. No liquidar aún._",
                "",
            ]
            continue
        lines += [
            f"### {mes}",
            "",
            "| Paso | Sol | Chris | Total |",
            "|---|---:|---:|---:|",
            f"| 1. Servicios CIUDAD | | | {fmt_ars(d['servicios'])} |",
            f"| 2. Cuota servicios (÷2) | {fmt_ars(d['cuota_servicios'])} | {fmt_ars(d['cuota_servicios'])} | {fmt_ars(d['servicios'])} |",
            f"| 3. Almacén pagado | {fmt_ars(d['almacen_sol'])} | {fmt_ars(d['almacen_chris'])} | {fmt_ars(d['almacen_total'])} |",
            f"| 4. Cuota almacén (÷2) | {fmt_ars(d['cuota_almacen'])} | {fmt_ars(d['cuota_almacen'])} | {fmt_ars(d['almacen_total'])} |",
            f"| 5. A cargo de cada uno | {fmt_ars(d['cada_uno'])} | {fmt_ars(d['cada_uno'])} | {fmt_ars(d['total'])} |",
            f"| 6. Ya pagó | {fmt_ars(d['pagado_sol'])} | {fmt_ars(d['pagado_chris'])} | {fmt_ars(d['total'])} |",
            f"| 7. Saldo (+ debe / − a favor) | {fmt_ars(d['pendiente_sol'])} | {fmt_ars(d['pendiente_chris'])} | |",
            "",
            f"**Reconciliación:** {d['veredicto']}",
            "",
        ]
        if fam[mes]:
            lines += [
                "#### Almacén por familia",
                "",
                "| Familia | Almacén Sol | Almacén Chris | Total |",
                "|---|---:|---:|---:|",
            ]
            for fam_name, vals in sorted(fam[mes].items(), key=lambda x: -x[1].get("total", 0)):
                lines.append(
                    f"| {fam_name} | {fmt_ars(vals.get('sol', 0))} | "
                    f"{fmt_ars(vals.get('chris', 0))} | {fmt_ars(vals.get('total', 0))} |"
                )
            lines.append(
                f"| **TOTAL** | **{fmt_ars(d['almacen_sol'])}** | **{fmt_ars(d['almacen_chris'])}** | "
                f"**{fmt_ars(d['almacen_total'])}** |"
            )
            lines.append("")

    lines += [
        "## Fuera de esta rendición",
        "",
        "- Gasto Personal Sol",
        "- MP No Gasto / inversión / retiros",
        "- Propiedades BONORINO / OHIGGINS / AVA",
        "",
        "## Regenerar",
        "",
        "```bash",
        "python3 CIUDAD/build_rendicion_ciudad.py",
        "```",
        "",
    ]
    OUT_MD.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    wb = load_workbook(SOURCE, data_only=False)
    matrix_rows = load_matrix_rows(wb)
    alm_sol, alm_chris, fam, items = load_almacen(wb)
    rendicion = resolve_month_values(matrix_rows, alm_sol, alm_chris)

    build_excel(rendicion, fam, items)
    build_markdown(rendicion, fam)

    print(f"Wrote {OUT_XLSX.relative_to(ROOT.parent)}")
    print(f"Wrote {OUT_MD.relative_to(ROOT.parent)}")
    print()
    print(f"{'Mes':<12} {'Total':>14} {'Cada uno':>14} {'Chris→Sol':>14}")
    for mes in MONTHS:
        d = rendicion[mes]
        if not d["tiene_datos"]:
            continue
        cargo = d["transferencia"] if d.get("direccion") == "Chris→Sol" else 0
        print(
            f"{mes:<12} {fmt_ars(d['total']):>14} {fmt_ars(d['cada_uno']):>14} "
            f"{fmt_ars(cargo):>14}"
        )


if __name__ == "__main__":
    main()
