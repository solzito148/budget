#!/usr/bin/env python3
"""Verificación Gmail Compra aprobada BBVA + cuotas BBVA/AVA + Chris no pagó septiembre.

Fuente oficial: alertas BBVA 2026-08-01..2026-10-23 y mails Ava Palpa.
Base: Gastos_CIUDAD_1132_2026.xlsx (Drive 2026-10-12) ya copiado al canónico.
"""
from __future__ import annotations

import csv
import math
from copy import copy
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path("/workspace")
CIUDAD = ROOT / "CIUDAD" / "Gastos_CIUDAD_1132_2026.xlsx"
AVA_XLSX = ROOT / "AVA" / "Gastos_AVA_2026.xlsx"
AVA_CSV = ROOT / "AVA" / "cuotas_fideicomiso_ava.csv"
AVA_MD = ROOT / "AVA" / "RESUMEN_CUOTAS_FIDEICOMISO.md"
CASHFLOW_CSV = ROOT / "CIUDAD" / "_inbox" / "tarjetas_bbva" / "cashflow_cuotas_bbva.csv"
OHIGGINS = ROOT / "OHIGGINS" / "Gastos_OHIGGINS_2026.xlsx"


def ceil_ars(x: float) -> int:
    return int(math.ceil(float(x) - 1e-9))


def month_name(d: date) -> str:
    names = [
        "",
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
    return names[d.month]


def copy_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def next_empty_row(ws, start: int, key_col: int = 1) -> int:
    r = max(start, ws.max_row + 1)
    # prefer first fully empty row at/after start within used range
    for i in range(start, ws.max_row + 2):
        if all(ws.cell(i, c).value is None for c in range(1, 7)):
            return i
    return r


def append_gps(ws, fecha: date, concepto: str, monto, medio: str, comentario: str):
    # Find last used data row (>=26)
    last = 25
    for r in range(26, ws.max_row + 1):
        if ws.cell(r, 1).value is not None or ws.cell(r, 3).value is not None:
            last = r
    r = last + 1
    template = ws.cell(last, 1) if last >= 26 else None
    values = [datetime(fecha.year, fecha.month, fecha.day), month_name(fecha), concepto, monto, medio, comentario]
    for c, v in enumerate(values, 1):
        cell = ws.cell(r, c, v)
        if template is not None:
            copy_style(ws.cell(last, c), cell)
    if isinstance(values[0], datetime):
        ws.cell(r, 1).number_format = "YYYY-MM-DD"
    return r


def find_almacen_slot(ws, mes: str) -> int:
    """First empty Octubre/Septiembre/Agosto data row (has mes, empty fecha)."""
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 1).value == mes and ws.cell(r, 2).value is None and ws.cell(r, 3).value is None:
            return r
    raise RuntimeError(f"No empty Almacen slot for {mes}")


def write_almacen(ws, mes: str, fecha: date, detalle: str, importe_sol: int, comentario: str = ""):
    r = find_almacen_slot(ws, mes)
    ws.cell(r, 1, mes)
    ws.cell(r, 2, datetime(fecha.year, fecha.month, fecha.day))
    ws.cell(r, 2).number_format = "YYYY-MM-DD"
    ws.cell(r, 3, detalle)
    ws.cell(r, 4, importe_sol)
    if comentario:
        ws.cell(r, 7, comentario)
    return r


def append_detalle(ws, fecha: date, tipo: str, desc: str, monto: int, forma: str, cliente, notas: str, mes: str):
    last = 3
    for r in range(4, ws.max_row + 1):
        if any(ws.cell(r, c).value is not None for c in range(1, 9)):
            last = r
    r = last + 1
    ws.cell(r, 1, datetime(fecha.year, fecha.month, fecha.day)).number_format = "YYYY-MM-DD"
    ws.cell(r, 2, tipo)
    ws.cell(r, 3, desc)
    ws.cell(r, 4, monto)
    ws.cell(r, 5, forma)
    if cliente:
        ws.cell(r, 6, cliente)
    ws.cell(r, 8, notas)
    ws.cell(r, 9, mes)
    return r


def mark_cashflow_paid(ws, months: set[str]) -> int:
    n = 0
    for r in range(2, ws.max_row + 1):
        am = str(ws.cell(r, 9).value or "")
        est = ws.cell(r, 10).value
        if am in months and est == "PENDIENTE":
            ws.cell(r, 10, "PAGADO")
            n += 1
    return n


def mark_compromisos_paid(ws, months: set[str]) -> int:
    n = 0
    for r in range(2, ws.max_row + 1):
        fuente = str(ws.cell(r, 1).value or "")
        am = str(ws.cell(r, 9).value or "")
        est = ws.cell(r, 10).value
        if "BBVA" in fuente and am in months and est == "PENDIENTE":
            ws.cell(r, 10, "PAGADO")
            n += 1
    return n


def update_gps_proyectado(ws, months_es: set[str]) -> int:
    n = 0
    for r in range(26, ws.max_row + 1):
        mes = ws.cell(r, 2).value
        medio = str(ws.cell(r, 5).value or "")
        com = str(ws.cell(r, 6).value or "")
        if mes in months_es and "proyectado" in medio.lower() and "PROYECTADO" in com:
            ws.cell(r, 5, medio.replace(" · proyectado", "").replace("· proyectado", "").strip())
            ws.cell(r, 6, com.replace("PROYECTADO · ", "COBRADO resumen BBVA · ").replace("PROYECTADO", "COBRADO"))
            n += 1
    return n


def update_cashflow_csv(months: set[str]) -> int:
    rows = []
    with CASHFLOW_CSV.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames
        n = 0
        for row in reader:
            if row.get("anio_mes") in months and row.get("estado") == "PENDIENTE":
                row["estado"] = "PAGADO"
                n += 1
            rows.append(row)
    with CASHFLOW_CSV.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)
    return n


def main():
    wb = load_workbook(CIUDAD)
    ws2026 = wb["2026"]
    wsc = wb["Cashflow Cuotas BBVA"]
    wscomp = wb["Cashflow Compromisos"]
    wsa = wb["Almacen Detalle"]
    wsd = wb["Detalle"]
    wsg = wb["Gasto Personal Sol"]

    paid_months = {"2026-07", "2026-08", "2026-09"}
    n_cf = mark_cashflow_paid(wsc, paid_months)
    n_comp = mark_compromisos_paid(wscomp, paid_months)
    n_csv = update_cashflow_csv(paid_months)
    n_gps_proj = update_gps_proyectado(wsg, {"Julio", "Agosto", "Septiembre"})

    # --- Chris no pagó septiembre ---
    note = (
        "• LIQUIDACIÓN SEPTIEMBRE 2026: Christian (Chris) NO pagó su parte de los gastos de la casa "
        "(Pagado Christian = $0 / sin ALMACEN CHRIS ni Extras Chris). "
        "Pendiente Chris = fila CHRIS columna Septiembre (Cada uno − ALMACEN CHRIS). "
        "Registrado 2026-10-23 por verificación Sol."
    )
    # find next empty note row
    note_row = None
    for r in range(27, 120):
        if ws2026.cell(r, 1).value is None:
            note_row = r
            break
    if note_row is None:
        note_row = ws2026.max_row + 1
    ws2026.cell(note_row, 1, note)

    # --- Agosto utilidades (mails BBVA) ---
    # Telecentro 21/08 82363.35
    ws2026.cell(6, 11, ceil_ars(82363.35))  # Ago TV
    # replace Detalle Aug placeholder 50000
    for r in range(4, wsd.max_row + 1):
        if wsd.cell(r, 2).value == "Telecentro" and wsd.cell(r, 9).value == "Agosto" and wsd.cell(r, 4).value == 50000:
            wsd.cell(r, 1, datetime(2026, 8, 21))
            wsd.cell(r, 3, "Telecentro CIUDAD — débito BBVA 21/08/2026")
            wsd.cell(r, 4, ceil_ars(82363.35))
            wsd.cell(r, 5, "Débito BBVA")
            wsd.cell(r, 8, "Gmail Compra aprobada BBVA · ARS 82.363,35 · mes=cobro")
            break
    # Metrogas Aug: 05/08 73694.19 + 17/08 73186.92
    gas_aug = ceil_ars(73694.19) + ceil_ars(73186.92)
    ws2026.cell(7, 11, gas_aug)
    append_detalle(
        wsd,
        date(2026, 8, 5),
        "Metrogas",
        "Gas Metrogas CIUDAD — débito BBVA 05/08/2026",
        ceil_ars(73694.19),
        "Débito BBVA",
        "30010996459",
        "Gmail Compra aprobada · medidor 40049923",
        "Agosto",
    )
    append_detalle(
        wsd,
        date(2026, 8, 17),
        "Metrogas",
        "Gas Metrogas CIUDAD — débito BBVA 17/08/2026",
        ceil_ars(73186.92),
        "Débito BBVA",
        "30010996459",
        "Gmail Compra aprobada · segundo cobro agosto",
        "Agosto",
    )
    # Luz CIUDAD 12/08 183271.35
    ws2026.cell(8, 11, ceil_ars(183271.35))
    append_detalle(
        wsd,
        date(2026, 8, 12),
        "Edenor",
        "Luz Edenor CIUDAD — débito BBVA 12/08/2026 · Sanes",
        ceil_ars(183271.35),
        "Débito BBVA",
        "00008917502309",
        "Gmail Compra aprobada BBVA",
        "Agosto",
    )
    # Spotify Aug 28/08 4399
    ws2026.cell(12, 11, 4399)
    append_detalle(
        wsd,
        date(2026, 8, 28),
        "Streaming",
        "Streaming música — Spotify · BBVA VISA 28/08/2026",
        4399,
        "BBVA VISA",
        None,
        "Gmail Compra aprobada · ARS 4.399",
        "Agosto",
    )

    # --- Octubre utilidades ---
    # Telecentro 23/10 70289 — replace placeholder 50000
    ws2026.cell(6, 13, ceil_ars(70289))
    for r in range(4, wsd.max_row + 1):
        if wsd.cell(r, 2).value == "Telecentro" and wsd.cell(r, 9).value == "Octubre" and wsd.cell(r, 4).value == 50000:
            wsd.cell(r, 1, datetime(2026, 10, 23))
            wsd.cell(r, 3, "Telecentro CIUDAD — débito BBVA 23/10/2026")
            wsd.cell(r, 4, ceil_ars(70289))
            wsd.cell(r, 5, "Débito BBVA")
            wsd.cell(r, 8, "Gmail Compra aprobada BBVA · ARS 70.289 · mes=cobro; deja de aplicar fijo $50.000")
            break
    # note Telecentro fijo
    for r in range(27, 120):
        v = ws2026.cell(r, 1).value
        if v and "Telecentro: $50.000" in str(v):
            ws2026.cell(
                r,
                1,
                "• Telecentro: cobros BBVA reales (sep $69.659 · oct $70.289). El fijo $50.000 ya no aplica desde sep-2026.",
            )
            break

    ws2026.cell(7, 13, ceil_ars(106266.4))  # Metrogas 18/10
    append_detalle(
        wsd,
        date(2026, 10, 18),
        "Metrogas",
        "Gas Metrogas CIUDAD — débito BBVA 18/10/2026",
        ceil_ars(106266.4),
        "Débito BBVA",
        "30010996459",
        "Gmail Compra aprobada BBVA",
        "Octubre",
    )
    ws2026.cell(8, 13, ceil_ars(66399.5))  # Edenor CIUDAD 12/10
    append_detalle(
        wsd,
        date(2026, 10, 12),
        "Edenor",
        "Luz Edenor CIUDAD — débito BBVA 12/10/2026 · Sanes",
        ceil_ars(66399.5),
        "Débito BBVA",
        "00008917502309",
        "Gmail Compra aprobada BBVA",
        "Octubre",
    )

    # --- Almacén: Aug Ramón + Oct Farmacity ---
    write_almacen(
        wsa,
        "Agosto",
        date(2026, 8, 1),
        "RAMON · Comida",
        ceil_ars(176000),
        "MERPAGO*DRBARF · BBVA VISA · Gmail Compra aprobada",
    )
    write_almacen(
        wsa,
        "Agosto",
        date(2026, 8, 22),
        "RAMON · Otros",
        ceil_ars(77200),
        "Paws Pet Market · BBVA VISA · Gmail Compra aprobada",
    )
    write_almacen(
        wsa,
        "Octubre",
        date(2026, 10, 17),
        "FARMACIA · Farmacity",
        ceil_ars(81924),
        "MERPAGO*FARMACITY · BBVA VISA · Gmail Compra aprobada",
    )

    # --- GPS October (y Aug faltantes claros) ---
    gps_rows = [
        # Aug
        (date(2026, 8, 28), "Eventos deportivos - River", 35430, "BBVA VISA", "Eventos deportivos - River · CLUB ATLETICO RIVE · Gmail Compra aprobada"),
        (date(2026, 8, 6), "SUSCRIPCIONES · IA", None, "BBVA VISA · USD", "SUSCRIPCIONES · IA · Google One / Gemini · USD 19.99 · alerta Gmail · ver resumen BBVA / Sol Dollar Mdp"),
        (date(2026, 8, 11), "SUSCRIPCIONES · IA", None, "BBVA VISA · USD", "SUSCRIPCIONES · IA · Vercel Domains · USD 11.25 · alerta Gmail · ver resumen BBVA"),
        (date(2026, 8, 24), "SUSCRIPCIONES · IA", None, "BBVA VISA · USD", "SUSCRIPCIONES · IA · Google One · USD 4.99 · alerta Gmail · ver resumen BBVA"),
        (date(2026, 8, 27), "SUSCRIPCIONES · IA", None, "BBVA VISA · USD", "SUSCRIPCIONES · IA · Cursor · USD 60.00 · alerta Gmail · ver resumen BBVA / ya hay fila si está cobrada en ARS"),
        (date(2026, 8, 31), "SUSCRIPCIONES · IA", None, "BBVA VISA · USD", "SUSCRIPCIONES · IA · ChatGPT / OpenAI · USD 20.00 · alerta Gmail · ver resumen BBVA"),
        (date(2026, 8, 8), "TRAÍDO DE MIAMI", None, "BBVA VISA · USD", "TRAÍDO DE MIAMI · DHgate · USD 196.75 · alerta Gmail · REVISAR vs resumen BBVA"),
        # Oct
        (date(2026, 10, 2), "Eventos deportivos - River", 36030, "BBVA VISA", "Eventos deportivos - River · CLUB ATLETICO RIVE · Gmail Compra aprobada"),
        (date(2026, 10, 15), "TRANSPORTE · Cabify", ceil_ars(10823.43), "BBVA VISA", "TRANSPORTE · Cabify · Gmail Compra aprobada"),
        (date(2026, 10, 15), "TRANSPORTE · Cabify", ceil_ars(13446.6), "BBVA VISA", "TRANSPORTE · Cabify · Gmail Compra aprobada"),
        (date(2026, 10, 1), "ROPA", ceil_ars(710739.16), "BBVA VISA", "ROPA · GlobalE / Onitsuka · Gmail Compra aprobada · REVISAR ítem"),
        (date(2026, 10, 7), "COMPRAS", ceil_ars(151999), "BBVA VISA", "COMPRAS · MERPAGO*DEERTECH · Gmail Compra aprobada · REVISAR ítem"),
        (date(2026, 10, 11), "COMPRAS", ceil_ars(1143000), "BBVA VISA", "COMPRAS · MERPAGO*PATAGONIAC · Gmail Compra aprobada · REVISAR ítem"),
        (date(2026, 10, 11), "COMPRAS", ceil_ars(402000), "BBVA VISA", "COMPRAS · BDP NIBEPO AIKE SC · Gmail Compra aprobada · REVISAR ítem"),
        (date(2026, 10, 17), "COMPRAS", ceil_ars(199900), "BBVA VISA", "COMPRAS · MERPAGO*DOTECOMSAS · Gmail Compra aprobada · REVISAR ítem"),
        (date(2026, 10, 20), "PROFESIONAL", ceil_ars(54700), "BBVA VISA", "PROFESIONAL · CPCECABA DEP-DAU · Gmail Compra aprobada"),
    ]
    # Avoid duplicating Cursor/OpenAI if already present as Sep ARS conversions — Aug USD stubs OK
    for fecha, concepto, monto, medio, com in gps_rows:
        # skip if exact date+concepto+monto already exists
        exists = False
        for r in range(26, wsg.max_row + 1):
            d = wsg.cell(r, 1).value
            if isinstance(d, datetime) and d.date() == fecha and wsg.cell(r, 3).value == concepto:
                m = wsg.cell(r, 4).value
                if monto is None or m == monto or (isinstance(m, (int, float)) and monto and abs(m - monto) <= 1):
                    exists = True
                    break
        if not exists:
            append_gps(wsg, fecha, concepto, monto if monto is not None else None, medio, com)

    wb.save(CIUDAD)

    # --- OHIGGINS Edenor small amounts ---
    woh = load_workbook(OHIGGINS)
    mat = woh["2026"]
    det = woh["Detalle"]
    # col mapping: D=Ene ... K=Ago(11), L=Sep(12), M=Oct(13) — check headers
    headers = {mat.cell(4, c).value: c for c in range(4, 16)}
    # Agosto Edenor Chagas 09/08 19492; Sep 09/09 18236; Oct 09/10 12004
    oh_luz = [
        (date(2026, 8, 9), "Agosto", ceil_ars(19491.89)),
        (date(2026, 9, 9), "Septiembre", ceil_ars(18235.37)),
        (date(2026, 10, 9), "Octubre", ceil_ars(12003.68)),
    ]
    for fecha, mes, monto in oh_luz:
        col = headers.get(mes)
        if col:
            prev = mat.cell(8, col).value  # LUZ row
            mat.cell(8, col, (prev or 0) + monto if isinstance(prev, (int, float)) else monto)
        # detalle dedupe
        exists = False
        for r in range(4, det.max_row + 1):
            d = det.cell(r, 1).value
            if isinstance(d, datetime) and d.date() == fecha and det.cell(r, 4).value == monto:
                exists = True
                break
        if not exists:
            r = det.max_row + 1
            det.cell(r, 1, datetime(fecha.year, fecha.month, fecha.day)).number_format = "YYYY-MM-DD"
            det.cell(r, 2, "SERVICIOS")
            det.cell(r, 3, "LUZ")
            det.cell(r, 4, f"Edenor — Marta Chagas · O'Higgins · débito BBVA {fecha.strftime('%d/%m/%Y')}")
            det.cell(r, 5, monto)
            det.cell(r, 6, "Débito BBVA Visa")
            det.cell(r, 7, "9267222037 · Gmail Compra aprobada")
    woh.save(OHIGGINS)

    # --- AVA cuotas 10-13 ---
    # MEP Rava cierre (alineado a práctica previa sep=1525.39)
    ava_cuotas = [
        # n, mes, monto, vto, mep, mep_fecha, estado, comentario
        (10, "agosto 2026", 6178522, date(2026, 8, 5), 1518.45, date(2026, 8, 5), "PAGADO", "Gmail Sancor 22/07 · vto 05/08 · MEP Rava cierre"),
        (11, "septiembre 2026", 6269440, date(2026, 9, 4), 1525.39, date(2026, 9, 4), "PAGADO", "Gmail Sancor 25/08 · vto 04/09 · MEP Rava · ya en GPS (renum cuota 11)"),
        (12, "octubre 2026", 6550413, date(2026, 10, 5), 1538.39, date(2026, 10, 5), "PAGADO", "Gmail Sancor 24/09 · vto 05/10 · MEP Rava cierre"),
        (13, "noviembre 2026", 6739503, date(2026, 11, 4), None, None, "PENDIENTE", "Gmail Sancor 23/10 · vto 04/11 · TC MEP pendiente hasta rueda del vencimiento"),
    ]

    # Fix GPS Sept AVA comment cuota 10 → 11
    wb2 = load_workbook(CIUDAD)
    wsg2 = wb2["Gasto Personal Sol"]
    for r in range(26, wsg2.max_row + 1):
        com = str(wsg2.cell(r, 6).value or "")
        if "AVA Fideicomiso · cuota 10 · vto 2026-09-04" in com:
            wsg2.cell(r, 6, com.replace("cuota 10", "cuota 11"))
            wsg2.cell(r, 3, "AVA FIDEICOMISO")
    # Add missing Aug + Oct AVA 50% Sol to GPS if absent
    def has_ava(nro, total):
        for r in range(26, wsg2.max_row + 1):
            com = str(wsg2.cell(r, 6).value or "")
            if f"cuota {nro}" in com and str(total) in com.replace(",", ""):
                return True
            if f"total {total}" in com.replace(",", ""):
                return True
        return False

    for n, mes, monto, vto, mep, mep_f, estado, com in ava_cuotas:
        sol = ceil_ars(monto / 2)
        if estado == "PAGADO" and not has_ava(n, monto):
            usd = round(monto / mep, 2) if mep else None
            append_gps(
                wsg2,
                vto,
                "AVA FIDEICOMISO",
                sol,
                "Mercado Pago / TC MEP",
                f"AVA Fideicomiso · cuota {n} · vto {vto.isoformat()} · total {monto} · TC MEP venta {mep} · USD {usd} · 50% Sol · {com}",
            )
    wb2.save(CIUDAD)

    # Update AVA workbook Cuotas Fideicomiso
    wba = load_workbook(AVA_XLSX)
    wcu = wba["Cuotas Fideicomiso"]
    # headers row1; data starts row2; find by cuota nro col A
    for n, mes, monto, vto, mep, mep_f, estado, com in ava_cuotas:
        row = None
        for r in range(2, wcu.max_row + 1):
            if wcu.cell(r, 1).value == n:
                row = r
                break
        if row is None:
            row = wcu.max_row + 1
            wcu.cell(row, 1, n)
        sol = monto // 2
        chris = monto - sol
        usd = round(monto / mep, 2) if mep else None
        wcu.cell(row, 2, mes)
        wcu.cell(row, 3, monto)
        wcu.cell(row, 4, datetime(vto.year, vto.month, vto.day))
        wcu.cell(row, 4).number_format = "YYYY-MM-DD"
        if mep:
            wcu.cell(row, 5, mep)
            wcu.cell(row, 6, datetime(mep_f.year, mep_f.month, mep_f.day))
            wcu.cell(row, 6).number_format = "YYYY-MM-DD"
            wcu.cell(row, 7, usd)
        wcu.cell(row, 8, sol)
        wcu.cell(row, 9, chris)
        # try to find Estado column
        headers = [wcu.cell(1, c).value for c in range(1, 20)]
        if "Estado" in headers:
            wcu.cell(row, headers.index("Estado") + 1, estado)
        if "Comentario" in headers:
            wcu.cell(row, headers.index("Comentario") + 1, com)
        # matrix 2026
    wmat = wba["2026"]
    # row 5 EXPENSAS — map months
    month_cols = {wmat.cell(4, c).value: c for c in range(4, 16)}
    for n, mes, monto, vto, mep, mep_f, estado, com in ava_cuotas:
        # mes like 'agosto 2026' -> Agosto
        mname = mes.split()[0].capitalize()
        if mname == "Septiembre" or True:
            mname = mes.split()[0]
            mname = mname[:1].upper() + mname[1:]
        col = month_cols.get(mname)
        if col:
            wmat.cell(5, col, monto)
    # Detalle AVA
    wdet = wba["Detalle"]
    for n, mes, monto, vto, mep, mep_f, estado, com in ava_cuotas:
        exists = False
        for r in range(4, wdet.max_row + 1):
            if wdet.cell(r, 5).value == monto and isinstance(wdet.cell(r, 1).value, datetime) and wdet.cell(r, 1).value.date() == vto:
                exists = True
                break
        if not exists:
            r = wdet.max_row + 1
            sol = monto // 2
            wdet.cell(r, 1, datetime(vto.year, vto.month, vto.day)).number_format = "YYYY-MM-DD"
            wdet.cell(r, 2, "EXPENSAS")
            wdet.cell(r, 3, "Fideicomiso Ava Palpa")
            wdet.cell(r, 4, f"Cuota {n} · {mes}")
            wdet.cell(r, 5, monto)
            wdet.cell(r, 6, "Mercado Pago / TC MEP" if estado == "PAGADO" else "PENDIENTE factura")
            wdet.cell(r, 9, f"Total cuota. 50% Sol=${sol:,} en Gasto Personal Sol (CIUDAD). {com}".replace(",", ""))
    wba.save(AVA_XLSX)

    # Update AVA CSV
    with AVA_CSV.open(newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
        fields = list(rows[0].keys()) if rows else []
    # ensure fieldnames from header
    with AVA_CSV.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fields = reader.fieldnames
        rows = list(reader)

    by_n = {int(r["cuota_nro"]): r for r in rows if r.get("cuota_nro")}
    # compute deltas vs previous
    prev_monto = int(by_n[9]["monto_ars"]) if 9 in by_n else None
    for n, mes, monto, vto, mep, mep_f, estado, com in ava_cuotas:
        sol = monto // 2
        chris = monto - sol
        pct = None
        if prev_monto:
            pct = round((monto / prev_monto - 1) * 100, 2)
        prev_monto = monto
        row = by_n.get(n, {k: "" for k in fields})
        row["cuota_nro"] = str(n)
        row["mes_periodo"] = mes
        row["monto_ars"] = str(monto)
        row["fecha_vencimiento"] = vto.isoformat()
        row["tc_mep_venta"] = "" if mep is None else str(mep)
        row["fecha_cotizacion_mep"] = "" if mep_f is None else mep_f.isoformat()
        row["valor_usd"] = "" if mep is None else str(round(monto / mep, 2))
        row["sol_50"] = str(sol)
        row["chris_50"] = str(chris)
        row["estado"] = estado
        row["comentario"] = com
        if pct is not None and "pct_delta_cuota" in fields:
            row["pct_delta_cuota"] = str(pct)
        by_n[n] = row
    ordered = sorted(by_n.values(), key=lambda r: int(r["cuota_nro"]))
    with AVA_CSV.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(ordered)

    # Refresh AVA markdown briefly
    lines = [
        "# Cuotas Fideicomiso AVA Palpa",
        "",
        "**Norma TC:** siempre valorizar con **dólar MEP tipo vendedor (venta)** del día de **vencimiento** de la cuota (fuente: Rava / argentinadatos / dólar bolsa). Si no hay rueda, último hábil anterior.",
        "",
        "| Cuota | Mes | Monto ARS | Venc. | MEP venta | USD | Estado | Nota |",
        "|------:|-----|----------:|-------|----------:|----:|--------|------|",
    ]
    for n, mes, monto, vto, mep, mep_f, estado, com in [
        (i, by_n[i]["mes_periodo"], int(by_n[i]["monto_ars"]), by_n[i]["fecha_vencimiento"], by_n[i].get("tc_mep_venta"), by_n[i].get("fecha_cotizacion_mep"), by_n[i]["estado"], by_n[i].get("comentario", ""))
        for i in sorted(by_n)
    ]:
        usd = "" if not mep else round(int(monto) / float(mep), 2)
        lines.append(f"| {n} | {mes} | {monto} | {vto} | {mep or 'TBD'} | {usd or 'TBD'} | {estado} | {com} |")
    lines.append("")
    AVA_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(
        {
            "cashflow_paid": n_cf,
            "compromisos_paid": n_comp,
            "csv_paid": n_csv,
            "gps_proyectado_updated": n_gps_proj,
            "chris_note_row": note_row,
            "ava_cuotas": [a[0] for a in ava_cuotas],
        }
    )


if __name__ == "__main__":
    main()
