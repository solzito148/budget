#!/usr/bin/env python3
"""Alert when a closed month is missing a required utility/service row.

A month is "closed" once we are past its last calendar day (ART).
For the current month, only alert if --include-current is set.

Usage:
  python3 scripts/audit_servicios_mensuales.py
  python3 scripts/audit_servicios_mensuales.py --year 2026 --through 2026-09
"""

from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

ART = ZoneInfo("America/Argentina/Buenos_Aires")
ROOT = Path(__file__).resolve().parents[1]

MESES = [
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

# Required services: one amount per closed month (unless noted).
# row = 1-based Excel row in hoja 2026; col Enero = 4.
CHECKS = [
    {
        "prop": "CIUDAD",
        "path": "CIUDAD/Gastos_CIUDAD_1132_2026.xlsx",
        "row": 5,
        "name": "Expensas",
        "allow_empty": set(),  # all closed months required
    },
    {
        "prop": "CIUDAD",
        "path": "CIUDAD/Gastos_CIUDAD_1132_2026.xlsx",
        "row": 6,
        "name": "Telecentro",
        "allow_empty": set(),
    },
    {
        "prop": "CIUDAD",
        "path": "CIUDAD/Gastos_CIUDAD_1132_2026.xlsx",
        "row": 7,
        "name": "Metrogas",
        # Jul 2026: confirmed no BBVA cobro that month (two charges fell in Aug)
        "allow_empty": {7},
        "allow_empty_note": {7: "sin cobro BBVA ese mes (confirmado Gmail)"},
    },
    {
        "prop": "CIUDAD",
        "path": "CIUDAD/Gastos_CIUDAD_1132_2026.xlsx",
        "row": 8,
        "name": "Edenor",
        "allow_empty": set(),
    },
    {
        "prop": "CIUDAD",
        "path": "CIUDAD/Gastos_CIUDAD_1132_2026.xlsx",
        "row": 12,
        "name": "Spotify",
        "allow_empty": set(),
    },
    {
        "prop": "OHIGGINS",
        "path": "OHIGGINS/Gastos_OHIGGINS_2026.xlsx",
        "row": 5,
        "name": "Expensas",
        "allow_empty": set(),
    },
    {
        "prop": "OHIGGINS",
        "path": "OHIGGINS/Gastos_OHIGGINS_2026.xlsx",
        "row": 6,
        "name": "Personal TV",
        # Feb: factura Gmail $35.784 venc 23/02 sin pago MP hallado
        # Mar: débito rechazado 09/04; saldo incluido en pago MP Abr $101.114
        # Ago: factura venc 03/08 $56.200 pagada MP 14/07 (cargada en Julio)
        "allow_empty": {2, 3, 8},
        "allow_empty_note": {
            2: "factura Gmail $35.784 venc 23/02; sin pago MP (revisar si entró en Abr $101.114)",
            3: "débito rechazado 09/04; incluido en pago MP Abr $101.114",
            8: "factura venc 03/08 $56.200 pagada MP 14/07 (en Julio)",
        },
    },
    {
        "prop": "OHIGGINS",
        "path": "OHIGGINS/Gastos_OHIGGINS_2026.xlsx",
        "row": 8,
        "name": "Edenor Chagas",
        # Ene/Feb: sin factura ni cobro BBVA (primer cobro 12/03 $116.475)
        "allow_empty": {1, 2},
        "allow_empty_note": {
            1: "sin factura Gmail ni cobro BBVA",
            2: "sin cobro BBVA (primer cobro 12/03)",
        },
    },
    {
        "prop": "OHIGGINS",
        "path": "OHIGGINS/Gastos_OHIGGINS_2026.xlsx",
        "row": 9,
        "name": "Metrogas INVALMAR",
        # Mar 2026 factura $0; Ene/Feb sin factura digital en Gmail
        "allow_empty": {1, 2, 3},
        "allow_empty_note": {
            1: "sin factura Gmail",
            2: "sin factura Gmail",
            3: "factura $0",
        },
    },
    {
        "prop": "BONORINO",
        "path": "BONORINO/Gastos_BONORINO_2026.xlsx",
        "row": 5,
        "name": "Expensas",
        # series starts May 2026 in workbook
        "allow_empty": {1, 2, 3, 4},
        "allow_empty_note": {m: "sin liquidación en archivo (pre-May)" for m in (1, 2, 3, 4)},
    },
    {
        "prop": "BONORINO",
        "path": "BONORINO/Gastos_BONORINO_2026.xlsx",
        "row": 6,
        "name": "Edesur",
        # Jul: factura emisión 21/07 $2.074 venc 03/08 → cobro Agosto (ya cargado)
        "allow_empty": {1, 2, 3, 7},
        "allow_empty_note": {
            **{m: "sin factura en archivo (pre-Abr)" for m in (1, 2, 3)},
            7: "factura emisión 21/07 $2.074 venc 03/08 → cobro Agosto",
        },
    },
    {
        "prop": "AVA",
        "path": "AVA/Gastos_AVA_2026.xlsx",
        "row": 5,
        "name": "Cuota Fideicomiso",
        "allow_empty": set(),
    },
]


def closed_months(today: date, year: int, include_current: bool) -> list[int]:
    """Return 1-based month numbers that are closed (and optionally current)."""
    out = []
    for m in range(1, 13):
        if year < today.year:
            out.append(m)
            continue
        if year > today.year:
            continue
        # same year
        if m < today.month:
            out.append(m)
        elif m == today.month and include_current:
            out.append(m)
    return out


def cell_empty(v) -> bool:
    return v is None or v == "" or v == 0


def audit(year: int, today: date, include_current: bool) -> list[dict]:
    months = closed_months(today, year, include_current)
    alerts = []
    allowed = []
    for chk in CHECKS:
        path = ROOT / chk["path"]
        if not path.exists():
            alerts.append(
                {
                    "level": "ALERTA",
                    "prop": chk["prop"],
                    "name": chk["name"],
                    "mes": "—",
                    "msg": f"archivo no encontrado: {chk['path']}",
                }
            )
            continue
        wb = load_workbook(path, data_only=True)
        ws = wb["2026"]
        allow = chk.get("allow_empty") or set()
        notes = chk.get("allow_empty_note") or {}
        for m in months:
            # only months of the audited year
            val = ws.cell(chk["row"], 3 + m).value  # Enero=col4 → 3+m
            if not cell_empty(val):
                continue
            mes_name = MESES[m - 1]
            if m in allow:
                allowed.append(
                    {
                        "level": "OK-VACÍO",
                        "prop": chk["prop"],
                        "name": chk["name"],
                        "mes": mes_name,
                        "msg": notes.get(m, "vacío permitido"),
                    }
                )
            else:
                alerts.append(
                    {
                        "level": "ALERTA",
                        "prop": chk["prop"],
                        "name": chk["name"],
                        "mes": mes_name,
                        "msg": f"mes cerrado sin monto en matriz 2026 (fila {chk['row']})",
                    }
                )
    return alerts, allowed


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--year", type=int, default=datetime.now(ART).year)
    p.add_argument(
        "--include-current",
        action="store_true",
        help="Also flag gaps in the current (open) month",
    )
    p.add_argument(
        "--date",
        type=str,
        default=None,
        help="Override today as YYYY-MM-DD (ART)",
    )
    args = p.parse_args()
    today = datetime.now(ART).date()
    if args.date:
        today = date.fromisoformat(args.date)

    alerts, allowed = audit(args.year, today, args.include_current)
    closed = closed_months(today, args.year, args.include_current)
    print(
        f"# Audit servicios mensuales — {args.year} · hoy ART {today.isoformat()}"
    )
    print(
        f"# Meses cerrados evaluados: {', '.join(MESES[m-1] for m in closed) or '(ninguno)'}"
    )
    print()

    if alerts:
        print(f"## ALERTAS ({len(alerts)}) — mes cerrado sin servicio")
        for a in alerts:
            print(f"- ⚠️  {a['prop']} · {a['name']} · {a['mes']}: {a['msg']}")
        print()
    else:
        print("## ALERTAS: ninguna ✅")
        print()

    if allowed:
        print(f"## Vacíos permitidos ({len(allowed)})")
        for a in allowed:
            print(f"- {a['prop']} · {a['name']} · {a['mes']}: {a['msg']}")
        print()

    print(
        "Regla: si un mes ya cerró y falta el servicio → AVISAR a Sol "
        "(no inventar monto; buscar Gmail o pedir el comprobante)."
    )
    return 1 if alerts else 0


if __name__ == "__main__":
    sys.exit(main())
