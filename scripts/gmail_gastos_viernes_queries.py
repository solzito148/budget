#!/usr/bin/env python3
"""Print Gmail search queries + checklist for the Friday expense sync.

Usage:
  python3 scripts/gmail_gastos_viernes_queries.py
  python3 scripts/gmail_gastos_viernes_queries.py --days 10
"""

from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

ART = ZoneInfo("America/Argentina/Buenos_Aires")


def friday_window(today: date, days: int | None) -> tuple[date, date]:
    """Default: previous Friday → today (ART). Or rolling N days."""
    if days is not None:
        start = today - timedelta(days=days)
        return start, today
    # Most recent Friday on or before today, then go back 7 days for start
    # Window: previous Friday inclusive through today
    weekday = today.weekday()  # Mon=0 … Fri=4
    days_since_fri = (weekday - 4) % 7
    this_or_last_fri = today - timedelta(days=days_since_fri)
    start = this_or_last_fri - timedelta(days=7)
    return start, today


def gmail_day(d: date) -> str:
    return d.strftime("%Y/%m/%d")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--days",
        type=int,
        default=None,
        help="Rolling lookback in days instead of previous-Friday window",
    )
    p.add_argument(
        "--date",
        type=str,
        default=None,
        help="Override 'today' as YYYY-MM-DD (ART)",
    )
    args = p.parse_args()

    now = datetime.now(ART).date()
    if args.date:
        now = date.fromisoformat(args.date)
    start, end = friday_window(now, args.days)
    # Gmail before: is exclusive → use day after end
    before = end + timedelta(days=1)
    after = gmail_day(start)
    before_s = gmail_day(before)

    print(f"# Gmail gastos viernes — ventana ART {start.isoformat()} → {end.isoformat()}")
    print(f"# after:{after}  before:{before_s} (before exclusive)\n")

    queries = [
        f'(subject:expensas OR from:kalmus OR from:filippo OR consorcio OR "estudio") after:{after} before:{before_s}',
        f'(Telecentro OR Edenor OR Edesur OR Metrogas OR Personal OR AGIP OR ABL) after:{after} before:{before_s}',
        f'("Fideicomiso" OR "Ava Palpa" OR Sancor OR cuota) after:{after} before:{before_s}',
        f'("Mercado Pago" OR "resumen de cuenta" OR "pago de resumen" OR BBVA OR "Visa Signature") after:{after} before:{before_s}',
        f'(subject:comprobante OR subject:factura OR has:attachment filename:pdf) after:{after} before:{before_s}',
    ]
    print("## Queries")
    for i, q in enumerate(queries, 1):
        print(f"{i}. {q}")

    print(
        """
## Destinos
- CIUDAD  → CIUDAD/Gastos_CIUDAD_1132_2026.xlsx
- BONORINO → BONORINO/Gastos_BONORINO_2026.xlsx
- OHIGGINS → OHIGGINS/Gastos_OHIGGINS_2026.xlsx
- AVA      → AVA/Gastos_AVA_2026.xlsx + AVA/cuotas_fideicomiso_ava.csv
- Luego    → python3 SOL/build_gastos_sol.py

## Checklist
[ ] Expensas CIUDAD / OHIGGINS / BONORINO
[ ] Utilidades (Telecentro, Edenor×2, Metrogas×2, Edesur, Personal, ABL)
[ ] AVA cuota + MEP venta + 50% Sol + Cashflow Compromisos
[ ] MP No Gasto / pagos resumen / retiros
[ ] Almacén + GPS
[ ] build_gastos_sol.py
[ ] commit / push / PR

Skill: .cursor/skills/gmail-gastos-viernes/SKILL.md
"""
    )


if __name__ == "__main__":
    main()
