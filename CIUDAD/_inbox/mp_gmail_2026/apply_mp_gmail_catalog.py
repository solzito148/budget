#!/usr/bin/env python3
"""Catalog all Mercado Pago Gmail movement emails and load missing GPS rows."""
from __future__ import annotations

import csv
import json
import math
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[2]  # CIUDAD/
INBOX = Path(__file__).resolve().parent
XLSX = ROOT / "Gastos_CIUDAD_1132_2026.xlsx"
CATALOGO = ROOT / "_inbox" / "catalogo_mercadopago.csv"

MES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def norm(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9áéíóúüñ ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def ceil_int(v) -> int | None:
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return int(math.ceil(float(v)))
    s = str(v).replace("US$", "").replace("$", "").strip()
    try:
        if "," in s:
            return int(math.ceil(float(s.replace(".", "").replace(",", "."))))
        return int(math.ceil(float(s)))
    except ValueError:
        return None


# Explicit loads: missing from Excel (not covered by cuotas / sept official PDF)
LOAD_GPS = [
    # Agosto — dinero en cuenta / BBVA one-shot no cuotificado en GPS
    {
        "fecha": "2026-08-03",
        "cat": "COMIDA · Oficina",
        "monto": 11010,
        "medio": "Mercado Pago",
        "comment": "COMIDA · Oficina · Copetin Catering · Gmail Pago aprobado",
    },
    {
        "fecha": "2026-08-05",
        "cat": "VESTIMENTA",
        "monto": 152000,
        "medio": "Mercado Pago",
        "comment": "VESTIMENTA · Renová tu vestidor · Gmail Pago aprobado · gasto personal Sol",
    },
    {
        "fecha": "2026-08-09",
        "cat": "SALUD",
        "monto": 22222,
        "medio": "Mercado Pago",
        "comment": "SALUD · Siendo Consciencia · Gmail Pago aprobado · REVISAR ítem",
    },
    {
        "fecha": "2026-08-09",
        "cat": "SALUD",
        "monto": 24000,
        "medio": "Mercado Pago",
        "comment": "SALUD · Lorena Ferreyra · Gmail Pago aprobado · REVISAR ítem",
    },
    {
        "fecha": "2026-08-09",
        "cat": "SALUD",
        "monto": 75000,
        "medio": "Mercado Pago",
        "comment": "SALUD · Lorena Ferreyra · Gmail Pago aprobado · REVISAR ítem",
    },
    {
        "fecha": "2026-08-10",
        "cat": "SALUD",
        "monto": 356000,
        "medio": "BBVA Crédito ****3200",
        "comment": "SALUD · Reiki · Gmail Pago aprobado · gasto personal Sol",
    },
    {
        "fecha": "2026-08-10",
        "cat": "COMIDA · Oficina",
        "monto": 13600,
        "medio": "TC Mercado Pago",
        "comment": "COMIDA · Oficina · Copetin Catering · Gmail Pago aprobado",
    },
    {
        "fecha": "2026-08-21",
        "cat": "SALUD",
        "monto": 27000,
        "medio": "Mercado Pago",
        "comment": "SALUD · Karen · Gmail Pago aprobado · REVISAR ítem",
    },
    {
        "fecha": "2026-08-28",
        "cat": "EDUCACIÓN",
        "monto": 30000,
        "medio": "Mercado Pago",
        "comment": "EDUCACIÓN · Bitgenia · Gmail Pago aprobado · gasto personal Sol",
    },
    {
        "fecha": "2026-08-01",
        "cat": "VESTIMENTA",
        "monto": 176000,
        "medio": "BBVA Crédito ****3200",
        "comment": "VESTIMENTA · Gibiana Cova · Gmail Pago aprobado · distinto de transferencia $131.000 · gasto personal Sol",
    },
    # Octubre
    {
        "fecha": "2026-10-04",
        "cat": "FARMACIA SALUD",
        "monto": 17038,
        "medio": "Mercado Pago",
        "comment": "FARMACIA SALUD · Hiperfarmaceutica Temperley · Gmail Pago aprobado",
    },
    {
        "fecha": "2026-10-05",
        "cat": "VIAJES",
        "monto": 1430000,
        "medio": "TC Mercado Pago",
        "comment": "VIAJES · regalo Chris Cumple 50 · MARPATAG SRL · Gmail Pago aprobado · gasto personal Sol",
    },
    {
        "fecha": "2026-10-09",
        "cat": "PELUQUERÍA",
        "monto": 7000,
        "medio": "Mercado Pago",
        "comment": "PELUQUERÍA · Soy lashista · Gmail Pago aprobado (2 mails mismo monto; se carga 1)",
    },
    {
        "fecha": "2026-10-11",
        "cat": "VIAJES",
        "monto": 1143000,
        "medio": "BBVA Crédito ****3200",
        "comment": "VIAJES · regalo Chris Cumple 50 · Patagonia Chic Excursiones · Gmail Pago aprobado · 1 cuota · gasto personal Sol",
    },
]

# MP No Gasto — pagos de resumen / crédito from Gmail not already mirrored
LOAD_MP_NO_GASTO = [
    {
        "fecha": "2026-09-15",
        "tipo": "Pago de resumen TC",
        "concepto": "Pagaste una parte del resumen de tu tarjeta · Gmail",
        "monto": 1214000,
    },
]


def gps_exists(ws, fecha: str, monto: int, needles: list[str]) -> bool:
    target = datetime.fromisoformat(fecha)
    for r in range(26, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if not isinstance(d, datetime):
            continue
        if abs((d.date() - target.date()).days) > 5:
            continue
        m = ws.cell(r, 4).value
        if ceil_int(m) != monto:
            continue
        text = norm(" ".join(str(ws.cell(r, c).value) for c in range(1, 7) if ws.cell(r, c).value))
        if any(norm(n) in text for n in needles):
            return True
    return False


def append_gps(ws, rows: list[dict]) -> int:
    # find last used row
    last = 25
    for r in range(26, ws.max_row + 1):
        if ws.cell(r, 1).value is not None:
            last = r
    added = 0
    for item in rows:
        fecha = item["fecha"]
        needles = re.findall(r"[A-Za-zÁÉÍÓÚáéíóúüñ]{4,}", item["comment"])
        if gps_exists(ws, fecha, item["monto"], needles[:4] or [item["cat"]]):
            print("skip existing GPS", fecha, item["monto"], item["comment"][:50])
            continue
        last += 1
        dt = datetime.fromisoformat(fecha)
        ws.cell(last, 1).value = dt
        ws.cell(last, 2).value = MES[dt.month]
        ws.cell(last, 3).value = item["cat"]
        ws.cell(last, 4).value = item["monto"]
        ws.cell(last, 5).value = item["medio"]
        ws.cell(last, 6).value = item["comment"]
        added += 1
        print("GPS+", fecha, item["monto"], item["cat"])
    return added


def append_mp_no_gasto(ws, rows: list[dict]) -> int:
    # find header Detalle
    header = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "Fecha" and ws.cell(r, 3).value == "Tipo":
            header = r
            break
    if not header:
        raise RuntimeError("MP No Gasto header not found")
    last = header
    existing = set()
    for r in range(header + 1, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if isinstance(d, datetime):
            last = r
            existing.add((d.date().isoformat(), ceil_int(ws.cell(r, 5).value), norm(str(ws.cell(r, 3).value or ""))))
    added = 0
    for item in rows:
        key = (item["fecha"], item["monto"], norm(item["tipo"]))
        if any(e[0] == key[0] and e[1] == key[1] for e in existing):
            print("skip existing MPNG", item)
            continue
        last += 1
        dt = datetime.fromisoformat(item["fecha"])
        ws.cell(last, 1).value = dt
        ws.cell(last, 2).value = MES[dt.month]
        ws.cell(last, 3).value = item["tipo"]
        ws.cell(last, 4).value = item["concepto"]
        ws.cell(last, 5).value = item["monto"]
        ws.cell(last, 6).value = "ARS"
        added += 1
        print("MPNG+", item)
    return added


def update_catalogo(classified: list[dict]) -> int:
    existing = list(csv.DictReader(open(CATALOGO)))
    keys = set()
    for r in existing:
        keys.add((r["fecha"], norm(r["comercio"]), str(ceil_int(r["monto"]) or ""), r["categoria"]))

    added = 0
    with open(CATALOGO, "a", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["fecha", "comercio", "monto", "categoria", "comentario"])
        for m in classified:
            comercio = m.get("comercio") or m.get("subject") or "MP Gmail"
            monto = m.get("monto")
            if monto is None:
                monto_s = ""
            else:
                monto_s = str(monto)
            cat = m.get("cat") or m.get("destino") or "REVISAR"
            # map destino into catalog category style
            if m.get("destino") == "MP_NO_GASTO":
                cat = "PAGO_TARJETA"
            elif m.get("destino") == "EXCLUIR":
                cat = m.get("cat") or "EXCLUIR"
            elif m.get("destino") == "ALMACEN":
                cat = "ALMACEN_SOL"
            elif m.get("destino") == "CIUDAD":
                cat = m.get("cat") or "CIUDAD"
            elif m.get("destino") == "OHIGGINS":
                cat = "EXPENSAS_OHIGGINS" if "EXPENSAS" in (m.get("cat") or "") else "OHIGGINS"
            elif m.get("destino") == "BONORINO":
                cat = "EXPENSAS_BONORINO"
            elif m.get("destino") == "GPS":
                cat = "GASTO_PERSONAL_SOL"
            comment = m.get("comment") or ""
            comment = f"{comment} · Gmail:{m.get('id')}"
            key = (m["fecha"], norm(comercio), monto_s, cat)
            # softer dedupe: fecha+monto+comercio
            soft = (m["fecha"], norm(comercio)[:20], monto_s)
            if any(soft == (k[0], k[1][:20], k[2]) for k in keys):
                continue
            w.writerow(
                {
                    "fecha": m["fecha"],
                    "comercio": comercio,
                    "monto": monto_s if monto_s else (f"US${monto}" if m.get("moneda") == "USD" and monto else ""),
                    "categoria": cat,
                    "comentario": comment[:240],
                }
            )
            keys.add(key)
            added += 1
    return added


def main() -> None:
    classified = json.loads((INBOX / "mp_gmail_classified.json").read_text())
    wb = load_workbook(XLSX)
    n_gps = append_gps(wb["Gasto Personal Sol"], LOAD_GPS)
    n_mpng = append_mp_no_gasto(wb["MP No Gasto"], LOAD_MP_NO_GASTO)
    wb.save(XLSX)
    n_cat = update_catalogo(classified)
    print("DONE gps+", n_gps, "mpng+", n_mpng, "catalog+", n_cat)

    # summary md
    by_dest = Counter(m["destino"] for m in classified)
    by_status = Counter(m["status"] for m in classified)
    dudas = [
        m
        for m in classified
        if m.get("duda")
        or (m.get("comercio") or "").upper().startswith("CORREO OFICIAL")
        or m.get("cat") == "REVISAR"
    ]
    lines = [
        "# Catálogo Gmail Mercado Pago — movimientos 2026",
        "",
        f"Fuente: Gmail `from:mercadopago.com` · {len(classified)} mails de movimiento (Pago aprobado, transferencias, pagos de resumen/crédito, devoluciones).",
        "",
        "**Nota:** La mayoría de los movimientos de cuenta MP (QR, súper, Uber, etc.) **no generan mail**. Esos siguen saliendo del settlement/PDF oficial. Septiembre cuenta ya está cargado desde PDF oficial.",
        "",
        "## Totales por destino (clasificación mail)",
        "| Destino | Cant |",
        "|---|---:|",
    ]
    for k, v in by_dest.most_common():
        lines.append(f"| {k} | {v} |")
    lines += [
        "",
        "## Cruce vs planillas",
        "| Status | Cant |",
        "|---|---:|",
    ]
    for k, v in by_status.most_common():
        lines.append(f"| {k} | {v} |")
    lines += [
        "",
        "## Cargado en esta pasada (faltantes reales)",
        f"- GPS nuevos: **{n_gps}** (ago + oct)",
        f"- MP No Gasto: **{n_mpng}**",
        f"- Filas nuevas en `catalogo_mercadopago.csv`: **{n_cat}**",
        "",
        "### GPS cargados",
    ]
    for x in LOAD_GPS:
        lines.append(f"- {x['fecha']} · ${x['monto']:,} · {x['cat']} · {x['comment']}".replace(",", "."))
    lines += [
        "",
        "## Dudas / REVISAR (Sol)",
        "",
    ]
    if not dudas:
        lines.append("- (ninguna)")
    else:
        for m in dudas:
            lines.append(
                f"- {m['fecha']} · {m.get('comercio') or m.get('subject')} · ${m.get('monto')} · "
                f"{m.get('destino')}/{m.get('cat')} — {m.get('comment')}"
            )
    lines += [
        "",
        "## Ya cubiertos por cuotas TC (no se duplicó el total del mail)",
        "- CENTRAL STORE $350.580 (ago) → cuotas TC MP",
        "- ETIQUETA NEGRA $432.000 (ago) → cuotas TC MP",
        "- Juan Cornacchini $57.725 → ya en GPS cierre sept",
        "- Pandora / Zadig / Bimba / Boken / Loccitane / Frivilli / Promarine / Itsclassic / etc. → cuotas BBVA o TC MP ya en GPS",
        "",
        "## Archivos",
        "- `mp_gmail_moves.csv` — parse crudo",
        "- `mp_gmail_classified.csv` / `.json` — clasificación",
        "- `apply_mp_gmail_catalog.py` — este script",
        "",
    ]
    (INBOX / "RESUMEN_CATALOGO_GMAIL_MP.md").write_text("\n".join(lines) + "\n")
    print("Wrote RESUMEN_CATALOGO_GMAIL_MP.md")


if __name__ == "__main__":
    main()
