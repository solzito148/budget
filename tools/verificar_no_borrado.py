"""Verifica que una planilla nueva no haya borrado datos de meses anteriores.

Uso: python3 tools/verificar_no_borrado.py VIGENTE.xlsx NUEVO.xlsx
Compara hoja por hoja, mes por mes (cantidad de filas con importe y suma de
importes). Sale con código 1 si algún mes del NUEVO tiene menos filas o
menos importe que el VIGENTE en una hoja de detalle, o si una celda de matriz
con valor quedó vacía.
"""
import sys
from collections import defaultdict

from openpyxl import load_workbook

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def mes_de_fila(row):
    for v in row:
        if hasattr(v, "year") and v.year == 2026:
            return v.month
    for v in row[:2]:
        s = str(v or "").strip().lower()
        if s in MESES:
            return MESES.index(s) + 1
    return None


def resumen(path):
    wb = load_workbook(path, data_only=False)
    det = defaultdict(lambda: [0, 0])      # (hoja, mes) -> [filas, suma]
    mat = {}                               # (hoja, fila_desc, col) -> valor
    for ws in wb.worksheets:
        hdr = [str(c.value or "").lower() for c in next(ws.iter_rows(min_row=1, max_row=6))]
        es_matriz = ws.title in ("2026",)
        for r in ws.iter_rows(values_only=True):
            if es_matriz:
                desc = r[2] if len(r) > 2 else None
                if desc:
                    for i, v in enumerate(r[3:15]):
                        if isinstance(v, (int, float)) and v:
                            mat[(ws.title, str(desc), i)] = v
                continue
            m = mes_de_fila(r)
            if not m:
                continue
            nums = [v for v in r if isinstance(v, (int, float)) and not isinstance(v, bool)]
            if nums:
                det[(ws.title, m)][0] += 1
                det[(ws.title, m)][1] += sum(abs(x) for x in nums)
    return det, mat


def main(vig, new):
    d0, m0 = resumen(vig)
    d1, m1 = resumen(new)
    problemas = []
    for k, (n0, s0) in d0.items():
        n1, s1 = d1.get(k, (0, 0))
        if n1 < n0:
            problemas.append(f"{k[0]} · {MESES[k[1]-1]}: filas {n0} -> {n1}")
    for k, v in m0.items():
        if k not in m1:
            problemas.append(f"{k[0]} · {k[1]} · {MESES[k[2]].capitalize()}: {v} -> VACÍO")
    if problemas:
        print("ATENCIÓN: se perdieron datos de meses anteriores:")
        print("\n".join("  - " + p for p in problemas))
        sys.exit(1)
    print("OK: ningún mes perdió filas ni valores de matriz.")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
