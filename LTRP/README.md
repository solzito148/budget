# Long Term Retention Program (LTRP)

Carpeta especial para el cálculo, tracking y proyección de pagos del **Long Term Retention Program** (bono de retención de largo plazo de MELI).

**Última actualización:** agosto 2026

## Fórmula del pago (regla clave)

Cada plan tiene un **valor nominal** asignado ese año. Ese plan paga **Nominal ÷ 6** cada 31/01 durante 6 años.

El cobro de un 31/01 **no es solo el plan nuevo**: hay que **sumar el 1/6 de todos los planes vigentes desde 2022** que todavía estén en su ventana de 6 años.

```
Pago base (31/01/AAAA)  =  Σ  (Valor_nominal_plan ÷ 6)
                           para cada plan LTRP 2022…AAAA-1
                           que aún no haya completado sus 6 tramos
```

Ejemplo — **pago 31/01/2027** (incluye el plan 2026 de 70.000):

| Concepto | Valor nominal | ÷ 6 (aporte al pago) |
|----------|--------------:|---------------------:|
| LTRP 2022 | 30.000 | 5.000 |
| LTRP 2023 | 30.000 | 5.000 |
| LTRP 2024 | 30.000 | 5.000 |
| LTRP 2025 | 36.822 | 6.137 |
| LTRP 2026 | 70.000 | 11.666,67 |
| **Suma = pago base** | **196.822** | **32.803,67** |

```
70.000 ÷ 6 = 11.666,67   ← solo el plan 2026
+ 5.000 + 5.000 + 5.000 + 6.137   ← 1/6 de 2022…2025
= 32.803,67 USD base
```

## Composición 50% fijo / 50% variable

Sobre ese total (o tramo a tramo), el pago se compone:

| Componente | Peso | Descripción |
|------------|-----:|-------------|
| Fijo | 50% | Monto estable sobre el 1/6 nominal |
| Variable | 50% | Sujeto a la variación de la acción **MELI** |

```
Pago final = (suma 1/6) × 50%  +  (suma 1/6) × 50% × factor_MELI
```

## Fecha de pago

El bono se paga el **31/01** del año siguiente al período del plan.

| Plan / período | Primer 1/6 se cobra el |
|----------------|------------------------|
| LTRP 2026 (70.000) | **31/01/2027** |
| LTRP 2025 | 31/01/2026 |
| LTRP 2024 | 31/01/2025 |
| … | … |

## Planes vigentes (valor asignado por año)

| Concepto | Valor nominal (USD) | 1/6 anual (USD) |
|----------|--------------------:|----------------:|
| LTRP 2022 | 30.000 | 5.000 |
| LTRP 2023 | 30.000 | 5.000 |
| LTRP 2024 | 30.000 | 5.000 |
| LTRP 2025 | 36.822 | 6.137 |
| LTRP 2026 | **70.000** | **11.666,67** |
| **Suma (si todos aportan)** | **196.822** | **32.803,67** |

Fuente: `planes_vigentes.csv`.

## Calendario: qué 1/6 entra en cada 31/01

Cada plan aporta su **Nominal÷6** durante **6 cobros consecutivos**.

| Plan | 31/01/23 | 24 | 25 | 26 | **27** | 28 | 29 | 30 | 31 | 32 |
|------|:--------:|:--:|:--:|:--:|:------:|:--:|:--:|:--:|:--:|:--:|
| LTRP 2022 | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | | | | |
| LTRP 2023 | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | | | |
| LTRP 2024 | | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | | |
| LTRP 2025 | | | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | |
| LTRP 2026 | | | | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| **# planes** | 1 | 2 | 3 | 4 | **5** | 5 | 4 | 3 | 2 | 1 |

En **31/01/2027** entran los cinco planes → suma de cinco 1/6 = **32.803,67** base.

## Simulación — estimación portal 2027

**Última actualización:** agosto 2026

| Concepto | USD |
|----------|----:|
| Suma de 1/6 (base, factor MELI = 1,0) | 32.803,67 |
| de lo cual fijo 50% | 16.401,83 |
| de lo cual variable 50% base | 16.401,83 |
| **Estimación portal MELI** | **35.332** |
| Factor MELI implícito (sobre la mitad variable) | ≈ 1,154 |

## Archivos

| Archivo | Uso |
|---------|-----|
| `README.md` | Norma y fórmula |
| `planes_vigentes.csv` | Valor nominal asignado por año |
| `calendario_vesting.csv` | Qué planes aportan en cada 31/01 |
| `simulacion_pagos.csv` | Proyección de la **suma de 1/6** por año |
| `build_ltrp.py` | Regenera el Excel |
| `LTRP_Sol.xlsx` | Libro con planes, suma por cobro y detalle 2027 |

## Cómo regenerar

```bash
python3 LTRP/build_ltrp.py
```

## Notas

- El valor de cada fila (30.000, 70.000, etc.) es el **plan asignado ese año**, no el cobro anual.
- El cobro anual de un plan = ese valor ÷ 6.
- El cobro del 31/01 = **suma** de esos ÷6 de todos los planes aún activos desde 2022.
- El 50% variable mueve el total respecto de la base (ej. portal 35.332 vs base 32.803,67).
- Carpeta **independiente** de CIUDAD / propiedades / Gasto Personal Sol (ingreso por bono, no gasto).
