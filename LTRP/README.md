# Long Term Retention Program (LTRP)

Carpeta especial para el cálculo, tracking y proyección de pagos del **Long Term Retention Program** (bono de retención de largo plazo de MELI).

**Última actualización:** agosto 2026

## Qué es

El LTRP es una asignación en **dólares (USD)** que se acumula durante un período de **6 años**. Cada año se libera **1/6** del valor nominal del plan, y ese tramo se paga compuesto así:

| Componente | Peso | Descripción |
|------------|-----:|-------------|
| Fijo | 50% | Monto estable sobre el valor nominal del tramo |
| Variable | 50% | Sujeto a la variación de la acción de **MELI** |

```
Pago anual del plan = (Nominal ÷ 6) × 50%  +  (Nominal ÷ 6) × 50% × factor_MELI
```

donde `factor_MELI` refleja la variación de la acción respecto del referente del plan.

## Fecha de pago

El bono se paga el **31/01** del año siguiente al período calculado.

| Período calculado | Fecha de pago |
|-------------------|---------------|
| 2026 | **31/01/2027** |
| 2027 | 31/01/2028 |
| … | … |

## Planes vigentes

| Concepto | Valor nominal (USD) | 1/6 anual (USD) | Fijo 50% | Variable 50% (base) |
|----------|--------------------:|----------------:|---------:|--------------------:|
| LTRP 2022 | 30.000 | 5.000 | 2.500 | 2.500 |
| LTRP 2023 | 30.000 | 5.000 | 2.500 | 2.500 |
| LTRP 2024 | 30.000 | 5.000 | 2.500 | 2.500 |
| LTRP 2025 | 36.822 | 6.137 | 3.068,50 | 3.068,50 |
| LTRP 2026 | 70.000 | 11.666,67 | 5.833,33 | 5.833,33 |
| **Total** | **196.822** | **32.803,67** | **16.401,83** | **16.401,83** |

Fuentes de datos: `planes_vigentes.csv`.

## Calendario de vesting (1/6 por año)

Cada plan paga durante **6 años consecutivos**, acumulándose con los demás planes vigentes.

| Plan | 2023 | 2024 | 2025 | 2026 | 2027 | 2028 | 2029 | 2030 | 2031 | 2032 |
|------|:----:|:----:|:----:|:----:|:----:|:----:|:----:|:----:|:----:|:----:|
| LTRP 2022 | 1/6 | 2/6 | 3/6 | 4/6 | 5/6 | 6/6 | — | — | — | — |
| LTRP 2023 | — | 1/6 | 2/6 | 3/6 | 4/6 | 5/6 | 6/6 | — | — | — |
| LTRP 2024 | — | — | 1/6 | 2/6 | 3/6 | 4/6 | 5/6 | 6/6 | — | — |
| LTRP 2025 | — | — | — | 1/6 | 2/6 | 3/6 | 4/6 | 5/6 | 6/6 | — |
| LTRP 2026 | — | — | — | — | 1/6 | 2/6 | 3/6 | 4/6 | 5/6 | 6/6 |

La columna indica el **año de pago** (fecha efectiva **31/01** de ese año).

## Simulación de pagos

**Última actualización:** agosto 2026

### Estimación de pago para 2027 (31/01/2027)

| Concepto | USD |
|----------|----:|
| Base fija (50% de todos los tramos 2027) | 16.401,83 |
| Variable base (50%, factor MELI = 1,0) | 16.401,83 |
| **Base total (factor 1,0)** | **32.803,67** |
| **Estimación portal MELI** | **35.332** |
| Factor MELI implícito (solo sobre variable) | ≈ 1,154 |

Planes que aportan al pago 2027:

| Plan | Año de vesting | Tramo 1/6 | Fijo | Variable base |
|------|---------------:|----------:|-----:|--------------:|
| LTRP 2022 | 5/6 | 5.000 | 2.500 | 2.500 |
| LTRP 2023 | 4/6 | 5.000 | 2.500 | 2.500 |
| LTRP 2024 | 3/6 | 5.000 | 2.500 | 2.500 |
| LTRP 2025 | 2/6 | 6.137 | 3.068,50 | 3.068,50 |
| LTRP 2026 | 1/6 | 11.666,67 | 5.833,33 | 5.833,33 |

## Archivos

| Archivo | Uso |
|---------|-----|
| `README.md` | Norma y explicación del programa |
| `planes_vigentes.csv` | Valores nominales de cada LTRP |
| `calendario_vesting.csv` | Matriz año × plan (fracción 1/6) |
| `simulacion_pagos.csv` | Proyección de pagos por año (base + estimado) |
| `build_ltrp.py` | Regenera el Excel canónico desde los CSV |
| `LTRP_Sol.xlsx` | Libro Excel con planes, calendario y simulación |

## Cómo regenerar

```bash
python3 LTRP/build_ltrp.py
```

## Notas

- Montos en **USD**. El factor variable depende de la acción MELI; la estimación del portal puede diferir de la base a factor 1,0.
- Los planes se **acumulan**: en un mismo 31/01 pueden cobrar varios LTRP a la vez (hasta 6 tramos simultáneos cuando hay 6 planes activos).
- Esta carpeta es **independiente** del presupuesto hogareño CIUDAD / propiedades / Gasto Personal Sol. Es tracking de ingreso por bono, no un gasto.
