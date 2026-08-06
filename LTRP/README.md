# Long Term Retention Program (LTRP)

Excel canónico: **`LTRP.xlsx`** (aparte de CIUDAD / SOL / propiedades).

**Última actualización:** agosto 2026

## Regla general

1. Cada año te **otorgan** un plan con un valor nominal (ej. 2026 → **70.000 USD**).
2. Ese plan se cobra en **6 tramos** de **1/6** cada uno, siempre el **31/01**.
3. En cada enero se **acumulan como máximo 6 planes**: se suma el **1/6 de cada plan** que todavía esté dentro de su ventana de 6 cobros.
4. Cuando entra un plan nuevo, el más viejo que ya cumplió sus 6 cobros **sale** de la suma.

```
Pago enero AAAA  =  (1/6)·plan_(AAAA-1)  +  (1/6)·plan_(AAAA-2)  +  …  +  (1/6)·plan_(AAAA-6)
                   ↑ máximo 6 planes
```

### Ejemplo — otorgado 70.000 en 2026 → cobro **enero 2027**

En 2027 sumás el 1/6 de 2026 **más** los de 2025, 2024, 2023 y 2022:

| Plan | Otorgado (nominal) | 1/6 en el pago ene-2027 |
|------|-------------------:|------------------------:|
| LTRP 2026 | **70.000** | **11.666,67** |
| LTRP 2025 | 36.822 | 6.137 |
| LTRP 2024 | 30.000 | 5.000 |
| LTRP 2023 | 30.000 | 5.000 |
| LTRP 2022 | 30.000 | 5.000 |
| **Suma (pago base)** | | **32.803,67** |

Hoy hay **5** planes (aún no hay un sexto grant). Cuando exista LTRP 2027, el cobro de ene-2028 sumará **6** tramos (2022…2027); en ene-2029 el 2022 ya no entra y sí el 2028, siempre tope 6.

## Composición 50% fijo / 50% variable

Sobre la suma de los 1/6:

| Componente | Peso | Descripción |
|------------|-----:|-------------|
| Fijo | 50% | Estable sobre el 1/6 nominal |
| Variable | 50% | Sujeto a la acción **MELI** |

```
Pago final = (suma de 1/6) × 50%  +  (suma de 1/6) × 50% × factor_MELI
```

## Fecha de pago

Siempre **31/01**. El primer 1/6 de un plan otorgado en el año Y se cobra el **31/01/(Y+1)**.

| Plan otorgado | Primer cobro | Último cobro (6/6) |
|---------------|--------------|--------------------|
| LTRP 2026 (70.000) | 31/01/2027 | 31/01/2032 |
| LTRP 2025 | 31/01/2026 | 31/01/2031 |
| LTRP 2022 | 31/01/2023 | 31/01/2028 |

## Planes vigentes

| Concepto | Valor nominal (USD) | 1/6 anual (USD) |
|----------|--------------------:|----------------:|
| LTRP 2022 | 30.000 | 5.000 |
| LTRP 2023 | 30.000 | 5.000 |
| LTRP 2024 | 30.000 | 5.000 |
| LTRP 2025 | 36.822 | 6.137 |
| LTRP 2026 | **70.000** | **11.666,67** |

## Calendario: qué planes entran en cada enero (tope 6)

| Plan | ene23 | 24 | 25 | 26 | **27** | 28 | 29 | 30 | 31 | 32 |
|------|:-----:|:--:|:--:|:--:|:------:|:--:|:--:|:--:|:--:|:--:|
| 2022 | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | | | | |
| 2023 | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | | | |
| 2024 | | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | | |
| 2025 | | | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | |
| 2026 | | | | | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| **# planes** | 1 | 2 | 3 | 4 | **5** | 5* | 4* | 3 | 2 | 1 |

\*Si se otorga LTRP 2027 / 2028, esos enero pueden llegar a **6** planes.

## Simulación — estimación portal 2027

**Última actualización:** agosto 2026

| Concepto | USD |
|----------|----:|
| Suma de 1/6 (base, sin efecto acción) | 32.803,67 |
| Estimación portal MELI | **35.332** |
| **Diferencia** | **+2.528,33** |

Esa diferencia (**35.332 − 32.803,67**) se debe al **valor de la acción MELI** sobre el **50% variable**. El 50% fijo queda en 16.401,83; el 50% variable sube de 16.401,83 base a ≈ 18.930 por el precio de la acción (factor ≈ 1,154).

```
Pago estimado ≈  16.401,83 fijo  +  16.401,83 × factor_acción
```

## Archivos

| Archivo | Uso |
|---------|-----|
| **`LTRP.xlsx`** | Excel a parte con todo el cálculo |
| `build_ltrp.py` | Regenera `LTRP.xlsx` |
| `planes_vigentes.csv` | Valores nominales otorgados |
| `calendario_vesting.csv` | Qué planes entran cada enero |
| `simulacion_pagos.csv` | Proyección de la suma de 1/6 |

```bash
python3 LTRP/build_ltrp.py
```

## Notas

- Nominal (70.000) = plan **otorgado** ese año; cobro anual del plan = Nominal÷6.
- Cobro de enero = **suma de hasta 6** tramos de 1/6.
- Carpeta independiente de CIUDAD / propiedades / Gasto Personal Sol (ingreso por bono).
