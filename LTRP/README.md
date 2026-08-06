# Long Term Retention Program (LTRP)

Excel canónico: **`LTRP.xlsx`** (aparte de CIUDAD / SOL / propiedades).

**Última actualización:** agosto 2026

## Regla general

1. Cada año te **otorgan** un plan (nominal + precio de acción al otorgamiento).
2. Se cobra en **6 tramos** de **1/6**, siempre en **enero**.
3. Máximo **6 planes** acumulados por cobro: se suma el 1/6 de cada plan activo.
4. Cada tramo = **50% fijo + 50% variable** (precio MELI del año / precio al otorgamiento).

El plan de **70.000 de 2026** (edad Sol **45**) se completa **recién en 2032** (último 1/6).

## Simulación de pagos (con edad Sol)

| Edad | Concepto | Nominal | $/año aprox | P. otorg. | 2027 | 2028 | 2029 | 2030 | 2031 | 2032 |
|-----:|----------|--------:|------------:|----------:|-----:|-----:|-----:|-----:|-----:|-----:|
| 41 | LTRP 2022 | 30.000 | 5.000 | 1.391,81 | 5.670 | 5.924 | | | | |
| 42 | LTRP 2023 | 30.000 | 5.000 | 888,69 | 7.465 | 7.862 | 8.291 | | | |
| 43 | LTRP 2024 | 30.000 | 5.000 | 1.426,11 | 5.594 | 5.842 | 6.109 | 6.398 | | |
| 44 | LTRP 2025 | 36.822 | 6.137 | 1.944,47 | 5.854 | 6.077 | 6.317 | 6.577 | 6.858 | |
| 45 | LTRP 2026 | 70.000 | 11.666,67 | 2.094,65 | 10.749 | 11.142 | 11.567 | 12.026 | 12.521 | **13.056** |
| 46 | LTRP 2027* | 70.000 | 11.666,67 | — | | 11.666 | … | | | |
| … | LTRP 2028–2036* | 70.000 | 11.666,67 | — | | | proyección hasta 2037 | | | |

\*Proyectados: nominal 70.000 asumido, **1/6 ≈ 11.666** sin precio de acción aún.

| Total 2027 | Pendiente vigentes 2022–2026 |
|-----------:|-----------------------------:|
| **35.332** | **161.899** |

## Fórmula

```
Pago = ROUND( (N÷6)×50% + (N÷6)×50% × (P_año / P_otorgamiento) ; 0 )
```

Precio acción: 60 días +3% (2027), luego +8% anual. Ver `METODOLOGIA_ACCION.md`.

```bash
python3 LTRP/build_ltrp.py
```
