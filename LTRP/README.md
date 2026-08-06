# Long Term Retention Program (LTRP)

Excel canónico: **`LTRP.xlsx`** (aparte de CIUDAD / SOL / propiedades).

**Última actualización:** agosto 2026

## Regla general

1. Cada año te **otorgan** un plan con valor nominal (ej. 2026 → **70.000 USD**) y un **precio de acción al otorgamiento**.
2. Ese plan se cobra en **6 tramos** de **1/6**, siempre el **31/01**.
3. En cada enero se acumulan como **máximo 6 planes**: se suma el 1/6 de cada plan aún activo.
4. Cada tramo se paga **50% fijo + 50% variable** según el precio de la acción MELI.

### Ejemplo — otorgado 70.000 en 2026 → cobro **enero 2027**

```
ene-2027 = 1/6·2026 + 1/6·2025 + 1/6·2024 + 1/6·2023 + 1/6·2022
```

## Fórmula (simulación)

```
Pago = ROUND( (N÷6)×50% + (N÷6)×50% × (P_año / P_otorgamiento) ; 0 )
```

La diferencia entre la base sin acción (32.803,67) y el estimado (35.332) es el **valor de la acción**.

## Simulación de pagos (agosto 2026)

| Concepto | Nominal | P. otorgamiento | 2027 | 2028 | 2029 | 2030 | 2031 | 2032 |
|----------|--------:|----------------:|-----:|-----:|-----:|-----:|-----:|-----:|
| LTRP 2022 | 30.000 | 1.391,81 | 5.670 | 5.924 | | | | |
| LTRP 2023 | 30.000 | 888,69 | 7.465 | 7.862 | 8.291 | | | |
| LTRP 2024 | 30.000 | 1.426,11 | 5.594 | 5.842 | 6.109 | 6.398 | | |
| LTRP 2025 | 36.822 | 1.944,47 | 5.854 | 6.077 | 6.317 | 6.577 | 6.858 | |
| LTRP 2026 | 70.000 | 2.094,65 | 10.749 | 11.142 | 11.567 | 12.026 | 12.521 | 13.056 |
| **Total estimado** | | | **35.332** | **36.847** | **32.284** | **25.001** | **19.379** | **13.056** |

**Total acumulado pendiente de pago: USD 161.899**

### Valor de la acción estimado

| 2027 | 2028 | 2029 | 2030 | 2031 | 2032 |
|-----:|-----:|-----:|-----:|-----:|-----:|
| 1.765,02 | 1.906,22 | 2.058,72 | 2.223,42 | 2.401,29 | 2.593,39 |

- **2027:** promedio últimos 60 días de trading + ajuste **3%** cierre 2026.
- **2028–2032:** crecimiento anual **8%**.

Detalle: `METODOLOGIA_ACCION.md`.

## Archivos

| Archivo | Uso |
|---------|-----|
| **`LTRP.xlsx`** | Excel a parte con la simulación completa |
| `build_ltrp.py` | Regenera `LTRP.xlsx` |
| `planes_vigentes.csv` | Nominal + precio al otorgamiento |
| `accion_estimada.csv` | Precio MELI proyectado por año |
| `simulacion_matriz.csv` | Matriz de cobros (simulación Sol) |
| `calendario_vesting.csv` | Qué planes entran cada enero |

```bash
python3 LTRP/build_ltrp.py
```
