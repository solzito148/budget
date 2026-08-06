# Metodología — valor de la acción estimado (simulación Sol, ago 2026)

- **Año 1 (2027):** promedio de los últimos 60 días de trading al momento de actualizar la simulación, con un **ajuste de +3%** correspondiente al cierre del 2026.
- **Años siguientes (2028–2032):** crecimiento anual del **8%** sobre el precio del año anterior.

| Año pago | Precio estimado |
|---------:|----------------:|
| 2027 | 1.765,02 |
| 2028 | 1.906,22 |
| 2029 | 2.058,72 |
| 2030 | 2.223,42 |
| 2031 | 2.401,29 |
| 2032 | 2.593,39 |

## Fórmula de cada celda de pago

```
Pago plan/año = ROUND( (Nominal÷6)×50%  +  (Nominal÷6)×50% × (P_año / P_otorgamiento) ; 0 )
```

- 50% fijo sobre el 1/6 nominal
- 50% variable escalado por precio de acción del año de cobro vs precio al otorgamiento

**Total acumulado pendiente de pago (2027–2032):** USD **161.899**
