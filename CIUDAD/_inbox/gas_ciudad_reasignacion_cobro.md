# GAS CIUDAD — reasignación por fecha de cobro

Criterio (2026-07-19): el gasto se asigna al **mes de la fecha de cobro** (pago efectivo BBVA / vencimiento factura), **no** al período de consumo de la boleta.

## Matriz `2026` fila GAS (post-corrección)

| Mes | Monto | Fecha cobro BBVA |
|-----|------:|------------------|
| Enero | 26.412 | 19-ene-2026 |
| Febrero | 25.745 | 18-feb-2026 |
| Marzo | — | sin cobro BBVA |
| Abril | 37.569 | 20-abr-2026 |
| Mayo | 36.014 | 19-may-2026 |
| Junio | 30.643 | 17-jun-2026 |

## Montos que estaban corridos (~1 mes) y se corrigieron

| Antes (Excel por período) | Después (mes cobro BBVA) |
|---------------------------|--------------------------|
| Feb 37.316 | Feb 25.745 |
| Mar 37.569 | Abr 37.569 |
| Abr 36.014 | May 36.014 |
| May 30.643 | Jun 30.643 |
| Jun 73.695 | *quitado de matriz* — no figura en resumen BBVA Visa cierre 02-jul |

## Pendientes / revisar

- **$37.316**: estaba en Detalle (02-05) y en MP 03-abr como Metrogas; no coincide con ningún cobro BBVA DEB. No se sumó de nuevo a la matriz.
- **$73.695**: estaba en Detalle (05-jun); no está en Visa BBVA. No se sumó a la matriz hasta confirmar cobro.
