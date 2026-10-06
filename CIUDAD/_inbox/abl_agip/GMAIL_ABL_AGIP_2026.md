# ABL AGIP — mails de boleta electrónica (Gmail)

Fuente: `from:agip.gob.ar subject:ABL` · cuenta s.sanes@gmail.com · revisado 2026-10-06.

## Hallazgo clave
Los mails **Emisión de boleta digital | ABL** solo llegan para la partida adherida a boleta electrónica:

| Partida | Propiedad | Dirección (boleta) |
|--------|-----------|--------------------|
| **3977389** | **CIUDAD** | Ciudad de la Paz 1134 (casa) |
| 3016365 | OHIGGINS | (sin mail BE; consulta pública / PDF) |
| 1084069 | BONORINO | Bonorino 82 1º 005 (sin mail BE; PDF ConsultaABL) |

## Mails 2026 (partida 3977389)
| Msg date | Cuota | Vto 1° | Estado al emitir |
|----------|-------|--------|------------------|
| 2026-06-03 | 6 | 12/06/2026 | PENDIENTE |
| 2026-08-08 | 8 | 14/08/2026 | PENDIENTE |
| 2026-09-09 | 9 | 14/09/2026 | PENDIENTE |

Los links `ConsultaABL/?p1=partida&token=...` vencen; al 2026-10-06 ya no descargan PDF.
Montos de cuota 08 tomados de descarga previa del mismo mail; cuota 09 confirmada por pago MP oficial $27.634,08 → ceil **27635**.

## Carga aplicada
Ver Excel CIUDAD / OHIGGINS / BONORINO y `abl_partidas.csv`.

## Actualización AGIP ConsultaABL (2026-10-06)
Ver `AGIP_DEUDA_2026-10-06.md`. Resumen:
- **CIUDAD** deuda vencida cuotas **08** (Act 51964) y **09** (Act 51756); pendientes 10–12 (42960 / 43680 / 44712).
- El MP $27.634,08 del 14/09 **no** era CIUDAD → **BONORINO** cuota 09.
- **OHIGGINS** sin deuda vencida; pendientes 10–12 a 88891.
- **BONORINO** pendientes 10–12 (28215 / 28723 / 29556).
