# MP dinero en cuenta — Septiembre 2026

Fuente oficial: `mp_exports/MercadoPago_septiembre__07de.pdf` (cuenta ARS 1–30 sep + tenencias USD).
Parse: `mp_exports/mp_cuenta_oficial_sept.json` (316 movimientos ARS).
Script: `apply_mp_oficial_sept.py`.

## Cobertura
- Resumen oficial completo **1–30 sep** (reemplaza el gap de capturas 28–30).
- Devoluciones aplicadas (neto 0): Didi $7.600 (2/9), Ebanx/Uber $6.607 (2/9), PedidosYa $38.099 (12/9), Ebanx/Uber $6.053 (18/9).
- Capturas 1–27 cruzadas; prevalece el PDF ante conflictos.

## Cargado desde PDF oficial
| Destino | Cant. / monto | Notas |
|---|---|---|
| Almacén 28–30 | Súper Pollo $10.900 + Shell $65.001 | ceil |
| GPS 28–30 | Uber $7.230 · Didi $9.700 · Camilo $69.900 · Kelly $106.000 | Kelly/Camilo REVISAR |
| MP No Gasto | 160 movs sept | reservas / retirado / reservado / inversión / pago resumen |
| Dinero Retirado | 17 | espejo |
| Pago de Tarjetas | 1 × $500.000 (30/09) | espejo |
| Inversión | 4 rescates | 3 ventas MEP + Retiro Bonos $12.361.244 |
| Transferencias recibidas | 7 | incl. propias Sanes (no son gasto) |
| Sol Dollar Mdp | saldo **US$ 1,68** al 30/09 | ajuste REVISAR vs detalle PDF (cierra 2002,21) |

## Totales Almacén Sol sept (post-PDF)
**$668.405** (antes capturas/TC $630.603; − PedidosYa anulado $38.099 + pollo/Shell $75.901).

## Excluidos (siguen fuera de GPS/Almacén)
- Transferencias propias Sanes Maria Sole / Maria Soledad Sanes
- Christian Ventura $10
- Reservas / dinero reservado / retirado / inversión / pago resumen → solo MP No Gasto (+ espejos)
- Créditos MP (si aparecen)

## REVISAR
1. Kelly ($20k+$86k) y Camilo Carlos Ihan ($69.900) — ítem
2. Sol Dollar Mdp: saldo final header US$ 1,68 vs detalle movimientos US$ 2002,21 (ajuste −2000,53)
3. Ítems previos: Luxodia, Markova, La Obanesad, pagos genéricos 13/09
