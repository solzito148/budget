# 06 — MP No Gasto, exclusiones e inversiones

**GitHub:** https://github.com/solzito148/budget/tree/master/CLAUDE · **Drive Excels:** https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29 · **Drive docs:** https://drive.google.com/drive/folders/1Lfu_3vQjZQ3L3tiBI5eAk661bnyC1UnW

Norma escrita: `CIUDAD/_inbox/NORMA_MP_NO_GASTO.md`.

## Tabla destino

| Concepto MP | Destino | Espejo |
|-------------|---------|--------|
| Dinero retirado | `MP No Gasto` | `Dinero Retirado` |
| Reserva automática | `MP No Gasto` | — |
| Dinero reservado | `MP No Gasto` | — |
| Pago de deuda (Mercado Pago) | `MP No Gasto` | — |
| Pago de resumen de tarjeta | `MP No Gasto` | `Pago de Tarjetas` |
| Inversión (USD, bonos, PF, fondos) | `MP No Gasto` | `Inversión` |
| Transferencia recibida / Ingreso | `Transferencias recibidas` | — |

## Prohibido

- Cargar cualquiera de lo anterior en `Almacen Detalle`  
- Sumarlos a ALMACEN SOL / ALMACEN CHRIS / total `2026` como gasto nuevo  

## Otras exclusiones

- Transferencias propias: Sanes Maria Sole / Maria Soledad Sanes  
- Cuotas de préstamo / créditos Mercado Pago  
- Cancelados / devueltos (usar neto; neto 0 = no cargar)

## Inversión — detalle

Movimientos tipificados: `Inversión` / `Rescate` / `Rendimiento`.  
Incluye venta/compra de dólares. No es gasto de casa.

## Implicancia analítica

Si un retiro o pago de resumen se mete en Almacén o en la matriz, el “consumo de la casa” queda **artificialmente inflado**. Por eso la separación `MP No Gasto` es obligatoria al parsear extractos MP.
