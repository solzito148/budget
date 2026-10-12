# Auditoría Gmail → Almacén CIUDAD (2026-10-12)

Revisión amplia de mails (Mercado Pago, BBVA Compra aprobada, TiendaNube Dr.Barf, Carrefour, Sodastream, etc.) cruzada contra `Almacen Detalle`.

## Hallazgo principal
La mayoría del Almacén (QR, PedidosYa, súper, combustible) **no genera mail**. Sale del settlement/PDF oficial de cuenta MP.
- **Septiembre:** cargado desde PDF oficial.
- **Agosto:** casi vacío en planilla (solo verdulería + lo cargado abajo). Falta el CSV oficial.

## Reporte MP agosto (pendiente descargar con sesión)
Mail 02/09/2026 · `Ya podés conciliar todas tus transacciones` · período **01-08-2026 → 01-09-2026**  
Link CSV (requiere login MP):  
`https://www.mercadopago.com.ar/balance/reports/settlement_v2/statements/7992b462-9433-4a8d-aecf-5627fb7e43e0/download?format=csv`

## Cargado desde emails en esta pasada
| Fecha | Monto | Destino | Fuente |
|---|---:|---|---|
| 2026-05-11 | 24.852 | `COMIDA · Pescadería Belgrano` | MP Pago aprobado · estaba mal en GPS |
| 2026-06-02 | 32.046 | `COMIDA · Pescadería Belgrano` | MP Pago aprobado · estaba mal en GPS |
| 2026-07-04 | 149.500 | `RAMON · Comida` · DR BARF #462 | Gmail + checkout success total |
| 2026-08-16 | 1.060.501 | `ENTRETENIMIENTO · Social` · Gardiner | BBVA Compra aprobada |
| 2026-08-22 | 77.200 | `RAMON · Otros` · Paws | BBVA Compra aprobada (3 cuotas) |

## Ya estaban / no duplicar
- Dr.Barf 01/08 $176.000 y 10/09 $131.000 → Almacén
- Sodastream 08/09 $88.659 · Pescadería 04/09 $6.213 · Superpollo 18/07 $34.293 · SUPERDIA 03/09 $41.139
- Carrefour mail 03/06 $79.809: settlement oficial cuenta solo $23.943 ese día; hay devolución Ecommerce 02/07 $75.417 neto 0 → **no se carga el $79.809 del mail**

## No Almacén (GPS u otros)
- Luxodia 27/09 $334.288 + $49.492 → GPS VESTIMENTA
- Suscripciones USD (Cursor/OpenAI/Google One), ML, ropa, etc. → GPS

## Siguiente paso crítico
Descargar el CSV de conciliación agosto (link arriba) o el PDF de cuenta MP agosto y cargar el Almacén Sol faltante (PedidosYa, súper, QR, combustible, etc.).
