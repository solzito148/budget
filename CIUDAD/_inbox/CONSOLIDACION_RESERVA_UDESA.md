# Consolidación — Reserva UDESA / Chris (monotributo LDO Pelesson)

Fecha de armado: 2026-07-23 (act. cobro → Reserva Chris)  
Fuente facturación: `SOL/Facturacion_Monotributo_SOL_2026.xlsx`  
Fuente cobros: extracto MP + confirmación Sol (cobro 23-jul a **Reserva Chris**; OpID pendiente)  
Fuente retiros: lista indicada por Sol, cruzada con `Dinero retirado Chris`

## Criterio

Reserva **UDESA** = reserva MP etiquetada **Chris**.  
Saldo = cobros LDO Pelesson ruteados a esa reserva − retiros de Chris.

No suma a la matriz `2026` ni a Almacén. Es bolsillo / reserva aparte (`Dinero reservado Chris` / `Dinero retirado Chris`).

## 0) Facturación monotributo Sol → LDO Pelesson

| Fecha | Comp. | Período | Cliente | CUIT | Importe ARS | CAE | Estado | Destino |
|---|---|---|---|---|---|---|---|---|
| 2026-07-23 | Factura C **00001-00000027** | 01/06/2026–30/06/2026 | PELESSON ALDO OSCAR | 23-14614689-9 | 4.900.000 | 86305115481677 | **COBRADO** 23/07/2026 MP | **Reserva Chris** |
| | | | | **Total facturado** | **4.900.000** | | | |

Concepto: honorarios profesionales · consultoría de sistemas (control interno / políticas y procedimientos).  
PDF: `SOL/_inbox/facturas_monotributo/27290390791_011_00001_00000027.pdf`  
Registro canónico: `SOL/Facturacion_Monotributo_SOL_2026.xlsx` · `SOL/facturacion_monotributo_sol_2026.csv`  
Nota: OpID del cobro/reserva 23-jul pendiente hasta el próximo extracto MP (el actual corta el 18-jul).

## 1) Cobros MP — LDO Pelesson

| Fecha | OpID | Concepto | Monto ARS |
|---|---|---|---|
| 2026-05-15 | 158620980755 | Transferencia recibida PELESSON ALDO OSCAR | 5.000.000 |
| 2026-06-12 | 163006931853 | Transferencia recibida PELESSON ALDO OSCAR | 4.750.000 |
| 2026-07-23 | *(pendiente extracto)* | Cobro MP Factura C 00001-00000027 | 4.900.000 |
| | | **Total cobrado** | **14.650.000** |

## 1b) Ruteo a Reserva Chris (`Dinero reservado Chris`)

| Fecha | OpID | Concepto | Monto ARS | Nota |
|---|---|---|---|---|
| 2026-05-16 | 158758469069 | Dinero reservado Chris | 5.000.000 | Cobro Pelesson 15-may |
| 2026-06-12 | 163024353651 | Dinero reservado Chris | 4.500.000 | Parcial del cobro 4.750.000 |
| 2026-07-23 | *(pendiente extracto)* | Dinero reservado Chris | 4.900.000 | Confirmación Sol: cobro factura va a Reserva Chris |
| | | **Total ruteado a Chris (Pelesson/factura)** | **14.400.000** | |

Cobro 23-jul: abona la factura C 00001-00000027 y **entra a Reserva Chris / UDESA**.

## 2) Retiros a descontar (`Dinero retirado Chris`)

| Mes (según Sol) | Fecha MP (extracto) | OpID | Concepto MP | Monto ARS |
|---|---|---|---|---|
| Junio | 2026-06-10 | 163418872656 | Dinero retirado Chris | 500.000 |
| Julio* | 2026-06-20 | 165013974648 | Dinero retirado Chris | 20.000 |
| Julio* | 2026-06-23 | 165530199038 | Dinero retirado Chris | 150.000 |
| Julio* | 2026-06-25 | 165824407784 | Dinero retirado Chris | **170.000** |
| | | | **Total retiros** | **840.000** |

\* En el mensaje figuraban como “Julio”; en el extracto oficial caen a fines de **junio**. El retiro de 170.000 queda consolidado en ese monto por indicación de Sol.

## 3) Resultado consolidado (Reserva Chris / UDESA)

| Concepto | ARS |
|---|---|
| (+) Total cobrado LDO Pelesson | 14.650.000 |
| (−) Retiros Chris | 840.000 |
| **(=) Saldo Reserva UDESA / Chris** | **13.810.000** |

| Concepto facturación | ARS |
|---|---|
| Facturado (registro monotributo) | 4.900.000 |
| de ese total, cobrado → Reserva Chris | 4.900.000 |
| Pendiente de cobro | **0** |

## 4) Lectura rápida

- Factura C 00001-00000027 · **$4.900.000** · COBRADO hoy con MP → **Reserva Chris**
- Cobros Pelesson totales: **$14.650.000**
- Retiros Chris descontados: **$840.000**
- Saldo reserva UDESA / Chris: **$13.810.000**
- Pendiente de cobro: **$0**

Archivos espejo: `consolidacion_reserva_udesa.csv` · `Consolidacion_Reserva_UDESA.xlsx` · `SOL/Facturacion_Monotributo_SOL_2026.xlsx`
