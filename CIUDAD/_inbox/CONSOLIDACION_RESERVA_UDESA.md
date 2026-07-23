# Consolidación — Reserva UDESA (monotributo LDO Pelesson)

Fecha de armado: 2026-07-23 (act. factura 23-jul)  
Fuente facturación: `SOL/Facturacion_Monotributo_SOL_2026.xlsx`  
Fuente cobros: extracto MP cuenta (`movimientos_cuenta_mp_2026.csv`)  
Fuente retiros: lista indicada por Sol, cruzada con `Dinero retirado Chris` del mismo extracto

## Criterio

Reserva **UDESA** = lo **cobrado** de facturas monotributo a LDO Pelesson  
menos los **retiros** indicados.

La **facturación** se registra aparte (puede estar PENDIENTE de cobro).  
No suma a la matriz `2026` ni a Almacén. Es bolsillo / reserva aparte (espejo operativo de la reserva MP etiquetada **Chris**).

## 0) Facturación monotributo Sol → LDO Pelesson

| Fecha factura | Período servicio | Cliente | CUIT | Importe ARS | Estado |
|---|---|---|---|---|---|
| 2026-07-23 | Junio 2026 | Aldo Pelesson / LDO Pelesson | 23-14614689-9 | 4.900.000 | PENDIENTE |
| | | | **Total facturado** | **4.900.000** | |

Registro canónico: `SOL/Facturacion_Monotributo_SOL_2026.xlsx` · `SOL/facturacion_monotributo_sol_2026.csv`

## 1) Cobros MP — LDO Pelesson (aún sin factura asociada en el registro)

| Fecha | OpID | Concepto | Monto ARS |
|---|---|---|---|
| 2026-05-15 | 158620980755 | Transferencia recibida PELESSON ALDO OSCAR | 5.000.000 |
| 2026-06-12 | 163006931853 | Transferencia recibida PELESSON ALDO OSCAR | 4.750.000 |
| | | **Total cobrado (extracto)** | **9.750.000** |

Nota: estos cobros no matchean 1:1 con la factura del 23-jul ($4.900.000, servicio junio). Quedan como cobros históricos Pelesson en la reserva hasta asociar facturas previas.  
Ruteo a reserva MP Chris: 16-may **5.000.000**; 12-jun **4.500.000** (del cobro 4.750.000). Diff no reservada del 2.º cobro: **250.000**.

## 2) Retiros a descontar

| Mes (según Sol) | Fecha MP (extracto) | OpID | Concepto MP | Monto Sol | Monto extracto |
|---|---|---|---|---|---|
| Junio | 2026-06-10 | 163418872656 | Dinero retirado Chris | 500.000 | 500.000 |
| Julio* | 2026-06-20 | 165013974648 | Dinero retirado Chris | 20.000 | 20.000 |
| Julio* | 2026-06-23 | 165530199038 | Dinero retirado Chris | 150.000 | 150.000 |
| Julio* | 2026-06-25 | 165824407784 | Dinero retirado Chris | 170.000 | **170.100** |
| | | | **Total retiros (lista Sol)** | **840.000** | |
| | | | **Total retiros (prevalece extracto)** | | **840.100** |

\* En el mensaje figuraban como “Julio”; en el extracto oficial caen a fines de **junio**. Prevalece el extracto.

## 3) Resultado consolidado (caja / cobrado)

| Concepto | ARS |
|---|---|
| (+) Total cobrado LDO Pelesson (extracto) | 9.750.000 |
| (−) Retiros (extracto) | 840.100 |
| **(=) Saldo Reserva UDESA (caja)** | **8.909.900** |

| Concepto facturación | ARS |
|---|---|
| Facturado (registro monotributo) | 4.900.000 |
| de ese total, cobrado | 0 |
| Pendiente de cobro | **4.900.000** |

## 4) Lectura rápida

- Factura nueva 23-jul · servicio junio · Aldo Pelesson CUIT 23-14614689-9: **$4.900.000** (PENDIENTE)
- Cobros Pelesson en MP (históricos): **$9.750.000**
- Retiros descontados: **$840.100**
- Saldo reserva UDESA (caja): **$8.909.900**
- A cobrar (factura julio): **$4.900.000**

Archivos espejo: `consolidacion_reserva_udesa.csv` · `Consolidacion_Reserva_UDESA.xlsx` · `SOL/Facturacion_Monotributo_SOL_2026.xlsx`
