# Consolidación — Reserva UDESA (monotributo LDO Pelesson)

Fecha de armado: 2026-07-23  
Fuente cobros: extracto MP cuenta (`movimientos_cuenta_mp_2026.csv`)  
Fuente retiros: lista indicada por Sol, cruzada con `Dinero retirado Chris` del mismo extracto

## Criterio

Reserva **UDESA** = lo **cobrado / facturado en monotributo a LDO Pelesson (Aldo Oscar Pelesson)**  
menos los **retiros** indicados (junio / julio).

No suma a la matriz `2026` ni a Almacén. Es bolsillo / reserva aparte (espejo operativo de la reserva MP etiquetada **Chris**, que es donde entró el dinero de Pelesson).

## 1) Ingresos — facturado / cobrado LDO Pelesson

| Fecha | OpID | Concepto | Monto ARS |
|---|---|---|---|
| 2026-05-15 | 158620980755 | Transferencia recibida PELESSON ALDO OSCAR | 5.000.000 |
| 2026-06-12 | 163006931853 | Transferencia recibida PELESSON ALDO OSCAR | 4.750.000 |
| | | **Total facturado / cobrado** | **9.750.000** |

Nota de ruteo a reserva MP: el 16-may se reservó **5.000.000** a Chris; el 12-jun se reservó **4.500.000** a Chris (del cobro de 4.750.000). Diferencia no reservada del 2.º cobro: **250.000**.

## 2) Retiros a descontar

| Mes (según Sol) | Fecha MP (extracto) | OpID | Concepto MP | Monto Sol | Monto extracto |
|---|---|---|---|---|---|
| Junio | 2026-06-10 | 163418872656 | Dinero retirado Chris | 500.000 | 500.000 |
| Julio* | 2026-06-20 | 165013974648 | Dinero retirado Chris | 20.000 | 20.000 |
| Julio* | 2026-06-23 | 165530199038 | Dinero retirado Chris | 150.000 | 150.000 |
| Julio* | 2026-06-25 | 165824407784 | Dinero retirado Chris | 170.000 | **170.100** |
| | | | **Total retiros (lista Sol)** | **840.000** | |
| | | | **Total retiros (prevalece extracto)** | | **840.100** |

\* En el mensaje figuraban como “Julio”; en el extracto oficial caen a fines de **junio**. Prevalece el extracto para fecha y para el centavo del último retiro (170.100).

## 3) Resultado consolidado

| Concepto | ARS |
|---|---|
| (+) Total facturado / cobrado LDO Pelesson | 9.750.000 |
| (−) Retiros (extracto) | 840.100 |
| **(=) Saldo Reserva UDESA** | **8.909.900** |

Con los montos redondeados de la lista (170.000 en lugar de 170.100):

| Concepto | ARS |
|---|---|
| (+) Total facturado / cobrado | 9.750.000 |
| (−) Retiros (lista) | 840.000 |
| **(=) Saldo (lista)** | **8.910.000** |

## 4) Lectura rápida

- Cobros Pelesson (monotributo): **$9.750.000**
- Retiros descontados: **$840.100** (extracto) / **$840.000** (lista)
- Saldo reserva UDESA: **$8.909.900** (extracto) ≈ **$8.910.000** (lista)

Archivos espejo: `consolidacion_reserva_udesa.csv` · `Consolidacion_Reserva_UDESA.xlsx`
