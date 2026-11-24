# 05 — Propiedades, utilidades y expensas

**GitHub:** https://github.com/solzito148/budget/tree/master/CLAUDE · **Drive Excels:** https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29 · **Drive docs:** https://drive.google.com/drive/folders/1Lfu_3vQjZQ3L3tiBI5eAk661bnyC1UnW

## Servicios: importe todos los meses (regla general)

- **Cada servicio de cada propiedad tiene que tener su importe cargado en todos los meses** (expensas, luz, gas, ABL, TV/internet, limpieza, streaming). Un mes vacío es un faltante a resolver, no un cero.
- **Expensas:** se pagan por **transferencia desde Mercado Pago** al consorcio. El respaldo es el **mail enviado** (Gmail `in:sent`) donde Sol reenvía la liquidación al administrador con el **comprobante de transferencia** adjunto (imagen MP: fecha, monto, CBU, n.º de operación). Mes de la matriz = **período** de la liquidación (asunto del mail), no la fecha de pago.
  - CIUDAD → De Filippo (`expensasdefilippo@gmail.com`, asunto "Envío expensa CIUDAD DE LA PAZ 1132")
  - OHIGGINS → Kalmus (`expensas@estudiokalmus.com.ar`, asunto "Expensas O´HIGGINS 2219 - {mes}") · desde ago-2026 cuenta nueva Galicia CBU `0070692220000003412822`
  - BONORINO → `pereyragladys@hotmail.com` (asunto "Liquidacion del mes de MM/2026")
- **BONORINO** también tiene ABL partida **1084069** y Metrogas cliente **30010713734** (pagos por MP). Un ABL de ~$28.000 o un Metrogas pagado por MP **no** son automáticamente CIUDAD: CIUDAD paga Metrogas por débito BBVA.
- **Antes de subir un Excel a Drive, bajar la versión actual del Drive y partir de ella.** Nunca pisar el Drive con una versión vieja del repo.

## Expensas — mapear por consorcio / CBU / CUIT

| Señal | Propiedad | Archivo |
|-------|-----------|---------|
| Consorcio De Propiet / Consorcio De Propietarios | CIUDAD 1132 | CIUDAD |
| Banco Patagonia · CBU `0340116900116010765007` · CUIL `30714962473` | CIUDAD (aunque el nombre varíe) | CIUDAD |
| Cons De Prop E Bonorino 82 | BONORINO | BONORINO |
| Consorcio De Copropi · CUIT `30-58711816-5` · Estudio Kalmus · O’Higgins 2219 8 | OHIGGINS | OHIGGINS |
| Fideicomiso Ava Palpa | AVA | AVA |

También hoja `Expensas Consorcios` en CIUDAD. **Nunca** Almacén ni GPS.

### Período vs fecha de pago (Gmail Sent / Defilippo)

A veces el comprobante de mayo se envía en junio. **Usar el período del cupón/mail/adjunto**, no solo la fecha de pago MP. Ejemplo: pago 10-jun → expensas mayo; pago 9-jul → junio.

## Utilidades — cuentas fijas

| Servicio | Propiedad | Identificador | Medio típico |
|----------|-----------|---------------|--------------|
| Telecentro | CIUDAD | cliente ~3491192 | BBVA Visa DEB |
| Edenor LUZ | CIUDAD | titular Sol · cuenta `00008917502309` | BBVA Visa |
| Edenor LUZ | OHIGGINS | Marta Chagas · `9267222037` · O’Higgins 2225 8 | BBVA Visa |
| Edenor / INVALMAR | OHIGGINS | INVALMAR/Ivilmar | — |
| Metrogas GAS | CIUDAD | cliente `30010996459` · medidor `40049923` | BBVA DEB |
| Metrogas GAS | OHIGGINS | cliente `10139236800` · medidor `8683287` · INVALMAR | MP |
| Edesur LUZ | BONORINO | cliente `3932799` | MP |
| Personal TV/Internet | OHIGGINS | Personal pago directo | — |

Referencia CSV: `CIUDAD/_inbox/utilidades_cuentas.csv`.

### Mes del gasto (GAS y débitos TC)

Asignar al mes de la **fecha de cobro** (pago efectivo BBVA / vencimiento), **no** al período de consumo de la boleta. Ver `CIUDAD/_inbox/gas_ciudad_reasignacion_cobro.md`.

## ABL (AGIP)

El settlement MP **no trae partida** → mapear a mano:

| Propiedad | Partida | Regla monto | Archivo |
|-----------|---------|-------------|---------|
| CIUDAD | `397789` | ≤ $80.000 | CIUDAD |
| OHIGGINS | `3016365` | > $80.000 | OHIGGINS |
| AVA / BONORINO | pendiente | — | — |

## Limpieza CIUDAD

**Susana Carmen Ortiz** → CIUDAD · `LIMPIEZA CASA` / `Limpieza Susana` (Detalle + matriz `2026`). Fuente única: Google Sheet `ORIGINAL_Tabla_Susi` (id `173WtFMCAwHJvcDUe4cfqkyeaV-J5-UQbfAqNop4v7bU`), solapa del año. Monto mensual de Limpieza = columna Y "Total Limpieza Casa Ciudad (sueldo + ARCA)" de la fila Total del mes (sueldo + aportes/contribuciones/ART), ceil. Las compras de almacén que hace Susana (columna V "Super", total mensual en columna Z) van a `Almacen Detalle` (ALMACEN SOL) como `SUPERMERCADO · Compras Susana`, una fila por semana. Las transferencias MP a Susana NO se suman desde el extracto: ya están en la Tabla Susi.

## PedidosYa Plus / Spotify

- PedidosYa Plus → CIUDAD `SUSCRIPCIONES` (todos los meses activos)  
- Spotify → CIUDAD `STREAMING` (USD × TC MEP del mes)
