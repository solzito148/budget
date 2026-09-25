---
name: gmail-gastos-viernes
description: >-
  Every Friday (America/Argentina/Buenos_Aires), extract payments/utilities/transfers
  from Gmail (s.sanes@gmail.com via Composio) and register them into CIUDAD, BONORINO,
  OHIGGINS, AVA workbooks plus GPS/Almacén/MP No Gasto, then regenerate SOL master.
  Use when asked to cargar gastos from mails, Friday expense sync, or weekly Gmail→Excel.
---

# Gmail → planillas de gastos (viernes)

## Cuándo

- **Todos los viernes, fin de día ART** (`America/Argentina/Buenos_Aires`).
- También al pedir: “revisá los mails”, “cargar gastos”, “sync Gmail”, “gastos de la semana”.

## Objetivo

Registrar en las planillas **solo lo confirmado en resúmenes / comprobantes oficiales** de Gmail:

| Destino | Archivo |
|---------|---------|
| Hogar CIUDAD 1132 | `CIUDAD/Gastos_CIUDAD_1132_2026.xlsx` |
| BONORINO | `BONORINO/Gastos_BONORINO_2026.xlsx` |
| OHIGGINS | `OHIGGINS/Gastos_OHIGGINS_2026.xlsx` |
| AVA Fideicomiso | `AVA/Gastos_AVA_2026.xlsx` + `AVA/cuotas_fideicomiso_ava.csv` |
| Maestro Sol | regenerar `SOL/Gastos_SOL_2026.xlsx` con `python3 SOL/build_gastos_sol.py` |

Seguir siempre `.cursor/rules/contexto-hogar-propiedades.mdc` y `.cursor/rules/ciudad-categorias-mercadopago.mdc`.

## Prerrequisitos

1. Composio Gmail conectado a `s.sanes@gmail.com` (alias `sol-gastos`).
2. `openpyxl` instalado (`pip install openpyxl` si falta).
3. Ventana de búsqueda: **desde el viernes anterior (inclusive) hasta hoy**, o desde la última fecha ya cargada en Detalle / Expensas si es más reciente.

Helper opcional: `python3 scripts/gmail_gastos_viernes_queries.py` (imprime queries Gmail y checklist).

## Paso 1 — Buscar mails (Composio)

Usar `COMPOSIO_SEARCH_TOOLS` → `GMAIL_FETCH_EMAILS` (paginar con `nextPageToken`; `include_payload=false` / listado liviano primero; hidratar con `GMAIL_FETCH_MESSAGE_BY_MESSAGE_ID` format=`full` + `GMAIL_GET_ATTACHMENT` cuando haga falta).

**Queries canónicas** (ajustar `after:` / `before:` a la ventana):

```
(subject:expensas OR from:kalmus OR from:filippo OR "consorcio" OR "estudio") after:YYYY/MM/DD
(Telecentro OR Edenor OR Edesur OR Metrogas OR "Personal" OR AGIP OR ABL) after:YYYY/MM/DD
("Fideicomiso" OR "Ava Palpa" OR Sancor OR "cuota") after:YYYY/MM/DD
(Mercado Pago OR "resumen de cuenta" OR "pago de resumen" OR BBVA OR "Visa Signature") after:YYYY/MM/DD
(subject:comprobante OR subject:factura OR has:attachment filename:pdf) after:YYYY/MM/DD
```

Remitentes / temas habituales:

- **Expensas CIUDAD** — Filippo / Consorcio / Patagonia CBU `0340116900116010765007`
- **Expensas OHIGGINS** — Estudio Kalmus / Consorcio De Copropi CUIT `30-58711816-5`
- **Expensas BONORINO** — Cons De Prop E Bonorino 82
- **AVA** — Sancor / Estudio CMS / factura cuota fideicomiso
- **Utilidades** — Telecentro (CIUDAD), Edenor Sol (CIUDAD) / Edenor Chagas (OHIGGINS), Metrogas CIUDAD vs INVALMAR (OHIGGINS), Edesur (BONORINO), Personal (OHIGGINS), ABL AGIP (partida)
- **Pagos** — Mercado Pago resumen / pago de deuda / retiros; BBVA débitos automáticos

## Paso 2 — Clasificar (no inventar)

- Resúmenes oficiales **ganan** sobre el chat.
- Montos enteros, **ceil** hacia arriba.
- **Fecha obligatoria** en toda fila (fecha de gasto / cobro según el extracto).
- Periodo de expensas: usar el **período del comprobante/mail**, no solo la fecha de pago.
- Utilidades con débito TC: mes = **fecha de cobro** BBVA, no el período de consumo.
- ABL: partida `397789` ≤$80k → CIUDAD; `3016365` >$80k → OHIGGINS.
- AVA: 50% Sol / 50% Chris; TC MEP **venta** del día de vencimiento (`api.argentinadatos.com` dólar bolsa); actualizar CSV + Excel + `RESUMEN_CUOTAS_FIDEICOMISO.md` + Cashflow Compromisos.
- MP No Gasto: retiros, reservas, pago de deuda/resumen, inversiones — **nunca** Almacén ni total `2026`.
- Almacén vs GPS: reglas fijas (Ramón→Almacén, Feed→GPS, PedidosYa Plus→CIUDAD SUSCRIPCIONES, etc.).

## Paso 3 — Cargar en Excel

Por cada ítem confirmado y **aún no cargado** (dedupe por fecha+monto+comercio/concepto):

1. Hoja `Detalle` (y matriz `2026` si suma al total de la propiedad).
2. `Expensas Consorcios` cuando aplique (CIUDAD).
3. `Almacen Detalle` / `Gasto Personal Sol` / `MP No Gasto` / `Pago de Tarjetas` / `Dinero Retirado` / `Inversión` según reglas.
4. Propiedad correcta: **nunca** mezclar CIUDAD ↔ BONORINO ↔ OHIGGINS ↔ AVA.

No cargar: cancelados/devueltos (neto 0), transferencias propias Sol↔Sol, cuotas préstamo MP, ABL solo “emisión” sin pago, facturas sin cobro confirmado (dejar en checklist “pendiente”).

## Paso 4 — Regenerar maestro SOL

```bash
python3 SOL/build_gastos_sol.py
```

Verificar que GPS, Importe Sol Almacén, percepciones USD, Cashflow BBVA y AVA 50% Sol queden consistentes.

## Paso 5 — Commit / PR

Branch `cursor/<descriptive>-548e` (o la del agente), commit descriptivo, push, actualizar PR. En el body del PR listar:

- Qué se cargó (propiedad · concepto · mes · monto)
- Qué quedó **pendiente** (sin pago / sin monto / solo emisión)

## Checklist rápido

- [ ] Ventana Gmail definida y paginada completa
- [ ] Expensas CIUDAD / OHIGGINS / BONORINO
- [ ] Utilidades por propiedad (Telecentro, Edenor×2, Metrogas×2, Edesur, Personal, ABL)
- [ ] AVA cuota(s) del período + MEP venta + 50% Sol GPS + Cashflow Compromisos
- [ ] MP No Gasto / pagos de resumen / retiros
- [ ] Almacén + GPS nuevos del período
- [ ] `build_gastos_sol.py` corrido
- [ ] Commit + push + PR actualizado

## Automatización Cursor

Crear (o verificar) una **Cloud Automation** en Cursor con cron viernes ~18:00 ART cuyo prompt sea:

> Ejecutá el skill `.cursor/skills/gmail-gastos-viernes/SKILL.md`: revisá Gmail de s.sanes@gmail.com, cargá gastos faltantes en CIUDAD, BONORINO, OHIGGINS y AVA, regenerá SOL, commit/push/PR.

Si no hay Automation, al arrancar un agente el viernes (o al pedir “gastos del viernes”) leer este skill y ejecutarlo de punta a punta.
