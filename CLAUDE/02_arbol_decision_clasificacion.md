# 02 — Árbol de decisión e implicancias

**GitHub:** https://github.com/solzito148/budget/tree/master/CLAUDE · **Drive Excels:** https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29 · **Drive docs:** https://drive.google.com/drive/folders/1Lfu_3vQjZQ3L3tiBI5eAk661bnyC1UnW

Usar este orden **siempre**. No saltear pasos.

## Paso 0 — ¿Es movimiento válido?

1. ¿Figura en un **resumen oficial** (MP / BBVA / banco / liquidación / mail con cupón)?  
   - **No** → no inventar. Pedir extracto o marcar `REVISAR`.
2. ¿Cancelado / devuelto / neto 0? → **ignorar** (no cargar).
3. ¿Transferencia propia Sol ↔ Sol (`Sanes Maria Sole` / `Maria Soledad Sanes`)? → **EXCLUIR**.
4. ¿Cuota préstamo / crédito Mercado Pago? → **EXCLUIR**.
5. ¿Falta fecha? → no cargar (o `REVISAR` con fecha tentativa + comentario explícito).

## Paso 1 — ¿Es “no gasto”?

Si el concepto es:

- Dinero retirado  
- Reserva automática / Dinero reservado  
- Pago de deuda MP  
- Pago de resumen de tarjeta  
- Inversión / venta-compra USD / bonos / PF / fondos  

→ `MP No Gasto` (+ espejo `Dinero Retirado` / `Pago de Tarjetas` / `Inversión`).  
**Nunca** Almacén ni matriz `2026` como gasto nuevo.

Si es **transferencia recibida / ingreso** → `Transferencias recibidas`.

## Paso 2 — ¿Es de una propiedad (no hogar / no personal)?

| Señal | Destino |
|-------|---------|
| Consorcio De Propiet / Patagonia CBU `0340116900116010765007` / CUIL `30714962473` | Expensas **CIUDAD** |
| Cons De Prop E Bonorino 82 | Expensas **BONORINO** |
| Consorcio De Copropi CUIT `30-58711816-5` / Estudio Kalmus / O’Higgins | Expensas **OHIGGINS** |
| Fideicomiso Ava Palpa / Sancor cuotas | **AVA** (50% Sol → GPS) |
| Telecentro | CIUDAD · TV CABLE |
| Personal / Personal pago directo | **OHIGGINS** · TV CABLE-INTERNET |
| Edesur cliente `3932799` | **BONORINO** · LUZ |
| Edenor titular Sol cuenta `8917502309` | CIUDAD · LUZ |
| Edenor Marta Chagas / INVALMAR | **OHIGGINS** · LUZ |
| Metrogas cliente `30010996459` | CIUDAD · GAS |
| Metrogas cliente `10139236800` (INVALMAR) | **OHIGGINS** · GAS |
| ABL ≤ $80.000 / partida `397789` | CIUDAD · ABL |
| ABL > $80.000 / partida `3016365` | **OHIGGINS** · ABL |
| Susana Carmen Ortiz | CIUDAD · LIMPIEZA CASA |
| PedidosYa Plus | CIUDAD · SUSCRIPCIONES |
| Spotify | CIUDAD · STREAMING |

**Implicancia:** estas filas **no** van a Almacén ni a GPS (salvo el 50% AVA Sol).

## Paso 3 — ¿Es consumo de casa / Ramón / combustible / entretenimiento hogareño?

Sí → `Almacen Detalle` (ALMACEN SOL). Familias:

- SUPERMERCADO  
- COMIDA (resto/café/verdulería/delivery comida; **no** Feed/Oporto)  
- Combustible (YPF/Shell/Generacion 20)  
- `RAMON · Comida|Veterinario|Otros`  
- ENTRETENIMIENTO (recitales/teatro/cine/cenas/social/salidas)  
- STREAMING TV (Directv)  
- FARMACIA · Farmacity / REPUESTOS · Sodastream  

Ver `03_almacen_taxonomia.md`.

## Paso 4 — Todo lo demás personal de Sol → GPS

Transporte, ropa/vestimenta, suscripciones IA/Apple, telefonía, farmacia salud (Maure/Selma), ML, envíos internacionales, ACA, River, Zurich, Feed/Copetin/Oporto, Promarine, peluquería, educación, funeral, transferencias a terceros sin regla de casa, etc.

Ver `04_gasto_personal_sol.md`.

## Implicancias de clasificación (resumen)

| Destino | Suma total casa `2026` | Entra a maestro SOL | Drive |
|---------|------------------------|---------------------|-------|
| Matriz CIUDAD / Detalle casa | Sí | Indirecto (casa ≠ GPS) | Sí CIUDAD |
| Almacen Detalle | Sí (vía ALMACEN) | Sí (`Almacén Sol` / Importe Sol) | Sí CIUDAD + regenerar SOL |
| Gasto Personal Sol | **No** | Sí | Sí CIUDAD + regenerar SOL |
| MP No Gasto / espejos | **No** | Sí (hoja MP No Gasto) | Sí |
| OHIGGINS / BONORINO / AVA | En su matriz | AVA 50% Sol sí en GPS | Su Excel + Drive |
| EXCLUIR / cancelados | No | No | No |

## Después de cargar

1. Actualizar Excel(s) canónicos de la propiedad.  
2. Si tocó GPS / Almacén Sol / AVA / BBVA / percepciones → `python3 SOL/build_gastos_sol.py`.  
3. Overwrite en Drive (raíz `03_Finanzas_y_Gastos`).  
4. Montos: enteros, **ceil** hacia arriba. Fecha obligatoria en toda fila.
