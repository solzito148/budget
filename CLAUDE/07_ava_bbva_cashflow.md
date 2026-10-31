# 07 — AVA, BBVA y cashflow de compromisos

**GitHub:** https://github.com/solzito148/budget/tree/master/CLAUDE · **Drive Excels:** https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29 · **Drive docs:** https://drive.google.com/drive/folders/1Lfu_3vQjZQ3L3tiBI5eAk661bnyC1UnW

## AVA Fideicomiso (Ava Palpa)

- Archivo: `AVA/Gastos_AVA_2026.xlsx`  
  Hojas: `Cuotas Fideicomiso`, `Análisis Cuota vs IPC+CAC`, `2026`, `Detalle`
- Fuente CSV: `AVA/cuotas_fideicomiso_ava.csv`
- Resumen humano: `AVA/RESUMEN_CUOTAS_FIDEICOMISO.md`

### Split 50/50

Cada cuota PAGADA: **50% Sol / 50% Chris**.  
El 50% Sol → `Gasto Personal Sol` con comentario `AVA Fideicomiso · cuota N`.  
Ese 50% ya está en GPS: **no duplicarlo** como gasto nuevo en Cashflow Compromisos.

### TC MEP (obligatorio)

Al cargar/actualizar cada cuota:

1. Buscar **dólar MEP tipo vendedor (venta / bolsa)** de la **fecha de vencimiento**.  
2. Si no hay rueda → último hábil anterior.  
3. `Valor USD = Monto ARS ÷ TC MEP venta`.  
4. Nunca reusar un TC fijo ni asumir USD constante.  
5. Fuente habitual: `api.argentinadatos.com` (dólar bolsa) / Rava.

Cuotas futuras PENDIENTE: sin monto USD hasta factura + TC del vencimiento.

### Análisis vs IPC+CAC

Mes a mes: `% Δ` cuota vs anterior, comparado con **IPC INDEC + índice CAC/CAMARCO** del mes de referencia (mes calendario anterior al período de la cuota).  
Veredicto: subió **MÁS** o **MENOS** que `IPC+CAC`.

## BBVA — Cashflow Cuotas

Hoja `Cashflow Cuotas BBVA` (CIUDAD y espejo SOL): estados **PAGADO** / **PENDIENTE** según resúmenes oficiales Visa/Mastercard.

Utilidades con débito automático (Telecentro, Edenor, Metrogas CIUDAD) se asignan por **fecha de cobro** del resumen.

## Cashflow Compromisos

Une:

1. BBVA TC (PAGADO/PENDIENTE)  
2. AVA (PAGADO con 50% Sol visible; PENDIENTE futuras sin monto hasta factura)

Sirve para ver **caja comprometida**, no para re-sumar gastos ya cargados en GPS/matriz.

## Percepciones USD en resúmenes BBVA / TC MP

Van a GPS (`PERCEPCIÓN USD · …`), no a casa. Espejos CSV en `CIUDAD/_inbox/tarjetas_bbva/` y `tarjetas_mp/`.

## Liquidación Chris (contexto operativo)

Cuando se analiza “cuánto puso Chris / cuánto debe del mes”, usar matriz casa + Almacén Chris si existe + liquidaciones; no mezclar con GPS de Sol ni con AVA 50% Chris salvo que Sol lo pida explícito.
