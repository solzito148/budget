# 11 — Ciclo mensual (cómo funciona mes a mes)

Indicado por Sol el 2026-11-24. **Obligatorio.**

## Reglas duras

1. **Nunca borrar meses anteriores.** Toda carga es **agregar**: filas nuevas al final, celdas vacías que se completan. No se vacían meses, no se reemplazan bloques, no se sube una versión con menos datos que la vigente.
   - Si un gasto está mal ubicado, se **agrega** en el destino correcto y se saca **solo esa fila**, dejando constancia en el comentario y en el commit.
   - Antes de guardar, correr `python3 tools/verificar_no_borrado.py <vigente.xlsx> <nuevo.xlsx>`: si algún mes perdió filas o importes, **no se guarda**.
2. **Todos los meses con todos los valores.** Cada servicio de cada propiedad tiene importe en cada mes cerrado. Un mes vacío es un faltante: se busca (extracto MP, resumen BBVA, Gmail recibidos/enviados) o se le pregunta a Sol. Lista en `Consolidado_2026.xlsx › Faltantes`.
3. **Rendición de CIUDAD todos los meses.** Al cerrar cada mes se arma la rendición CIUDAD 50/50 (hoja `Rendición CIUDAD` del consolidado + `Rendiciones_2026.xlsx › Rendición Chris`).
4. **Drive siempre al día.** Se parte de la versión vigente en Drive (`Budget_2026_vigente`) y se guarda ahí al terminar. Sol no sube archivos.

## Paso a paso de cada mes (M)

1. **Bajar** las planillas vigentes del Drive (`CIUDAD/_inbox/drive_excel_canonico.json › vigente.files`).
2. **Fuentes del mes M:**
   - Extracto de Mercado Pago (Sol lo pasa por chat).
   - Resúmenes BBVA Visa / Mastercard y TC Mercado Pago.
   - Gmail **recibidos**: facturas Edenor, Edesur, Metrogas (CIUDAD 30010996459 · INVALMAR/OHIGGINS 10139236800 · BONORINO 30010713734), Telecentro, Personal Flow, liquidaciones de expensas.
   - Gmail **enviados**: comprobantes de expensas (De Filippo = CIUDAD, Kalmus = OHIGGINS, pereyragladys = BONORINO) y cuota AVA (Sancor).
   - `ORIGINAL_Tabla_Susi` → Limpieza Susana (col. Y) y compras Susana.
   - **Gastos de Chris**: Sol los pasa por chat (mes + detalle + importe, sin fecha) → Almacén Chris.
2b. **Registro Gmail:** toda factura / comprobante / boleta leída de Gmail se agrega a `REGISTROS/registro_gmail_2026.csv` (propiedad, clave, concepto, mes, emisión, vencimiento, monto, pagado, tipo, fuente, nota, gmail_id) y se sube a `Budget_2026_vigente`. Metrogas y Edenor traen el importe solo en el HTML del mail (no en el resumen).
3. **Clasificar** con el árbol (`02_arbol_decision_clasificacion.md`); dudas → tabla por pantalla a Sol.
4. **Cargar agregando** en cada propiedad (matriz `2026` + `Detalle`), Almacén Sol/Chris, Gasto Personal Sol, MP No Gasto.
5. **AVA:** cuando Sol pasa el aviso de la cuota → registrarla PENDIENTE (monto, vencimiento, 50/50) **y crear el evento de pago en su Google Calendar** el día del vencimiento (todo el día, recordatorios 2 días antes y el mismo día). Al pagarse: PAGADO, MEP vendedor del vencimiento (argentinadatos), matriz del mes y 50% Sol en Gasto Personal Sol.
6. **Regenerar:** `python3 SOL/build_gastos_sol.py` y `python3 tools/consolidar_2026.py`.
7. **Controlar:**
   - `tools/verificar_no_borrado.py` contra la versión vigente (ningún mes anterior pierde datos).
   - `Consolidado › Faltantes`: resolver o preguntar.
8. **Rendición CIUDAD del mes** (ver abajo) y avisarle a Sol el saldo.
9. **Guardar en Drive** (`Budget_2026_vigente`), actualizar ids en el JSON, commit + push.
10. Anotar en `10_aprendizajes_sol.md` cualquier indicación nueva de Sol.

## Rendición CIUDAD (cálculo)

| Paso | Cálculo |
|---|---|
| Servicios CIUDAD | expensas + luz + gas + ABL + Telecentro + limpieza + streaming + suscripciones de casa |
| Cuota servicios | Servicios ÷ 2 |
| Cuota almacén | (Almacén Sol + Almacén Chris) ÷ 2 |
| A cargo de Chris | cuota servicios + cuota almacén |
| Chris debe a Sol | A cargo de Chris − Almacén Chris |

- Ene–abr 2026: se calcula y registra, **no se cobra**.
- AVA, OHIGGINS y BONORINO **no entran** (AVA Chris lo paga aparte).
- Pendiente de definición con Sol (24/11): qué parte del Almacén Sol es compartida (¿solo `GASTOS2026 › Detalle 2026` o todo lo clasificado como almacén?).
