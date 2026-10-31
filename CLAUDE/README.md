# Paquete Claude — Budget familiar Sol 2026

Documentación operativa del sistema de gastos (CIUDAD / SOL / OHIGGINS / BONORINO / AVA) para usar en **Claude Projects** o como knowledge base.

## Cómo usarlo en Claude

1. Creá un **Project** en Claude.
2. Pegá el contenido completo de [`PROMPT_SISTEMA.md`](PROMPT_SISTEMA.md) en **Project instructions** (o Custom instructions).
3. Subí como knowledge / archivos del Project:
   - este directorio `CLAUDE/` completo
   - opcional: `.cursor/rules/*.mdc`, `CIUDAD/_inbox/NORMA_MP_NO_GASTO.md`, `CIUDAD/_inbox/utilidades_cuentas.csv`, `CIUDAD/_inbox/abl_partidas.csv`, `AVA/RESUMEN_CUOTAS_FIDEICOMISO.md`
4. Cuando cargues extractos (PDF/CSV MP, BBVA, mails), adjuntá el resumen oficial: **siempre prevalece sobre el chat**.

**Copia en Drive:** [CLAUDE_DOCS_Budget_Sol](https://drive.google.com/drive/folders/1Lfu_3vQjZQ3L3tiBI5eAk661bnyC1UnW) (dentro de `03_Finanzas_y_Gastos`).

## Índice de archivos

| Archivo | Contenido |
|---------|-----------|
| [`PROMPT_SISTEMA.md`](PROMPT_SISTEMA.md) | **Prompt completo** listo para pegar |
| [`01_contexto_y_arquitectura.md`](01_contexto_y_arquitectura.md) | Quién es quién, archivos, hojas, Drive |
| [`02_arbol_decision_clasificacion.md`](02_arbol_decision_clasificacion.md) | Flujo de decisión + implicancias |
| [`03_almacen_taxonomia.md`](03_almacen_taxonomia.md) | ALMACEN SOL: etiquetas y excepciones |
| [`04_gasto_personal_sol.md`](04_gasto_personal_sol.md) | GPS: categorías personales |
| [`05_propiedades_utilidades_expensas.md`](05_propiedades_utilidades_expensas.md) | Expensas, luz, gas, TV, ABL |
| [`06_mp_no_gasto_inversiones.md`](06_mp_no_gasto_inversiones.md) | Retiros, reservas, pagos TC, inversión |
| [`07_ava_bbva_cashflow.md`](07_ava_bbva_cashflow.md) | AVA 50/50, MEP, BBVA, compromisos |
| [`08_drive_formato_fuentes.md`](08_drive_formato_fuentes.md) | Jerarquía de fuentes, formato, Drive |
| [`09_mapa_comercios.md`](09_mapa_comercios.md) | Lookup rápido comercio → destino |

## Fuente de verdad en el repo

Las reglas canónicas viven también en:

- `.cursor/rules/ciudad-categorias-mercadopago.mdc`
- `.cursor/rules/contexto-hogar-propiedades.mdc`
- `.cursor/rules/drive-excel-canonico.mdc`
- `CIUDAD/_inbox/drive_excel_canonico.json`
- `CIUDAD/_inbox/NORMA_MP_NO_GASTO.md`

Si hay conflicto entre un MD de `CLAUDE/` y esas reglas `.mdc`, **prevalecen las `.mdc` + resúmenes oficiales**.
