# Automation: Gmail gastos — todos los viernes

Configuración lista para crear en [cursor.com/automations](https://cursor.com/automations).

## Settings

| Campo | Valor |
|-------|--------|
| **Name** | `Gmail gastos viernes — CIUDAD / BONORINO / OHIGGINS / AVA` |
| **Trigger** | Scheduled (cron) |
| **Cron** | `CRON_TZ=America/Argentina/Buenos_Aires 0 18 * * 5` |
| **Equiv. UTC** | `0 21 * * 5` (viernes 18:00 ART = 21:00 UTC) |
| **Repository** | `solzito148/budget` (**obligatorio** — debe editar Excel) |
| **Environment** | el del repo budget (Composio Gmail activo) |
| **Tools** | Gmail/Composio MCP; git push OK |

## Prompt (copiar tal cual)

```
Ejecutá de punta a punta el skill del repo:

  .cursor/skills/gmail-gastos-viernes/SKILL.md

Tarea semanal (viernes ART):
1. Revisá Gmail de s.sanes@gmail.com (Composio, conexión sol-gastos) desde el viernes anterior hasta hoy.
2. Extraé pagos, transferencias, servicios, expensas, utilidades, cuotas AVA y settlements MP/BBVA.
3. Registrá solo montos confirmados en resúmenes/comprobantes oficiales en:
   - CIUDAD/Gastos_CIUDAD_1132_2026.xlsx  (Ciudad de la Paz 1132)
   - BONORINO/Gastos_BONORINO_2026.xlsx
   - OHIGGINS/Gastos_OHIGGINS_2026.xlsx
   - AVA/Gastos_AVA_2026.xlsx + AVA/cuotas_fideicomiso_ava.csv
4. Seguí .cursor/rules/contexto-hogar-propiedades.mdc y .cursor/rules/ciudad-categorias-mercadopago.mdc (no mezclar propiedades; Almacén vs GPS; MP No Gasto).
5. Corré: python3 SOL/build_gastos_sol.py
6. Commit, push y abrí/actualizá el PR con lo cargado y lo pendiente.

Helper de queries: python3 scripts/gmail_gastos_viernes_queries.py
No inventes movimientos. Toda fila lleva fecha. Montos enteros ceil.
```

## Cómo activarla (1 vez)

1. Abrí https://cursor.com/automations → **New automation**
2. Trigger → **Schedule** → cron de arriba
3. Pegá el prompt
4. Elegí repo **budget**
5. Save + Enable

Sin este paso el skill queda listo pero no corre solo; con la Automation activa, cada viernes un Cloud Agent hace la carga.
