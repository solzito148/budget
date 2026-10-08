# Budget 2026 por propiedad

Cada carpeta representa una propiedad independiente. No se deben mezclar
gastos, fórmulas, clientes, medidores ni detalles entre propiedades.

## Estructura

- `CIUDAD/Gastos_CIUDAD_1132_2026.xlsx`
- `AVA/Gastos_AVA_2026.xlsx`
- `OHIGGINS/Gastos_OHIGGINS_2026.xlsx`
- `BONORINO/Gastos_BONORINO_2026.xlsx`

## Reglas generales

- Filas: tipo, categoría y descripción del gasto.
- Columnas: enero a diciembre de 2026.
- Los montos se muestran sin decimales y se redondean hacia arriba.
- La hoja `Detalle` registra fecha, forma de pago, número de cliente y
  número de medidor cuando estén disponibles.
- Cada archivo solo puede contener datos de su propia propiedad.

La planilla de CIUDAD se mantiene sincronizada con su archivo canónico en
Google Drive.

## Sync semanal Gmail → planillas

Todos los **viernes** (fin de día ART) un agente debe leer
`.cursor/skills/gmail-gastos-viernes/SKILL.md`, revisar Gmail
(`s.sanes@gmail.com`), cargar gastos faltantes en CIUDAD / BONORINO /
OHIGGINS / AVA y regenerar `SOL/Gastos_SOL_2026.xlsx`.

Helper de queries:

```bash
python3 scripts/gmail_gastos_viernes_queries.py
```
