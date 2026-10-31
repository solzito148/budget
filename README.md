# Budget 2026 por propiedad

Cada carpeta representa una propiedad independiente. No se deben mezclar
gastos, fórmulas, clientes, medidores ni detalles entre propiedades.

## Estructura

- `CIUDAD/Gastos_CIUDAD_1132_2026.xlsx` — hogar (Ciudad de la Paz 1132)
- `SOL/Gastos_SOL_2026.xlsx` — maestro bolsillo Sol (regenerar con `SOL/build_gastos_sol.py`)
- `AVA/Gastos_AVA_2026.xlsx` — fideicomiso Ava Palpa
- `OHIGGINS/Gastos_OHIGGINS_2026.xlsx`
- `BONORINO/Gastos_BONORINO_2026.xlsx`

## Enlaces canónicos

| Qué | Link |
|-----|------|
| **GitHub (este repo)** | https://github.com/solzito148/budget |
| **GitHub · docs Claude** | https://github.com/solzito148/budget/tree/master/CLAUDE |
| **GitHub · prompt** | https://github.com/solzito148/budget/blob/master/CLAUDE/PROMPT_SISTEMA.md |
| **Drive · Excels `03_Finanzas_y_Gastos`** | https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29 |
| **Drive · docs Claude** | https://drive.google.com/drive/folders/1Lfu_3vQjZQ3L3tiBI5eAk661bnyC1UnW |

## Claude / documentación operativa

Paquete listo para Project instructions + knowledge:

- **`CLAUDE/PROMPT_SISTEMA.md`** — prompt completo para pegar en Claude
- **`CLAUDE/`** — lógicas, árbol de decisión, taxonomías, mapa de comercios, Drive

Reglas canónicas del agente: `.cursor/rules/*.mdc`.

## Reglas generales

- Filas: tipo, categoría y descripción del gasto.
- Columnas: enero a diciembre de 2026.
- Los montos se muestran sin decimales y se redondean hacia arriba.
- La hoja `Detalle` registra fecha, forma de pago, número de cliente y
  número de medidor cuando estén disponibles.
- Cada archivo solo puede contener datos de su propia propiedad.
- Toda alta/edición se overwrite también en Drive: https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29
