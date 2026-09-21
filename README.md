# Budget 2026 por propiedad

Cada carpeta representa una propiedad independiente. No se deben mezclar
gastos, fórmulas, clientes, medidores ni detalles entre propiedades.

## Estructura (repo)

- `CIUDAD/Gastos_CIUDAD_1132_2026.xlsx`
- `AVA/Gastos_AVA_2026.xlsx`
- `OHIGGINS/Gastos_OHIGGINS_2026.xlsx`
- `BONORINO/Gastos_BONORINO_2026.xlsx`

## Archivos canónicos (iCloud) — actualizar todos los meses

Ruta base:

`/Users/sol/Library/Mobile Documents/com~apple~CloudDocs/GastosCiudad/`

- **BONORINO** → `…/BONORINO/Gastos_BONORINO_2026.xlsx` (todos los gastos de la propiedad, mes a mes)
- **CIUDAD** → `…/CIUDAD/Gastos_CIUDAD_1132_2026.xlsx`
- **OHIGGINS** → `…/OHIGGINS/Gastos_OHIGGINS_2026.xlsx`

Tras editar en el repo (p. ej. desde un Cloud Agent), en la Mac:

```bash
python3 scripts/sync_gastos_icloud.py --only BONORINO
# o todos:
python3 scripts/sync_gastos_icloud.py
```

## Reglas generales

- Filas: tipo, categoría y descripción del gasto.
- Columnas: enero a diciembre de 2026.
- Los montos se muestran sin decimales y se redondean hacia arriba.
- La hoja `Detalle` registra fecha, forma de pago, número de cliente y
  número de medidor cuando estén disponibles.
- Cada archivo solo puede contener datos de su propia propiedad.
- El Excel en iCloud `GastosCiudad/` es la fuente de verdad operativa.
