# Gastos — carpeta canónica (espejo Google Drive)

Carpeta de Drive:
https://drive.google.com/drive/folders/1cWkY56sBG-k7ZmdzIdRImUr4EyHHSE7x

Folder ID: `1cWkY56sBG-k7ZmdzIdRImUr4EyHHSE7x`

## Archivos en esta carpeta

| Archivo | Origen operativo |
|---------|------------------|
| `Gastos_CIUDAD_1132_2026.xlsx` | `CIUDAD/` |
| `Gastos_BONORINO_2026.xlsx` | `BONORINO/` |
| `Gastos_OHIGGINS_2026.xlsx` | `OHIGGINS/` |
| `Gastos_SOL_2026.xlsx` | `SOL/` (maestro personal Sol; regenerar con `python3 SOL/build_gastos_sol.py`) |

## Flujo de trabajo

1. Editar el Excel de la propiedad (`CIUDAD/`, `BONORINO/`, `OHIGGINS/`) o regenerar `SOL/`.
2. Correr `python3 scripts/sync_gastos_drive.py` para copiar a `Gastos/` y, si hay credenciales, subir a Drive.
3. La carpeta Drive es la copia compartida de referencia; no mezclar propiedades en un solo libro.
