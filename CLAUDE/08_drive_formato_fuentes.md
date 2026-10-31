# 08 — Fuentes, formato y Drive

## Jerarquía de fuentes (obligatoria)

1. **Resúmenes oficiales** (PDF/CSV MP, TC débito/crédito MP, BBVA, extractos bancarios, liquidaciones, cupones de expensas)  
2. Texto del chat (contexto, reglas, aclaraciones)  
3. Inferencias del asistente  

Si hay conflicto de monto / fecha / comercio / neto / cancelación / mes → **prevalece el resumen**.  
Solo Sol puede autorizar corregir un error del extracto.

**No inventar movimientos.** Solo lo que figura en un resumen. Toda fila con fecha.

## Formato de carga

- Montos **enteros**, redondeo **ceil** (hacia arriba)  
- Fecha del gasto = fecha del movimiento / cobro según resumen  
- Editar **solo** el Excel canónico de la propiedad correspondiente  
- Comentario con comercio/detalle cuando la categoría sea genérica (farmacia, AVA cuota N, etc.)

## Flujo de trabajo completo

```
Extracto oficial
  → clasificar (árbol 02)
  → escribir Excel propiedad (CIUDAD / OHIGGINS / BONORINO / AVA)
  → si tocó Sol/Almacén/AVA/BBVA/percepciones: regenerar SOL/build_gastos_sol.py
  → overwrite mismos file ids en Drive raíz 03_Finanzas_y_Gastos
  → verificar listado en carpeta
  → recién ahí dar por cerrado
```

## Drive

- Carpeta: https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29  
- Inventario: `CIUDAD/_inbox/drive_excel_canonico.json`  
- Regla: `.cursor/rules/drive-excel-canonico.mdc`  
- xlsx en la **raíz** (no en subcarpeta `Gastos/`)  
- Archivo nuevo → crear en esa carpeta y registrar file id en regla + JSON  

## Qué responder al usuario

Al cerrar una carga, informar de forma corta:

- Qué se cargó (archivo + hoja + mes + monto)  
- Qué se excluyó y por qué  
- Si quedó algo `REVISAR` / `A CONFIRMAR`  
- Confirmación de sync Drive (si aplica en el entorno)
