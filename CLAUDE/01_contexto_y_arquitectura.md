# 01 — Contexto y arquitectura

## Personas y roles

| Persona | Rol | Implicancia |
|---------|-----|-------------|
| **Sol** (Maria Soledad Sanes) | Titular / dueña del sistema | Decisiones finales; bolsillo personal = GPS |
| **Christian (Chris)** | Marido; vive en CIUDAD | Gastos de casa en CIUDAD; liquidación de su parte; AVA 50% |
| **Ramón** | Perro del hogar CIUDAD | **Siempre** Almacén (`RAMON · …`), nunca GPS |

## Propiedades / vehículos (nunca mezclar)

| Código | Qué es | Archivo | Qué registra |
|--------|--------|---------|--------------|
| **CIUDAD 1132** | Hogar donde viven Sol, Chris y Ramón · Ciudad de la Paz 1132 | `CIUDAD/Gastos_CIUDAD_1132_2026.xlsx` | Casa + almacén + utilidades CIUDAD + limpieza + TV Telecentro + matrices |
| **OHIGGINS** | Propiedad (no hogar) · O’Higgins 2219/2225 piso 8 | `OHIGGINS/Gastos_OHIGGINS_2026.xlsx` | Expensas, Personal TV, Edenor Chagas, Metrogas INVALMAR, ABL grande |
| **BONORINO** | Propiedad (no hogar) · E Bonorino 82 | `BONORINO/Gastos_BONORINO_2026.xlsx` | Expensas Bonorino, Edesur |
| **AVA** | Fideicomiso Ava Palpa | `AVA/Gastos_AVA_2026.xlsx` | Cuotas fideicomiso + análisis IPC+CAC |
| **SOL** | Maestro bolsillo Sol | `SOL/Gastos_SOL_2026.xlsx` | Consolida GPS + Almacén Sol + percepciones + cashflows + MP No Gasto |

## Hojas clave — CIUDAD

| Hoja | Función | ¿Suma a total casa `2026`? |
|------|---------|----------------------------|
| `2026` | Matriz mensual por categoría | **Sí** (total casa) |
| `Detalle` | Detalle operativo casa (fecha, medio, cliente, medidor) | Alimenta matriz |
| `Almacen Detalle` | Consumo hogareño Sol (súper, comida, Ramón, combustible, entretenimiento, Farmacity, Directv…) | Sí vía fila ALMACEN |
| `Gasto Personal Sol` | Bolsillo personal Sol | **No** |
| `MP No Gasto` | Retiros, reservas, pago deuda/resumen, inversión | **No** |
| `Dinero Retirado` / `Inversión` / `Pago de Tarjetas` | Espejos de MP No Gasto | **No** |
| `Transferencias recibidas` | Ingresos / transferencias recibidas | **No** (ingreso) |
| `Expensas Consorcios` | Tracker expensas | Matriz propiedad / CIUDAD |
| `Cashflow Cuotas BBVA` | Cuotas BBVA PAGADO/PENDIENTE | Compromiso, no gasto nuevo duplicado |
| `Cashflow Compromisos` | BBVA + AVA unidos | Visión de caja |

## Hojas clave — SOL (maestro)

Regenerar con `SOL/build_gastos_sol.py` después de actualizar fuentes CIUDAD/AVA.

- `Resumen`, `Movimientos Sol`, `Gastos Personales`, `Almacén Sol`
- `Cashflow Cuotas BBVA`, `Cashflow Compromisos`, `MP No Gasto`
- `Categorías`, `Percepciones USD`, `Control y fuentes`

## Drive canónico (desde 2026-10-23)

Carpeta raíz: https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29  
(`03_Finanzas_y_Gastos` — **no** la subcarpeta `Gastos/`)

| Local | Nombre en Drive | file id |
|-------|-----------------|---------|
| `CIUDAD/Gastos_CIUDAD_1132_2026.xlsx` | `Gastos_CIUDAD_1132_2026.xlsx` | `1xnT1W5oObGF8VLho511-P_KxIBev70c7` |
| `SOL/Gastos_SOL_2026.xlsx` | `Gastos_Personales_SOL_2026.xlsx` | `1V2VcaieVaKNIKOa5I9LXHxWrgaHIUWy_` |
| `AVA/Gastos_AVA_2026.xlsx` | `Gastos_AVA_2026.xlsx` | `1ArzZnX0C7w1Ff5F5OXhLfhd-XcSCE844` |
| `OHIGGINS/Gastos_OHIGGINS_2026.xlsx` | `Gastos_OHIGGINS_2026.xlsx` | `1qkYD_3Zvxq3vGLJnBCm1B2RtFFpgpK_c` |
| `BONORINO/Gastos_BONORINO_2026.xlsx` | `Gastos_BONORINO_2026.xlsx` | `1dqtXPaEgn5e34BcHzaVKL8lA-qhqzjlA` |

Sheets nativos en la misma carpeta: `GASTOS2026`, `Gastos Personales` (legacy).

**Regla dura:** ninguna tarea de gastos queda completa si solo cambió git/local. Overwrite en Drive + verificar.
