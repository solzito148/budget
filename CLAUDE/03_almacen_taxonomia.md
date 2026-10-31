# 03 — ALMACEN SOL · taxonomía

**GitHub:** https://github.com/solzito148/budget/tree/master/CLAUDE · **Drive Excels:** https://drive.google.com/drive/folders/1fc6P038_VTOXtH_GzZRti1rv_UD14w29 · **Drive docs:** https://drive.google.com/drive/folders/1Lfu_3vQjZQ3L3tiBI5eAk661bnyC1UnW

Hoja: `Almacen Detalle` en CIUDAD. Columna Importe Sol alimenta el maestro SOL.

## Qué entra (y por qué)

Es el **consumo del hogar** que Sol paga y se analiza mes a mes: comida, súper, Ramón, nafta, salidas hogareñas, Farmacity, Directv, Sodastream.

## SUPERMERCADO

Normalizar siempre a:

| Etiqueta canónica | Comercios / señales |
|-------------------|---------------------|
| `SUPERMERCADO · Coto` | Coto |
| `SUPERMERCADO · Carrefour` | Carrefour |
| `SUPERMERCADO · Día` | Día / SUPERDIA / Día Online / Supermercados Día |
| `PEDIDOS YA · SUPERMERCADO` (+ `· propina`) | PedidosYa Market / `DLO*PEDIDOSYA MARKET` / `PEDIDOSYA*MARKET` (MP o BBVA), incl. propina del mismo pedido |

**Nunca** GPS ni `PEDIDOS YA · COMIDA` para Market.

## COMIDA

Etiqueta: `COMIDA · {comercio}` (o fijas abajo).

Incluye: Todo Tartas, Huevos Point, Open 25, Nemuri Sushi (`COMIDA · SUSHI`), Valenti SA, Carnes Tito, Bee Coffee, Le Pain Quotidien, Tienda De Café, verdulería (`COMIDA · VERDULERÍA` · Noelia Caballero), Súper Pollo (`COMIDA · pollo`), Havanna (`COMIDA · HAVANNA`), Buenos Muchachos, Soraya, Nicky Cheese, Vitalcer, Dean & Dennys, Green Eat, Dos Escudos, Burger King, Chocolates CHIAZZA (Florencia Londero), Enerfit (`COMIDA · Barritas proteicas`), PedidosYa comida (`PEDIDOS YA · COMIDA` + propina).

### Excepciones — NO Almacén COMIDA

| Comercio | Destino real | Categoría |
|----------|--------------|-----------|
| Feed | GPS | `COMIDA · Oficina` |
| Copetin Catering | GPS | `COMIDA · Oficina` |
| Oporto Almacen | GPS | `ENTRETENIMIENTO · Comidas` |
| Kalbby / Compraenkalbb / DR BARF | Almacén | `RAMON · Comida` (no COMIDA genérica) |

## Ramón (perro) — SIEMPRE Almacén, NUNCA GPS

| Etiqueta | Comercios |
|----------|-----------|
| `RAMON · Comida` | Kalbby / Compraenkalbb / DR BARF |
| `RAMON · Veterinario` | Valeria Mele / El Perro Verde / Leocan |
| `RAMON · Otros` | Paws / Paws Pet Market y afines |

Otras mascotas (no Ramón) → GPS.

## Combustible

YPF, Shell, Generacion 20 Combustibles → Almacén detalle tipo `Combustible — Shell/YPF`.  
**No** Extras, **no** AUTO, **no** GPS (salvo ACA → GPS AUTO).

## Entretenimiento / salidas (Almacén)

| Etiqueta | Comercio |
|----------|----------|
| `ENTRETENIMIENTO · Recitales` | Movistar Arena |
| `ENTRETENIMIENTO · Teatro` | Plateanet |
| `ENTRETENIMIENTO · Cine` | Hoyts / MOD*HOYTS |
| `ENTRETENIMIENTO · Cenas` | Cosi Mi Piace |
| `ENTRETENIMIENTO · Social` | Belisario Roldán / Restaurant Roldan |
| `ENTRETENIMIENTO · Salidas comida` | Hotel Madero (BBVA HOTEL MADERO) |

## Streaming / farmacia / repuestos en Almacén

| Etiqueta | Comercio | Nota |
|----------|----------|------|
| `STREAMING TV` | Directv | No Telecentro ni GPS |
| `FARMACIA · Farmacity` | Farmacity | Suma vía ALMACEN; no fila FARMACIA matriz aparte |
| `REPUESTOS · Sodastream` | Sodastream | No fila CASA/REPUESTOS aparte ni GPS |

## PedidosYa Plus

**No** Almacén. → CIUDAD matriz `SUSCRIPCIONES` + `Detalle`.

## Spotify

**No** Almacén ni GPS. → CIUDAD `2026` · `STREAMING` (USD × TC MEP del mes).
