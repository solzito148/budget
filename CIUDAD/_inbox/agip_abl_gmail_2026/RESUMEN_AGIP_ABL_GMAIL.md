# AGIP / ABL — inventario Gmail por propiedad

Corte: 2026-10-12 · buzón `s.sanes@gmail.com`.

## Regla de asignación
| Propiedad | Partida AGIP | Regla monto |
|---|---|---|
| **CIUDAD 1132** | `397789` | ≤ $80.000 |
| **OHIGGINS** | `3016365` | > $80.000 |
| BONORINO / AVA | pendiente | — |

Identificador BBVA: `CIUDADABL 0070039773890` → partida **397789** (CIUDAD).

## Boletas digitales AGIP (mails `Emisión de boleta digital | ABL`)
Solo llegan boletas de **CIUDAD**. El cuerpo no trae partida ni importe (link “DESCARGAR AHORA”).

| Mail | Cuota | Vto | Propiedad |
|---|---:|---|---|
| 03/06/2026 | 6 | 12/06/2026 | CIUDAD |
| 08/08/2026 | 8 | 14/08/2026 | CIUDAD |
| 09/09/2026 | 9 | 14/09/2026 | CIUDAD |
| 09/10/2026 | 10 | 14/10/2026 | CIUDAD |

No hay en este buzón boletas AGIP 2026 de cuotas 1–5 ni 7, ni de partida **3016365** (OHIGGINS).

## Avisos BBVA “Tu servicio está por vencer” (CIUDADABL)
| Mail | Vto | Importe (ceil) | Cruce |
|---|---|---:|---|
| 12/08 | 14/08 | 41.344 | = pago CIUDAD 14/08 |
| 29/08 | 31/08 | 49.478 | 2º vencimiento; sin pago aparte cargado |
| 12/09 | 14/09 | 42.188 | pago MP fue **27.635** (prevalece extracto) |
| 28/09 | 30/09 | 50.487 / 50.721 | dos avisos; sin pago oct/cierre |
| 12/10 | 14/10 | 42.960 | cuota 10 · **pendiente de pago** |

## Pagos MP ya en planillas
### CIUDAD (397789)
| Fecha | Monto | Mes matriz |
|---|---:|---|
| 07/01 | 12.435 | Enero |
| 10/05 ×5 | 27.739+27.439+27.164+24.745+25.512 = **132.599** | Mayo |
| 14/07 | 26.175 | Julio |
| 14/08 | 41.344 | Agosto |
| 14/09 | 27.635 | Septiembre |

### OHIGGINS (3016365) — sin boleta Gmail; por monto >80k
| Fecha | Monto | Mes matriz |
|---|---:|---|
| 13/02 | 83.757 | Febrero |
| 03/03 | 88.141 | Marzo |
| 30/06 | 100.779 | Junio |
| 14/07 | 88.891 | Julio |

### BONORINO / AVA
Sin boletas ni pagos ABL identificados en Gmail/MP 2026.

## Limpieza hecha
- CIUDAD `Detalle` tenía filas feb/mar $83.757 / $88.141 (son OHIGGINS): marcadas SUPERSEDE / no sumar.

## Pendiente / dudas
1. **Octubre CIUDAD** cuota 10 · vto 14/10 · aviso BBVA **$42.960** — aún no hay pago MP; no cargar hasta cobro.
2. Avisos BBVA 31/08 y 30/09 (2º vencimientos) — ¿intereses/mora o boletas distintas?
3. ¿Las boletas OHIGGINS llegan a otro mail (sucesión Chagas / Roberto)?
