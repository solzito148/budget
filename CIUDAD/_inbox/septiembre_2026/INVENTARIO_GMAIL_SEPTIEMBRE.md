# Inventario Gmail — rendición septiembre 2026

Fuente: Gmail `s.sanes@gmail.com` (Composio) · corte 2026-10-04.

## Estado del Excel antes de esta pasada
- Matriz CIUDAD `2026` septiembre: solo Telecentro placeholder $50.000
- Almacén Detalle septiembre: 0 filas con fecha
- GPS septiembre: solo cuotas BBVA **PROYECTADO**
- Catálogo MP / settlements locales: cortan en **julio 2026**
- Agosto también casi vacío (solo Telecentro placeholder)

## Bloqueante para cerrar Almacén + GPS completo
Faltan resúmenes oficiales que **no llegan como adjunto** al correo:
1. **Settlement / export cuenta Mercado Pago** (ago + sept 2026) — CSV como los de `_inbox/mp_exports/`
2. **Resúmenes PDF BBVA Visa + Mastercard** del cierre de sept (mails del 4-oct solo avisan “ya podés descargar”, sin PDF)
3. **Resumen TC Mercado Pago** septiembre

Sin esos tres, no se puede cargar el grueso de súper/comida/Ramón/transporte ni cruzar neto de cancelaciones.

## Confirmado desde Gmail (oficial)

### Expensas — período vs pago
| Propiedad | Período | Monto (ceil) | Pago | Evidencia |
|-----------|---------|--------------|------|-----------|
| CIUDAD 1132 | Agosto 2026 | 300850 | Pagado ~01–10/09 (Gmail Sent 10/09) | cupón/aviso De Filippo |
| CIUDAD 1132 | Septiembre 2026 | 328200 | Pendiente (vto 10/10) | mail 02/10 |
| OHIGGINS UF8 | Agosto 2026 | 722979 | Pagado 10/09 | recibo Kalmus $722.978,08 |
| OHIGGINS UF8 | Septiembre 2026 | 722979 | Pendiente (liq. 01/10) | liquidación sep26 |
| BONORINO UF5 | Agosto 2026 | 222825 | Pagado 01–07/09 | liq 09/2026 vto 07/09 $222.824,05 |
| BONORINO UF5 | Septiembre 2026 | 226137 | Pendiente (vto 06/10) | liq 10/2026 $226.136,05 |

### Utilidades (mes = fecha de cobro BBVA cuando aplica)
| Destino | Concepto | Monto ceil | Fecha cobro/venc | Evidencia |
|---------|----------|------------|------------------|-----------|
| CIUDAD | Edenor LUZ 8917502309 | 63980 | 11/09 BBVA | mail + Compra aprobada |
| CIUDAD | Metrogas 30010996459 | 105839 | 15/09 BBVA | mail + Compra aprobada |
| CIUDAD | Telecentro | 69659 | 25/09 BBVA | factura 10/2026 + débito; **reemplaza placeholder 50000** |
| OHIGGINS | Edenor Chagas 9267222037 | 18236 | 09/09 BBVA | mail + Compra aprobada |
| OHIGGINS | Metrogas INVALMAR 10139236800 | 5589 | factura 01/10 (monto en mail) | mail Metrogas — mes cobro a confirmar |
| OHIGGINS | Personal TV | 76044 | vto 01/10 | mail 10/09 — mes = cobro (pendiente si no debitó en sept) |
| BONORINO | Edesur 3932799 | 2009 | vto 30/09 | mail Edesur |

### AVA Fideicomiso
- Factura/recibo 07/09/2026 · **$6.269.440** · vto cuota 04/09/2026
- TC MEP venta 04/09/2026: **1525.39** (Rava cierre / bolsa)
- Valor USD = 6269440 / 1525.39 ≈ **4110.04**
- 50% Sol = **3.134.720** → GPS + AVA
- Cuota siguiente a la 9 (julio): cargar como **cuota 10** con mes **septiembre 2026** (vto sept). Revisar si faltó cuota agosto.

### ABL AGIP
- Mail 09/09: cuota 9 vto 14/09 — **sin monto ni partida** en el mail → mapear con settlement MP / portal AGIP

### Alertas BBVA Visa (parcial; no reemplazan resumen)
Incluye: ChatGPT USD20, Spotify 5999, Cursor USD60, Google One USD4.99, Telecentro, Metrogas, Edenor x2, SUPERDIA, ITSCLASSIC, Markova, LunaHogar, Luxodia, dhgate, etc.
No hay alertas Mastercard en el período buscado.

### Alertas MP “Pago aprobado” (parcial; no reemplazan settlement)
ViennaHogar, Yenny, Promarine, EducaciónIT, ITSCLASSIC, Sodastream, Pescadería Belgrano, Correo, Copetin Catering.

## PDFs bajados
Carpeta: `CIUDAD/_inbox/septiembre_2026/gmail_pdfs/`

## Cargado en esta pasada (desde Gmail)
- CIUDAD: Edenor, Metrogas, Telecentro (cobro), Expensas ago+sep, Spotify, Almacén parcial (Día/Sodastream/Pescadería), GPS (AVA 50%, Promarine, Itsclassic, EducaciónIT, Copetin)
- OHIGGINS: Expensas ago+sep, Edenor Chagas; Detalle Personal/Metrogas pendientes de cobro
- BONORINO: Expensas ago+sep, Edesur
- AVA: cuota 10 $6.269.440 · TC MEP 1525.39 · 50% Sol en GPS
- Maestro SOL regenerado

## Pendiente del usuario
1. Export settlement MP ago+sept (CSV)
2. PDFs resumen BBVA Visa + Mastercard cierre sept
3. Resumen TC Mercado Pago septiembre
4. Confirmar ABL cuota 9 (partida/monto) y Limpieza Susana septiembre
5. Confirmar si hubo cuota AVA de agosto aparte de la de vto 04/09
