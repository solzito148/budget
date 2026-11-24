# 10 — Bitácora de indicaciones de Sol

Todo lo que Sol indica se anota acá con fecha y se pasa a las reglas correspondientes. **Leer siempre antes de trabajar.**

## 2026-11-24

- **Regla madre:** una planilla por propiedad (CIUDAD, OHIGGINS, BONORINO, AVA) + una de gastos personales de Sol. Cada gasto es de una propiedad **o** personal de Sol.
- **CIUDAD = casa compartida Sol + Chris + Ramón.** Servicios + Almacén Sol + Almacén Chris se suman por mes y se dividen 50/50. Sol paga casi todos los servicios; Chris rinde a fin de mes.
- **Gastos de Chris:** Chris le pasa sus compras a Sol y **Sol las informa por chat** → todas a `ALMACEN CHRIS` (CIUDAD · `Almacen Detalle` col. Importe Chris). Solo **mes + detalle + importe**, **sin fecha**.
- **Liquidación Chris:** ene–abr se calcula pero no se cobra; AVA no entra (Chris lo paga aparte); OHIGGINS y BONORINO = 100% Sol.
- **Gastos personales de Sol** salen de resúmenes (MP, TC MP, BBVA) y de compras en los mails.
- **Ante la duda** (¿personal o propiedad? ¿qué propiedad?) → **preguntar a Sol por pantalla** con tabla; no cargar por descarte.
- **Servicios: importe todos los meses.** Mes vacío = faltante a resolver.
- **Expensas:** se pagan por **transferencia desde Mercado Pago**; el respaldo es el **mail enviado** al administrador con el comprobante adjunto (De Filippo = CIUDAD, Kalmus = OHIGGINS, pereyragladys = BONORINO).
- **Metrogas de OHIGGINS** = la factura a nombre de **INVALMAR / INVILMAR** (cliente 10139236800).
- **Edesur = luz de BONORINO** (cliente 3932799). BONORINO también: ABL partida 1084069 y Metrogas 30010713734.
- **Sol no sube archivos a mano.** Claude guarda todo en Drive. Como el conector no puede sobrescribir los xlsx viejos de la raíz, las planillas vigentes están en la carpeta **`Budget_2026_vigente`** (ids en `CIUDAD/_inbox/drive_excel_canonico.json`).
- **Nunca pisar el Drive con una versión vieja:** antes de editar, bajar lo último del Drive y combinar (el 24/11 a las 00:25 se perdieron datos de OHIGGINS/BONORINO por eso; se reconstruyeron desde Rendiciones/Control + Gmail).
- **Consolidado por propiedad / mes / tipo de gasto:** `python3 tools/consolidar_2026.py` → `CONSOLIDADO/Consolidado_2026.xlsx` (hojas Consolidado, Faltantes, Liquidación Chris).
- **`GASTOS2026` (Google Sheet) es la planilla de CIUDAD hecha por Sol**; su solapa **`Detalle 2026`** es la que dice los gastos de **almacén** (columnas Importe Sol / Importe Chris).
- **Chris octubre 2026 = $335.669 en 26 gastos** (confirmado por Sol, coincide con `Almacen Detalle`). Los gastos de Chris rondan ~$360.000/mes: el $1.201.880 NO son gastos de Chris sino el saldo calculado de la liquidación → revisar qué entra como almacén compartido (pendiente de definición con Sol).
- **Almacén Chris 2026 confirmado (ramas GitHub + Rendiciones):** may $176.932 (80.000 + 96.932; el 80.000 primero se cargó en agosto y Sol lo corrigió a mayo) · jun $101.000 · jul $155.800 · ago **$0 (confirmado por Sol 24/11)** · sep $237.104 (24 ítems) · oct $335.669 (26 ítems). La columna Chris de `GASTOS2026 › Detalle 2026` NO coincide (oct $403.110 vs $335.669 confirmado por Sol) → no usarla como fuente de Chris sin confirmar.
- Rendición histórica: `CIUDAD/Rendicion_CIUDAD_2026.xlsx` + `build_rendicion_ciudad.py` (rama, 11/09) quedaron reemplazados por `Rendiciones_2026.xlsx`.
- **Ciclo mensual (24/11):** nunca borrar meses anteriores; siempre agregar; todos los meses con todos los valores; armar la **rendición de CIUDAD todos los meses**. Procedimiento completo en `11_ciclo_mensual.md`. Control automático: `tools/verificar_no_borrado.py` (vigente vs nuevo) antes de guardar.
- **Luz OHIGGINS (Edenor Chagas):** mes de la matriz = **mes de emisión de la factura** (criterio original de la planilla; no correr meses ya cargados).
- **Almacén Chris 2026 confirmado por Sol (24/11):** ene–abr $0 · may $176.932 · jun $101.000 · jul $155.800 · ago $0 · sep $237.104 (24) · oct $335.669 (26).
