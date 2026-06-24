# Crear confirmación

- Base: `appnS0geRMpHaxib9`
- Tabla: `Contratos de venta`
- Botón: `Crear confirmación`
- Campo destino: `Documentos confirmación`

Genera un único adjunto desde el servicio `steel-trade-invoice-service`:

- `.docx` desde `src/templates/confirmacion-pedido.docx`
- tabla de mercancía insertada en el Word usando el formato de `src/templates/formatos-confirmacion.xlsx`

Regla de formato:

- Si algún ítem vinculado tiene `Ítems de compra.Tipo de material` igual a `Chapa` o `Chapa grande`, usa `formato3`.
- Si no, pregunta al usuario entre `Agrupado` y `Detalle`.

El script sustituye el contenido actual de `Documentos confirmación` por el Word generado.

Antes de pegarlo en Airtable, cambiar `SERVICE_BASE_URL` por la URL pública real
del servicio desplegado en Render. Airtable no puede adjuntar archivos desde
`localhost`.
