const SERVICE_BASE_URL = "https://steel-trade-invoice-service.onrender.com";

const tablaContratos = base.getTable("Contratos de venta");
const tablaItemsVenta = base.getTable("Ítems de venta");
const tablaItemsCompra = base.getTable("Ítems de compra");

if (SERVICE_BASE_URL.includes("TU-SERVICIO-RENDER")) {
  output.text("Falta configurar SERVICE_BASE_URL con la URL pública del servicio en Render.");
  return;
}

const contrato = await input.recordAsync("Selecciona el contrato de venta", tablaContratos);

if (!contrato) {
  output.text("Operación cancelada.");
  return;
}

const items = contrato.getCellValue("Ítems de venta") || [];
if (!items.length) {
  output.text("Este contrato no tiene ítems de venta.");
  return;
}

const tieneChapa = await contratoTieneChapa(items);
let modo = "formato3";

if (!tieneChapa) {
  modo = await input.buttonsAsync("Selecciona el formato de confirmación:", [
    { label: "Agrupado", value: "grouped" },
    { label: "Detalle", value: "detail" }
  ]);
}

const marca = Date.now();
const baseUrl = `${SERVICE_BASE_URL.replace(/\/$/, "")}/api/confirmations/${contrato.id}`;
const nombreBase = safeName(contrato.getCellValueAsString("Contrato") || contrato.id);

await tablaContratos.updateRecordAsync(contrato.id, {
  "Documentos confirmación": [
    {
      url: `${baseUrl}/${nombreBase}-confirmacion.docx?mode=${modo}&v=${marca}`,
      filename: `${nombreBase}-confirmacion.docx`
    }
  ]
});

output.text(
  `Confirmación Word creada en formato ${modo === "formato3" ? "3" : modo === "grouped" ? "agrupado" : "detalle"}.`
);

async function contratoTieneChapa(itemsVenta) {
  for (const item of itemsVenta) {
    const itemVenta = await tablaItemsVenta.selectRecordAsync(item.id, {
      fields: ["Ítem de compra"]
    });
    const compras = itemVenta?.getCellValue("Ítem de compra") || [];

    for (const compra of compras) {
      const itemCompra = await tablaItemsCompra.selectRecordAsync(compra.id, {
        fields: ["Tipo de material"]
      });
      const tipos = itemCompra?.getCellValue("Tipo de material") || [];
      if (tipos.some((tipo) => esChapa(tipo.name))) {
        return true;
      }
    }
  }
  return false;
}

function esChapa(valor) {
  const normalizado = String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
  return normalizado === "chapa" || normalizado === "chapa grande";
}

function safeName(value) {
  return String(value || "confirmacion")
    .replace(/[^\w.-]+/g, "-")
    .replace(/^-+|-+$/g, "");
}
