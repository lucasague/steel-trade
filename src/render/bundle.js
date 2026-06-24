import JSZip from "jszip";
import { renderPackingExcel } from "./excel.js";
import { renderInvoicePdf } from "./pdf.js";

export async function renderInvoiceResponse(invoice, { mode, format }) {
  const safeInvoiceNumber = invoice.invoiceNumber.replace(/[^\w.-]+/g, "-");

  if (format === "xlsx") {
    return {
      filename: `${safeInvoiceNumber}.xlsx`,
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: await renderPackingExcel(invoice)
    };
  }

  const pdf = await renderInvoicePdf(invoice, { mode });
  if (format === "zip" || (mode === "grouped" && format !== "pdf")) {
    const zip = new JSZip();
    zip.file(`${safeInvoiceNumber}.pdf`, pdf);
    if (mode === "grouped") {
      zip.file(`${safeInvoiceNumber}.xlsx`, await renderPackingExcel(invoice));
    }
    return {
      filename: `${safeInvoiceNumber}.zip`,
      contentType: "application/zip",
      buffer: await zip.generateAsync({ type: "nodebuffer" })
    };
  }

  return {
    filename: `${safeInvoiceNumber}.pdf`,
    contentType: "application/pdf",
    buffer: pdf
  };
}

