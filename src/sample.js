import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { buildSampleInvoice } from "./sample-data.js";
import { renderPackingExcel } from "./render/excel.js";
import { renderInvoicePdf } from "./render/pdf.js";
import { renderInvoiceResponse } from "./render/bundle.js";

const outputDir = path.resolve("output");
await mkdir(outputDir, { recursive: true });

const invoice = buildSampleInvoice();
await writeFile(
  path.join(outputDir, "sample-grouped.pdf"),
  await renderInvoicePdf(invoice, { mode: "grouped" })
);
await writeFile(
  path.join(outputDir, "sample-detail.pdf"),
  await renderInvoicePdf(invoice, { mode: "detail" })
);
await writeFile(path.join(outputDir, "sample-grouped.xlsx"), await renderPackingExcel(invoice));
const zip = await renderInvoiceResponse(invoice, { mode: "grouped", format: "zip" });
await writeFile(path.join(outputDir, zip.filename), zip.buffer);

console.log(`Wrote sample files to ${outputDir}`);

