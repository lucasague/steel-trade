import JSZip from "jszip";
import { readFileSync } from "node:fs";
import { formatNumber } from "../domain/format.js";
import { getConfirmationLines } from "../services/confirmation-service.js";

const TEMPLATE_URL = new URL("../templates/confirmacion-pedido.docx", import.meta.url);

export async function renderConfirmationDocx(confirmation, { mode }) {
  const zip = await JSZip.loadAsync(readFileSync(TEMPLATE_URL));
  const replacements = buildReplacements(confirmation, mode);

  await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^word\/.*\.xml$/.test(name))
      .map(async (name) => {
        const file = zip.file(name);
        if (!file) return;
        let xml = await file.async("string");
        xml = replaceTextInParagraphs(xml, (text) =>
          text.trim() === "FECHA" ? formatDateEs(confirmation.date) : text
        );
        for (const [needle, replacement] of replacements) {
          xml = replaceTextInParagraphs(xml, (text) =>
            text.includes(needle) ? text.split(needle).join(replacement || "") : text
          );
        }
        zip.file(name, xml);
      })
  );

  return zip.generateAsync({ type: "nodebuffer" });
}

function buildReplacements(confirmation, mode) {
  const customer = confirmation.customer;
  const lines = getConfirmationLines(confirmation, mode);
  const summary = summarizeLines(lines);
  const packing = confirmation.hasSheetMaterial
    ? "PACKING: Según condiciones del pedido"
    : "PACKING: Standard export packing";

  return [
    ["CLIENTE XXXXX", customer.fiscalName || customer.commercialName || ""],
    ["DIRECCIÓN CLIENTE XXXX", customerAddress(customer)],
    ["CIF CLLIENTE XXX", customer.taxId || ""],
    ["CONFIRMACIÓN DE PEDIDO: STA – 2026-XXXX", `CONFIRMACIÓN DE PEDIDO: ${confirmation.contractNumber}`],
    ["CONFIRMACIÓN DE PEDIDO: STA - 2026-XXXX", `CONFIRMACIÓN DE PEDIDO: ${confirmation.contractNumber}`],
    ["MERCANCIA :", `MERCANCÍA: ${summary}`],
    ["ORIGEN: FABRICA Y PAÍS", "ORIGEN: Según contrato de compra"],
    [
      "CANTIDAD TOTAL: 600,000 MT (+ / - 10%)",
      `CANTIDAD TOTAL: ${formatNumber(confirmation.totalQuantity, 3)} MT ${formatTolerance(confirmation)}`
    ],
    [
      "CONDICIONES DE ENTREGA: INCOTERM DE LA VENTA",
      `CONDICIONES DE ENTREGA: ${confirmation.deliveryTerms || ""}`
    ],
    [
      "PESO BOBINA: RANGO DEL ITEM DE COMPRA",
      confirmation.hasSheetMaterial ? "" : `PESO BOBINA: ${coilWeightRanges(confirmation)}`
    ],
    [
      "PACKING:  Standard export packing  ( PARA TODO MENOS PARA CHAPA Y  CALIENTE",
      packing
    ],
    [
      "CONDICIONES DE PAGO :  LA FORMA DE PAGO DEL PEDIDO",
      `CONDICIONES DE PAGO: ${confirmation.paymentTerms || ""}`
    ],
    ["TECHOS FALSTECH", customer.fiscalName || customer.commercialName || ""]
  ];
}

function summarizeLines(lines) {
  return lines
    .map((line) => line.specification)
    .filter(Boolean)
    .filter((value, index, all) => all.indexOf(value) === index)
    .join("; ");
}

function customerAddress(customer) {
  return [
    customer.address,
    [customer.postalCode, customer.city].filter(Boolean).join(" "),
    customer.province,
    customer.country
  ]
    .filter(Boolean)
    .join(", ");
}

function formatTolerance(confirmation) {
  const minus = toPercent(confirmation.toleranceMinus);
  const plus = toPercent(confirmation.tolerancePlus);
  if (!minus && !plus) return "";
  if (minus && plus && minus === plus) return `(+ / - ${plus})`;
  return `(- ${minus || "0%"} / + ${plus || "0%"})`;
}

function toPercent(value) {
  if (value === undefined || value === null) return "";
  return `${Math.round(value * 100)}%`;
}

function coilWeightRanges(confirmation) {
  const ranges = confirmation.items
    .filter((item) => item.minNet || item.maxNet)
    .map((item) => {
      if (item.minNet && item.maxNet) {
        return `${formatNumber(item.minNet, 3)}-${formatNumber(item.maxNet, 3)} MT`;
      }
      return `${formatNumber(item.minNet || item.maxNet, 3)} MT`;
    });
  return ranges.length ? [...new Set(ranges)].join(" / ") : "Según ítem de compra";
}

function formatDateEs(value) {
  if (!value) return "";
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : String(value);
}

function replaceTextInParagraphs(xml, replaceText) {
  return xml.replace(/<w:p[\s\S]*?<\/w:p>/g, (paragraph) => {
    const runs = [...paragraph.matchAll(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g)];
    if (!runs.length) return paragraph;

    const text = runs.map((run) => unescapeXml(run[2])).join("");
    const replaced = replaceText(text);
    if (replaced === text) return paragraph;

    let wroteFirstRun = false;
    return paragraph.replace(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_match, attrs = "") => {
      if (!wroteFirstRun) {
        wroteFirstRun = true;
        return `<w:t${attrs}>${escapeXml(replaced)}</w:t>`;
      }
      return `<w:t${attrs}></w:t>`;
    });
  });
}

function escapeXml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function unescapeXml(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}
