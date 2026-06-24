import JSZip from "jszip";
import { readFileSync } from "node:fs";
import { formatNumber, roundMoney } from "../domain/format.js";
import { getConfirmationLines } from "../services/confirmation-service.js";

const TEMPLATE_URL = new URL("../templates/confirmacion-pedido.docx", import.meta.url);
const VAT_RATE = 0.21;
const TABLE_WIDTHS = [650, 2350, 950, 1200, 1100, 1100, 1300, 1200];

export async function renderConfirmationDocx(confirmation, { mode }) {
  const zip = await JSZip.loadAsync(readFileSync(TEMPLATE_URL));
  const replacements = buildReplacements(confirmation, { mode });
  const merchandiseTable = buildMerchandiseTableXml(confirmation, mode);

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
        if (name === "word/document.xml") {
          xml = insertTableAfterMerchandise(xml, merchandiseTable);
          const showsStorageLine = confirmation.hasSheetMaterial && mode === "formato3";
          xml = updateStorageLine(xml, showsStorageLine);
        }
        zip.file(name, xml);
      })
  );

  return zip.generateAsync({ type: "nodebuffer" });
}

function buildReplacements(confirmation, { mode }) {
  const customer = confirmation.customer;
  const origin = confirmation.origin || "Según contrato de compra";
  const packing = confirmation.hasSheetMaterial
    ? "PACKING: Según condiciones del pedido"
    : "PACKING: Standard export packing";
  const showsBankDetails = isTransferPaymentTerm(confirmation.paymentTerms);
  const showsStorageLine = confirmation.hasSheetMaterial && mode === "formato3";

  return [
    ["CLIENTE XXXXX", customer.fiscalName || customer.commercialName || ""],
    ["DIRECCIÓN CLIENTE XXXX", customerAddress(customer)],
    ["CIF CLLIENTE XXX", customer.taxId || ""],
    ["CONFIRMACIÓN DE PEDIDO: STA – 2026-XXXX", `CONFIRMACIÓN DE PEDIDO: ${confirmation.contractNumber}`],
    ["CONFIRMACIÓN DE PEDIDO: STA - 2026-XXXX", `CONFIRMACIÓN DE PEDIDO: ${confirmation.contractNumber}`],
    ["MERCANCIA :", "MERCANCÍA"],
    ["MERCANCIA:", "MERCANCÍA"],
    ["MARCANCIA :", "MERCANCÍA"],
    ["MARCANCIA:", "MERCANCÍA"],
    ["ORIGEN: FABRICA Y PAÍS", `ORIGEN: ${origin}`],
    [
      "CANTIDAD TOTAL: 600,000 MT (+ / - 10%)",
      `CANTIDAD TOTAL: ${formatNumber(confirmation.totalQuantity, 3)} MT ${formatTolerance(confirmation)}`
    ],
    [
      "CONDICIONES DE ENTREGA: INCOTERM DE LA VENTA",
      `CONDICIONES DE ENTREGA: ${confirmation.deliveryTerms || ""}`
    ],
    ["PESO BOBINA: RANGO DEL ITEM DE COMPRA", ""],
    [
      "PACKING:  Standard export packing  ( PARA TODO MENOS PARA CHAPA Y  CALIENTE",
      packing
    ],
    [
      "CONDICIONES DE PAGO :  LA FORMA DE PAGO DEL PEDIDO",
      `CONDICIONES DE PAGO: ${confirmation.paymentTerms || ""}`
    ],
    ["CUANDO EL PAGO ES POR TRANSFERENCIA", showsBankDetails ? "CUANDO EL PAGO ES POR TRANSFERENCIA" : ""],
    [
      "CAIXA BANK - ES40 2100 6428 2213 0012 3884",
      showsBankDetails ? "CAIXA BANK - ES40 2100 6428 2213 0012 3884" : ""
    ],
    ["ALMACENAJES:", showsStorageLine ? "ALMACENAJES:" : ""],
    ["TECHOS FALSTECH", customer.fiscalName || customer.commercialName || ""]
  ];
}

function isTransferPaymentTerm(paymentTerms) {
  const normalized = String(paymentTerms || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
  return (
    normalized === "transferencia" ||
    normalized.includes("transferencia")
  );
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

function formatCoilWeightRange(minNet, maxNet) {
  if (minNet === undefined && maxNet === undefined) return "";
  if (minNet === undefined) return `${formatNumber(maxNet, 3)} MT`;
  if (maxNet === undefined || maxNet === minNet) {
    return `${formatNumber(minNet, 3)} MT`;
  }
  return `${formatNumber(minNet, 3)} - ${formatNumber(maxNet, 3)} MT`;
}

function formatDateEs(value) {
  if (!value) return "";
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : String(value);
}

function buildMerchandiseTableXml(confirmation, mode) {
  const lines = getConfirmationLines(confirmation, mode);
  const subtotal = sum(lines, "amount");
  const vat = roundMoney(subtotal * VAT_RATE);
  const total = roundMoney(subtotal + vat);
  const rows = [
    buildHeaderRow(mode),
    ...lines.map((line, index) => buildLineRow(line, index + 1, mode)),
    [
      cell("", { span: 5 }),
      cell(formatNumber(sum(lines, "quantity"), 3), { align: "right", bold: true }),
      cell(""),
      cell(formatMoney(subtotal), { align: "right", bold: true })
    ],
    [
      cell("IVA 21%", { span: 7, align: "right", bold: true }),
      cell(formatMoney(vat), { align: "right", bold: true })
    ],
    [cell(formatMoney(total), { span: 8, align: "right", bold: true, shade: "EDEDED" })]
  ];

  return [
    "<w:tbl>",
    "<w:tblPr>",
    '<w:tblW w:w="5000" w:type="pct"/>',
    '<w:tblLayout w:type="fixed"/>',
    '<w:tblCellMar><w:top w:w="45" w:type="dxa"/><w:left w:w="45" w:type="dxa"/><w:bottom w:w="45" w:type="dxa"/><w:right w:w="45" w:type="dxa"/></w:tblCellMar>',
    "</w:tblPr>",
    `<w:tblGrid>${TABLE_WIDTHS.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>`,
    rows.map(rowXml).join(""),
    "</w:tbl>"
  ].join("");
}

function buildHeaderRow(mode) {
  if (mode === "detail") {
    return [
      cell("ITEM", { bold: true, align: "center", shade: "EDEDED" }),
      cell("ESPECIFICACIÓN", { bold: true, align: "center", shade: "EDEDED" }),
      cell("NÚMERO DE BOBINA", { span: 2, bold: true, align: "center", shade: "EDEDED" }),
      cell("PESO BOBINA (MT)", { bold: true, align: "center", shade: "EDEDED" }),
      cell("CANTIDAD (MT)", { bold: true, align: "center", shade: "EDEDED" }),
      cell("PRECIO (EUR/MT)", { bold: true, align: "center", shade: "EDEDED" }),
      cell("TOTAL EUR", { bold: true, align: "center", shade: "EDEDED" })
    ];
  }
  if (mode === "formato3") {
    return [
      cell("ITEM", { bold: true, align: "center", shade: "EDEDED" }),
      cell("ESPECIFICACIÓN", { span: 2, bold: true, align: "center", shade: "EDEDED" }),
      cell("UNIDADES", { bold: true, align: "center", shade: "EDEDED" }),
      cell("PESO BOBINA (MT)", { bold: true, align: "center", shade: "EDEDED" }),
      cell("CANTIDAD (MT)", { bold: true, align: "center", shade: "EDEDED" }),
      cell("PRECIO (EUR/MT)", { bold: true, align: "center", shade: "EDEDED" }),
      cell("TOTAL EUR", { bold: true, align: "center", shade: "EDEDED" })
    ];
  }
  return [
    cell("ITEM", { bold: true, align: "center", shade: "EDEDED" }),
    cell("ESPECIFICACIÓN", { span: 3, bold: true, align: "center", shade: "EDEDED" }),
    cell("PESO BOBINA (MT)", { bold: true, align: "center", shade: "EDEDED" }),
    cell("CANTIDAD (MT)", { bold: true, align: "center", shade: "EDEDED" }),
    cell("PRECIO (EUR/MT)", { bold: true, align: "center", shade: "EDEDED" }),
    cell("TOTAL EUR", { bold: true, align: "center", shade: "EDEDED" })
  ];
}

function buildLineRow(line, index, mode) {
  const itemNumber = line.itemNumber || index;
  if (mode === "detail") {
    return [
      cell(itemNumber, { align: "center" }),
      cell(line.specification),
      cell(line.factoryId || "", { span: 2 }),
      cell(formatCoilWeightRange(line.minNet, line.maxNet), { align: "right" }),
      cell(formatNumber(line.quantity, 3), { align: "right" }),
      cell(formatMoney(line.price), { align: "right" }),
      cell(formatMoney(line.amount), { align: "right" })
    ];
  }
  if (mode === "formato3") {
    return [
      cell(itemNumber, { align: "center" }),
      cell(line.specification, { span: 2 }),
      cell(formatUnits(line.units), { align: "right" }),
      cell(formatCoilWeightRange(line.minNet, line.maxNet), { align: "right" }),
      cell(formatNumber(line.quantity, 3), { align: "right" }),
      cell(formatMoney(line.price), { align: "right" }),
      cell(formatMoney(line.amount), { align: "right" })
    ];
  }
  return [
    cell(itemNumber, { align: "center" }),
    cell(line.specification, { span: 3 }),
    cell(formatCoilWeightRange(line.minNet, line.maxNet), { align: "right" }),
    cell(formatNumber(line.quantity, 3), { align: "right" }),
    cell(formatMoney(line.price), { align: "right" }),
    cell(formatMoney(line.amount), { align: "right" })
  ];
}

function cell(text, options = {}) {
  return { text: text ?? "", span: 1, align: "left", bold: false, shade: "", ...options };
}

function rowXml(row) {
  let columnIndex = 0;
  const cells = row
    .map((tableCell) => {
      const xml = cellXml(tableCell, columnIndex);
      columnIndex += tableCell.span;
      return xml;
    })
    .join("");
  return `<w:tr>${cells}</w:tr>`;
}

function cellXml({ text, span, align, bold, shade }, columnIndex) {
  const width = TABLE_WIDTHS.slice(columnIndex, columnIndex + span).reduce(
    (total, value) => total + value,
    0
  );
  const props = [
    `<w:tcW w:w="${width}" w:type="dxa"/>`,
    span > 1 ? `<w:gridSpan w:val="${span}"/>` : "",
    '<w:vAlign w:val="center"/>',
    shade ? `<w:shd w:fill="${shade}"/>` : "",
    '<w:tcBorders><w:top w:val="single" w:sz="4" w:color="808080"/><w:left w:val="single" w:sz="4" w:color="808080"/><w:bottom w:val="single" w:sz="4" w:color="808080"/><w:right w:val="single" w:sz="4" w:color="808080"/></w:tcBorders>'
  ].join("");
  const runProps = [
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>',
    '<w:sz w:val="16"/>',
    bold ? "<w:b/>" : ""
  ].join("");

  return [
    "<w:tc>",
    `<w:tcPr>${props}</w:tcPr>`,
    "<w:p>",
    `<w:pPr><w:jc w:val="${align}"/></w:pPr>`,
    `<w:r><w:rPr>${runProps}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`,
    "</w:p>",
    "</w:tc>"
  ].join("");
}

function insertTableAfterMerchandise(xml, tableXml) {
  let inserted = false;
  const result = xml.replace(/<w:p[\s\S]*?<\/w:p>/g, (paragraph) => {
    if (inserted) return paragraph;
    const text = normalizeForMatch(paragraphText(paragraph));
    if (!text.includes("mercancia") && !text.includes("marcancia")) return paragraph;
    inserted = true;
    return `${paragraph}${tableXml}`;
  });
  return inserted ? result : xml;
}

function updateStorageLine(xml, showStorageLine) {
  return xml.replace(/<w:p[\s\S]*?<\/w:p>/g, (paragraph) => {
    if (!normalizeForMatch(paragraphText(paragraph)).includes("almacenajes")) return paragraph;
    if (!showStorageLine) {
      return "";
    }

    return paragraph.replace(/<w:color w:val="EE0000"\/>/g, '<w:color w:val="000000"/>');
  });
}

function paragraphText(paragraph) {
  return [...paragraph.matchAll(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((run) => unescapeXml(run[2]))
    .join("");
}

function normalizeForMatch(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function sum(lines, key) {
  return roundMoney(lines.reduce((total, line) => total + (Number(line[key]) || 0), 0));
}

function formatMoney(value) {
  return new Intl.NumberFormat("es-ES", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(value || 0);
}

function formatUnits(value) {
  if (value === undefined || value === null || value === "") return "";
  return new Intl.NumberFormat("es-ES", {
    maximumFractionDigits: 0
  }).format(value);
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
