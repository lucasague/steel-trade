import test from "node:test";
import assert from "node:assert/strict";
import JSZip from "jszip";
import { renderConfirmationDocx } from "../src/render/confirmation-docx.js";

test("inserts the merchandise table inside the confirmation Word", async () => {
  const buffer = await renderConfirmationDocx(fakeConfirmation(), {
    mode: "detail"
  });
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const materialIndex = documentXml.indexOf("SIGUIENTE MATERIAL:");
  const tableIndex = documentXml.indexOf("<w:tbl>", materialIndex);

  assert.ok(materialIndex > -1);
  assert.ok(tableIndex > materialIndex);
  assert.match(documentXml, /S235JR/);
  assert.doesNotMatch(documentXml, /MERCANC(?:IA|\u00cdA)\s*:/);
  assert.match(documentXml, /PESO \(MT\)/);
  assert.doesNotMatch(documentXml, /CANTIDAD \(MT\)/);
  assert.doesNotMatch(documentXml, /PESO BOBINA \(MT\)/);
  assert.doesNotMatch(documentXml, /10,000 - 15,000 MT/);
});

test("shows coil weight range only in grouped format", async () => {
  const groupedBuffer = await renderConfirmationDocx(fakeConfirmation(), {
    mode: "formato1"
  });
  const groupedDoc = await JSZip.loadAsync(groupedBuffer);
  const groupedXml = await groupedDoc.file("word/document.xml").async("string");

  assert.match(groupedXml, /RANGO \(MT\)/);
  assert.doesNotMatch(groupedXml, /PESO BOBINA \(MT\)/);
  assert.match(groupedXml, /10,000 - 15,000/);

  const detailBuffer = await renderConfirmationDocx(fakeConfirmation(), {
    mode: "detail"
  });
  const detailDoc = await JSZip.loadAsync(detailBuffer);
  const detailXml = await detailDoc.file("word/document.xml").async("string");

  assert.doesNotMatch(detailXml, /PESO BOBINA \(MT\)/);
  assert.doesNotMatch(detailXml, /10,000 - 15,000 MT/);

  const sheetBuffer = await renderConfirmationDocx(fakeConfirmation(), {
    mode: "formato3"
  });
  const sheetDoc = await JSZip.loadAsync(sheetBuffer);
  const sheetXml = await sheetDoc.file("word/document.xml").async("string");

  assert.doesNotMatch(sheetXml, /PESO BOBINA \(MT\)/);
  assert.doesNotMatch(sheetXml, /10,000 - 15,000 MT/);
});

test("formats quantities and money with thousands separators", async () => {
  const buffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      totalQuantity: 12345.678,
      items: [
        {
          ...fakeConfirmation().items[0],
          quantity: 12345.678,
          price: 12000,
          amount: 148148.14
        }
      ]
    },
    {
      mode: "formato1"
    }
  );
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");

  assert.match(documentXml, /CANTIDAD TOTAL: 12\.345,678 MT/);
  assert.match(documentXml, /12\.000,00/);
  assert.match(documentXml, /148\.148,14/);
});

test("keeps long customer header name in one right-aligned cell", async () => {
  const customerName = "Arcelormital Distribuci\u00f3n, S.L.";
  const buffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      customer: {
        ...fakeConfirmation().customer,
        fiscalName: customerName,
        address: "P.I. Silvota, C/ Pe\u00f1a Ten Parcela 13",
        postalCode: "33690",
        city: "Llanera",
        province: "Asturias",
        country: "Espa\u00f1a",
        taxId: "B83481002"
      }
    },
    {
      mode: "detail"
    }
  );
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const paragraphs = [...documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const logoParagraph = paragraphs.find((paragraph) => paragraph.includes("<w:drawing"));
  const tables = [...documentXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const headerTable = tables.find((table) => extractParagraphText(table).includes(customerName));

  assert.ok(headerTable, "The customer name should be rendered in the header table");
  assert.ok(logoParagraph, "The header should keep the STA logo");
  assert.match(logoParagraph, /<w:ind w:left="-432"\/>/);
  const nameParagraphs = [...headerTable.matchAll(/<w:p[\s\S]*?<\/w:p>/g)]
    .map((match) => match[0])
    .filter((paragraph) => extractParagraphText(paragraph).includes(customerName));
  const headerParagraphs = [...headerTable.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const nameParagraphIndex = headerParagraphs.findIndex((paragraph) =>
    extractParagraphText(paragraph).includes(customerName)
  );

  assert.equal(logoParagraph.includes(customerName), false);
  assert.equal(nameParagraphs.length, 1);
  assert.equal(extractParagraphText(nameParagraphs[0]), customerName);
  assert.ok(nameParagraphIndex > 0, "The customer name should follow the top spacer");
  assert.equal(extractParagraphText(headerParagraphs[nameParagraphIndex - 1]), "");
  assert.match(
    headerParagraphs[nameParagraphIndex - 1],
    /<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"\/>/
  );
  assert.match(headerTable, /<w:noWrap\/>/);
  assert.match(nameParagraphs[0], /<w:jc w:val="right"\/>/);
});

test("replaces merchandise header and origin line", async () => {
  const buffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      contractNumber: "STA-2026-030",
      origin: "Planta Madrid (ES / FR)",
      deliveryTerms: "DDP - Delivered Duty Paid",
      items: [
        {
          ...fakeConfirmation().items[0],
          sheetUnits: 20,
          quantity: 50,
          existences: []
        }
      ]
    },
    {
      mode: "formato1"
    }
  );
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const paragraphs = [...documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const titleIndex = paragraphs.findIndex(
    (paragraph) => extractParagraphText(paragraph) === "CONFIRMACI\u00d3N DE PEDIDO STA-2026-030"
  );
  const materialIntro = paragraphs.find((paragraph) =>
    extractParagraphText(paragraph).includes("SIGUIENTE MATERIAL:")
  );

  assert.ok(titleIndex > 2, "The confirmation title should have three blank lines before it");
  assert.equal(extractParagraphText(paragraphs[titleIndex - 1]), "");
  assert.equal(extractParagraphText(paragraphs[titleIndex - 2]), "");
  assert.equal(extractParagraphText(paragraphs[titleIndex - 3]), "");
  assert.equal(extractParagraphText(paragraphs[titleIndex + 1]), "");
  assert.equal(extractParagraphText(paragraphs[titleIndex + 2]), "");
  assert.match(paragraphs[titleIndex - 1], /<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"\/>/);
  assert.match(paragraphs[titleIndex + 1], /<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"\/>/);
  assert.doesNotMatch(documentXml, /CONFIRMACI\u00d3N DE PEDIDO:/);
  assert.ok(materialIntro, "The introduction paragraph should end with SIGUIENTE MATERIAL:");
  assert.doesNotMatch(documentXml, /MERCANCIA:/);
  assert.doesNotMatch(documentXml, /MERCANC\u00cdA:/);
  assert.match(
    documentXml,
    /SIGUIENTE MATERIAL:[\s\S]*<w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"\/>[\s\S]*<w:tbl>/
  );
  assert.match(documentXml, /ORIGEN: Planta Madrid \(ES \/ FR\)/);
    assert.match(documentXml, /CONDICIONES DE ENTREGA: DDP - Delivered Duty Paid Madrid/);
});

test("keeps body paragraph spacing tight with a blank line after long paragraphs", async () => {
  const buffer = await renderConfirmationDocx(fakeConfirmation(), { mode: "formato3" });
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const paragraphs = [...documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const introParagraph = paragraphs.find((paragraph) =>
    extractParagraphText(paragraph).includes("LE AGRADECEMOS SU PEDIDO")
  );
  const reclamacionesParagraph = paragraphs.find((paragraph) =>
    extractParagraphText(paragraph).startsWith("RECLAMACIONES:")
  );

  assert.ok(introParagraph, "The introduction sentence should be present");
  assert.match(introParagraph, /<w:sz w:val="19"\/>/);
  assert.match(introParagraph, /<w:szCs w:val="19"\/>/);
  assert.ok(reclamacionesParagraph, "The reclamaciones paragraph should be present");
  assert.match(
    reclamacionesParagraph,
    /<w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="auto"\/>/
  );
});

test("uses origin column when multiple origins are present", async () => {
  const buffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      contractNumber: "STA-ORIG-01",
      origins: ["Planta Norte (ES)", "Planta Sur (FR)"],
      items: [
        {
          ...fakeConfirmation().items[0],
          number: 1,
          specification: "S235JR 2,00 x 1000 x 2000",
          sheetUnits: 50,
          quantity: 80,
          price: 860,
          amount: 68800,
          origin: "Planta Norte (ES)",
          minNet: 10,
          maxNet: 15,
          existences: []
        },
        {
          ...fakeConfirmation().items[0],
          number: 2,
          specification: "S235JR 2,50 x 1250 x 2500",
          sheetUnits: 25,
          quantity: 120,
          price: 900,
          amount: 108000,
          origin: "Planta Sur (FR)",
          minNet: 12,
          maxNet: 18,
          existences: []
        }
      ],
      totalQuantity: 200,
      hasSheetMaterial: false
    },
    {
      mode: "formato1"
    }
  );
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");

  assert.doesNotMatch(documentXml, /ORIGEN:\s*Planta Norte/);
  assert.doesNotMatch(documentXml, /ORIGEN:\s*Planta Sur/);
  assert.match(documentXml, /ORIGEN/);
  assert.match(documentXml, /Planta Norte \(ES\)/);
  assert.match(documentXml, /Planta Sur \(FR\)/);
});

test("adds bank details only when payment is transfer", async () => {
  const transferBuffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      paymentTerms: "Transferencia"
    },
    {
      mode: "formato3"
    }
  );
  const transferDoc = await JSZip.loadAsync(transferBuffer);
  const transferXml = await transferDoc.file("word/document.xml").async("string");

  assert.match(transferXml, /DETALLES BANCARIOS:?\s*/);
  assert.match(transferXml, /CAIXA BANK - ES40 2100 6428 2213 0012 3884/);
  assert.doesNotMatch(transferXml, /CUANDO EL PAGO ES POR TRANSFERENCIA/);
  const transferParagraphs = [...transferXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((paragraph) => paragraph[0]);
  const transferBankLineIndex = transferParagraphs.findIndex((paragraph) => paragraph.includes("CAIXA BANK - ES40 2100 6428 2213 0012 3884"));
  const transferBankHeaderIndex = transferParagraphs
    .map((paragraph, index) => ({ paragraph, index }))
    .filter((item) => /DETALLES BANCARIOS:/.test(item.paragraph))
    .map((item) => item.index)
    .filter((index) => index < transferBankLineIndex)
    .at(-1);
  assert.equal(transferBankHeaderIndex, transferBankLineIndex - 1);

  const nonTransferBuffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      paymentTerms: "Confirming"
    },
    {
      mode: "formato3"
    }
  );
  const nonTransferDoc = await JSZip.loadAsync(nonTransferBuffer);
  const nonTransferXml = await nonTransferDoc.file("word/document.xml").async("string");

  assert.doesNotMatch(nonTransferXml, /CAIXA BANK - ES40 2100 6428 2213 0012 3884/);
  assert.doesNotMatch(nonTransferXml, /DETALLES BANCARIOS/);
});

test("uses final reclamaciones text", async () => {
  const buffer = await renderConfirmationDocx(fakeConfirmation(), { mode: "formato3" });
  const doc = await JSZip.loadAsync(buffer);
  const documentXml = await doc.file("word/document.xml").async("string");

  assert.match(documentXml, /RECLAMACIONES: Si se encuentran da\u00f1os en las condiciones de los bienes[\s\S]*rfernandez@steeltradeadvisors\.com\./);
});

test("removes any packing line in all formats", async () => {
  const modes = ["formato1", "formato2", "formato3"];
  for (const mode of modes) {
    const buffer = await renderConfirmationDocx(fakeConfirmation(), { mode });
    const zip = await JSZip.loadAsync(buffer);
    const documentXml = await zip.file("word/document.xml").async("string");

    assert.match(documentXml, /PACKING LIST/);
    assert.doesNotMatch(documentXml, /PACKING:\s*Standard export packing/i);
    assert.doesNotMatch(documentXml, /standard export packing/i);
  }
});

test("always keeps fixed documentos lines in all formats", async () => {
  const modes = ["formato1", "formato2", "formato3"];
  for (const mode of modes) {
    const buffer = await renderConfirmationDocx(fakeConfirmation(), { mode });
    const zip = await JSZip.loadAsync(buffer);
    const documentXml = await zip.file("word/document.xml").async("string");

    const paragraphs = [...documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) =>
      match[0].replace(/<[^>]+>/g, "").replace(/\s+/g, " ").trim()
    );
    const documentosIndex = paragraphs.findIndex((paragraph) => paragraph.includes("DOCUMENTOS:"));
    const facturaIndex = paragraphs.findIndex((paragraph, index) =>
      index > documentosIndex && paragraph.includes("FACTURA COMERCIAL ORIGINAL")
    );
    const packingIndex = paragraphs.findIndex((paragraph, index) =>
      index > documentosIndex && paragraph.includes("PACKING LIST")
    );
    const millIndex = paragraphs.findIndex((paragraph, index) =>
      index > documentosIndex && paragraph.includes("MILL TEST CERTIFICADO 3.1 ACORDE A")
    );

    assert.ok(documentosIndex > -1);
    assert.ok(facturaIndex > documentosIndex);
    assert.ok(packingIndex > facturaIndex);
    assert.ok(millIndex > packingIndex);
  }
});

test("places client name in the right signature column", async () => {
  const buffer = await renderConfirmationDocx(fakeConfirmation(), { mode: "formato1" });
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const tables = [...documentXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const signatureTable = tables.find(
    (table) =>
      table.includes("FOR AND ON BEHALF OF") &&
      table.includes("STEEL TRADE ADVISORS, S.L.U.") &&
      table.includes(fakeConfirmation().customer.fiscalName)
  );

  assert.ok(signatureTable, "The final signature block should be a table with client name");
  assert.match(signatureTable, /<w:gridCol/);
  assert.match(signatureTable, /<w:jc w:val="left"\/>/);
  assert.match(signatureTable, /<w:jc w:val="right"\/>/);
});

test("normalizes client name in signature to avoid embedded line breaks", async () => {
  const signatureName = "ACME\r\nGLOBAL\u2028S.A.";
  const buffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      customer: {
        ...fakeConfirmation().customer,
        fiscalName: signatureName
      }
    },
    {
      mode: "formato1"
    }
  );
  const doc = await JSZip.loadAsync(buffer);
  const documentXml = await doc.file("word/document.xml").async("string");
  const tables = [...documentXml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map((match) => match[0]);
  const signatureTable = tables.find(
    (table) =>
      /STEEL TRADE ADVISORS, S\.L\.U\./.test(table) &&
      /ACME GLOBAL S\.A\./.test(table)
  );
  assert.ok(signatureTable, "The signature table should include normalized client name");

  const signatureText = extractParagraphText(signatureTable);
  assert.equal(signatureText.includes(signatureName), false);
  assert.equal(signatureText.includes("ACME GLOBAL S.A."), true);
  assert.match(signatureTable, /<w:noWrap\/>/);
});

test("shows storage line in all formats with rate by material type", async () => {
  const modes = ["formato1", "formato2", "formato3"];
  for (const mode of modes) {
    const coilBuffer = await renderConfirmationDocx(
      {
        ...fakeConfirmation(),
        hasSheetMaterial: false
      },
      { mode }
    );
    const coilDoc = await JSZip.loadAsync(coilBuffer);
    const coilXml = await coilDoc.file("word/document.xml").async("string");
    const coilStorageParagraph = storageParagraph(coilXml);

    assert.ok(coilStorageParagraph, `ALMACENAJES line should be present in ${mode} for coil orders`);
    assert.equal(
      extractParagraphText(coilStorageParagraph),
      "ALMACENAJES: 30 DÍAS LIBRES A PARTIR DE LA FECHA FACTURA, TRANSCURRIDO ESE PERIODO, SE FACTURARÁ A 0,15 €/MT POR DÍA."
    );
    assert.doesNotMatch(coilStorageParagraph, /w:color w:val="EE0000"/);
    assert.match(coilStorageParagraph, /w:color w:val="000000"/);
    assert.doesNotMatch(coilXml, /PARA LA BOBINA|PARA LA CHAPA/);

    const sheetBuffer = await renderConfirmationDocx(
      {
        ...fakeConfirmation(),
        hasSheetMaterial: true
      },
      { mode }
    );
    const sheetDoc = await JSZip.loadAsync(sheetBuffer);
    const sheetXml = await sheetDoc.file("word/document.xml").async("string");
    const sheetStorageParagraph = storageParagraph(sheetXml);

    assert.ok(sheetStorageParagraph, `ALMACENAJES line should be present in ${mode} for sheet orders`);
    assert.equal(
      extractParagraphText(sheetStorageParagraph),
      "ALMACENAJES: 30 DÍAS LIBRES A PARTIR DE LA FECHA FACTURA, TRANSCURRIDO ESE PERIODO, SE FACTURARÁ A 0,22 €/MT POR DÍA."
    );
    assert.doesNotMatch(sheetStorageParagraph, /w:color w:val="EE0000"/);
    assert.match(sheetStorageParagraph, /w:color w:val="000000"/);
    assert.doesNotMatch(sheetXml, /PARA LA BOBINA|PARA LA CHAPA/);
  }
});

function fakeConfirmation() {
  return {
    contractNumber: "STA-TEST",
    date: "2026-06-24",
    hasSheetMaterial: true,
    totalQuantity: 200,
    toleranceMinus: 0.1,
    tolerancePlus: 0.1,
    deliveryTerms: "CPT Madrid",
    paymentTerms: "Transferencia",
    customer: {
      fiscalName: "Cliente Test",
      commercialName: "",
      taxId: "B00000000",
      address: "Calle Test 1",
      postalCode: "28001",
      city: "Madrid",
      province: "Madrid",
      country: "España"
    },
    items: [
      {
        number: 1,
        specification: "S235JR 2,00 x 1000 x 2000",
        sheetUnits: 50,
        quantity: 200,
        price: 860,
        amount: 172000,
        minNet: 10,
        maxNet: 15,
        existences: []
      }
    ]
  };
}

function extractParagraphText(paragraph) {
  return [...paragraph.matchAll(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((run) => unescapeXml(run[2]))
    .join("")
    .replace(/\s+/g, " ")
    .trim();
}

function storageParagraph(documentXml) {
  return [...documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)]
    .map((match) => match[0])
    .find((paragraph) => extractParagraphText(paragraph).startsWith("ALMACENAJES:"));
}

function unescapeXml(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}
