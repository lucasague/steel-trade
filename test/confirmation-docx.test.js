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
  const merchandiseIndex = documentXml.indexOf("MERCANC");
  const tableIndex = documentXml.indexOf("<w:tbl>", merchandiseIndex);

  assert.ok(merchandiseIndex > -1);
  assert.ok(tableIndex > merchandiseIndex);
  assert.match(documentXml, /S235JR/);
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

test("replaces merchandise header and origin line", async () => {
  const buffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
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

  assert.match(documentXml, /MERCANC/);
  assert.doesNotMatch(documentXml, /MERCANCIA:/);
  assert.match(documentXml, /ORIGEN: Planta Madrid \(ES \/ FR\)/);
    assert.match(documentXml, /CONDICIONES DE ENTREGA: DDP - Delivered Duty Paid Madrid/);
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
  const paragraphs = [...documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const signatureParagraph = paragraphs.find(
    (paragraph) =>
      paragraph.includes("STEEL TRADE ADVISORS, S.L.U.") &&
      paragraph.includes(fakeConfirmation().customer.fiscalName)
  );

  assert.ok(signatureParagraph, "The final signature paragraph should include client name");
  assert.match(signatureParagraph, /STEEL TRADE ADVISORS, S\.L\.U\./);
  assert.ok(
    /w:tab\/>[\s\S]*Cliente Test/.test(signatureParagraph),
    "Client name should be after a tab in the signature line"
  );
});

test("normalizes client name in signature to avoid embedded line breaks", async () => {
  const signatureName = "ACME\nGLOBAL SL";
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
  const paragraphs = [...documentXml.matchAll(/<w:p[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const signatureParagraph = paragraphs.find(
    (paragraph) =>
      /STEEL TRADE ADVISORS, S\.L\.U\./.test(paragraph) &&
      /ACME GLOBAL SL/.test(paragraph)
  );
  assert.ok(signatureParagraph);

  const signatureText = extractParagraphText(signatureParagraph);
  assert.equal(signatureText.includes(signatureName), false);
  assert.equal(signatureText.includes("ACME GLOBAL SL"), true);
});

test("shows storage line only for formato 3 and keeps it black", async () => {
  const sheetBuffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      hasSheetMaterial: true
    },
    {
      mode: "formato3"
    }
  );
  const sheetDoc = await JSZip.loadAsync(sheetBuffer);
  const sheetXml = await sheetDoc.file("word/document.xml").async("string");
  const sheetStorageParagraph = sheetXml.match(
    /<w:p[\s\S]*?<\/w:p>/g
  );
  const storageParagraph = [...sheetStorageParagraph || []].find((paragraph) =>
    /ALMACENAJES:[\s\S]*30 DIAS LIBRES/i.test(paragraph)
  );

  assert.ok(storageParagraph, "ALMACENAJES line should be present in formato 3");
  assert.doesNotMatch(
    storageParagraph,
    /w:color w:val="EE0000"/,
    "Storage line should not keep red color in formato 3"
  );
  assert.match(storageParagraph, /w:color w:val="000000"/);

  const nonSheetBuffer = await renderConfirmationDocx(
    {
      ...fakeConfirmation(),
      hasSheetMaterial: false
    },
    {
      mode: "formato1"
    }
  );
  const nonSheetDoc = await JSZip.loadAsync(nonSheetBuffer);
  const nonSheetXml = await nonSheetDoc.file("word/document.xml").async("string");

  assert.doesNotMatch(
    nonSheetXml,
    /ALMACENAJES: 30 DIAS LIBRES/
  );
});

test("never includes the 30 días libres storage text when not chapa", async () => {
  const nonChapaModes = ["formato1", "formato2"];
  for (const mode of nonChapaModes) {
    const nonSheetBuffer = await renderConfirmationDocx(
      {
        ...fakeConfirmation(),
        hasSheetMaterial: false
      },
      { mode }
    );
    const nonSheetDoc = await JSZip.loadAsync(nonSheetBuffer);
    const nonSheetXml = await nonSheetDoc.file("word/document.xml").async("string");

    assert.doesNotMatch(
      nonSheetXml,
      /30 DIAS LIBRES A PARTIR DE LA FECHA FACTURA/
    );
    assert.doesNotMatch(nonSheetXml, /30 DIAS LIBRES/);
    assert.doesNotMatch(nonSheetXml, /ALMACENAJES:/);
  }

  const chapaLikeModes = ["formato1", "formato2"];
  for (const mode of chapaLikeModes) {
    const sheetButNonStorageBuffer = await renderConfirmationDocx(
      {
        ...fakeConfirmation(),
        hasSheetMaterial: true
      },
      { mode }
    );
    const sheetButNonStorageDoc = await JSZip.loadAsync(sheetButNonStorageBuffer);
    const sheetButNonStorageXml = await sheetButNonStorageDoc
      .file("word/document.xml")
      .async("string");

    assert.doesNotMatch(
      sheetButNonStorageXml,
      /30 DIAS LIBRES A PARTIR DE LA FECHA FACTURA/
    );
    assert.doesNotMatch(sheetButNonStorageXml, /ALMACENAJES:/);
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

function unescapeXml(value) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}
