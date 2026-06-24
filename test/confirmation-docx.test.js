import test from "node:test";
import assert from "node:assert/strict";
import JSZip from "jszip";
import { renderConfirmationDocx } from "../src/render/confirmation-docx.js";

test("inserts the merchandise table inside the confirmation Word", async () => {
  const buffer = await renderConfirmationDocx(fakeConfirmation(), {
    mode: "formato3"
  });
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const merchandiseIndex = documentXml.indexOf("MERCANC");
  const tableIndex = documentXml.indexOf("<w:tbl>", merchandiseIndex);

  assert.ok(merchandiseIndex > -1);
  assert.ok(tableIndex > merchandiseIndex);
  assert.match(documentXml, /UNIDADES/);
  assert.match(documentXml, /S235JR/);
  assert.match(documentXml, /PESO BOBINA \(MT\)/);
  assert.match(documentXml, /10,000 - 15,000 MT/);
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

  assert.match(documentXml, /MERCANCÍA/);
  assert.doesNotMatch(documentXml, /MERCANCÍA:/);
  assert.match(documentXml, /ORIGEN: Planta Madrid \(ES \/ FR\)/);
  assert.match(documentXml, /CONDICIONES DE ENTREGA: DDP/);
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

  assert.match(transferXml, /CAIXA BANK - ES40 2100 6428 2213 0012 3884/);
  assert.match(transferXml, /CUANDO EL PAGO ES POR TRANSFERENCIA/);

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
