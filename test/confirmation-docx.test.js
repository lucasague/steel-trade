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
  assert.match(documentXml, /FORMATO 3 - CHAPA AGRUPADO/);
  assert.match(documentXml, /UNIDADES/);
  assert.match(documentXml, /S235JR/);
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
        existences: []
      }
    ]
  };
}
