import { finalizeInvoice } from "./domain/normalize.js";
import { getConfig } from "./config.js";

export function buildSampleInvoice() {
  const config = getConfig({});
  const baseLines = [
    ["HRC", "2,10 x 1250", "S235JR", 23.285, 23.295, 485, "PL 25"],
    ["HRC", "2,10 x 1250", "S235JR", 23.21, 23.22, 485, "PL 25"],
    ["HRC", "2,40 x 1219", "S235JR", 22.0, 22.01, 485, "PL 25"],
    ["HRC", "2,40 x 1219", "S235JR", 24.38, 24.39, 485, "PL 25"],
    ["HRC", "2,30 x 1296", "S235JR", 23.45, 23.46, 485, "PL 26"],
    ["HRC", "2,75 x 1250", "S275JR", 24.12, 24.13, 494, "PL 27"],
    ["HRC", "3,65 x 1500", "S275JR", 24.7, 24.71, 494, "PL 27"],
    ["HRC", "7,55 x 1510", "S355J2", 25.2, 25.21, 524, "PL 28"]
  ];
  const lines = [];

  for (let round = 0; round < 5; round++) {
    for (const [material, measure, quality, net, gross, price, packing] of baseLines) {
      const factoryId = `${2600000000 + round * 100 + lines.length}`;
      lines.push({
        recordId: `recSample${lines.length}`,
        factoryId,
        productType: "Bobina",
        material,
        measure,
        thickness: Number(String(measure).split(" ")[0].replace(",", ".")),
        width: String(measure).split(" x ")[1],
        length: "",
        quality,
        purchaseItemIds: [`STA-SAMPLE-${material}-${measure}-${quality}`],
        saleItemIds: [`STA-2026-0002-${round + 1}`],
        quantity: net,
        net,
        gross,
        units: 1,
        price,
        amount: Math.round(net * price * 100) / 100,
        packingList: `${packing} - MV PHILIPP BAY`,
        description: [measure, quality].join(" "),
        detailDescription: [
          material,
          measure,
          quality,
          factoryId,
          `${packing} - MV PHILIPP BAY`
        ].join(" ")
      });
    }
  }

  return finalizeInvoice({
    recordId: "recSampleInvoice",
    company: config.company,
    invoiceNumber: "2026-SAMPLE",
    date: "2026-01-30",
    dueDate: "2026-01-30",
    description: "HOT ROLLED STEEL COILS",
    customer: {
      commercialName: "Cliente Demo",
      fiscalName: "SteelInvest Iberia S.L.",
      address: "Calle Villanueva 23, 1 Centro",
      postalCode: "28001",
      city: "Madrid",
      province: "Madrid",
      country: "Espa\u00f1a",
      taxId: "B70864863"
    },
    saleRef: "STA-2026-0002",
    packingList:
      "PL 25 - MV PHILIPP BAY / PL 26 - MV PHILIPP BAY / PL 27 - MV PHILIPP BAY / PL 28 - MV PHILIPP BAY",
    deliveryTerms: "CFR BILBAO",
    paymentTerms: "Transferencia a la vista fecha factura",
    lines
  });
}

