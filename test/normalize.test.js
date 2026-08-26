import test from "node:test";
import assert from "node:assert/strict";
import { finalizeInvoice } from "../src/domain/normalize.js";

test("groups invoice lines by purchase item, description and price", () => {
  const invoice = finalizeInvoice({
    invoiceNumber: "T-1",
    lines: [
      line("A", "2,10 x 1250 S235JR", 10, 485),
      line("A", "2,10 x 1250 S235JR", 15, 485),
      line("B", "2,40 x 1219 S235JR", 20, 490)
    ]
  });

  assert.equal(invoice.groupedLines.length, 2);
  assert.equal(invoice.groupedLines[0].quantity, 25);
  assert.equal(invoice.groupedLines[0].amount, 12125);
  assert.equal(invoice.subtotal, 21925);
});

function line(purchaseItemId, description, quantity, price) {
  return {
    purchaseItemIds: [purchaseItemId],
    description,
    quantity,
    price,
    amount: quantity * price
  };
}

