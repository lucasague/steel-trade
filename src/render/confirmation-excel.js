import ExcelJS from "exceljs";
import { getConfirmationLines } from "../services/confirmation-service.js";
import { roundMoney } from "../domain/format.js";

const TEMPLATE_URL = new URL("../templates/formatos-confirmacion.xlsx", import.meta.url);
const VAT_RATE = 0.21;
const BLOCKS = {
  grouped: { start: 4, label: "FORMATO 1", subtitle: "AGRUPADO" },
  detail: { start: 14, label: "FORMATO 2", subtitle: "DETALLADO" },
  formato3: { start: 24, label: "FORMATO 3", subtitle: "CHAPA", subtitle2: "AGRUPADO" }
};

export async function renderConfirmationExcel(confirmation, { mode }) {
  const source = new ExcelJS.Workbook();
  await source.xlsx.readFile(TEMPLATE_URL);
  const sourceSheet = source.getWorksheet(1);
  const block = BLOCKS[mode] || BLOCKS.grouped;
  const lines = getConfirmationLines(confirmation, mode);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Steel Trade confirmation service";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("Confirmacion", {
    pageSetup: {
      paperSize: 9,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0
    }
  });

  copyColumns(sourceSheet, sheet);
  copyRow(sourceSheet, sheet, block.start, 1);
  copyRow(sourceSheet, sheet, block.start + 2, 3);

  sheet.getCell("A1").value = block.label;
  sheet.getCell("C1").value = block.subtitle;
  sheet.getCell("D1").value = block.subtitle2 || null;
  applyHeaderMerges(sheet, mode, 3);

  let rowIndex = 4;
  lines.forEach((line, index) => {
    copyRow(sourceSheet, sheet, block.start + 3, rowIndex);
    writeLine(sheet, rowIndex, line, index + 1, mode);
    applyLineMerge(sheet, mode, rowIndex);
    rowIndex += 1;
  });

  const subtotalRow = rowIndex++;
  copyRow(sourceSheet, sheet, block.start + 6, subtotalRow);
  mergeRow(sheet, subtotalRow, 1, 4);
  sheet.getCell(subtotalRow, 5).value = sum(lines, "quantity");
  sheet.getCell(subtotalRow, 7).value = sum(lines, "amount");

  const vatRow = rowIndex++;
  copyRow(sourceSheet, sheet, block.start + 7, vatRow);
  mergeRow(sheet, vatRow, 1, 6);
  sheet.getCell(vatRow, 1).value = "IVA 21%";
  sheet.getCell(vatRow, 7).value = roundMoney(sum(lines, "amount") * VAT_RATE);

  const totalRow = rowIndex++;
  copyRow(sourceSheet, sheet, block.start + 8, totalRow);
  mergeRow(sheet, totalRow, 1, 7);
  sheet.getCell(totalRow, 1).value = roundMoney(sum(lines, "amount") * (1 + VAT_RATE));

  sheet.getColumn(5).numFmt = "#,##0.000";
  sheet.getColumn(6).numFmt = "#,##0.00";
  sheet.getColumn(7).numFmt = "#,##0.00";

  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function writeLine(sheet, rowIndex, line, index, mode) {
  sheet.getCell(rowIndex, 1).value = line.itemNumber || index;
  sheet.getCell(rowIndex, 2).value = line.specification;
  if (mode === "detail") {
    sheet.getCell(rowIndex, 3).value = line.factoryId;
  }
  if (mode === "formato3") {
    sheet.getCell(rowIndex, 4).value = line.units || "";
  }
  sheet.getCell(rowIndex, 5).value = line.quantity || 0;
  sheet.getCell(rowIndex, 6).value = line.price || 0;
  sheet.getCell(rowIndex, 7).value = line.amount || roundMoney((line.quantity || 0) * (line.price || 0));
}

function copyColumns(sourceSheet, targetSheet) {
  for (let index = 1; index <= 7; index += 1) {
    targetSheet.getColumn(index).width = sourceSheet.getColumn(index).width;
  }
}

function copyRow(sourceSheet, targetSheet, sourceRowIndex, targetRowIndex) {
  const sourceRow = sourceSheet.getRow(sourceRowIndex);
  const targetRow = targetSheet.getRow(targetRowIndex);
  targetRow.height = sourceRow.height;
  for (let col = 1; col <= 7; col += 1) {
    const sourceCell = sourceRow.getCell(col);
    const targetCell = targetRow.getCell(col);
    targetCell.value = sourceCell.value;
    targetCell.style = clone(sourceCell.style);
  }
}

function applyHeaderMerges(sheet, mode, row) {
  if (mode === "grouped" || mode === "detail") {
    mergeRow(sheet, row, 2, 4);
  }
}

function applyLineMerge(sheet, mode, row) {
  if (mode === "grouped") {
    mergeRow(sheet, row, 2, 4);
  } else if (mode === "detail") {
    mergeRow(sheet, row, 3, 4);
  }
}

function mergeRow(sheet, row, left, right) {
  if (right > left) {
    sheet.mergeCells(row, left, row, right);
  }
}

function sum(lines, key) {
  return roundMoney(lines.reduce((total, line) => total + (Number(line[key]) || 0), 0));
}

function clone(value) {
  return value ? JSON.parse(JSON.stringify(value)) : value;
}
