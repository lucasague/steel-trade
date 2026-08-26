import ExcelJS from "exceljs";
import { readFileSync } from "node:fs";

const LOGO_PATH = new URL("../assets/sta-imagotipo-fondo-blanco.png", import.meta.url);
const LOGO_SIZE = { width: 120, height: 104 };
const COLUMNS = [
  { label: "Uds", key: "units", width: 8 },
  { label: "Material", key: "material", width: 14 },
  { label: "Espesor", key: "thickness", width: 10 },
  { label: "Ancho", key: "width", width: 10 },
  { label: "Largo", key: "length", width: 10 },
  { label: "Calidad", key: "quality", width: 16 },
  { label: "Recubrimiento", key: "coating", width: 14 },
  { label: "Acabado", key: "finish", width: 14 },
  { label: "Color", key: "color", width: 12 },
  { label: "ID Fábrica", key: "factoryId", width: 20 },
  { label: "Neto", key: "net", width: 12 },
  { label: "Bruto", key: "gross", width: 12 }
];
const COLORS = {
  header: "FF7B0000",
  total: "FFA75454",
  itemTotal: "FFD6AFAF",
  white: "FFFFFFFF",
  black: "FF000000",
  border: "FFDDDDDD"
};

export async function renderPackingExcel(invoice) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Steel Trade invoice service";
  workbook.created = new Date();

  const sheet = workbook.addWorksheet("Sheet1", {
    pageSetup: {
      paperSize: 9,
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0
    }
  });

  sheet.columns = COLUMNS.map(({ key, width }) => ({ key, width }));
  applyReferenceSheetLayout(sheet);

  const logoId = workbook.addImage({
    buffer: readFileSync(LOGO_PATH),
    extension: "png"
  });
  sheet.addImage(logoId, {
    tl: { col: 1, row: 1 },
    ext: LOGO_SIZE
  });

  sheet.mergeCells("B2:B6");
  sheet.mergeCells("B7:G7");
  sheet.mergeCells("B8:K8");
  sheet.mergeCells("H7:L7");
  sheet.mergeCells("H9:L9");

  sheet.getCell("C2").value = invoice.company.name;
  sheet.getCell("C3").value = invoice.company.address;
  sheet.getCell("C4").value = invoice.company.city;
  sheet.getCell("C5").value = invoice.company.phone;
  sheet.getCell("C6").value = invoice.company.taxId;
  styleCompanyHeader(sheet);

  sheet.getCell("H3").value = invoice.customer.fiscalName;
  sheet.getCell("H4").value = invoice.customer.address;
  sheet.getCell("H5").value = [
    invoice.customer.postalCode,
    invoice.customer.city,
    invoice.customer.province,
    invoice.customer.country
  ]
    .filter(Boolean)
    .join(" ");
  sheet.getCell("H6").value = invoice.customer.taxId;
  styleCustomerHeader(sheet);

  sheet.getCell("L8").value = invoice.date || null;
  sheet.getCell("L8").numFmt = "d/MM/yyyy";
  sheet.getCell("L8").font = { name: "Arial", size: 12 };
  sheet.getCell("L8").alignment = { vertical: "middle" };

  sheet.getCell("B8").value = `${invoice.packingList || ""} (Commercial Invoice Number: ${invoice.invoiceNumber})`;
  sheet.getCell("B8").font = { name: "Arial", size: 12, bold: true };
  sheet.getCell("B8").alignment = { horizontal: "center", vertical: "middle" };

  const headerRow = sheet.getRow(10);
  COLUMNS.forEach((column, index) => {
    headerRow.getCell(index + 1).value = column.label;
  });
  styleHeaderRow(headerRow);

  const totalRow = sheet.getRow(11);
  totalRow.getCell(1).value = "TOTAL";
  totalRow.getCell(11).value = invoice.lines.reduce((sum, line) => sum + line.net, 0);
  totalRow.getCell(12).value = invoice.lines.reduce((sum, line) => sum + line.gross, 0);
  styleTotalRow(totalRow);

  let rowIndex = 12;
  for (const group of invoice.groupedLines) {
    const groupRow = sheet.getRow(rowIndex++);
    groupRow.getCell(2).value = group.material;
    groupRow.getCell(3).value = group.lines[0]?.thickness;
    groupRow.getCell(4).value = group.lines[0]?.width;
    groupRow.getCell(5).value = group.lines[0]?.length;
    groupRow.getCell(6).value = group.quality;
    groupRow.getCell(11).value = group.lines.reduce((sum, line) => sum + line.net, 0);
    groupRow.getCell(12).value = group.lines.reduce((sum, line) => sum + line.gross, 0);
    styleItemTotalRow(groupRow);

    for (const line of group.lines) {
      const row = sheet.getRow(rowIndex++);
      row.values = [
        line.units,
        line.material,
        line.thickness,
        line.width,
        line.length,
        line.quality,
        "",
        "",
        "",
        line.factoryId,
        line.net,
        line.gross
      ];
    }
  }

  for (const row of sheet.getRows(10, Math.max(1, rowIndex - 10)) || []) {
    for (let colIndex = 1; colIndex <= COLUMNS.length; colIndex += 1) {
      const cell = row.getCell(colIndex);
      cell.border = {
        top: { style: "thin", color: { argb: "FFDDDDDD" } },
        bottom: { style: "thin", color: { argb: "FFDDDDDD" } }
      };
      cell.font = cell.font?.bold ? cell.font : { name: "Arial" };
      cell.alignment = { horizontal: "center", vertical: "middle" };
    }
  }

  for (const col of [3, 10, 11, 12]) {
    sheet.getColumn(col).numFmt = "0.000";
  }
  sheet.getColumn(4).numFmt = "#,##0";
  sheet.getColumn(5).numFmt = "#,##0";

  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function applyReferenceSheetLayout(sheet) {
  const widths = [10, 18, 9.13, 10, 13, 15.5, 13, 17.63, 18.88, 19.25, 13.75, 13];
  widths.forEach((width, index) => {
    sheet.getColumn(index + 1).width = width;
  });
  sheet.getRow(1).height = 27.75;
}

function styleCompanyHeader(sheet) {
  sheet.getCell("C2").font = {
    name: "Lexend",
    size: 13,
    bold: true,
    color: { argb: COLORS.header }
  };
  for (const address of ["C3", "C4", "C5", "C6"]) {
    sheet.getCell(address).font = { name: "Arial", size: 11 };
  }
}

function styleCustomerHeader(sheet) {
  sheet.getCell("H3").font = { name: "Arial", size: 12, bold: true };
  for (const address of ["H4", "H5", "H6"]) {
    sheet.getCell(address).font = { name: "Arial", size: 11 };
  }
}

function styleHeaderRow(row) {
  row.eachCell((cell) => {
    cell.fill = solidFill(COLORS.header);
    cell.font = { name: "Arial", bold: true, color: { argb: COLORS.white } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
}

function styleTotalRow(row) {
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = solidFill(COLORS.total);
    cell.font = {
      name: "Arial",
      size: 12,
      bold: true,
      color: { argb: COLORS.white }
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
}

function styleItemTotalRow(row) {
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = solidFill(COLORS.itemTotal);
    cell.font = { name: "Arial", bold: true, color: { argb: COLORS.black } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
}

function solidFill(argb) {
  return {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb }
  };
}
