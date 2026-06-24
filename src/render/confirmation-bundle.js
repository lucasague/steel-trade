import { renderConfirmationDocx } from "./confirmation-docx.js";
import { renderConfirmationExcel } from "./confirmation-excel.js";

export async function renderConfirmationResponse(confirmation, { mode, format }) {
  const safeContractNumber = confirmation.contractNumber.replace(/[^\w.-]+/g, "-");

  if (format === "xlsx") {
    return {
      filename: `${safeContractNumber}-confirmacion.xlsx`,
      contentType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: await renderConfirmationExcel(confirmation, { mode })
    };
  }

  return {
    filename: `${safeContractNumber}-confirmacion.docx`,
    contentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await renderConfirmationDocx(confirmation, { mode })
  };
}
