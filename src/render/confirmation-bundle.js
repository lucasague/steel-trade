import { renderConfirmationDocx } from "./confirmation-docx.js";

export async function renderConfirmationResponse(confirmation, { mode }) {
  const safeContractNumber = confirmation.contractNumber.replace(/[^\w.-]+/g, "-");

  return {
    filename: `${safeContractNumber}-confirmacion.docx`,
    contentType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await renderConfirmationDocx(confirmation, { mode })
  };
}
