function main(workbook: ExcelScript.Workbook) {

  const ws: ExcelScript.Worksheet = workbook.getWorksheets()[0];
  const PASSWORD = "PASSWORD";

  if (ws.getProtection().getProtected()) {
    ws.getProtection().unprotect(PASSWORD);
  }

  const allCells: ExcelScript.Range = ws.getUsedRange();
  allCells.getFormat().getProtection().setLocked(false);

}
