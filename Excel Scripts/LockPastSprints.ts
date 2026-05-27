function main(workbook: ExcelScript.Workbook) {

  const ws: ExcelScript.Worksheet = workbook.getWorksheets()[0];
  const PASSWORD = "dswd"; // Store once, use everywhere

  // Change value on excel every sprint manually
  const CURRENT_SPRINT: number = ws.getRange("B2").getValue() as number; 

  // Unlock entire file
  if (ws.getProtection().getProtected()) {
    ws.getProtection().unprotect(PASSWORD);
  }

  const allCells: ExcelScript.Range = ws.getUsedRange();
  allCells.getFormat().getProtection().setLocked(false);

  const lastRow: number = ws.getUsedRange().getRowCount();
  const lastCol: number = ws.getUsedRange().getColumnCount();

  if (lastRow < 3) {
    ws.getProtection().protect({ allowSort: true, allowAutoFilter: true }, PASSWORD);
    return;
  }

  // Lock rows that are under sprint threshold 
  for (let i: number = 2; i < lastRow; i++) {
    const cell: ExcelScript.Range = ws.getRangeByIndexes(i, 0, 1, 1);
    const cellVal: string | number | boolean = cell.getValue();
    const row: ExcelScript.Range = ws.getRangeByIndexes(i, 0, 1, lastCol);

    if (cellVal === "") {
      row.getFormat().getProtection().setLocked(false);

    } else if (typeof cellVal === "number" && Number.isInteger(cellVal)) {
      if (cellVal < CURRENT_SPRINT) {
        row.getFormat().getProtection().setLocked(true);
      } else {
        row.getFormat().getProtection().setLocked(false);
      }

    } else {
      row.getFormat().getProtection().setLocked(false);
    }
  }

  ws.getProtection().protect({ allowSort: true, allowAutoFilter: true }, PASSWORD);
}
