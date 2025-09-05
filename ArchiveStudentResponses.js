function onEdit(e) {
  // Get the active sheet and the edited range
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Check if we're in the "Student Work" sheet and column H was edited
  if (sheet.getName() !== 'Student Work' || range.getColumn() !== 8) {
    return;
  }
  
  // Check if the checkbox was checked
  if (range.getValue() !== true) {
    return;
  }
  
  // Get the row number that was edited
  const rowNum = range.getRow();
  
  // Get reference to the archive sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName('Student Response Archive');
  
  // Get the full row data
  const sourceRow = sheet.getRange(rowNum, 1, 1, 8).getValues()[0];
  
  // Add timestamp to the archived data
  sourceRow.push(new Date());
  
  // Copy to archive sheet (in the next empty row)
  const lastArchiveRow = archiveSheet.getLastRow();
  const targetRow = lastArchiveRow + 1;
  archiveSheet.getRange(targetRow, 1, 1, sourceRow.length).setValues([sourceRow]);
  
  // Clear cells C through G in the source row
  for (let col = 3; col <= 7; col++) {
    sheet.getRange(rowNum, col).setValue('');
  }
  
  // Uncheck the checkbox
  range.setValue(false);
  
  // Force update
  SpreadsheetApp.flush();
}