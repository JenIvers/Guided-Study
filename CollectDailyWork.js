/**
 * Collects the most recent daily work from student documents
 */
function collectStudentDailyWork() {
  try {
    // Get the active spreadsheet and the Student Work sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student Work');
   
    if (!sheet) {
      throw new Error('Could not find sheet named "Student Work"');
    }


    // Get all data
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { // Only header row
      throw new Error('No student data found in sheet');
    }


    // Skip header row, process each student row
    for (let row = 1; row < data.length; row++) {
      const docId = data[row][1]; // Column B (index 1) contains Doc ID
     
      if (!docId) {
        continue; // Skip rows without doc ID
      }


      try {
        // Try to open the document
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();


        // Find the "Daily Work" heading
        const searchResult = body.findText("Daily Work");
        if (!searchResult) {
          Logger.log(`No "Daily Work" heading found in doc ${docId}`);
          continue;
        }


        // Get the element containing "Daily Work"
        const headingElement = searchResult.getElement().getParent();
        const headingIndex = body.getChildIndex(headingElement);


        // Find the first table after the heading
        let dailyWorkTable = null;
        for (let i = headingIndex + 1; i < body.getNumChildren(); i++) {
          const element = body.getChild(i);
          if (element.getType() === DocumentApp.ElementType.TABLE) {
            dailyWorkTable = element.asTable();
            break;
          }
        }


        if (!dailyWorkTable) {
          Logger.log(`No table found after Daily Work heading in doc ${docId}`);
          continue;
        }


        // Find the most recent non-empty task entry
        let mostRecentTask = '';
        let mostRecentDate = '';
        const numRows = dailyWorkTable.getNumRows();


        // Start from row 1 (skip header) and look for most recent task
        for (let i = 1; i < numRows; i++) {
          const row = dailyWorkTable.getRow(i);
          if (row.getNumCells() < 2) continue;


          const dateCell = row.getCell(0); // First column contains date
          const tasksCell = row.getCell(1); // Second column contains tasks
          const taskText = tasksCell.getText().trim();


          if (taskText) {
            // Get the date text (first line only, in case it includes day of week)
            const dateCellText = dateCell.getText().trim();
            mostRecentDate = dateCellText.split('\n')[0]; // Take first line only
            mostRecentTask = taskText;
            break; // Found the most recent task
          }
        }


        // Update the spreadsheet with the most recent task and date
        if (mostRecentTask && mostRecentDate) {
          sheet.getRange(row + 1, 3).setValue(mostRecentDate); // Column C for date
          sheet.getRange(row + 1, 4).setValue(mostRecentTask); // Column D for task
          Logger.log(`Updated tasks for doc ${docId}`);
        }


      } catch (docError) {
        Logger.log(`Error processing doc ${docId}: ${docError.message}`);
        continue; // Skip to next document on error
      }
    }


    SpreadsheetApp.getUi().alert('Student daily work collection completed.');


  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}