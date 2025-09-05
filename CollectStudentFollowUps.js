/**
 * Collects student follow-up responses from their documents
 */
function collectStudentFollowUps() {
  try {
    // Get the active spreadsheet and the Student Work sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(STUDENT_WORK_SHEET);
   
    if (!sheet) {
      throw new Error('Could not find sheet named "Student Work"');
    }


    // Get all data
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { // Only header row
      throw new Error('No student data found in sheet');
    }


    let updatedCount = 0;
    let errors = [];


    // Skip header row, process each student row
    for (let row = 1; row < data.length; row++) {
      const docId = data[row][STUDENT_WORK_DOC_ID_COLUMN - 1]; // Column B
      const date = data[row][STUDENT_WORK_DATE_COLUMN - 1]; // Column C
     
      // Skip if no doc ID or date
      if (!docId || !date) {
        continue;
      }


      try {
        // Format date to match document format (M/d/yyyy)
        const formattedDate = Utilities.formatDate(
          new Date(date),
          Session.getScriptTimeZone(),
          "M/d/yyyy"
        );


        // Open the document and find the Daily Work table
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();


        // Find the "Daily Work" heading
        const headingSearch = body.findText('Daily Work');
        if (!headingSearch) {
          throw new Error('Could not find "Daily Work" heading');
        }


        // Find the table after the heading
        let table = null;
        const headingElement = headingSearch.getElement().getParent();
        const headingIndex = body.getChildIndex(headingElement);


        for (let i = headingIndex + 1; i < body.getNumChildren(); i++) {
          const element = body.getChild(i);
          if (element.getType() === DocumentApp.ElementType.TABLE) {
            table = element.asTable();
            break;
          }
        }


        if (!table) {
          throw new Error('Could not find table after "Daily Work" heading');
        }


        // Find the row with matching date
        let found = false;
        for (let i = 1; i < table.getNumRows(); i++) {
          const tableRow = table.getRow(i);
          if (tableRow.getCell(0).getText().includes(formattedDate)) {
            // Get the student response cell (fourth column - index 3)
            const responseCell = tableRow.getCell(3);
            const responseText = responseCell.getText().trim();
           
            // Update the spreadsheet with the response (column F - index 5)
            sheet.getRange(row + 1, 6).setValue(responseText);
           
            found = true;
            updatedCount++;
            break;
          }
        }


        if (!found) {
          throw new Error(`Could not find row for date ${formattedDate}`);
        }


      } catch (docError) {
        errors.push(`Error processing document ${docId}: ${docError.message}`);
        continue;
      }
    }


    // Show completion message
    const ui = SpreadsheetApp.getUi();
    if (errors.length === 0) {
      ui.alert(
        'Success',
        `Successfully collected ${updatedCount} student follow-up response(s).`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Completed with errors',
        `Collected ${updatedCount} response(s).\n\nErrors:\n${errors.join('\n')}`,
        ui.ButtonSet.OK
      );
    }


  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to collect follow-ups: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}