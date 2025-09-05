// SendTeacherFeedback.gs
// ====================

/**
 * Sends teacher feedback from spreadsheet to student documents
 */
function sendTeacherFeedback() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(STUDENT_WORK_SHEET);
    
    if (!sheet) {
      throw new Error(`Sheet "${STUDENT_WORK_SHEET}" not found`);
    }

    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    
    let updatedCount = 0;
    let errors = [];

    // Start from row 1 (skip header)
    for (let row = 1; row < data.length; row++) {
      const docId = data[row][STUDENT_WORK_DOC_ID_COLUMN - 1];
      const date = data[row][STUDENT_WORK_DATE_COLUMN - 1];
      const comment = data[row][STUDENT_WORK_COMMENTS_COLUMN - 1];

      // Skip if any required field is empty
      if (!docId || !date || !comment) continue;

      try {
        // Format date to match the document format (M/d/yyyy)
        const formattedDate = Utilities.formatDate(new Date(date),
          Session.getScriptTimeZone(), "M/d/yyyy");

        // Open the document and find the Daily Work table
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();

        // Find the "Daily Work" heading
        const headingSearch = body.findText('Daily Work');
        if (!headingSearch) {
          throw new Error(`Could not find "Daily Work" heading`);
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
          throw new Error(`Could not find table after "Daily Work" heading`);
        }

        // Find the row with matching date
        let found = false;
        for (let i = 1; i < table.getNumRows(); i++) {
          const row = table.getRow(i);
          const dateText = row.getCell(0).getText().trim();
          if (dateText.includes(formattedDate)) {
            // Get the teacher comments cell (third column)
            const commentsCell = row.getCell(2);
            
            // Replace with the new comment
            const newText = comment;
            
            // Completely clear the cell
            commentsCell.clear();
            
            // Remove any existing paragraphs
            while (commentsCell.getNumChildren() > 0) {
              commentsCell.removeChild(commentsCell.getChild(0));
            }
            
            // Add the new text with proper formatting
            const paragraph = commentsCell.appendParagraph(newText);
            paragraph
              .setFontFamily('Nunito')
              .setFontSize(12)
              .setForegroundColor('#cc0000') // Match the red color from student doc
              .setBold(true);                // Match the bold style from student doc

            found = true;
            updatedCount++;
            break;
          }
        }

        if (!found) {
          throw new Error(`Could not find row for date ${formattedDate}`);
        }

      } catch (docError) {
        errors.push(`Error processing student document: ${docError.message}`);
        continue;
      }
    }

    // Show completion message
    const ui = SpreadsheetApp.getUi();
    if (errors.length === 0) {
      ui.alert('Success',
        `Successfully added feedback to ${updatedCount} student document(s).`,
        ui.ButtonSet.OK);
    } else {
      ui.alert('Completed with errors',
        `Added feedback to ${updatedCount} document(s).\n\nErrors:\n${errors.join('\n')}`,
        ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('Error',
      `Failed to send feedback: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}