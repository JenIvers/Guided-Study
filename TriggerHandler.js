// TriggerHandler.gs

const SHEET_ACTIVATED_KEY = 'isActivated';

/**
 * Menu-driven activation function - this should only be called from the menu
 */
function activateSheetFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    // Perform the actual activation
    activateSheet();
    
    // If successful, show instructions
    ui.alert(
      'Sheet activated! Important: You must now set up the automatic trigger.\n\n' +
      '1. Click on "Extensions" in the menu\n' +
      '2. Click "Apps Script"\n' +
      '3. In the Apps Script editor, click on "Triggers" (clock icon on left)\n' +
      '4. Click "+ Add Trigger" (bottom right)\n' +
      '5. Set up the trigger with these settings:\n' +
      '   - Choose function to run: onEdit\n' +
      '   - Select event source: From spreadsheet\n' +
      '   - Select event type: On edit\n' +
      '6. Click Save'
    );
    
  } catch (e) {
    ui.alert('Failed to activate sheet: ' + e.toString());
    PropertiesService.getDocumentProperties().setProperty(SHEET_ACTIVATED_KEY, 'false');
  }
}

/**
 * Core activation function - can be called from any context
 */
function activateSheet() {
  try {
    // Test basic spreadsheet access
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(STUDENT_WORK_SHEET);
    if (!sheet) {
      throw new Error('Could not find Student Work sheet');
    }
    
    // Set activation status
    PropertiesService.getDocumentProperties().setProperty(SHEET_ACTIVATED_KEY, 'true');
    
    return true;
  } catch (e) {
    console.error('Activation failed:', e);
    PropertiesService.getDocumentProperties().setProperty(SHEET_ACTIVATED_KEY, 'false');
    throw e; // Re-throw to allow caller to handle
  }
}

// Utility function to check activation status
function isSheetActivated() {
  return PropertiesService.getDocumentProperties().getProperty(SHEET_ACTIVATED_KEY) === 'true';
}

// Combined edit trigger handler

// Handler for teacher feedback
function handleTeacherFeedback(range, sheet) {
  const row = range.getRow();
  const newComment = range.getValue();
  
  // Only proceed if there's actual content
  if (newComment && newComment.trim() !== '') {
    try {
      // Get the doc ID from Column B of this row
      const docId = sheet.getRange(row, STUDENT_WORK_DOC_ID_COLUMN).getValue();
      const date = sheet.getRange(row, STUDENT_WORK_DATE_COLUMN).getValue();
      
      if (!docId || !date) {
        console.error('Missing doc ID or date in row', row);
        return;
      }

      // Send feedback only to this specific document
      sendTeacherFeedbackToDoc(docId, date, newComment);
    } catch (error) {
      console.error('Error sending teacher feedback:', error);
    }
  }
}

// Handler for archive checkbox
function handleArchive(range, sheet) {
  const isChecked = range.getValue() === true;
  if (isChecked) {
    try {
      // Get the row data
      const row = range.getRow();
      const archiveSheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(STUDENT_RESPONSE_ARCHIVE_SHEET);
      
      if (!archiveSheet) {
        console.error('Archive sheet not found');
        return;
      }
      
      // Get all values from the row
      const rowValues = sheet.getRange(row, 1, 1, ARCHIVE_CHECKBOX_COLUMN).getValues()[0];
      
      // Add timestamp to archive entry
      const timestamp = new Date();
      rowValues.push(timestamp);
      
      // Add to archive sheet
      archiveSheet.appendRow(rowValues);
      
      // Clear the checkbox and response
      sheet.getRange(row, ARCHIVE_CHECKBOX_COLUMN).setValue(false);
      sheet.getRange(row, STUDENT_WORK_RESPONSE_COLUMN).clearContent();
      
    } catch (error) {
      console.error('Error archiving response:', error);
    }
  }
}

// Function to send feedback to a specific document
function sendTeacherFeedbackToDoc(docId, date, comment) {
  try {
    // Format date to match document format (M/d/yyyy)
    const formattedDate = Utilities.formatDate(new Date(date),
      Session.getScriptTimeZone(), "M/d/yyyy");

    // Open the specific document
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Find the "Daily Work" heading
    const headingSearch = body.findText('Daily Work');
    if (!headingSearch) {
      throw new Error(`Could not find "Daily Work" heading in doc ${docId}`);
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
      throw new Error(`Could not find table after "Daily Work" heading in doc ${docId}`);
    }

    // Find the row with matching date
    let found = false;
    for (let i = 1; i < table.getNumRows(); i++) {
      const row = table.getRow(i);
      if (row.getCell(0).getText().includes(formattedDate)) {
        // Get the teacher comments cell (third column)
        const commentsCell = row.getCell(2);
        
        // Replace with the new comment
        commentsCell.clear();
        
        // Remove any existing paragraphs
        while (commentsCell.getNumChildren() > 0) {
          commentsCell.removeChild(commentsCell.getChild(0));
        }
        
        // Add the new text with proper formatting
        const paragraph = commentsCell.appendParagraph(comment);
        paragraph
          .setFontFamily('Nunito')
          .setFontSize(12)
          .setForegroundColor('#cc0000')
          .setBold(true);

        found = true;
        break;
      }
    }

    if (!found) {
      throw new Error(`Could not find row for date ${formattedDate} in doc ${docId}`);
    }

  } catch (error) {
    console.error(`Error processing document ${docId}:`, error.message);
    throw error;
  }
}

// Check activation status for documents and update spreadsheet
function checkActivationStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    console.error(`Sheet "${SHEET_NAME}" not found`);
    return;
  }

  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  
  // Skip header row
  for (let row = 1; row < data.length; row++) {
    const docId = data[row][DOC_ID_COLUMN - 1];
    
    if (!docId) continue;
    
    try {
      console.log(`Checking document ID: ${docId}`);
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      
      const today = new Date();
      const formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "M/d/yyyy");
      
      const searchResult = body.findText(formattedToday);
      const isActivated = searchResult !== null;
      console.log(`Activation check for doc ${docId}: ${isActivated ? 'Document is activated' : 'Document not activated'}`);
      
      if (isActivated && !data[row][ACTIVATION_DATE_COLUMN - 1]) {
        const file = DriveApp.getFileById(docId);
        const activationDate = file.getLastUpdated();
        
        const formattedDate = Utilities.formatDate(
          activationDate,
          Session.getScriptTimeZone(),
          "MM/dd/yyyy HH:mm:ss"
        );
        
        sheet.getRange(row + 1, ACTIVATION_DATE_COLUMN).setValue(formattedDate);
      }
    } catch (error) {
      console.error(`Error checking doc ${docId}: ${error.message}`);
      continue;
    }
  }
}

// Handle time-based triggers
function setupTimedTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkActivationStatus') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('checkActivationStatus')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
}