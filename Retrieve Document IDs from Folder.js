// Global variables
var SPREADSHEET_ID = '';
var FOLDER_IDS = {};

function onOpen() {
  createMenu();
}

function createMenu() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Teacher Tools');
  
  if (isScriptActivated()) {
    menu.addItem('Link Schoology folder from Drive', 'linkSchoologyFolder')
        .addItem('Collect document IDs', 'collectDocumentIDs')
        .addSeparator()
        .addItem('Instructions', 'showInstructions');
  } else {
    menu.addItem('Activate Script', 'activateScript');
  }
  
  menu.addToUi();
}

function isScriptActivated() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('SPREADSHEET_ID') !== null;
}

function activateScript() {
  var ui = SpreadsheetApp.getUi();
  SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
  FOLDER_IDS = {};
  PropertiesService.getScriptProperties().setProperty('FOLDER_IDS', JSON.stringify(FOLDER_IDS));
  ui.alert('Script Activated', 'The script has been successfully activated for this spreadsheet. The menu will update after you click "OK".', ui.ButtonSet.OK);
  
  // Trigger a reload of the spreadsheet to update the menu
  SpreadsheetApp.getActiveSpreadsheet().updateMenu();
}

function linkSchoologyFolder() {
  if (!isScriptActivated()) {
    showActivationAlert();
    return;
  }
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Link Schoology folder from Drive',
    'On Schoology, copy the exact title of your course [i.e. CLASS TITLE A (S1) Last, F CLASS TITLE A(2)]\n' +
    'In Google Drive, paste the course title from Schoology.\n' +
    'Right click on the folder, select \'Share\', and choose \'Copy Link\'. Then paste that link below:',
    ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    var folderUrl = result.getResponseText().trim();
    var folderId = extractFolderIdFromUrl(folderUrl);
    if (folderId) {
      var folder = DriveApp.getFolderById(folderId);
      var folderName = folder.getName();
      createOrUpdateSheet(folderName, folderId);
      ui.alert('Success', 'Folder linked successfully: ' + folderName, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Invalid folder URL. Please try again.', ui.ButtonSet.OK);
    }
  }
}

function extractFolderIdFromUrl(url) {
  var match = url.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

function createOrUpdateSheet(sheetName, folderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange('A1:C1').setValues([['Name', 'Assignment Title', 'Document ID']]);
  }
  FOLDER_IDS[sheetName] = folderId;
  PropertiesService.getScriptProperties().setProperty('FOLDER_IDS', JSON.stringify(FOLDER_IDS));
}

function collectDocumentIDs() {
  if (!isScriptActivated()) {
    showActivationAlert();
    return;
  }
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Collect document IDs',
    'Enter the name of the assignment on Schoology (without any numbers at the end):',
    ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() == ui.Button.OK) {
    var assignmentName = result.getResponseText().trim();
    processAllSheets(assignmentName);
  }
}

function processAllSheets(assignmentName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var folderIds = JSON.parse(PropertiesService.getScriptProperties().getProperty('FOLDER_IDS') || '{}');

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    if (folderIds[sheetName]) {
      processSheet(sheet, folderIds[sheetName], assignmentName);
    }
  }

  SpreadsheetApp.getUi().alert('Document ID collection complete');
}

function processSheet(sheet, folderId, assignmentName) {
  var folder = DriveApp.getFolderById(folderId);
  var subfolders = folder.getFolders();
  var matchingFolder = null;
  
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var subfolderName = subfolder.getName();
    if (subfolderName.startsWith(assignmentName)) {
      matchingFolder = subfolder;
      break;
    }
  }
  
  if (matchingFolder) {
    var files = matchingFolder.getFiles();
    var data = [];

    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      var nameParts = fileName.split(' ');
      if (nameParts.length < 2) continue;

      var studentName = nameParts[0] + ' ' + nameParts[1];
      var docId = file.getId();
      data.push([studentName, assignmentName, docId]);

      if (data.length == 10) {
        appendDataToSheet(sheet, data);
        data = [];
        Utilities.sleep(1000); // Sleep for 1 second to avoid rate limiting
      }
    }

    if (data.length > 0) {
      appendDataToSheet(sheet, data);
    }
    
    SpreadsheetApp.getUi().alert('Success', 'Document IDs collected for assignment: ' + assignmentName, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('Error', 'Subfolder not found for assignment: ' + assignmentName + ' in folder: ' + folder.getName(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function appendDataToSheet(sheet, data) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(lastRow + 1, 1, data.length, 3);
  range.setValues(data);
}

function showInstructions() {
  var ui = SpreadsheetApp.getUi();
  var instructions = 
    'Instructions for using Teacher Tools:\n\n' +
    '1. Link Schoology folder from Drive:\n' +
    '   - Copy the exact title of your course from Schoology\n' +
    '   - Find the corresponding folder in Google Drive\n' +
    '   - Right-click the folder, select "Share", and copy the link\n' +
    '   - Paste the link when prompted by the script\n\n' +
    '2. Collect document IDs:\n' +
    '   - Enter the exact name of the assignment as it appears in Schoology\n' +
    '   - The script will find the corresponding folder in Drive and extract document IDs\n' +
    '   - Results will be added to the appropriate sheet in this spreadsheet';
  
  ui.alert('Teacher Tools Instructions', instructions, ui.ButtonSet.OK);
}

function showActivationAlert() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Script Not Activated', 'Please use the "Activate Script" option in the Teacher Tools menu first.', ui.ButtonSet.OK);
}