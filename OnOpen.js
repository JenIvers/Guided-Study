// OnOpen.gs
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Admin Tools');
 
  // Check if sheet is activated
  const isActivated = isSheetActivated();
 
  // Add activation option only if not activated
  if (!isActivated) {
    menu.addItem('Activate Sheet', 'activateSheet');
  }
 
  // Add other menu items
  menu.addItem('Collect Student Daily Work', 'collectStudentDailyWork')
      .addItem('Send Teacher Feedback', 'sendTeacherFeedback')
      .addItem('Collect Student Follow-ups', 'collectStudentFollowUps');
 
  menu.addToUi();
}
