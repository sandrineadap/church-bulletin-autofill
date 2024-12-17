/*
 *  Create Bulletin from an active (highlighted) row.
 *  Creates a new Google doc file from a template based on the ProgramType column.
 * 
 */
function createBulletinFromRow() {

  const destinationFolder = DriveApp.getFolderById(FOLDERID)
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule');
  const rows = sheet.getDataRange().getValues();
  const namedRanges = SpreadsheetApp.getActive().getNamedRanges();

  const highlighted = sheet.getActiveRange().getValues();
  const week = highlighted[0]
  const weekIndex = Math.floor(sheet.getActiveRange().getRowIndex()) - 1
  
  dataToBulletin(destinationFolder, rows, namedRanges, week, weekIndex);

}