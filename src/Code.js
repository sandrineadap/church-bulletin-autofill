/**
 * Bulletin Autofill Script. 
 * By Sandrine Adap.
 * 
 * IMPORTANT !!!!!!!!!!!!!!!
 * If you modify FOLDERID, you must manually run updateHyperlink() at least once to update the Instruction sheet.
 */
const FOLDERID = PropertiesService.getScriptProperties().getProperty('folder_id')  // if you modify the shared google drive folder, run updateHyperlink() !!

// Google Doc ID's of templates

const TRADITIONALID = PropertiesService.getScriptProperties().getProperty('traditional_id')
const CREATIVEID = PropertiesService.getScriptProperties().getProperty('creative_id')
const COMMUNIONID = PropertiesService.getScriptProperties().getProperty('communion_id')


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Create Bulletin');
  menu.addItem('Create Next Bulletin', 'createBulletin')
  // menu.addItem('Create Bulletin from selected row', 'createBulletinFromRow'); // for testing purposes only
  menu.addToUi();
}

/**
 * Updates the "Bulletins Folder Link" hyperlink on the Instructions sheet.
 * 
 * Must be run manually when FOLDERID is modified.
 */
function updateHyperlink() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Instructions');
  const url = "https://drive.google.com/drive/folders/" + FOLDERID
  const hyperlink = `=HYPERLINK("${url}", "Bulletins Folder Link")`
  sheet.getRange(4, 2).setValue(hyperlink)
}

/**
 * Create Bulletin for next Sabbath after today.
 * Calls dataToBulletin()
 * 
 */
function createBulletin() {
  const destinationFolder = DriveApp.getFolderById(FOLDERID)
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schedule');
  const rows = sheet.getDataRange().getValues();
  const namedRanges = SpreadsheetApp.getActive().getNamedRanges();


  // get this week's index

  nextSabbath = getNextSabbath();
  // Logger.log("nextSabbath = " + nextSabbath)
  let weekIndex;

  for (let i = 1; i < rows.length; i++){  // start on 1 to skip header row
    if (rows[i][0].toLocaleDateString() == nextSabbath) {
      weekIndex = i;
    }
  }

  const week = rows[weekIndex];

  dataToBulletin(destinationFolder, rows, namedRanges, week, weekIndex);

}





