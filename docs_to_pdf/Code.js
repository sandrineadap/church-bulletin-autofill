/**
 * Google Doc to PDF in Google Drive.
 * By Sandrine Adap
 * 
 * Creates a PDF of the current Google Doc and saves it to the parent Google Drive folder 
 * without needing to download it locally.
 * 
 */

function onOpen() {
  const ui = DocumentApp.getUi();
  const menu = ui.createMenu('Create PDF');
  menu.addItem('Save PDF to Google Drive', 'createPDF')
  menu.addToUi();
}

/**
 * Creates a PDF file of the current document.
 */
function createPDF() {
  const doc = DocumentApp.getActiveDocument();
  const file = DriveApp.getFileById(doc.getId());
  const folderId = file.getParents().next().getId();  // gets the parent folder of the current doc
  const folder = DriveApp.getFolderById(folderId);

  const pdf = doc.getAs('application/pdf');
  pdf.setName(doc.getName());
  folder.createFile(pdf);  
}