/**
 * Make bulletin from given data.
 * Creates a new Google doc file from a template based on the ProgramType column.
 * Saves the new doc to a specified Google Drive folder.
 * 
 * @param   {Object}  destinationFolder Google Drive folder where to save doc.
 * @param   {Object}  rows              Which rows in the sheet.
 * @param   {Object}  namedRanges       Named columns in active sheet.
 * @param   {Object}  week              Array of participants for this week (singular row).
 * @param   {int}     weekIndex         The index of the week row.
 * 
 */
function dataToBulletin(destinationFolder, rows, namedRanges, week, weekIndex){
  // determine what template to use

  const SabbathDate = week[0]
  const ProgramType = week[1]
  let bulletinTemplate

  if (ProgramType == 'Traditional'){
    bulletinTemplate = DriveApp.getFileById(TRADITIONALID)
  }
  else if (ProgramType == 'Creative'){
    bulletinTemplate = DriveApp.getFileById(CREATIVEID)
  }
  else if (ProgramType == 'Communion'){
    bulletinTemplate = DriveApp.getFileById(COMMUNIONID)
  }
  else {
    Logger.log("No valid week row selected.")
    return;
  }

  // get shortened date for file name

  const sabbathDateShort = SabbathDate.toLocaleString('default', { month: 'short' }) + "-" + SabbathDate.getDate().toString() + "-" + SabbathDate.getFullYear().toString()


  // get long date for bulletin

  const sabbathDateLong = SabbathDate.toLocaleString('default', { month: 'long' }) + " " + SabbathDate.getDate().toString() + ", " + SabbathDate.getFullYear().toString()


  // create a copy of the selected template

  const copy = bulletinTemplate.makeCopy(`Bulletin${ProgramType}_${sabbathDateShort}`, destinationFolder);
  const doc =  DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  const header = doc.getHeader();


  // replace date tag in the header
  first_header = header.getParent().getChild(3);
  first_header.replaceText("{{SabbathDate}}", sabbathDateLong)


  // replace tags in template based on range names in sheet

  const mainSchedule = new Set(["DWPresidingElder", "SSSuperintendent", "DWDeacons", "SSUshers", "DWWorshipCoordinator", "SSChorister", "DWChorister", "DWPianist", "LaptopOperators", "AudioOperator", "ChildrenStory", "FamilyPrayer"]);

  for (let i = 0; i < namedRanges.length; i++) {
    let col = namedRanges[i].getRange().getColumn() - 1;
    body.replaceText(`{{${namedRanges[i].getName()}}}`, week[col]);
    

    // replace tags in schedule

    if (mainSchedule.has(namedRanges[i].getName())) {
      for (let j = 1; j <= 3; j++){
        body.replaceText(`{{${namedRanges[i].getName()}_${j}}}`, rows[weekIndex+j][col]);
      } 
    }

  }

  // replace dates in schedule

  for (let i = 1; i <= 3; i++){
    let nextDate = rows[weekIndex+i][0]
    nextDate = nextDate.toLocaleString('default', { month: 'short' }) + ". " + nextDate.getDate().toString()

    body.replaceText(`{{SabbathDate_${i}}}`, nextDate);
  } 


  // compute sunset times

  Sunset = calcSunset(SabbathDate);  // this week
  Sunset_1 = calcSunset(rows[weekIndex + 1][0]); // next week

  body.replaceText("{{Sunset}}", Sunset);
  body.replaceText("{{Sunset_1}}", Sunset_1);


  doc.saveAndClose();
}