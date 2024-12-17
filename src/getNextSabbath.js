/**
 * Gets the next Sabbath after today.
 * If today is a Saturday, gets the next one.
 * 
 * @returns  {string} LocaleDateString of next Saturday.
 */
function getNextSabbath() {
  let today = new Date();
  if (today.getDay() == 6) {
    today.setDate(today.getDate() + 7)
  }
  else { 
    today.setDate(today.getDate() + (6-today.getDay()))
  }
  
  Logger.log(today.toLocaleDateString())
  return today.toLocaleDateString()
}