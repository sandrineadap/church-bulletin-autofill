/**
 * Calculate sunset time at the church's location.
 *  
 * @param   {Object} sabbathDate  Date of sunset time to calculate.
 * 
 * @returns {string} Short form of time, e.g. 6:32 PM
 */
function calcSunset(sabbathDate) {
  let sunset = getSunset(43.82577110739776, -79.50532049120118, sabbathDate); // latitude and longitude of the church
  sunset = sunset.toLocaleTimeString([], {timeStyle: 'short'})

  return sunset;
}