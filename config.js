/**
 * @fileoverview config.gs
 *  This script is used to retrieve data from the config sheets. 
 * @author Eric Nguyen, eric.nguyen.kt424@gmail.com / enguyen@odysseyhouse.org
 * @lastmodified 7/24/25
 */

/**
 * Returns the list of columns to copy over
 */
function getColsToCopy() {
  const config = getSheetFromName("_programlist");
  const data = config.getDataRange().getValues().filter((items, index) => (index == 0 || index == 2)).map(row => row.slice(2));

  return data[0].filter((item, index) => data[1][index]);
}

/**
 * Returns a list of the columns from which to calculate totals of
 */
function getColsToTotal() {
  const config = getSheetFromName("_programlist");
  const data = config.getDataRange().getValues().filter((items, index) => (index == 0 || index == 1)).map(row => row.slice(1));
  return data[0].filter((item, index) => data[1][index]);
}

/**
 * Returns a list of every program in the order given in _programlist
 */
function getPrograms() {
  const config = getSheetFromName("_programlist");
  return config.getDataRange().getValues().map(row => row[0]).filter(column => !(['Program List',
  '1Calculate Totals?',
  '2Copy into next week?',].includes(column)));
}

/**
 * Returns the name of the data sheet. For example, "FY26 Tracker"
 */
function getDataSheetName() {
  const sheetConfig = getSheetFromName("_sheetconfig");
  const data = sheetConfig.getDataRange().getValues()[1];
  return data[0];
}

/**
 * This function will return all the data entry types listed in the _programlist sheet as an array.
 */
function getValues() {
  const config = getSheetFromName("_programlist");
  const headers = new Set(config.getRange(1, 1, 1, config.getLastColumn()).getValues()[0]);
  headers.delete("Program List");
  return [...headers];
}

function test() {
  console.log(getColsToTotal());
}

