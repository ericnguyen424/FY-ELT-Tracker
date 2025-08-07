/**
 * @fileoverview helper_functions.gs
 *  This script contains helper functions used throughout the program 
 * @author Eric Nguyen, eric.nguyen.kt424@gmail.com / enguyen@odysseyhouse.org
 * @lastmodified 7/25/25
 */

/**
 * Used in stat calculation. Returns the date object in the date column
 */
function getDateRange() {
  const statSheet = getSheetFromName("Stats");
  const dateCell = statSheet.getRange("B1");

  return dateCell.getValue();
}

/**
 * Used in stat calculation. Returns the value in the dropdown for values.
 */
function getValueType() {
  const statSheet = getSheetFromName("Stats");
  const valueTypeCell = statSheet.getRange("C1");

  return valueTypeCell.getValue();
}

/**
 * Gives the 0 based index of the requested column.
 * columnName - string of requested column name
 * throws an error if not found
 */
function getColIndexOf(columnName) {
  const sheet = getSheetFromName(getDataSheetName());

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // first row
  const colIndex = headers.indexOf(columnName);

  if (colIndex === -1) {
    throw new Error(`Column "${columnName}" not found`);
  }

  return colIndex;
}

/**
 * Assumes sheet to retrieve from is dataSheet. Given a 0-based index, will return the name of the column header.
 */
function getColHeader(colIndex) {
  const dataSheet = getSheetFromName(getDataSheetName());
  const headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues();
  return headers[0][colIndex];
}

/**
 * Will return true if the user has sufficient permissions 
 * Conditions: user is owner (Seth Mower) or script editor (Eric Nguyen)
 */
function checkUserAccess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const owner = ss.getOwner();
  const activeUser = Session.getActiveUser();
  
  const ownerEmail = owner ? owner.getEmail() : "Unknown";
  const userEmail = activeUser.getEmail();

  const editor = "enguyen@odysseyhouse.org";

  if (userEmail === ownerEmail || userEmail === editor) {
    return true;
  } 
  else {
    return false;
  }
}

/**
 * Returns the active spreadsheet
 */
function getCurrentSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

/**
 * Returns the sheet object if existing
 * sheetName - name of sheet as a string
 */
function getSheetFromName(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(sheetName);
}

/**
 * Converts the given letter into index
 * A -> 1
 * Z - 26
 * AA - 27
 */
function columnLetterToIndex(letter) {
  letter = letter.toUpperCase();
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index *= 26;
    index += letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return index;
}

/**
 * Displays to the screen. For testing
 */
function alert(message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(message, ui.ButtonSet.OK);
}

/**
 * This function adds cooldowns to button clicks. 
 * delay - amount of time to delay in seconds
 * returns true if the cooldown has passed, false otherwise.
 */
function checkButtonCooldown(delay) {
  const props = PropertiesService.getUserProperties();
  const lastClick = Number(props.getProperty("lastClick")) || 0;
  const now = Date.now();

  if (now - lastClick < (delay * 1000)) {
    const timeElapsed = Math.floor((now - lastClick) / 1000);
    const timeRemaining = delay - timeElapsed;
    SpreadsheetApp.getActiveSpreadsheet().toast(("Please wait " +  timeRemaining + "s before clicking again."), "Error!", 5)
    return false;
  }

  // Update timestamp
  props.setProperty("lastClick", now);

  return true;
}

/**
 * Compares the current sheet and checks if it is the given sheet
 * sheetName - string of other sheet
 */
function validateSheet(sheetName) {
  return (getCurrentSheet().getName() === sheetName);
} 