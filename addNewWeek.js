/**
 * @fileoverview addNewWeek.gs
 *  This script automates the insertion of new weeks into the ELT Weekly tracker. 
 *  A trigger will call newWeek() every Sunday to add the newest week. A row for every program listed in _config will be added *  7 days forward.
 * @author Eric Nguyen, eric.nguyen.kt424@gmail.com / enguyen@odysseyhouse.org
 * @lastmodified 7/24/25
 * 
 * 
 */

// TODO: Handle the case in which the first week is being added

function newWeek() {
  const dataSheetName = getDataSheetName();
  const sheet = getSheetFromName(dataSheetName);
  try {
    addNewWeek(sheet);
  }
  catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(err.message, "Error!", 3);
    Logger.log("Error occured when adding new sheet! Undoing...\n" + err.message);
  }
}

/**
 * Given the sheet to update, will insert a row for every program listed in _config with the updated  * week, and the correct format copied over. 
 */
function addNewWeek(sheet) {
  try {
    if (!sheet) 
    {
      throw new Error("Error accessing sheet!");
    }

    var backupData = sheet.getDataRange().getValues();

    const oldRange = getOldRange(sheet);

    // insert the new range, then store it in a variable
    const newRange = sheet.insertRowsAfter(oldRange.getLastRow(), getPrograms().length + 1)
      .getRange(oldRange.getLastRow() + 1, 1, getPrograms().length + 1, oldRange.getLastColumn());

    // Prevent Google's "Smart Fill"
    preventSmartFill(newRange);

    // Map of data for ranges
    let newValues = newRange.getValues();
    let oldValues = oldRange.getValues();

    extendFormat(oldRange, newRange, oldValues, newValues, getColsToCopy());
    updateWeekColumn(oldValues, newValues);
    newRange.setValues(newValues);
    calculateTotals(sheet, newRange, getColsToTotal());

  }
  catch (err) 
  {
    if (sheet.getLastRow() != backupData.length) {
      console.log("Undoing!");
      sheet.deleteRows(backupData.length + 1, (sheet.getLastRow() - backupData.length))
    }
    sheet.getDataRange().setValues(backupData);

    throw new Error("Error!\n" + err);
  }
}

/**
 * Copies over the format of previous rows to the new rows
 * oldRange, newRange - ranges to copy from and to
 * oldValues, newValues - arrays of data values
 * colsToCopy - List of columns to copy over into the new range. Example: ["Full", "Cap"].
 */
function extendFormat(oldRange, newRange, oldValues, newValues, colsToCopy) {

  // Set borders of new rows
  newRange.setBorder(true, true, true, true, true, true, null, null);

  let newPrograms = getPrograms();
  let programIndex = getColIndexOf("Program");
  let columnIndexes = colsToCopy.map(column => getColIndexOf(column));
  let priorValues = {};
  oldValues.forEach(row => priorValues[row[0]] = row);

  // Insert the new programs (skip the "Totals" row) and copy over values if applicable
  newPrograms.forEach((program, index) => {
    {
      newValues[index][programIndex] = program;

      // Check if value existed before, if so copy, otherwise, leave empty.
      columnIndexes.forEach(col => {
        if (program in priorValues) {
          newValues[index][col] = priorValues[program][col];
        }
      })
    }});

  newValues[newValues.length - 1][programIndex] = "Total";
}

/**
 * Applies the sum formula to the totals row given the starting indexes
 * newRange - 
 * columns - 
 */
function calculateTotals(sheet, newRange, columns) {

  let columnIndexes = columns.map(column => (getColIndexOf(column) + 1));

  columnIndexes.forEach(col => {
    let colLetter = String.fromCharCode(64 + col);

    // Construct the formula string for this column
    const formula = `=SUM(${colLetter}${newRange.getRow()}:${colLetter}${newRange.getLastRow() - 1})`;

    // Set the formula in the correct cell in the total row
    sheet.getRange(newRange.getLastRow(), col).setFormula(formula);
  
  })

}


/**
 * When called, this function will update the given range 1 week forward
 * oldValues, newValues - arrays of the old & new ranges
 */

function updateWeekColumn(oldValues, newValues) {

  // Retrieve and validate the previous week value
  let previousWeekDate = oldValues[0][getColIndexOf("Week")];
  if (!previousWeekDate || isNaN(previousWeekDate.getTime())) 
  {
    throw new Error("Invalid or missing date in the last week's row.");
  }

  let newWeekDate = new Date(previousWeekDate);
  newWeekDate.setDate(newWeekDate.getDate() + 7);
  
  // Update the weeks
  newValues.forEach(row => row[getColIndexOf("Week")] = newWeekDate);  
}

/**
 * Used to clear Google's Smart Fill
 */
function preventSmartFill(newRange) {
  const blankFormulas = Array(newRange.getNumRows()).fill().map(() =>
    Array(newRange.getNumColumns()).fill("")
  );
  newRange.setFormulas(blankFormulas);
}

/**
 * Returns the range of last week
 */
function getOldRange(sheet) {
  const lastDate = sheet.getRange(sheet.getLastRow(), 2).getValue();
  const data = sheet.getDataRange().getValues();

  const rowCount = data.filter(row => 
    {
      let weekCell = row[1];
      return weekCell instanceof Date && weekCell.getTime() === lastDate.getTime();
    }).length;

  const startingRow = sheet.getLastRow() - rowCount + 1;
  return sheet.getRange(startingRow, 1, rowCount, sheet.getLastColumn());
}