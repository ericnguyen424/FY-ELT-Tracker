/**
 * @fileoverview averages.gs
 *  This script is used to calculate averages for the Census / Full % into the Stats Table. 
 * @author Eric Nguyen, eric.nguyen.kt424@gmail.com / enguyen@odysseyhouse.org
 * @lastmodified 7/25/25
 */

function onEdit(e) {
  if (validateSheet("Stats") && checkUserAccess())
  {
    calculateAverages();
  }
}

/**
 * When called, this function will calculate the averages based on the current month and update the table to reflect those averages
*/

function calculateAverages() {
  try {
  const statSheet = getSheetFromName("Stats");  
  const date = getDateRange();

  // const interval = getCurrentInterval();
  const programs = getPrograms();

  if (statSheet.getLastRow() - 2 != programs.length) {
    resetPrograms();
  }

  const tableData = statSheet.getRange(3, 2, programs.length, 2).getValues();
  updateData(tableData, date);
  statSheet.getRange(3, 2, programs.length, 2).setValues(tableData);
  SpreadsheetApp.getActiveSpreadsheet().toast("Averages updated.", "Success!", 3);
  }
  catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(err.message, "Error!", 3);

  }
} 

/**
 * Returns an object dict containing census and full sum and tally of every program found within the date range
 * For example: 
 * {
 *  "Martindale": [[row1], [row2], row3, etc....]
 * }
 * date - Date interval
 */
function programDataByDate(date) 
{
  const dataSheet = getSheetFromName(getDataSheetName());
  // Collect all data (assuming data starts at row 2 and headers in row 1)
  let data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
  let censusIndex = getColIndexOf("Census");
  let fullIndex = getColIndexOf("Full");
  // Use the reduce function to collect rows by date sorted into a dict with the program names as keys
  return data.reduce(
    ((filteredData, currentRow) => {

      // verify date
      let dateValue = currentRow[1];
      if (dateValue instanceof Date && dateMatches(dateValue, date)) 
      {
        let program = currentRow[0];

        // add program to dict if not present
        if (!(program in filteredData)) 
        {
          filteredData[program] = {full: currentRow[fullIndex], census: 0, count: 0};
          

        }

        // If census is a number, add and tally
        if (typeof currentRow[censusIndex] === 'number' && !(isNaN(currentRow[censusIndex]))) {
          filteredData[program].census += currentRow[censusIndex];
          filteredData[program].count++;
        }

      }
      return filteredData;
      
    }), {}
  )
}

/**
 * Checks if cellDate matches with given month
 * cellDate - Date object of cell
 * Date - Date object to compare with
 * returns true if matching, false otherwise
 */
function dateMatches(cellDate, date) {
  return cellDate.getMonth() === date.getMonth() && cellDate.getFullYear() === date.getFullYear();
}

/**
 * Given the table 2d array and date, will return a new 2D array containing all [programName, monthlyAverage]  pairs related to date
 * tableArray - Table 2d array
 * date - Date object 
 */
function updateData(tableArray, date) {

  // Get data related to date and program
  const data = programDataByDate(date);
  if (!data) {
    throw new Error("Accessing date that doesn't exist in data!");
  }
  // loop through every program in tableArray
  tableArray.forEach(tableRow => {
    let currentProgram = tableRow[0];
      if (!(currentProgram in data)) {
        tableRow[1] = "Not found in data!";
        return;
    }
    let currentData = data[currentProgram];

    // update 2nd index of tableArray to be average
    if (currentData.count == 0 || currentData.full == 0) 
    {
        tableRow[1] = "N/A";
    }
    else 
    {
        tableRow[1] = ((((currentData.census / currentData.count) / currentData.full) * 100).toFixed(2) + "%");
    }});
}

/**
 * Populate every program into the table
 */
function resetPrograms() {
  const statSheet = getSheetFromName("Stats");
  let programs = Array.from(getPrograms());

  console.log(statSheet.getRange(3, 2, programs.length, 1).getDisplayValues());
  statSheet.getRange(3, 2, programs.length, 1) // row 3, col 2 (B3), N rows, 1 col
         .setValues(programs.map(program => [program]));
}



