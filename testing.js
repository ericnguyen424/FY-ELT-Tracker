/**
 * @fileoverview testing.gs
 *  This script is used to test the functioning of the AppScript program.
 * @author Eric Nguyen, eric.nguyen.kt424@gmail.com / enguyen@odysseyhouse.org
 * @lastmodified 7/22/25
 */

function runTests() {
  test_addNewWeek_testingSheet();
}

function testingSheet() {
  addNewWeek(getSheetFromName("TESTING COPY"));
}

/**
 * Will create and return a testing sheet identical to the statSheet.
 */
function createTestingSheet(ss) {
  if (ss.getSheetByName("TESTING COPY")) {
    ss.deleteSheet(ss.getSheetByName("TESTING COPY"));
  }
  const dataSheet = ss.getSheetByName(getDataSheetName());

  // Copy the sheet as a new one at the end of the spreadsheet
  const copySheet = dataSheet.copyTo(ss);
  copySheet.setName("TESTING COPY"); 

  // Move the new sheet to the end
  ss.setActiveSheet(copySheet);
  return copySheet;

}

// Helper to build the test suite array
function buildTestSuite(tests) {
  return tests.map(fn => ({
    name: fn.name,
    description: fn.description || "Testing Description N/A.",
    fn
  }));
}

// Testing Suite for dataSheet_functions.gs
function test_addNewWeek_testingSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testingSheet = createTestingSheet(ss);

  // TODO: Test the retrieval of datasheet name

  addNewWeek(testingSheet);
  const values = testingSheet.getDataRange().getValues();
  let newRange = {start: testingSheet.getLastRow() - getPrograms().length - 1, end: testingSheet.getLastRow() - 1};
  let priorRange = {start: ((newRange.start - 1) - getPrograms().length), end: newRange.start - 1};

  // testSuite for addNewWeek
  // every item contains the test's name, description, and the function.
  const testSuite = buildTestSuite([
    checkPrograms,
    weekUpdated,
    checkEmptyData,
    checkRowsAndColCount,
    sumFormulaCorrectness,
  ]);

  let numTests = 0;
  let numPassed = 0;
  // Call every test in the testing suite and log the result.
  testSuite.forEach(({ name, description, fn }) => {
    numTests++;
    try {
      // sheet, valuesArray, new & prior ranges as dict objects containing 0-based indexes -> {start: x, end: y}
      fn(testingSheet, values, newRange, priorRange);
      numPassed++;
    } catch (err) {
      console.error(`${name}: ${description}\nTEST FAILED, ERR MESSAGE: ${err.message}`);
    }
  });

  console.log(`test_addNewWeek_testingSheet passed ${numPassed}/${numTests} tests.`);
  ss.deleteSheet(testingSheet);
}

// ------ TESTING FUNCTIONS HERE ------

// Tests if all the programs expected were added into the new week
function checkPrograms(sheet, values, newRange, priorRange) {
  const expectedPrograms = new Set(getPrograms());
  expectedPrograms.add("Total");
  const programCol = getColIndexOf("Program");

  // Loop through every row in values. If the program is invalid, throw an error. Otherwise, remove it from expectedPrograms.
  for (let row = newRange.start; row <= newRange.end; row++) 
  {
    var currentProgram = values[row][programCol];
    if (expectedPrograms.has(currentProgram)) 
    {
      expectedPrograms.delete(currentProgram);
    }
    else 
    {
      throw new Error(`Encountered the unexpected program "${currentProgram}" in the new week!`)
    }
  }

  // If the size of expectedPrograms is 0, then every program was accounted for. Otherwise, throw an error.
  if (expectedPrograms.size != 0) {
    throw new Error(`These expected programs weren't found in the new week: ${[...expectedPrograms]}`);
  }
}
checkPrograms.description = "Tests if all the programs expected were added into the new week";

// Tests to see if the new rows have new weeks
function weekUpdated(sheet, values, newRange, priorRange) {
  let weekCol = getColIndexOf("Week");
  let priorWeek = values[priorRange.start][weekCol];

  if (!priorWeek || isNaN(priorWeek.getTime())) 
  {
    throw new Error(`Invalid date in prior week's data at row ${priorRange.start + 1}`);
  }

  // Loop through prior week and validate
  for (let row = priorRange.start; row < newRange.start; row++) {
    let currentRow = values[row][weekCol];
    if (!currentRow || isNaN(currentRow.getTime())) {
      throw new Error(`Invalid date of ${currentRow} in prior weeks data at ${row}`);
    }

    if (currentRow.getDate() !== priorWeek.getDate()) {
      throw new Error(`Mismatched date at row ${row}. Expected ${priorWeek.getDate()} but found ${currentRow.getDate()}`);
    }

  }

  let expectedNewWeek = new Date(priorWeek);
  expectedNewWeek.setDate(expectedNewWeek.getDate() + 7);

  // Loop through current week and validate
  for (let row = newRange.start; row <= newRange.end; row++) {
    let currentRow = values[row][weekCol];
    if (!currentRow || isNaN(currentRow.getTime())) {
      throw new Error(`Invalid date of ${currentRow} in new weeks data at ${row}`);
    }

    if (currentRow.getDate() !== expectedNewWeek.getDate()) {
      throw new Error(`Mismatched date at row ${row}. Expected ${expectedNewWeek.getDate()} but found ${currentRow.getDate()}`);
    }
  }
}
weekUpdated.description = "Tests to see if the new rows have new weeks";

function checkEmptyData(sheet, values, newRange, priorRange) {
  let columnsToCheck = getValues();
  columnsToCheck = columnsToCheck.filter(column => !(["Week", "Program", "Full", "Cap"].includes(column)));  // These columns are allowed to have copied data
  
  columnsToCheck = columnsToCheck.map(column => getColIndexOf(column));

  // Loop through every col and check the totals row for invalid data (total != 0)
  columnsToCheck.forEach(col => {
      if (values[newRange.end][col] != 0)
      {
        throw new Error(`Unexpected data ${values[newRange.end][col]} tallied in the "Totals" row (index ${newRange.end}) in the "${getColHeader(col)}" column`);
      }
      });
}
checkEmptyData.description = "Tests to see if the new rows didn't copy over any unwanted data";

function checkRowsAndColCount(sheet, values, newRange, priorRange) {
  let expectedNewRowCount = getPrograms().length + 1;
  let actualNewRowCount = (newRange.end - newRange.start) + 1;
  if (actualNewRowCount != expectedNewRowCount) {
    throw new Error(`Invalid number of rows detected. Expected ${expectedNewRowCount} rows but counted ${actualNewRowCount} rows.`)
  }
  else if (values[newRange.start].length != getValues().length + 1) {
    throw new Error(`Invalid number of cols detected. Expected ${getValues().length + 1} cols but counted ${values[newRange.start].length} cols.`)
  }
}
checkRowsAndColCount.description = "Checks to see if the new row and column counts match the expected outcome";

function sumFormulaCorrectness(sheet, values, newRange, priorRange) {

  const lastRow = sheet.getRange(sheet.getLastRow(), 3, 1, 6);
  let expectedFormulas = [2, 3, 4, 5, 6, 7].map(colIndex => `=SUM(${String.fromCharCode(65 + colIndex)}${newRange.start + 1}:${String.fromCharCode(65 + colIndex)}${newRange.end})`);

  const expectedFormulaSet = new Set(expectedFormulas);

  lastRow.getFormulas()[0].forEach(formula => {
    if (!(expectedFormulaSet.has(formula))) {
      throw new Error(`Unexpected formula ${formula} encountered in new row!`);
    }
    expectedFormulaSet.delete(formula);
  })

  if (expectedFormulaSet.size != 0) {
      throw new Error(`New row is missing formulas!`);
  }

}
sumFormulaCorrectness.description = "Verifies the sum formula points to the correct ranges";

