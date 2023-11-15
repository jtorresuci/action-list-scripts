const TASK_SHEET_NAME = "Tasks List"
const BETA_LIST = "Beta"

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('GSOC: Utilities')
      .addItem('Move Completed Tasks', 'moveCompletedTasksv2')
      .addToUi();
}

function sortColumnF(sheet) {
  // Find the last row with data in column F (assuming column F is the column to sort)
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 4) {
    Logger.log("No data to sort below row 3.");
    return;
  }
  
  var rangeToSort = sheet.getRange(3, 1, lastRow - 3, sheet.getLastColumn()); // Assuming you want to sort all columns
  rangeToSort.sort({column: 6, ascending: true}); // Sort by column F (index 6)
}



function moveCompletedTasks2() {
  var sourceSheetName = BETA_LIST; // Replace with the name of your source sheet
  var targetSheetName = "Completed Tasks"; // Replace with the name of your target sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var targetSheet = ss.getSheetByName(targetSheetName);

  if (!sourceSheet || !targetSheet) {
    Logger.log("Source or target sheet not found.");
    return;
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  var newData = [];

  for (var i = 2; i < sourceData.length; i++) { // Start from row 3 (index 2)
    var row = sourceData[i];
    if (row[7] && row[7].toLowerCase() === "Completed") { // Column H (index 7) contains "completed"
      targetSheet.appendRow(row);
      Logger.log("Moved row to 'Completed Tasks': Row " + (i + 1));
    } else {
      newData.push(row);
    }
  }

  // Clear the source sheet starting from row 3 and set new data
  sourceSheet.getRange(3, 1, sourceData.length - 2, sourceData[0].length).clearContent();
  if (newData.length > 0) {
    sourceSheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);
    Logger.log("Cleared source sheet from row 3 onwards and set new data.");
  }
}

function duplicateSheet(sourceSheet, spreadsheet) {
  var newSheetName = "FAILSAFE LIST"; // Change to the desired name for the duplicated sheet

  // Check if a sheet with the same name exists
  var existingSheet = spreadsheet.getSheetByName(newSheetName);
  if (existingSheet) {
    // Delete the existing sheet
    spreadsheet.deleteSheet(existingSheet);
  }

  // Duplicate the source sheet
  var newSheet = sourceSheet.copyTo(spreadsheet);
  newSheet.setName(newSheetName);

  // Move the duplicated sheet to the right of the source sheet
  var sheets = spreadsheet.getSheets();
  var sourceSheetIndex = sheets.indexOf(sourceSheet);
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(sheets.length);
}


function deleteFailsafeCopySheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNameToDelete = "FAILSAFE LIST"; // Change to the name of the sheet you want to delete

  var sheetToDelete = spreadsheet.getSheetByName(sheetNameToDelete);

  if (sheetToDelete) {
    spreadsheet.deleteSheet(sheetToDelete);
    console.log('Sheet "' + sheetNameToDelete + '" has been deleted.');
  } else {
    console.log('Sheet "' + sheetNameToDelete + '" does not exist.');
  }
}

function flipDateFormat(sheet) {


  // Get the data in columns E and F
  var data = sheet.getRange("E:F").getValues();

  // Loop through the data and flip the date format for all strings
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (typeof data[i][j] === "string") {
        // Split the string assuming it's in the format "dd/mm/yy" and flip it
        var parts = data[i][j].split("/");
        if (parts.length === 3) {
          data[i][j] = parts[1] + "/" + parts[0] + "/" + parts[2];
        }
      }
    }
  }

  // Update the values in columns E and F with the flipped date format
  sheet.getRange("E:F").setValues(data);

}




function moveCompletedTasksv1() {
  logFunctionStart()
  var sourceSheetName = BETA_LIST; // Replace with the name of your source sheet
  var targetSheetName = "Completed Tasks"; // Replace with the name of your target sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var targetSheet = ss.getSheetByName(targetSheetName);

  if (!sourceSheet || !targetSheet) {
    Logger.log("Source or target sheet not found.");
    return;
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  var newData = [];
  var count = 0;
  var action = 0;

  duplicateSheet(sourceSheet, ss);

  for (var i = 2; i < sourceData.length; i++) { // Start from row 3 (index 2)
    count+=1
    var row = sourceData[i];
    if (row[7] && row[7].toLowerCase() === "completed") { // Column H (index 7) contains "completed"
      targetSheet.appendRow(row);
      action+=1;
      Logger.log("Moved row to 'Completed Tasks': Row " + (i + 1));
    } else {
      newData.push(row);
    }
  }

  // Clear the source sheet starting from row 3 and set new data
  sourceSheet.getRange(3, 1, sourceData.length - 2, sourceData[0].length).clearContent();
  console.log(sourceSheet.getRowHeight(3))
  if (newData.length > 0) {
    sourceSheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);
    Logger.log("Cleared source sheet from row 3 onwards and set new data.");
  }
  deleteFailsafeCopySheet()
  setRowHeightTo98(sourceSheet)
  insertIncreasingNumbers(sourceSheet)
  logFunctionEnd(count, action)

  
}

function moveCompletedTasksv2() {
  logFunctionStart();
  var sourceSheetName = TASK_SHEET_NAME; // Replace with the name of your source sheet
  var targetSheetName = "Completed Tasks"; // Replace with the name of your target sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var targetSheet = ss.getSheetByName(targetSheetName);

  if (!sourceSheet || !targetSheet) {
    Logger.log("Source or target sheet not found.");
    return;
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  var newData = [];
  var count = 0;
  var action = 0;

  duplicateSheet(sourceSheet, ss);

  for (var i = 2; i < sourceData.length; i++) { // Start from row 3 (index 2)
    count += 1;
    var row = sourceData[i];
      if(row[1] == ""){
      continue
    }
    var isMerged = isRowMerged(sourceSheet, i + 1); // Check if the current row is part of a merged range
    console.log("is merged: " + isMerged)
  

    if (!isMerged && row[7] && row[7].toLowerCase() === "completed") { // Column H (index 7) contains "completed"
      targetSheet.appendRow(row);
      action += 1;
      Logger.log("Moved row to 'Completed Tasks': Row " + (i + 1));
    } else {
      newData.push(row);
    }
  }

  // Clear the source sheet starting from row 3 and set new data
  sourceSheet.getRange(3, 1, sourceData.length - 2, sourceData[0].length).clearContent();
  console.log(sourceSheet.getRowHeight(3));

  if (newData.length > 0) {
    sourceSheet.getRange(3, 1, newData.length, newData[0].length).setValues(newData);
    Logger.log("Cleared source sheet from row 3 onwards and set new data.");
  }

  deleteFailsafeCopySheet();
  setRowHeightTo98(sourceSheet);
  sortColumnF(sourceSheet)
  insertIncreasingNumbers(sourceSheet);
  logFunctionEnd(count, action);
}

// Function to check if a specific row is part of a merged range
function isRowMerged(sheet, row) {
  var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  return range.isPartOfMerge();
}


function insertIncreasingNumbers(sheet) {

  var startingRow = 3; // Change to the starting row number

  // Get the last row with content in column A
  var lastRow = sheet.getLastRow();

  // Calculate the number of rows to insert
  var numRowsToInsert = lastRow - startingRow + 1;

  // Generate an array of increasing numbers starting from 1
  var numbersArray = [];
  for (var i = 1; i <= numRowsToInsert; i++) {
    numbersArray.push([i]);
  }

  // Insert the numbers into column A starting from the specified row
  sheet.getRange(startingRow, 1, numRowsToInsert, 1).setValues(numbersArray);
}


function setRowHeightTo98(sheet) {


  // Get the last row with content in the sheet
  var lastRow = sheet.getLastRow();

  // Loop through rows starting from row 3 to the last row
  for (var row = 3; row <= lastRow; row++) {
    sheet.setRowHeight(row, 98); // Set the row height to 98
  }
}


function logFunctionStart() {
  // Get the current spreadsheet.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 
  // Get or create a sheet named "LOG."
  var sheet = spreadsheet.getSheetByName("LOG");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("LOG");
    // Add headers to the "LOG" sheet.
    sheet.appendRow(["Timestamp", "Function Name", "Status"]);
  }
 
  // Get the current date and time.
  var currentDate = new Date();
 
  // Get the name of the currently executing function.
  var functionName = arguments.callee.name;
 
  // Log the start of the function.
  sheet.appendRow([currentDate, functionName, "Start"]);
}

function logFunctionEnd(count, action) {
  // Get the current spreadsheet.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 
  // Get the "LOG" sheet (assuming it already exists).
  var sheet = spreadsheet.getSheetByName("LOG");
 
  if (!sheet) {
    // If the "LOG" sheet does not exist, create it.
    sheet = spreadsheet.insertSheet("LOG");
    // Add headers to the "LOG" sheet.
    sheet.appendRow(["Timestamp", "Function Name", "Status"]);
  }
 
  // Get the current date and time.
  var currentDate = new Date();
 
  // Get the name of the currently executing function.
  var functionName = arguments.callee.name;
 
  // Log the end of the function.
  sheet.appendRow([currentDate, functionName, "End", count, action]);
}


function readTasksList(data) {
  // Loop through the rows and process the data
  var taskData = []
  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    if(row[1] ==""){
      break;
    }
    taskData.push(row)
    
    // Access each cell's value in the row
    for (var j = 0; j < row.length; j++) {
      var cellValue = row[j];
      // Process the cell value here
      Logger.log('Row ' + (i + 1) + ', Column ' + (j + 1) + ': ' + cellValue);
    }
  }
  return taskData;
}

function getTasksListData(sheet) {
  // Replace 'Tasks List' with the name of your spreadsheet

  if (!sheet) {
    Logger.log("Sheet '" + sheetName + "' not found.");
    return [];
  }

  // Get all data from the sheet
  var data = sheet.getDataRange().getValues();

  // Return the data as a 2D array
  return data;
}

function clearCellsFromRow3Down(sheet) {

  // if (!sheet) {
  //   Logger.log("Sheet '" + sheetName + "' not found.");
  //   return;
  // }

  // Get the range starting from row 3 to the last row in the sheet
  var lastRow = sheet.getLastRow();
  var rangeToClear = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn());

  // Clear the contents of the range
  rangeToClear.clearContent();

  Logger.log("Cleared cells from row 3 down in sheet '" + sheetName + "'.");
}



function main(){
  var betaSheet = BETA_LIST;
  var sheetName = TASK_SHEET_NAME;

  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet by name
  var sheet = ss.getSheetByName(sheetName);
  var betaSheet = ss.getSheetByName(betaSheet);

 flipDateFormat(betaSheet)
}

function betaMain(){
  var betaSheet = BETA_LIST;
  var sheetName = TASK_SHEET_NAME;

  // Get the active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet by name
  // var sheet = ss.getSheetByName(sheetName);
  var betaSheet = ss.getSheetByName(betaSheet);

  sortColumnF(betaSheet)

  // console.log(taskData)
}