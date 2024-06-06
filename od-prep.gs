// var tabName = "OD Scores | Jan 24";
// var ranges = getParameterRanges(tabName);

// New Test
function Final(){
  var tabName = "OD Scores | Apr 24";
  var ranges = getParameterRanges(tabName);
  transferDataToOutputSheet(tabName, ranges);
  rearrangedMonthlyDataSheet()
  DeleteODScores()

}

function transferDataToOutputSheet(tabName, parameterRanges) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = spreadsheet.getSheetByName('output');

  if (!outputSheet) {
    outputSheet = spreadsheet.insertSheet('output');
  }

  for (var parameter in parameterRanges) {
    var range = parameterRanges[parameter];
    var data = spreadsheet.getSheetByName(tabName).getRange(range).getValues();
    var numRows = data.length;
    var parameterColumn = Array(numRows).fill([parameter]);

    // Add parameter column to data
    for (var i = 0; i < numRows; i++) {
      data[i].push(parameterColumn[i][0]);
    }

    // Write data to output sheet
    outputSheet.getRange(outputSheet.getLastRow() + 1, 1, numRows, data[0].length).setValues(data);
  }
}

// Call the function with tabName and ranges
var tabName = "OD Scores | Apr 24";
var ranges = getParameterRanges(tabName);
transferDataToOutputSheet(tabName, ranges);


//2nd try

function getParameterRanges(tabName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(tabName);
  
  if (!sheet) {
    Logger.log("Tab with name '" + tabName + "' not found.");
    return;
  }
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var ranges = {};
  var parameters = ["oGV", "iGV", "oGTe", "iGTe", "oGTa", "iGTa", "DXP", "BD", "FnL", "TM", "Brand MKT", "EM", "EwA & PR", "IM"];
  // var parameters = ["TM"];
 
  for (var p = 0; p < parameters.length; p++) {
    var parameter = parameters[p];
    var parameterRange = '';
    var parameterFound = false;
    
    for (var i = 1; i <= lastRow; i++) {
      for (var j = 1; j <= lastColumn; j++) {
        var cellValue = sheet.getRange(i, j).getValue();
        
        if (cellValue === parameter) {
          var tableStartRow = i;
          var tableStartColumn = j;
          parameterFound = true;
          break;
        }
      }
      if (parameterFound) break;
    }
    
    if (parameterFound) {
      for (var r = tableStartRow; r <= lastRow; r++) {
        var cellValue = sheet.getRange(r, tableStartColumn).getValue();
        if (cellValue === "OD Score") {
          var tableEndRow = r;
          var tableEndColumn = tableStartColumn + 11; // Assuming tables always have 11 columns
          parameterRange = sheet.getRange(tableStartRow, tableStartColumn, tableEndRow - tableStartRow + 1, tableEndColumn - tableStartColumn + 1).getA1Notation();
          ranges[parameter] = parameterRange;
          break;
        }
      }
    }
  }
  
  Logger.log(ranges);
  return ranges;
}

function DeleteODScores() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = spreadsheet.getSheetByName('output');

  if (!outputSheet) {
    Logger.log("Output sheet not found.");
    return;
  }

  // Delete rows containing "CC"
  var dataRange = outputSheet.getDataRange();
  var values = dataRange.getValues();
  var numRows = values.length;
  var rowsToDelete = [];
  
  for (var i = 0; i < numRows; i++) {
    if (values[i].indexOf("OD Score") !== -1) {
      rowsToDelete.push(i + 1);
    }
  }
  
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    outputSheet.deleteRow(rowsToDelete[j]);
  }
}

function rearrangedMonthlyDataSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = spreadsheet.getSheetByName('output');

  if (!outputSheet) {
    Logger.log("Output sheet not found.");
    return;
  }

  // Delete rows containing "CC"
  var dataRange = outputSheet.getDataRange();
  var values = dataRange.getValues();
  var numRows = values.length;
  var rowsToDelete = [];
  
  for (var i = 0; i < numRows; i++) {
    if (values[i].indexOf("CC") !== -1) {
      rowsToDelete.push(i + 1);
    }
  }
  
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    outputSheet.deleteRow(rowsToDelete[j]);
  }

  // // Add new row with specific headers
  var headers = ["Criteria", "CC", "CN", "CS", "KANDY", "NIBM", "NSBM", "RUHUNA", "SLIIT", "USJ", "Rajarata", "Function"];
  // outputSheet.insertRowAfter(1);
  // outputSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  
}

function clearAllOutput() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = spreadsheet.getSheetByName('output');

  if (!outputSheet) {
    Logger.log("Output sheet not found.");
    return;
  }

  outputSheet.clear(); // Clear all data in the 'output' sheet
}
