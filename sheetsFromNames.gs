function createSheetsFromNamesUsingTemplate() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = spreadsheet.getSheetByName('Template'); // Ensure you have a template sheet named 'Template'
  if (!templateSheet) {
    throw new Error('Template sheet not found');
  }
  
  var namesSheet = spreadsheet.getActiveSheet(); // Assumes the list is on the active sheet
  var namesRange = namesSheet.getDataRange(); // Adjust the range as necessary
  var names = namesRange.getValues();

  names.forEach(function(row) {
    var sheetName = row[0]; // Assumes names are in the first column
    if (sheetName && !spreadsheet.getSheetByName(sheetName)) {
      // If a sheet with the name doesn't exist, create it by copying the template
      var newSheet = templateSheet.copyTo(spreadsheet);
      newSheet.setName(sheetName);
      newSheet.getRange("A1").setValue(sheetName);
      
    }
  });
  // names.forEach(function(row) {
  //   var sheetName = row[0]; // Assumes names are in the first column
  //   var sheet = spreadsheet.getSheetByName(sheetName);
  //     var cell = sheet.getRange('A1');
  //     if (cell.getFormula() === "=mySheetName()"){
  //         cell.setValue(sheetName);
  //     }
      
    
  // });
}

