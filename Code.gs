function getSheetByCaseInsensitiveName(source, editedValue) {
  // Get all sheets in the spreadsheet
  var sheets = source.getSheets();

  // Normalize the case for comparison
  var editedValueLower = editedValue.toLowerCase();

  // Iterate through all sheets
  for (var i = 0; i < sheets.length; i++) {
    // If a sheet name matches editedValue case-insensitively, return it
    if (sheets[i].getName().toLowerCase() === editedValueLower) {
      return sheets[i];
    }
  }

  // If no matching sheet is found, return null
  return null;
}
function eraseRemovedName(editedValue, source, oldValue, activeRange, activeSheet){
  var targetSheet = getSheetByCaseInsensitiveName(source, oldValue);
        if (targetSheet) {
            Logger.log("Clearing name from: " + oldValue + "'s sheet");
            var targetRange = targetSheet.getRange(activeRange.getA1Notation());
            // Check if the target cell contains the active sheet's name before clearing
            if (targetRange.getValue().toLowerCase() === activeSheet.getName().toLowerCase()) {
                targetRange.setValue("Free!");
            }
        }
}
function dealWithChange(editedValue, oldValue, activeRange, source, activeSheet){
  Logger.log(editedValue);

    if ((!editedValue || !getSheetByCaseInsensitiveName(source, editedValue)) && oldValue) {
        eraseRemovedName(editedValue, source, oldValue, activeRange, activeSheet);
        
        
    }else if (editedValue) {
      eraseRemovedName(editedValue, source, oldValue, activeRange, activeSheet);
      // Proceed with the original logic
      var targetSheet = getSheetByCaseInsensitiveName(source, editedValue);
      if (targetSheet) {
          Logger.log("target sheet exists");
          var targetRange = targetSheet.getRange(activeRange.getA1Notation());
          var activeSheetName = activeSheet.getName();

          // Check if the target cell is empty before attempting to set a new value
          if (targetRange.getValue() === '' || targetRange.getValue() == 'Free!') {
              targetRange.setValue(activeSheetName);
          } else {
              activeRange.setValue(oldValue);
              Logger.log("Target cell is already occupied. No changes made.");
              SpreadsheetApp.getUi().alert('The scheduling slot is already taken. Please choose a different slot.');
              // Optionally, you might want to notify the editor that the target slot is already taken
              // Note: Direct user notifications from onEdit triggers have limitations and may not always work as expected
          }
      }
    }
}
function onEdit(e) {

    Logger.log('onEdit');

    // Basic checks to ensure the edit event object 'e' is properly defined
    if (!e || !e.range) {
        Logger.log('The function must be triggered by an edit. Event object is not defined.');
        return;
    }
    
    var activeSheet = e.source.getActiveSheet();
    var activeRange = e.range;
    var editedValue = e.value; // The name of the person whose sheet to update
    var oldValue = e.oldValue;
    
    dealWithChange(editedValue, oldValue, activeRange, e.source, activeSheet);
  
}
