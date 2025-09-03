
function onEdit(e) {
  var sourceSheetName = "********************"; 
  var destinationSheetName = "**********************"; 
  var columnToWatch = 1; 

  var range = e.range;
  var sheet = range.getSheet();
  var column = range.getColumn();
  
  Logger.log("Sheet name: " + sheet.getName());
  Logger.log("Edited column: " + column);
  Logger.log("Edited value: " + range.getValue());

  if (sheet.getName() === sourceSheetName && column === columnToWatch) {
    var newValue = range.getValue().trim();
    Logger.log("New value: " + newValue);
    
    if (newValue !== "") {
      addToDestinationSheet(newValue, destinationSheetName);
    }
  }
}

function addToDestinationSheet(name, destSheetName) {
  Logger.log("Adding value to destination sheet: " + name);
  var destSpreadsheet = SpreadsheetApp.openById("*********************************************");
  var destSheet = destSpreadsheet.getSheetByName(destSheetName);
  var lastRow = destSheet.getLastRow();
  destSheet.getRange(lastRow + 1, 1).setValue(name);
}


function formatNamesInColumnA() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange("A2:A" + sheet.getLastRow()); // assumes headers in row 1
  const values = range.getValues();

  const lowercaseWords = ["de", "da", "do", "dos"];

  const formatted = values.map(row => {
    const name = row[0];
    if (typeof name === "string" && name.trim() !== "") {
      const formattedName = name
        .toLowerCase()
        .split(" ")
        .map(word => lowercaseWords.includes(word) 
          ? word 
          : word.charAt(0).toUpperCase() + word.slice(1))
        .join(" ");
      return [formattedName];
    } else {
      return [name];
    }
  });

  range.setValues(formatted);
}
