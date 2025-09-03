function onEdit(e) {
  var sourceSheetName = "********************";      // Source sheet name to watch
  var destinationSheetName = "**********************"; // Destination sheet name to append values into
  var columnToWatch = 1; // Column index to monitor (1 = column A)

  var range = e.range;              // Edited range
  var sheet = range.getSheet();     // Sheet where the edit occurred
  var column = range.getColumn();   // Column number of the edited cell
  
  // Log details about the edit (useful for debugging)
  Logger.log("Sheet name: " + sheet.getName());
  Logger.log("Edited column: " + column);
  Logger.log("Edited value: " + range.getValue());

  // Only trigger if the edit is on the specific source sheet and in the watched column
  if (sheet.getName() === sourceSheetName && column === columnToWatch) {
    var newValue = range.getValue().trim(); // Get new value, trimmed
    Logger.log("New value: " + newValue);
    
    // If the new value is not empty, add it to the destination sheet
    if (newValue !== "") {
      addToDestinationSheet(newValue, destinationSheetName);
    }
  }
}

function addToDestinationSheet(name, destSheetName) {
  // Append the given name into the destination sheet (column A, next empty row)
  Logger.log("Adding value to destination sheet: " + name);
  var destSpreadsheet = SpreadsheetApp.openById("*********************************************"); // Destination spreadsheet ID
  var destSheet = destSpreadsheet.getSheetByName(destSheetName);
  var lastRow = destSheet.getLastRow(); // Find last row with content
  destSheet.getRange(lastRow + 1, 1).setValue(name); // Write into next empty row in column A
}

function formatNamesInColumnA() {
  // Format all names in column A to proper case, except for certain lowercase words
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange("A2:A" + sheet.getLastRow()); // Range from A2 down to last row (skip header)
  const values = range.getValues();

  const lowercaseWords = ["de", "da", "do", "dos"]; // Words that should stay lowercase

  const formatted = values.map(row => {
    const name = row[0];
    if (typeof name === "string" && name.trim() !== "") {
      // Convert name to lowercase, split into words, capitalize except for "stop words"
      const formattedName = name
        .toLowerCase()
        .split(" ")
        .map(word => lowercaseWords.includes(word) 
          ? word 
          : word.charAt(0).toUpperCase() + word.slice(1))
        .join(" ");
      return [formattedName];
    } else {
      // If not a string or empty, keep as-is
      return [name];
    }
  });

  // Write the formatted names back into the sheet
  range.setValues(formatted);
}
