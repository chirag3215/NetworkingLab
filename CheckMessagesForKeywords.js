function checkMessagesForKeywords() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Assume "Messages" is in column B (2), adjust if it's in a different column
  var messagesRange = sheet.getRange(2, 2, lastRow - 1, 1); 
  var messages = messagesRange.getValues();
  
  var results = [];
  
  for (var i = 0; i < messages.length; i++) {
    var message = messages[i][0].toLowerCase(); // Convert message to lowercase to make the check case insensitive
    
    // Check if either "inspired by" or "honored" are present
    if (message.includes("inspired by") || message.includes("honored")) {
      results.push([0]);  // Set 1 if either word is found
    } else {
      results.push([1]);  // Set 0 if neither word is found
    }
  }
  
  // Write the results to the next column (Assuming column C (3))
  var resultsRange = sheet.getRange(2, 3, results.length, 1);
  resultsRange.setValues(results);
  
  Logger.log("Check completed successfully.");
}
