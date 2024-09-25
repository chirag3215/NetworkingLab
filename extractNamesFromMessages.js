function extractNamesFromMessages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();  // Get the last row with content

  // Loop through rows starting from row 85 (adjust this to your needs)
  for (var i = 2; i <= lastRow; i++) { //assumes row 1 is headers
    var message = sheet.getRange(i, 2).getValue();  // Get message in column B (2)

    // Initialize variables for the extracted names
    var nameAfterGreeting = "";
    var nameAfterSincerely = "";

    // Find the word immediately after "Dear" or "Hello"
    var greetingIndex = -1;
    var greetingWord = "";

    // Check for "Dear" first
    var dearIndex = message.indexOf("Dear");
    if (dearIndex !== -1) {
      greetingIndex = dearIndex;
      greetingWord = "Dear";
    }

    // If "Dear" not found, check for "Hello"
    if (greetingIndex === -1) {
      var helloIndex = message.indexOf("Hello");
      if (helloIndex !== -1) {
        greetingIndex = helloIndex;
        greetingWord = "Hello";
      }
    }

    // If we found a greeting, extract the name after it
    if (greetingIndex !== -1) {
      var afterGreeting = message.substring(greetingIndex + greetingWord.length).trim();  // Get text after "Dear " or "Hello "
      var firstComma = afterGreeting.indexOf(",");  // Find the first comma
      if (firstComma !== -1) {
        nameAfterGreeting = afterGreeting.substring(0, firstComma).trim();  // Extract text until the first comma
      } else {
        nameAfterGreeting = afterGreeting;  // If no comma, take all text after "Dear" or "Hello"
      }
    }

    // Find the word immediately after "Sincerely,"
    var sincerelyIndex = message.indexOf("Sincerely,");
    if (sincerelyIndex !== -1) {
      var afterSincerely = message.substring(sincerelyIndex + 10).trim();  // Get text after "Sincerely, "
      var firstSpace = afterSincerely.indexOf(" ");
      if (firstSpace !== -1) {
        nameAfterSincerely = afterSincerely.substring(0, firstSpace).trim();  // Extract first word after "Sincerely,"
      } else {
        nameAfterSincerely = afterSincerely;  // If no space, take all text after "Sincerely,"
      }
    }

    // Write the extracted names to columns C (3) and E (5)
    if (nameAfterGreeting) {
      sheet.getRange(i, 3).setValue(nameAfterGreeting);  // Write name after "Dear" or "Hello" to column C
    }

    if (nameAfterSincerely) {
      sheet.getRange(i, 5).setValue(nameAfterSincerely);  // Write name after "Sincerely" to column E
    }
  }

  Logger.log("Names have been extracted and written to columns C and E.");
}
