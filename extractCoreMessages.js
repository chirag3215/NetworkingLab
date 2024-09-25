function extractCoreMessages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  // Assume "Transcribed Text" is in column B (2), adjust if it's in a different column
  var range = sheet.getRange(2, 2, lastRow - 1, 1); 
  var transcribedTexts = range.getValues();
  var coreMessages = [];
  
  // Expanded greetings list with capitalized versions and added punctuation
  var greetings = ["hey", "Hey", "hello", "Hello", "dear", "Dear", "how", "How"];
  var signOffs = ["sincerely", "Sincerely", "from", "From", "best regards", "Best regards", "yours truly", "Yours truly"];
  
  for (var i = 0; i < transcribedTexts.length; i++) {
    var text = transcribedTexts[i][0]; // Original text
    var lowerCaseText = text.toLowerCase(); // Convert the text to lowercase
    var coreMessage = extractMessage(text, lowerCaseText, greetings, signOffs); // Search in both the original and lowercase
    
    // Remove special characters except letters, numbers, punctuation, and newlines
    coreMessage = coreMessage.replace(/[^a-zA-Z0-9.,!?'"()\- \n\r]/g, '').trim();
    
    coreMessages.push([coreMessage]);
  }
  
  // Write the core messages to the next column, assuming column C (3)
  var coreMessageRange = sheet.getRange(2, 3, coreMessages.length, 1);
  coreMessageRange.setValues(coreMessages);
  
  Logger.log("Core messages extracted successfully.");
}

// Function to extract the core message based on the first greeting and sign-off
function extractMessage(originalText, lowerCaseText, greetings, signOffs) {
  var start = -1;
  var end = -1;
  var greetingFound = "";

  // Find the **first** occurrence of the greeting
  for (var i = 0; i < greetings.length; i++) {
    var greetingIndex = lowerCaseText.indexOf(greetings[i].toLowerCase());  // Check in lowercase version
    if (greetingIndex !== -1) {
      start = greetingIndex;  // Start at the first greeting
      greetingFound = originalText.substring(greetingIndex, greetingIndex + greetings[i].length) + " ";  // Add space after the greeting
      break;
    }
  }

  // Find the **first** occurrence of the sign-off
  for (var j = 0; j < signOffs.length; j++) {
    var signOffIndex = lowerCaseText.indexOf(signOffs[j].toLowerCase());  // Check in lowercase version
    if (signOffIndex !== -1) {
      end = signOffIndex;  // End at the start of the first sign-off
      break;
    }
  }

  // Case 1: Both greeting and sign-off found
  if (start !== -1 && end !== -1 && start < end) {
    // Return the message, including the greeting and ending just before the sign-off
    return greetingFound + originalText.substring(start + greetingFound.length, end).trim() + '\n' + originalText.substring(end).trim();
  }
  // Case 2: Only greeting found (no sign-off/condolence)
  else if (start !== -1 && end === -1) {
    return greetingFound + originalText.substring(start + greetingFound.length).trim();  // Keep text after greeting
  }
  // Case 3: No greeting or sign-off found, return the whole text
  else {
    return originalText.trim();
  }
}
