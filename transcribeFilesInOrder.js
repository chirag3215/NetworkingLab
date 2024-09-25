function transcribeFilesInOrder() {
  var folderId = '1vNcilDeZgJcHCQWdEH7fMz2NinNpVBPF'; // Replace with your folder ID
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var transcribedTexts = [];
  var participantIDs = [];
  var allFiles = [];
  
  // Your Google Cloud Vision API key
  var apiKey = 'AIzaSyDvvUN7ac7HW1rKKUNueQuNaD3vAjFjdT4';  // Replace with your Vision API key
  
  // Collect all file information into an array
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    allFiles.push({ fileName: fileName, file: file });
  }
  
  // Sort files alphabetically by file name
  allFiles.sort(function(a, b) {
    return a.fileName.localeCompare(b.fileName, undefined, { sensitivity: 'base' });
  });
  
  // Process the sorted files
  for (var i = 0; i < allFiles.length; i++) {
    var file = allFiles[i].file;
    var fileName = allFiles[i].fileName;
    
    Logger.log("Processing file: " + fileName);
    
    // Extract participant ID from the file name
    var participantID = fileName;

    // Remove "R_" prefix if it exists
    if (participantID.startsWith("R_")) {
      participantID = participantID.substring(2);
    }

    // Remove everything after the first underscore and the underscore itself
    if (participantID.indexOf('_') !== -1) {
      participantID = participantID.split('_')[0];
    }

    // Remove the file extension (everything after and including the last period)
    if (participantID.indexOf('.') !== -1) {
      participantID = participantID.substring(0, participantID.lastIndexOf('.'));
    }

    // Add the participant ID to the list
    participantIDs.push([participantID]);

    // Only process .png, .jpg, .jpeg files for transcription
    if (fileName.toLowerCase().endsWith('.png') || fileName.toLowerCase().endsWith('.jpg') || fileName.toLowerCase().endsWith('.jpeg')) {
      var blob = file.getBlob();
      
      // Create the Vision API request payload
      var visionRequest = {
        "requests": [
          {
            "image": {
              "content": Utilities.base64Encode(blob.getBytes())
            },
            "features": [
              {
                "type": "TEXT_DETECTION"
              }
            ]
          }
        ]
      };
      
      try {
        // Make the API call
        var url = 'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey;
        var options = {
          'method': 'post',
          'contentType': 'application/json',
          'payload': JSON.stringify(visionRequest),
          'muteHttpExceptions': true
        };
        
        var response = UrlFetchApp.fetch(url, options);
        var json = JSON.parse(response.getContentText());

        // Check for errors in the API response
        if (json.error) {
          Logger.log("Error from Vision API: " + JSON.stringify(json.error));
          throw new Error("Vision API returned an error: " + json.error.message);
        }

        // Check if the response contains text
        if (json.responses && json.responses[0].fullTextAnnotation) {
          var fileText = json.responses[0].fullTextAnnotation.text;
          Logger.log('Transcribed text for ' + fileName + ': ' + fileText);
          
          // Append the transcribed content to the list
          transcribedTexts.push([fileText]);
        } else {
          Logger.log('No text found in ' + fileName);
          transcribedTexts.push(['No text detected.']);
        }

      } catch (error) {
        Logger.log("Error processing file: " + fileName + " - " + error.message);
        transcribedTexts.push(['Error during transcription.']);
      }
    } else {
      Logger.log('Skipping non-image file: ' + fileName);
      transcribedTexts.push(['Skipped non-image file']);
    }
  }

  // Write the participant IDs and transcribed content to the sheet
  var participantIDRange = sheet.getRange(2, 1, participantIDs.length, 1);  // Column A
  var transcribedTextRange = sheet.getRange(2, 2, transcribedTexts.length, 1);  // Column B
  participantIDRange.setValues(participantIDs);
  transcribedTextRange.setValues(transcribedTexts);

  Logger.log("All files have been processed, sorted alphabetically, and written to the sheet.");
}
