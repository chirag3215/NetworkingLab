function renameFilesInFolderWithParticipantID() {
  var folderId = '1vNcilDeZgJcHCQWdEH7fMz2NinNpVBPF'; // Your folder ID
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var existingNames = {}; // To keep track of existing participant IDs

  while (files.hasNext()) {
    var file = files.next();
    var oldFileName = file.getName();
    
    // Extract the file extension (e.g., ".png", ".PNG", ".jpg", ".jpeg")
    var fileExtension = oldFileName.substring(oldFileName.lastIndexOf('.')).toLowerCase();

    // Remove "R_" prefix if it exists and get the participant ID
    var participantID = oldFileName;
    if (participantID.startsWith("R_")) {
      participantID = participantID.substring(2);
    }

    // Remove everything after the first underscore (and the underscore itself)
    if (participantID.indexOf('_') !== -1) {
      participantID = participantID.split('_')[0];
    }

    // Check if the participant ID already exists
    if (existingNames[participantID]) {
      // Increment the counter for this participant ID and append the number
      existingNames[participantID]++;
      participantID = participantID + "_" + existingNames[participantID]; // Append number after ID
    } else {
      // First occurrence, set counter to 1
      existingNames[participantID] = 1;
    }

    // Create the new file name with the participant ID and original file extension
    var newFileName = participantID + fileExtension;
    Logger.log("Renaming file: " + oldFileName + " to " + newFileName);
    
    // Rename the file
    file.setName(newFileName);
  }

  Logger.log("All files have been renamed.");
}

