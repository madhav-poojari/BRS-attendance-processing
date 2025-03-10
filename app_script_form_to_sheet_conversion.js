function generateAndOrganizeResponseSheets() {
  var parentFolder = DriveApp.getFolderById("1pT9KRlLEwuYY5KWON0uNNTp3Uf-rUAqZ"); // Replace with your Folder ID
  var responseFolder = parentFolder.createFolder("Response Sheets - " + new Date().toISOString().slice(0, 10));

  var files = parentFolder.getFilesByType(MimeType.GOOGLE_FORMS);

  while (files.hasNext()) {
    var file = files.next();
    var formName = file.getName(); // Get the form's filename
    var form = FormApp.openById(file.getId());

    // Extract coach's full name from the form name
    var coachName = extractCoachName(formName);

    // Create a new spreadsheet for responses
    var sheetName = coachName + " Attendance Report";
    var newSheet = SpreadsheetApp.create(sheetName);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, newSheet.getId());

    // Move the new spreadsheet into the response folder
    var sheetFile = DriveApp.getFileById(newSheet.getId());
    sheetFile.moveTo(responseFolder);

    Logger.log("Created and moved response sheet: " + sheetFile.getName());
  }

  Logger.log("All response sheets are created and organized in: " + responseFolder.getName());
}

// Function to extract coach's full name from the form filename
function extractCoachName(fileName) {
  // Assuming form name is in format "Coach Firstname Lastname Attendance Sheet"
  var match = fileName.match(/Coach (.+?) Attendance Sheet/);
  return match ? match[1] : "Unknown Coach";
}
