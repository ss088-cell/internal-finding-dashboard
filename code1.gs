// Main function to run weekly
function runWeeklyReportProcess() {
  moveOldReport();  // Move the old report to Location Y
  createNewReport();  // Create the new report in Location X
}

// 1. Move the old report to Location Y with timestamp
function moveOldReport() {
  const platopsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Platops Internal Findings');
  
  if (!platopsSheet) {
    Logger.log('Platops Internal Findings sheet not found!');
    return;
  }

  const folderXId = 'YOUR_FOLDER_X_ID';  // Replace with the ID of Location X (current folder)
  const folderX = DriveApp.getFolderById(folderXId);
  
  const folderYId = 'YOUR_FOLDER_Y_ID';  // Replace with the ID of Location Y (archive folder)
  const folderY = DriveApp.getFolderById(folderYId);
  
  const date = new Date();
  const timestamp = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  
  const oldFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const newName = 'Platops Internal Findings_' + timestamp;
  oldFile.setName(newName);
  
  folderY.addFile(oldFile);
  folderX.removeFile(oldFile);  // Remove the file from Location X
  
  Logger.log('Report moved to Location Y with name: ' + newName);
}

// 2. Create the new report in Location X
function createNewReport() {
  // The URL of the source Google Sheet (view-only access)
  const sourceSheetUrl = 'https://docs.google.com/spreadsheets/d/your-spreadsheet-id/edit'; // Replace with your Google Sheet URL
  
  // Open the source spreadsheet by URL (view access)
  const sourceSpreadsheet = SpreadsheetApp.openByUrl(sourceSheetUrl);
  
  // Access the 'Detail Data' sheet
  const sourceSheet = sourceSpreadsheet.getSheetByName('Detail Data');
  
  if (!sourceSheet) {
    Logger.log('Detail Data sheet not found!');
    return;
  }

  // Create a filter view on the source sheet to filter "Bus App Name" for specific apps (ABC, BCA, DBA, EEE)
  const filter = sourceSheet.getFilter();
  if (!filter) {
    sourceSheet.createFilter();
  }

  // Apply a filter on the "Bus App Name" column (12th column, "Z")
  const columnToFilter = 12;  // Column "Z" (12th column)
  const filterCriteria = SpreadsheetApp.newFilterCriteria()
    .whenTextContains('ABC')
    .or().whenTextContains('bca')
    .or().whenTextContains('dba')
    .or().whenTextContains('eee')
    .build();

  filter.setColumnFilterCriteria(columnToFilter, filterCriteria);

  // Get the filtered data (only the rows that match the filter)
  const filteredRows = sourceSheet.getDataRange().getValues();

  // Now, we need to copy only the required columns to the newly created report
  const requiredColumns = [
    1, // Host Name
    2, // VPR
    3, // Plugin ID
    4, // Plugin name
    5, // IP
    6, // Description
    7, // Solution
    8, // First Discovered
    9, // Last Observed
    10, // Days Since First Discovered
    11, // Days Since Last Observed
    12, // Bus App Name
    13, // VPR Remediation Due Date
    14, // VPR Compliance
    15  // Risk Type
  ];

  // Prepare the filtered data for copying
  const filteredData = [];
  
  // Loop through the rows and extract only the required columns
  filteredData.push(filteredRows[0].filter((_, index) => requiredColumns.includes(index + 1)));  // Add header row
  
  for (let i = 1; i < filteredRows.length; i++) {
    filteredData.push(filteredRows[i].filter((_, index) => requiredColumns.includes(index + 1)));
  }

  // Create a new sheet and store the filtered data
  const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Platops Internal Findings');
  newSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

  // Store this new sheet in a specific folder (Location X)
  const folderXId = 'YOUR_FOLDER_X_ID';  // Replace with the ID of Location X
  const folderX = DriveApp.getFolderById(folderXId);
  
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  folderX.addFile(file);  // Add the new file to Location X
  Logger.log('New report created and stored in Location X');
  
  // Optionally, remove the file from the root folder after it's added to the specific folder
  const rootFolder = DriveApp.getRootFolder();
  rootFolder.removeFile(file);
  
  Logger.log('New report has been created and stored in Location X');
}
