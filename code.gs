// Main function to run weekly
function runWeeklyReportProcess() {
  moveOldReport();  // Move the old report to Location Y
  createNewReport();  // Create the new report in Location X
}

// 1. Move the old report to Location Y with timestamp
function moveOldReport() {
  // Access the 'Platops Internal Findings' sheet in Location X (current report)
  const platopsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Platops Internal Findings');
  
  if (!platopsSheet) {
    Logger.log('Platops Internal Findings sheet not found!');
    return;
  }
  
  // Folder X: The folder where the latest report is stored (current folder)
  const folderXId = 'YOUR_FOLDER_X_ID';  // Replace with the ID of Location X (current folder)
  const folderX = DriveApp.getFolderById(folderXId);
  
  // Folder Y: The folder where previous reports will be moved (archive folder)
  const folderYId = 'YOUR_FOLDER_Y_ID';  // Replace with the ID of Location Y (archive folder)
  const folderY = DriveApp.getFolderById(folderYId);
  
  // Create a timestamp for the report name
  const date = new Date();
  const timestamp = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
  
  // Rename the old report with a timestamp and move it to Location Y
  const oldFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const newName = 'Platops Internal Findings_' + timestamp;
  oldFile.setName(newName);
  
  // Move the file to Location Y
  folderY.addFile(oldFile);
  folderX.removeFile(oldFile);  // Remove the file from Location X
  
  // Optionally, log or notify
  Logger.log('Report moved to Location Y with name: ' + newName);
}

// 2. Create the new report in Location X
function createNewReport() {
  // The URL of the Google Sheet containing the 'Detailed Data' sheet
  const sourceSheetUrl = 'YOUR_GOOGLE_SHEET_URL'; // Replace with your Google Sheet URL
  
  // Open the source spreadsheet by URL
  const sourceSpreadsheet = SpreadsheetApp.openByUrl(sourceSheetUrl);
  
  // Access the 'Detailed Data' sheet
  const sourceSheet = sourceSpreadsheet.getSheetByName('Detailed Data');
  
  if (!sourceSheet) {
    Logger.log('Detailed Data sheet not found!');
    return;
  }

  // Get all the data from the 'Detailed Data' sheet
  const data = sourceSheet.getDataRange().getValues();
  
  // Create a temporary sheet in the current Google Sheet
  const tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('TempData');
  
  // Set headers (assuming headers are in the first row of 'Detailed Data')
  tempSheet.appendRow(data[0]); // Assuming first row is the header
  
  // Loop through the data and insert it into the temporary sheet
  for (let i = 1; i < data.length; i++) {
    tempSheet.appendRow(data[i]);
  }
  
  // Apply the filter function to select specific applications
  applyFilters(tempSheet);

  // After filtering, copy only the required columns to the newly created report
  createAndStoreNewReport(tempSheet);
}

// 3. Apply filters and select specific columns
function applyFilters(tempSheet) {
  const range = tempSheet.getDataRange();
  
  // Apply a filter for the "Bus App Name" column (assuming it's in the 13th column)
  const filterCriteria = SpreadsheetApp.newFilterCriteria()
    .whenTextContains('ABC')
    .or().whenTextContains('bca')
    .or().whenTextContains('dba')
    .or().whenTextContains('eee')
    .build();
  
  range.createFilter().setColumnFilterCriteria(13, filterCriteria);
}

// 4. Create and store the filtered data in a new sheet (with selected columns)
function createAndStoreNewReport(tempSheet) {
  // Create a new sheet named 'Platops Internal Findings'
  const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Platops Internal Findings');
  
  // Define the columns we want to copy (index is based on the original 'Detailed Data' sheet)
  const columnsToCopy = [
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

  // Filtered data will be stored here
  const filteredData = [];
  
  // Get the rows that match the filter
  const rows = tempSheet.getDataRange().getValues();
  
  // Push the headers first
  filteredData.push(rows[0].filter((_, idx) => columnsToCopy.includes(idx + 1)));
  
  // Loop through the data and filter only the selected columns
  for (let i = 1; i < rows.length; i++) {
    if (tempSheet.getFilter() && tempSheet.getFilter().getColumnFilterCriteria(13)) {
      filteredData.push(rows[i].filter((_, idx) => columnsToCopy.includes(idx + 1)));
    }
  }
  
  // Copy the filtered data to the new sheet
  newSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

  // Store this new sheet in a specific folder (Location X)
  const folderXId = 'YOUR_FOLDER_X_ID';  // Replace with the ID of Location X
  const folderX = DriveApp.getFolderById(folderXId);
  
  // Get the file and add it to Location X
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  folderX.addFile(file);  // Add the new file to Location X
  Logger.log('New report created and stored in Location X');
  
  // Optionally, remove the file from the root folder after it's added to the specific folder
  const rootFolder = DriveApp.getRootFolder();
  rootFolder.removeFile(file);
  
  // Optionally, log or notify
  Logger.log('New report has been created and stored in Location X');
}
