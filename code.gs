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

  // Get all the data from the 'Detail Data' sheet
  const data = sourceSheet.getDataRange().getValues();
  
  // Now we filter out only the necessary columns
  const columnsToFetch = [
    0, // Host Name
    1, // VPR
    2, // Plugin ID
    3, // Plugin name
    4, // IP
    5, // Description
    6, // Solution
    7, // First Discovered
    8, // Last Observed
    9, // Days Since First Discovered
    10, // Days Since Last Observed
    11, // Bus App Name
    12, // VPR Remediation Due Date
    13, // VPR Compliance
    14  // Risk Type
  ];
  
  // Prepare filtered data based on columnsToFetch
  const filteredData = [];
  
  // Push headers first
  filteredData.push(columnsToFetch.map(index => data[0][index]));  // Map headers from the selected columns
  
  // Loop through the data and only keep the required columns
  for (let i = 1; i < data.length; i++) {
    filteredData.push(columnsToFetch.map(index => data[i][index]));
  }

  // Create a temporary sheet in the current Google Sheet
  const tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('TempData');
  
  // Set the filtered headers
  tempSheet.appendRow(filteredData[0]); // Set headers
  
  // Loop through the filtered data and insert into the temporary sheet
  for (let i = 1; i < filteredData.length; i++) {
    tempSheet.appendRow(filteredData[i]);
  }
  
  // Apply the filter function to select specific applications
  applyFilters(tempSheet);

  // After filtering, copy the filtered data to the newly created report
  createAndStoreNewReport(tempSheet);
}

// 3. Apply filters and select specific columns (Bus App Name is in the "Z" column, i.e., the 12th column)
function applyFilters(tempSheet) {
  const range = tempSheet.getDataRange();
  
  // Apply a filter for the "Bus App Name" column (12th column, "Z")
  const filterCriteria = SpreadsheetApp.newFilterCriteria()
    .whenTextContains('ABC')
    .or().whenTextContains('bca')
    .or().whenTextContains('dba')
    .or().whenTextContains('eee')
    .build();
  
  range.createFilter().setColumnFilterCriteria(12, filterCriteria);
}

// 4. Create and store the filtered data in a new sheet (with selected columns)
function createAndStoreNewReport(tempSheet) {
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

  const filteredData = [];
  
  const rows = tempSheet.getDataRange().getValues();
  
  // Push the headers first
  filteredData.push(rows[0].filter((_, idx) => columnsToCopy.includes(idx + 1)));
  
  // Loop through the data and filter only the selected columns
  for (let i = 1; i < rows.length; i++) {
    if (tempSheet.getFilter() && tempSheet.getFilter().getColumnFilterCriteria(12)) {
      filteredData.push(rows[i].filter((_, idx) => columnsToCopy.includes(idx + 1)));
    }
  }
  
  newSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

  const folderXId = 'YOUR_FOLDER_X_ID';  // Replace with the ID of Location X
  const folderX = DriveApp.getFolderById(folderXId);
  
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  folderX.addFile(file);  // Add the new file to Location X
  Logger.log('New report created and stored in Location X');
  
  const rootFolder = DriveApp.getRootFolder();
  rootFolder.removeFile(file);
  
  Logger.log('New report has been created and stored in Location X');
}
