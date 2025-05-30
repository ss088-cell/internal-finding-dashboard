// Main function to run weekly
function runWeeklyReportProcess() {
  resetStartRow();  // Reset startRow to 1 before starting the process
  moveOldReport();  // Move the old report to Location Y if it exists, and save the new one at Location X
  processLargeData();  // Process the large data in batches
}

// 1. Reset startRow to 1 before starting the process
function resetStartRow() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('startRow');  // Reset the startRow to 1
  Logger.log("Start row reset to 1.");
}

// 2. Move the generated Platops Internal Findings report to Location Y with timestamp
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
  
  // Check if the Platops Internal Findings file exists at Location X
  const existingFiles = folderX.getFilesByName('Platops Internal Findings');
  if (existingFiles.hasNext()) {
    const oldFile = existingFiles.next();
    const newName = 'Platops Internal Findings_' + timestamp;
    oldFile.setName(newName);
    
    // Move the old file to Location Y
    folderY.addFile(oldFile);
    folderX.removeFile(oldFile);  // Remove the file from Location X

    Logger.log('Old report moved to Location Y with name: ' + newName);
  }
  
  // Ensure the new file is created and saved at Location X
  const newFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  folderX.addFile(newFile);
  Logger.log('New report saved to Location X with name: Platops Internal Findings');
}

// 3. Process the large data in chunks to avoid timeout
function processLargeData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sourceSheetUrl = 'https://docs.google.com/spreadsheets/d/your-spreadsheet-id/edit'; // Replace with your Google Sheet URL
  const sourceSpreadsheet = SpreadsheetApp.openByUrl(sourceSheetUrl);
  const sourceSheet = sourceSpreadsheet.getSheetByName('Detail Data');
  
  if (!sourceSheet) {
    Logger.log('Detail Data sheet not found!');
    return;
  }

  // Get the total number of rows in the source sheet
  const totalRows = sourceSheet.getLastRow();
  
  // Get the current row to start from (stored in properties)
  let startRow = parseInt(scriptProperties.getProperty('startRow') || '1');  // Starts from row 1

  // Process a batch of data
  const batchSize = 1000;  // Adjust the batch size for optimal performance (set a reasonable size based on your data)
  const endRow = Math.min(startRow + batchSize - 1, totalRows);  // Ensure we don't exceed total rows
  
  const dataRange = sourceSheet.getRange(startRow, 1, endRow - startRow + 1, 52);  // Columns A to AZ (52 columns)
  const data = dataRange.getValues();
  
  // Get the header row to map the column names to their indexes
  const headerRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  // Map the required column names to their respective indexes
  const columnNames = [
    "Host Name", "VPR", "Plugin ID", "Plugin name", "IP", "Description", "Solution",
    "First Discovered", "Last Observed", "Days Since First Discovered", "Days Since Last Observed",
    "Bus App Name", "VPR Remediation Due Date", "VPR Compliance", "Risk Type"
  ];

  // Create a map of column indexes based on the header row
  const columnIndexes = {};
  columnNames.forEach(function (columnName) {
    const index = headerRow.indexOf(columnName);
    if (index !== -1) {
      columnIndexes[columnName] = index;
    }
  });

  // Check if all columns were found
  if (Object.keys(columnIndexes).length !== columnNames.length) {
    Logger.log("Some required columns were not found.");
    return;
  }

  // Create the temporary Google Sheets file
  const tempSpreadsheet = SpreadsheetApp.create("Temp Report");
  const tempSheet = tempSpreadsheet.getSheets()[0];
  
  // Filter only necessary columns and apply the "Bus App Name" filter
  const filteredData = [];
  
  // Add headers to filteredData
  filteredData.push(columnNames);  // Use the column names as headers

  // Loop through each row and filter out the required columns
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const busAppName = row[columnIndexes["Bus App Name"]];
    
    // Apply the filter for "Bus App Name"
    if (busAppName === 'ABC' || busAppName === 'bca' || busAppName === 'dba' || busAppName === 'eee') {
      const filteredRow = columnNames.map(function (colName) {
        return row[columnIndexes[colName]]; // Extract the values for the required columns
      });
      filteredData.push(filteredRow);  // Add filtered row to the data
    }
  }

  // Set the filtered data into the temporary sheet
  tempSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

  // Process the data in the temp sheet (this can be customized as needed)
  Logger.log('Data successfully filtered in the temporary sheet...');
  
  // Ensure the Platops Internal Findings sheet exists and get it
  let platopsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard Sheet');
  
  // If the sheet doesn't exist, create it
  if (!platopsSheet) {
    platopsSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Dashboard Sheet');
  }
  
  // Append filtered data to the Dashboard Sheet
  if (filteredData.length > 1) {  // Ensure data exists before appending
    platopsSheet.getRange(platopsSheet.getLastRow() + 1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  }

  // Store the current row in script properties for the next run
  scriptProperties.setProperty('startRow', endRow + 1);  // Move to the next batch

  Logger.log(`Processed rows from ${startRow} to ${endRow}`);
  
  // Clear existing triggers before creating new ones
  deleteTriggers();

  // Re-run the function in 40 seconds to continue processing the next batch
  if (endRow < totalRows) {
    ScriptApp.newTrigger('processLargeData')
      .timeBased()
      .after(40 * 1000)  // Trigger the function after 40 seconds
      .create();
  } else {
    Logger.log('All rows processed.');
  }

  // Delete the temporary Google Sheets file after use
  Logger.log('Deleting temporary sheet...');
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
  Logger.log('Temporary sheet deleted.');
}

// Delete all existing triggers to prevent overload
function deleteTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}





