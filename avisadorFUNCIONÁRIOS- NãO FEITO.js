/**
 * Manually triggered function to check folders and update the sheet for each employee.
 */
function manualVerifyFiles() {
  const recFolderId = "1Bl8mVpTYeqzz-_X_Gyi6qJyInTU63BEl";
  const spreadsheetId = "1vmw6AWwH36EDW99kOoBwrLxHO2PAsydGwc5IdMyYahc";
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const recFolder = DriveApp.getFolderById(recFolderId);
  
  const sheets = spreadsheet.getSheets();
  
  sheets.forEach(sheet => {
    const year = sheet.getName();
    if (year === "Base-Sheet") return; // Skip the Base-Sheet

    let headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    let columns = {};
    
    for (let i = 0; i < headers.length; i++) {
      if (/^\d{2}\/\d{4}$/.test(headers[i])) {
        columns[headers[i]] = i + 1;
      }
    }

    let values = sheet.getRange(3, 2, sheet.getLastRow(), 1).getValues();
    let employees = values.map(row => row[0]).filter(initials => initials);

    employees.forEach((initials, rowIndex) => {
      Object.keys(columns).forEach(monthYear => {
        let [month, year] = monthYear.split("/");
        let cell = sheet.getRange(rowIndex + 3, columns[monthYear]);
        let fileExists = checkFileExists(recFolder, year, month, initials);
        
        if (fileExists) {
          cell.setFormula(`=HYPERLINK("${fileExists}", "LINK")`);
          cell.setBackground("green");
        } else {
          cell.setValue("em falta");
          cell.setBackground("red");
        }
      });
    });
  });
}

/**
 * Checks if a file exists in the correct subfolder.
 * @param {GoogleAppsScript.Drive.Folder} recFolder The root REC folder.
 * @param {string} year The year of the file.
 * @param {string} month The month of the file.
 * @param {string} initials The initials of the employee.
 * @return {string|boolean} The file URL if it exists, otherwise false.
 */
function checkFileExists(recFolder, year, month, initials) {
  const yearFolder = findOrCreateSubfolder(recFolder, year);
  const monthFolder = findOrCreateSubfolder(yearFolder, month);
  const files = monthFolder.getFiles();
  
  while (files.hasNext()) {
    let file = files.next();
    if (file.getName().includes(initials)) {
      return file.getUrl();
    }
  }
  return false;
}

/**
 * Finds or creates a subfolder inside a parent folder.
 * @param {GoogleAppsScript.Drive.Folder} parentFolder The parent folder.
 * @param {string} folderName The name of the subfolder.
 * @return {GoogleAppsScript.Drive.Folder} The subfolder.
 */
function findOrCreateSubfolder(parentFolder, folderName) {
  let subfolders = parentFolder.getFoldersByName(folderName);
  return subfolders.hasNext() ? subfolders.next() : parentFolder.createFolder(folderName);
}
