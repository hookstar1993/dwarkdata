/**
 * @license
 * Copyright 2024 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

// ============================================================================
//   CONFIGURATION
// ============================================================================
const SHEET_ID = '1OOthoN3XS_GfwYsbQC-WSpJer9FMZ-AYK7E4xuxE4y8';
const UNIQUE_VALUE_COLUMNS = ['guest_name', 'guest_initials', 'departure_city', 'arrival_city'];

// ============================================================================
//   HELPER FUNCTIONS
// ============================================================================
function getSheet() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheets()[0];
    if (!sheet) throw new Error("No sheets were found in the spreadsheet file.");
    console.log(`Successfully accessed sheet: "${sheet.getName()}"`);
    return sheet;
  } catch (e) {
    console.error(`CRITICAL ERROR in getSheet(): ${e.stack}`);
    return null;
  }
}

function getUniqueColumnValues(sheet, columnName) {
  // ROBUSTNESS CHECK: Prevent crashes if run manually from the editor.
  if (!sheet || typeof sheet.getLastRow !== 'function') {
      const message = "DEV INFO: getUniqueColumnValues was called with an invalid sheet object. This is expected if you are running this function manually from the script editor. Please only test the app via the deployed web URL.";
      console.error(message);
      // Return an empty array so the script doesn't crash.
      return [];
  }
  try {
    if (sheet.getLastRow() < 2) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headers.indexOf(columnName);
    if (colIndex === -1) return [];
    const allValues = sheet.getRange(2, colIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    return [...new Set(allValues.flat().filter(String))].sort();
  } catch (e) {
    console.error(`Error in getUniqueColumnValues for "${columnName}": ${e.stack}`);
    return [];
  }
}

function getSheetDataAsJSON(sheet) {
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data.shift() || [];
  const jsonRows = data.map((row, index) => {
    let obj = { rowIndex: index + 2 };
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
  const uniqueValues = {};
  UNIQUE_VALUE_COLUMNS.forEach(colName => {
    uniqueValues[colName] = getUniqueColumnValues(sheet, colName);
  });
  return { headers, rows: jsonRows, uniqueValues };
}

// ============================================================================
//   MAIN WEB APP FUNCTIONS (These are the only functions that are called by the web app)
// ============================================================================

function doGet(e) {
  console.log("doGet: Request received.");
  const sheet = getSheet();
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'ERROR', message: 'Could not access the spreadsheet.' })).setMimeType(ContentService.MimeType.JSON);
  }
  const data = getSheetDataAsJSON(sheet);
  return ContentService.createTextOutput(JSON.stringify({ status: 'SUCCESS', ...data })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const action = e.parameter.action;
  console.log(`doPost: Request received for action: "${action}"`);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let actionResult = {};
  let successfulSheet = null; 

  try {
    const sheet = getSheet();
    if (!sheet) throw new Error("Could not access spreadsheet to perform action.");
    successfulSheet = sheet; 

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const params = e.parameter;
    const lastRow = sheet.getLastRow();
    
    console.log(`Parameters received: ${JSON.stringify(params)}`);

    switch (action) {
      case 'create':
        const newRow = headers.map(header => params[header] || "");
        sheet.appendRow(newRow);
        console.log(`SUCCESS: New row added.`);
        actionResult = { status: 'SUCCESS', message: 'Record added successfully' };
        break;

      case 'update':
        const rowIndexToUpdate = parseInt(params.rowIndex);
        console.log(`Attempting to update row: ${rowIndexToUpdate}`);
        if (!rowIndexToUpdate || rowIndexToUpdate < 2 || rowIndexToUpdate > lastRow) {
            console.error(`VALIDATION FAILED: Invalid row index for update: ${rowIndexToUpdate}. Last row is ${lastRow}.`);
            throw new Error(`Invalid row index for update: ${rowIndexToUpdate}.`);
        }
        const updatedData = headers.map(header => params[header] || "");
        sheet.getRange(rowIndexToUpdate, 1, 1, headers.length).setValues([updatedData]);
        console.log(`SUCCESS: Row ${rowIndexToUpdate} updated.`);
        actionResult = { status: 'SUCCESS', message: 'Record updated successfully' };
        break;

      case 'delete':
        const rowIndexToDelete = parseInt(params.rowIndex);
        console.log(`Attempting to delete row: ${rowIndexToDelete}`);
        if (!rowIndexToDelete || rowIndexToDelete < 2 || rowIndexToDelete > lastRow) {
             console.error(`VALIDATION FAILED: Invalid row index for delete: ${rowIndexToDelete}. Last row is ${lastRow}.`);
            throw new Error(`Invalid row index for delete: ${rowIndexToDelete}.`);
        }
        sheet.deleteRow(rowIndexToDelete);
        console.log(`SUCCESS: Row ${rowIndexToDelete} deleted.`);
        actionResult = { status: 'SUCCESS', message: 'Record deleted successfully' };
        break;
      
      case 'duplicate':
        const rowIndexToDuplicate = parseInt(params.rowIndex);
        console.log(`Attempting to duplicate row: ${rowIndexToDuplicate}`);
        if (!rowIndexToDuplicate || rowIndexToDuplicate < 2 || rowIndexToDuplicate > lastRow) {
            console.error(`VALIDATION FAILED: Invalid row index for duplicate: ${rowIndexToDuplicate}. Last row is ${lastRow}.`);
            throw new Error(`Invalid row index for duplicate: ${rowIndexToDuplicate}.`);
        }
        const rowData = sheet.getRange(rowIndexToDuplicate, 1, 1, headers.length).getDisplayValues();
        sheet.appendRow(rowData[0]);
        console.log(`SUCCESS: Row ${rowIndexToDuplicate} duplicated.`);
        actionResult = { status: 'SUCCESS', message: 'Record duplicated successfully' };
        break;

      default:
        throw new Error(`Invalid action specified: "${action}"`);
    }
  } catch (error) {
    console.error(`CRITICAL ERROR during action "${action}": ${error.stack}`);
    actionResult = { status: 'ERROR', message: `Server-side error during '${action}': ${error.message}` };
  } finally {
    lock.releaseLock();
  }

  if (actionResult.status === 'SUCCESS' && successfulSheet) {
      console.log("Action was successful. Fetching fresh data to send back to client.");
      const freshData = getSheetDataAsJSON(successfulSheet);
      actionResult = {...actionResult, ...freshData};
  }
  
  return ContentService.createTextOutput(JSON.stringify(actionResult)).setMimeType(ContentService.MimeType.JSON);
}

