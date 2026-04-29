// --- CONFIGURATION ---
var SPREADSHEET_ID = '1AgRUT6d-XBhWj3V6w7Xu6sumJxlEwZ12m7oGHt8hvGU';
var DRIVE_FOLDER_ID = '1tu92dH38glARqlogXqftAeV8HG-8p9P4';

// --- HISTORY SHEET COLUMN CONSTANTS (#22) ---
var COL_DATE = 1;
var COL_ORDER_NO = 2;
var COL_PDF_LINK = 3;
var COL_STATUS = 4;
var COL_NOTES = 5;
var HISTORY_COLS = 5;

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Planning Sheet - Gorpara')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns the current server date formatted for the planning date field.
 */
function getServerDate() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy');
}

/**
 * Fetches Item and Process names from the Settings sheet for autocomplete.
 */
function getSettingsData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('settings');
  if (!sheet) return { items: [], processes: [] };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { items: [], processes: [] };

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var items = [];
  var processes = [];
  var seenItems = {};
  var seenProcesses = {};

  for (var i = 0; i < data.length; i++) {
    var item = String(data[i][0] || '').trim();
    var process = String(data[i][1] || '').trim();
    if (item && !seenItems[item]) { items.push(item); seenItems[item] = true; }
    if (process && !seenProcesses[process]) { processes.push(process); seenProcesses[process] = true; }
  }

  return { items: items, processes: processes };
}

/**
 * Returns all records from the History Sheets for the Planning Sheets popup.
 * Uses named column constants (#22).
 * Returns array of [dateStr, orderNo, status, notes, pdfUrl]
 */
function getPlanningSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('History Sheets');
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, HISTORY_COLS).getValues();

  var records = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateStr = '';
    if (row[COL_DATE - 1] instanceof Date) {
      dateStr = Utilities.formatDate(row[COL_DATE - 1], Session.getScriptTimeZone(), 'd/MMM/yyyy');
    } else {
      dateStr = String(row[COL_DATE - 1] || '');
    }

    records.push([
      dateStr,
      String(row[COL_ORDER_NO - 1] || ''),
      String(row[COL_STATUS - 1] || 'Pending'),
      String(row[COL_NOTES - 1] || ''),
      String(row[COL_PDF_LINK - 1] || '')
    ]);
  }
  return records;
}

/**
 * Updates the Status column for ALL rows matching the given Order No (#6).
 */
function updateSheetStatus(orderNo, newStatus) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('History Sheets');
    if (!sheet) return { success: false, message: 'Sheet not found' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'No records found' };

    var finder = sheet.getRange(2, COL_ORDER_NO, lastRow - 1, 1).createTextFinder(orderNo).matchEntireCell(true);
    var cell = finder.findNext();
    
    if (cell) {
      sheet.getRange(cell.getRow(), COL_STATUS).setValue(newStatus);
      return { success: true, updated: 1 };
    }

    return { success: false, message: 'Order not found' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Deletes a record from History Sheets and trashes the associated PDF from Google Drive.
 * Finds the row by Order No, extracts the PDF link, trashes the Drive file, then deletes the row.
 */
function deleteRecord(orderNo) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('History Sheets');
    if (!sheet) return { success: false, message: 'Sheet not found' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'No records found' };

    var finder = sheet.getRange(2, COL_ORDER_NO, lastRow - 1, 1).createTextFinder(String(orderNo)).matchEntireCell(true);
    var cell = finder.findNext();

    if (!cell) return { success: false, message: 'Order No "' + orderNo + '" not found.' };

    var rowNum = cell.getRow();
    var pdfUrl = String(sheet.getRange(rowNum, COL_PDF_LINK).getValue() || '');

    // Attempt to trash the PDF from Google Drive
    if (pdfUrl) {
      try {
        var fileId = extractFileId(pdfUrl);
        if (fileId) {
          var file = DriveApp.getFileById(fileId);
          file.setTrashed(true);
        }
      } catch (driveErr) {
        // PDF may have been manually deleted already — continue with row deletion
      }
    }

    // Delete the row from the sheet
    sheet.deleteRow(rowNum);

    return { success: true, message: 'Order "' + orderNo + '" deleted successfully.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Extracts a Google Drive file ID from a URL.
 */
function extractFileId(url) {
  if (!url) return null;
  // Pattern: /d/FILE_ID/ or id=FILE_ID
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  match = url.match(/id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  return null;
}

/**
 * Handles form submission with transaction safety (#4):
 * 1. Validate duplicate
 * 2. Create PDF first
 * 3. Write to sheet
 * 4. If sheet write fails, cleanup orphan PDF
 */
function processFormSubmission(formData) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('History Sheets');

  // Create sheet with headers if missing
  if (!sheet) {
    sheet = ss.insertSheet('History Sheets');
    sheet.appendRow(['Planning Date', 'Order No', 'Planning Sheet Link', 'Status', 'Notes']);
    var headerRange = sheet.getRange(1, 1, 1, HISTORY_COLS);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e3a8a');
    headerRange.setFontColor('#ffffff');
  }

  // 1. Check duplicate Order No
  var finder = sheet.createTextFinder(formData.orderNo).matchEntireCell(true);
  if (finder.findNext()) {
    return { success: false, message: 'Order No "' + formData.orderNo + '" already exists. Please use a unique Order Number.' };
  }

  // 2. Generate PDF first
  var pdfFile;
  try {
    pdfFile = createPDF(formData);
  } catch (pdfError) {
    return { success: false, message: 'PDF generation failed: ' + pdfError.toString() };
  }

  // 3. Write to sheet
  try {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, COL_DATE).setValue(formData.planningDate);
    sheet.getRange(newRow, COL_ORDER_NO).setValue(formData.orderNo).setNumberFormat('@');
    sheet.getRange(newRow, COL_PDF_LINK).setValue(pdfFile.getUrl());
    sheet.getRange(newRow, COL_STATUS).setValue('Pending');
    sheet.getRange(newRow, COL_NOTES).setValue(formData.remarks || '');
  } catch (sheetError) {
    // Cleanup: trash the orphan PDF (#4)
    try { pdfFile.setTrashed(true); } catch (cleanupErr) { /* best effort */ }
    return { success: false, message: 'Sheet write failed: ' + sheetError.toString() };
  }

  return { success: true, message: 'Saved successfully!', url: pdfFile.getUrl() };
}

/**
 * Generates a minimal black-and-white PDF (#8, #10).
 * - Only black text and borders, no colors
 * - No border-radius, no background colors
 * - Uses web-safe fonts only (#7)
 * - 1cm margins (#10)
 * - No unsupported CSS properties (#8)
 * - Footer pinned to bottom of the page
 */
function createPDF(data) {
  var template = HtmlService.createTemplateFromFile('pdf_template');
  template.data = data;
  template.generatedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy hh:mm:ss a');
  template.escapeHtml = escapeHtml; // Pass the escapeHtml function to the template

  var htmlTemplate = template.evaluate().getContent();
  var blob = Utilities.newBlob(htmlTemplate, 'text/html', 'temp.html');
  var pdfBlob = blob.getAs('application/pdf').setName(data.orderNo + '.pdf');
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  return folder.createFile(pdfBlob);
}

/**
 * Escapes HTML special characters to prevent injection in PDF (#4 related).
 */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}