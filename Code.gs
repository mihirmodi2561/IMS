/**
 * ========================================
 * CODE.GS - Main Functions
 * Based on working LATEST_code.gs
 * ========================================
 */

// Global variable to cache spreadsheet ID
var SPREADSHEET_ID = null;

// Serve the web app
function doGet() {
  // Cache the ID here — this is the only place getActiveSpreadsheet() works reliably
  SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Inventory Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include HTML files (for modular structure)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Get spreadsheet — works from both IDE and published web app
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Setup all sheets
function setupAllSheets() {
  setupUsersSheet();
  
  return {
    success: true,
    message: 'Users sheet created successfully!'
  };
}

// Get spreadsheet URL
function getSpreadsheetUrl() {
  try {
    var ss = getSpreadsheet();
    return {
      success: true,
      url: ss.getUrl()
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}