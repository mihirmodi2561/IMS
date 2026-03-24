/**
 * ========================================
 * CODE.GS - COMPLETE WORKING VERSION
 * Replace your entire Code.gs with this
 * ========================================
 */

// Global variable to cache spreadsheet ID
var SPREADSHEET_ID = null;

// Serve the web app
/*function doGet() {
  SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Adams Telecom Systems')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
*/

function doGet() {
  SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Adams Telecom Systems')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Get spreadsheet - SINGLE SOURCE OF TRUTH
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
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