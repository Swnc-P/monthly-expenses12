// --- Code.gs สำหรับ Google Apps Script ---
var SHEET_NAME = "Database";

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Email", "DataJson", "LastUpdated"]);
  }
}

function doPost(e) {
  try {
    setupSheet();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    var data = JSON.parse(e.postData.contents);
    var email = data.email;
    var payloadData = JSON.stringify(data.payload);
    var timestamp = new Date();
    
    var foundRow = -1;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === email) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow > -1) {
      sheet.getRange(foundRow, 2).setValue(payloadData);
      sheet.getRange(foundRow, 3).setValue(timestamp);
    } else {
      sheet.appendRow([email, payloadData, timestamp]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    setupSheet();
    var email = e.parameter.email;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    var rows = sheet.getDataRange().getValues();
    
    var payloadData = "{}";
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === email) {
        payloadData = rows[i][1];
        break;
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({status: "success", payload: JSON.parse(payloadData || "{}")}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "success", payload: {}}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
