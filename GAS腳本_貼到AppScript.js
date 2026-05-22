// PMO Vote System - Google Apps Script v3
// IMPORTANT: Replace YOUR_SHEET_ID below with your actual Google Sheet ID
// Sheet ID is between /d/ and /edit in your spreadsheet URL

var SHEET_ID = 'YOUR_SHEET_ID';

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName('votes') || ss.insertSheet('votes');
}

// GET: returns all votes as JSONP (supports ?callback=fn for cross-origin use)
// Even on error, returns JSONP so the browser callback always fires
function doGet(e) {
  var callback = '';
  try { callback = e.parameter.callback || ''; } catch(ignore) {}

  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    var votes = [];

    if (lastRow > 1) {
      var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < data.length; i++) {
        try {
          var obj = JSON.parse(data[i][0]);
          if (obj && obj.name) votes.push(obj);
        } catch (parseErr) {}
      }
    }

    var json = JSON.stringify(votes);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    // Always return valid JSONP even on error so browser callback fires
    var errJson = JSON.stringify({ error: err.toString() });
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + errJson + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errJson)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// POST: save / delete / clear
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    var sheet = getSheet();

    function findRow(name) {
      var last = sheet.getLastRow();
      if (last <= 1) return -1;
      var data = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < data.length; i++) {
        try {
          if (JSON.parse(data[i][0]).name === name) return i + 2;
        } catch (err) {}
      }
      return -1;
    }

    if (action === 'save') {
      var rec = JSON.stringify({
        name: payload.name,
        dates: payload.dates,
        restaurants: payload.restaurants,
        ts: payload.ts
      });
      var saveRow = findRow(payload.name);
      if (saveRow > 0) {
        sheet.getRange(saveRow, 1).setValue(rec);
      } else {
        sheet.appendRow([rec]);
      }

    } else if (action === 'delete') {
      var delRow = findRow(payload.name);
      if (delRow > 0) sheet.deleteRow(delRow);

    } else if (action === 'clear') {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    }

    return ContentService.createTextOutput('ok');

  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.toString());
  }
}
