/**
 * DTI Methodology Validation — Google Apps Script backend
 *
 * SETUP (5 minutes):
 * 1. Go to https://sheets.google.com → Create a blank spreadsheet
 *    Name it "DTI Validation Responses"
 * 2. In the spreadsheet, go to Extensions → Apps Script
 * 3. Delete everything in Code.gs and paste this entire file
 * 4. Click Deploy → New deployment
 *    - Click the gear icon → Web app
 *    - Description: "DTI Validation"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click Deploy → Authorize access → choose your Google account → Allow
 * 6. Copy the Web app URL (starts with https://script.google.com/macros/s/...)
 * 7. Send that URL to Festus — he will update the live form
 */

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = getOrCreateSheet();
    var entries = payload.entries || [];
    var details = payload.details || {};
    var respondentId = payload.respondentId || "";

    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i];
      var row = [
        entry.entryId || "",
        respondentId,
        details.name || "",
        details.org || "",
        details.contact || "",
        details.date || "",
        entry.submittedAt || new Date().toISOString(),
        entry.section || "",
        entry.item || "",
        entry.itemText || "",
        entry.type || "",
        entry.verdict || "",
        entry.comment || ""
      ];

      var existing = findRow(ss, entry.entryId);
      if (existing > -1) {
        ss.getRange(existing, 1, 1, row.length).setValues([row]);
      } else {
        ss.appendRow(row);
      }
    }

    SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify({ok: true, received: entries.length}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ok: true, status: "DTI Validation backend is running."}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Responses");
  if (!sheet) {
    sheet = ss.insertSheet("Responses");
    sheet.appendRow([
      "Entry ID",
      "Respondent ID",
      "Name",
      "Organisation",
      "Contact",
      "Date",
      "Submitted At",
      "Section",
      "Item",
      "Item Text",
      "Type",
      "Verdict",
      "Comment"
    ]);
    sheet.getRange(1, 1, 1, 13).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(10, 300);
    sheet.setColumnWidth(13, 400);
  }
  return sheet;
}

function findRow(sheet, entryId) {
  if (!entryId) return -1;
  var data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === entryId) return i + 1;
  }
  return -1;
}
