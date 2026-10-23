/**
 * Google Apps Script for Don Joe & Ann Mary Wedding RSVP & Wishes
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open Google Sheets (create a new sheet or open existing one).
 * 2. In Google Sheets menu, click: Extensions -> Apps Script.
 * 3. Delete any default code in Code.gs and paste this entire code.
 * 4. Click "Deploy" (top right) -> "New deployment".
 * 5. Click the gear icon next to "Select type" -> choose "Web app".
 * 6. Set:
 *    - Description: "Wedding RSVP Web App"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone" (crucial so guests can submit without signing into Google)
 * 7. Click "Deploy", authorize permissions when prompted.
 * 8. Copy the Web App URL (ends with /exec).
 * 9. Paste that URL into js/rsvp.js as SCRIPT_URL.
 */

var SHEET_NAME = 'DonJoe-AnnMary'; // Name of the tab in your spreadsheet

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000); // Wait up to 10 seconds for concurrent writes

    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(SHEET_NAME) || doc.getSheets()[0];

    // If the sheet is empty, create header columns
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Full Name', 'Phone Number', 'Attendance', 'Guests', 'Wishes / Message']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f3eaf5');
      sheet.setFrozenRows(1);
    }

    var params = e.parameter || {};
    var timestamp = params.timestamp || new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
    var name = params.name || '';
    var phone = params.phone || '';
    var attendance = params.attendance || 'Joyfully Accepts';
    var guests = params.guests || (attendance === 'Joyfully Accepts' ? '1' : '0');
    var wish = params.wish || '';

    // Append response to the sheet
    sheet.appendRow([timestamp, name, phone, attendance, guests, wish]);

    lock.releaseLock();

    return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName(SHEET_NAME) || doc.getSheets()[0];
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ wishes: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var rows = sheet.getDataRange().getValues();
    var wishes = [];

    // Skip header row
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var name = row[1];
      var wish = row[5];
      if (name && wish && String(wish).trim().length > 0) {
        wishes.push({
          name: name,
          wish: wish
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ wishes: wishes }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ wishes: [], error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
