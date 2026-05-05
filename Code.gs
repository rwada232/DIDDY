// EpsteinTalk — Google Apps Script Signaling Relay
// Deploy as Web App: Execute as Me, Access: Anyone
// Paste this entire file into your Apps Script editor.

var SHEET_NAME = 'signals';
var TTL_MS = 10 * 60 * 1000; // auto-clear signals older than 10 minutes

/* ── Entry point ── */
function doGet(e) {
  var p = e.parameter;
  var action = p.action || '';

  try {
    var result;
    if      (action === 'ping')  result = doPing();
    else if (action === 'write') result = doWrite(p.room, p.role, p.payload);
    else if (action === 'read')  result = doRead(p.room, p.role);
    else if (action === 'clear') result = doClear(p.room);
    else throw new Error('Unknown action: ' + action);

    return respond({ ok: true, result: result });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

/* ── Actions ── */
function doPing() {
  getSheet(); // will throw if sheet can't be accessed
  return { msg: 'pong' };
}

function doWrite(room, role, payload) {
  if (!room || !role || !payload) throw new Error('Missing room, role, or payload');
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var now = Date.now();

  // Overwrite existing row for this room+role
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === room && rows[i][1] === role) {
      sheet.getRange(i + 1, 3).setValue(payload);
      sheet.getRange(i + 1, 4).setValue(now);
      return { written: true };
    }
  }

  // Append new row
  sheet.appendRow([room, role, payload, now]);
  return { written: true };
}

function doRead(room, role) {
  if (!room || !role) throw new Error('Missing room or role');
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var now = Date.now();

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === room && rows[i][1] === role) {
      var ts = rows[i][3];
      if (now - ts > TTL_MS) return { payload: null }; // expired
      return { payload: rows[i][2] };
    }
  }
  return { payload: null };
}

function doClear(room) {
  if (!room) throw new Error('Missing room');
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();

  // Delete matching rows bottom-up so indexes stay valid
  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === room) {
      sheet.deleteRow(i + 1);
    }
  }
  return { cleared: true };
}

/* ── Helpers ── */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    // If no spreadsheet is bound, create one automatically
    ss = SpreadsheetApp.create('EpsteinTalk Signals');
    // Note: re-bind this script to the new sheet manually if needed
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Write header row
    sheet.appendRow(['room', 'role', 'payload', 'timestamp']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
