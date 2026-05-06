// ╔══════════════════════════════════════════════════════════╗
// ║  V O I D T A L K  —  Signaling Relay                    ║
// ║  Google Apps Script Backend                              ║
// ║  Deploy → Web App · Execute as: Me · Access: Anyone     ║
// ╚══════════════════════════════════════════════════════════╝

var SHEET_NAME  = 'signals';
var TTL_MS      = 15 * 60 * 1000;   // signals expire after 15 min
var MAX_ROOMS   = 500;               // cap rows to avoid runaway sheets
var VERSION     = '2.1.0';

// !! SET THIS — must match the key entered in the frontend !!
var AUTH_KEY    = 'CHANGE_ME';

/* ════════════════════════════════════════════
   ENTRY POINTS
   ════════════════════════════════════════════ */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var p      = (e && e.parameter) ? e.parameter : {};
  var action = p.action || '';

  // Auth check — reject anything without the correct key
  if (p.key !== AUTH_KEY) {
    return respond({ ok: false, error: 'Unauthorized' });
  }

  try {
    var result;
    switch (action) {
      case 'ping':   result = doPing();                              break;
      case 'write':  result = doWrite(p.room, p.role, p.payload);   break;
      case 'read':   result = doRead(p.room, p.role);               break;
      case 'clear':  result = doClear(p.room);                      break;
      case 'status': result = doStatus();                           break;
      default:       throw new Error('Unknown action: "' + action + '"');
    }
    return respond({ ok: true, version: VERSION, result: result });
  } catch (err) {
    return respond({ ok: false, version: VERSION, error: err.message });
  }
}

/* ════════════════════════════════════════════
   ACTIONS
   ════════════════════════════════════════════ */

function doPing() {
  getSheet();
  return { msg: 'pong', ts: Date.now() };
}

function doStatus() {
  var sheet = getSheet();
  var count = Math.max(0, sheet.getLastRow() - 1);
  return { rows: count, maxRooms: MAX_ROOMS, ts: Date.now() };
}

function doWrite(room, role, payload) {
  if (!room)    throw new Error('Missing: room');
  if (!role)    throw new Error('Missing: role');
  if (!payload) throw new Error('Missing: payload');

  room = sanitize(room);
  role = sanitize(role);

  var sheet = getSheet();
  pruneExpired(sheet);

  var rows = sheet.getDataRange().getValues();
  var now  = Date.now();

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === room && rows[i][1] === role) {
      sheet.getRange(i + 1, 3).setValue(payload);
      sheet.getRange(i + 1, 4).setValue(now);
      return { written: true, updated: true };
    }
  }

  if (rows.length - 1 >= MAX_ROOMS) {
    throw new Error('Relay at capacity. Try again later.');
  }

  sheet.appendRow([room, role, payload, now]);
  return { written: true, updated: false };
}

function doRead(room, role) {
  if (!room) throw new Error('Missing: room');
  if (!role) throw new Error('Missing: role');

  room = sanitize(room);
  role = sanitize(role);

  var sheet = getSheet();
  var rows  = sheet.getDataRange().getValues();
  var now   = Date.now();

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === room && rows[i][1] === role) {
      var ts = Number(rows[i][3]);
      if (now - ts > TTL_MS) return { payload: null, reason: 'expired' };
      return { payload: rows[i][2], age: now - ts };
    }
  }
  return { payload: null, reason: 'not_found' };
}

function doClear(room) {
  if (!room) throw new Error('Missing: room');
  room = sanitize(room);

  var sheet   = getSheet();
  var rows    = sheet.getDataRange().getValues();
  var deleted = 0;

  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === room) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return { cleared: true, deleted: deleted };
}

/* ════════════════════════════════════════════
   HELPERS
   ════════════════════════════════════════════ */

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.create('VoidTalk Signals');

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['room', 'role', 'payload', 'timestamp']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 400);
  }
  return sheet;
}

function pruneExpired(sheet) {
  var rows = sheet.getDataRange().getValues();
  var now  = Date.now();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (now - Number(rows[i][3]) > TTL_MS) sheet.deleteRow(i + 1);
  }
}

function sanitize(str) {
  return String(str).replace(/[^a-zA-Z0-9\-_]/g, '').substring(0, 64);
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
