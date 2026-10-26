/**
 * Team Day Survey — Google Apps Script backend
 * One sheet tab per team (TEAMS below): Timestamp | Name | Slots | Note
 *   Slots = comma-separated slot ids from index.html (e.g. "mon 14.30,wed 15.45")
 *
 * GET  ?team=chess              -> { ok, total, counts: { slotId: n } }   (no names)
 * GET  ?team=chess&name=Nok     -> also { mine: { slots, note } | null }
 * POST (text/plain JSON) { team, name, slots: [...], note }
 *                               -> same name (case-insensitive) is updated, otherwise appended
 * POST { team, name, remove: true } -> deletes that person's row
 */

var TEAMS = ['chess', 'boardgame'];   // must match CONFIG.teams ids in index.html
var HEADERS = ['Timestamp', 'Name', 'Slots', 'Note'];

function doGet(e) {
  try {
    var team = team_(e.parameter.team);
    var name = clean_(e.parameter.name, 60);
    return json_(result_(team, name));
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var team = team_(body.team);
    var name = clean_(body.name, 60);
    if (!name) throw new Error('name required');

    var sheet = sheet_(team);
    var row = findRow_(sheet, name);

    if (body.remove) {
      if (row) sheet.deleteRow(row);
    } else {
      var slots = (Array.isArray(body.slots) ? body.slots : [])
        .map(function (s) { return clean_(s, 40); })
        .filter(String);
      if (!slots.length) throw new Error('pick at least one slot');
      var values = [[new Date(), name, slots.join(','), clean_(body.note, 200)]];
      if (row) sheet.getRange(row, 1, 1, HEADERS.length).setValues(values);
      else sheet.appendRow(values[0]);
    }
    return json_(result_(team, name));
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function team_(t) {
  if (TEAMS.indexOf(t) < 0) throw new Error('unknown team');
  return t;
}

function sheet_(team) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(team) || ss.insertSheet(team);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

// Counts per slot for the whole team; the asking person's own answer only if a name is given
function result_(team, name) {
  var sheet = sheet_(team);
  var last = sheet.getLastRow();
  var rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var counts = {}, total = 0, mine = null, key = norm_(name);
  rows.forEach(function (r) {
    if (!r[1]) return;
    total++;
    var slots = String(r[2]).split(',').filter(String);
    slots.forEach(function (s) { counts[s] = (counts[s] || 0) + 1; });
    if (key && norm_(r[1]) === key) mine = { slots: slots, note: String(r[3] || '').replace(/^'/, '') };
  });
  return { ok: true, total: total, counts: counts, mine: mine };
}

function findRow_(sheet, name) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var key = norm_(name);
  var names = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (norm_(names[i][0]) === key) return i + 2;
  }
  return 0;
}

function norm_(v) {
  return String(v).replace(/^'/, '').trim().toLowerCase();
}

// Trim, cap length, and stop text being read as a formula (=, +, -, @)
function clean_(v, max) {
  var s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
