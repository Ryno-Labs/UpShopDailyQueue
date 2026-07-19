/**
 * Color New — Daily Queue Board  (backend)
 * Read-only JSON feed for the PWA. Reads the EXISTING sheet as-is:
 * no form changes, no new columns, no edits to your data.
 *
 * Deploy: Extensions > Apps Script > paste this > Deploy > New deployment
 *   Type: Web app  |  Execute as: Me  |  Who has access: Anyone
 * Copy the /exec URL and paste it into index.html (SCRIPT_URL).
 */

// Tab names are matched loosely (trimmed), so the trailing spaces in your
// sheet names ("Dealer Upholstery1 ", "Retail Appointments ") don't matter.
var DEALER_TAB = 'Dealer Upholstery1';
var RETAIL_TAB = 'Retail Appointments';

function doGet() {
  var out = { ok: true, generatedAt: new Date().toISOString(), dealer: [], retail: [] };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    out.dealer = readDealer_(findSheet_(ss, DEALER_TAB), tz);
    out.retail = readRetail_(findSheet_(ss, RETAIL_TAB), tz);
  } catch (err) {
    out.ok = false;
    out.error = String(err);
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- readers ---------- */

function readDealer_(sheet, tz) {
  var rows = getRows_(sheet);
  if (!rows) return [];
  var h = rows.header;
  var iDealer = col_(h, 'Dealer');
  var iVehicle = col_(h, 'Year and Model');
  var iDesc = col_(h, 'Brief Description of Repair');
  var iPrice = col_(h, 'Price Quoted');
  var iStamp = col_(h, 'Timestamp'); // entry date drives the dealer queue

  var list = [];
  rows.data.forEach(function (r) {
    var dealer = str_(r[iDealer]);
    var vehicle = str_(r[iVehicle]);
    var desc = str_(r[iDesc]);
    // skip fully blank rows
    if (!dealer && !vehicle && !desc) return;
    list.push({
      dealer: dealer,
      vehicle: vehicle,
      description: desc,
      price: num_(r[iPrice]),
      entryDate: toISODate_(r[iStamp], tz)
    });
  });
  return list;
}

function readRetail_(sheet, tz) {
  var rows = getRows_(sheet);
  if (!rows) return [];
  var h = rows.header;
  var iName = col_(h, 'Name');
  var iVehicle = col_(h, 'Vehicle Year & Model');
  var iDesc = col_(h, 'Brief Description of Repair');
  var iAppt = col_(h, 'Preferred Appointment date');
  var iSrc = col_(h, 'How did you hear about us?');
  // NOTE: phone ("Best Contact Number") and email are intentionally NOT read.
  // Those columns stay in the sheet and never enter this public feed.

  var list = [];
  rows.data.forEach(function (r) {
    var name = str_(r[iName]);
    var vehicle = str_(r[iVehicle]);
    var desc = str_(r[iDesc]);
    if (!name && !vehicle && !desc) return;
    list.push({
      name: name,
      vehicle: vehicle,
      description: desc,
      apptDate: toISODate_(r[iAppt], tz),
      source: str_(r[iSrc])
    });
  });
  return list;
}

/* ---------- helpers ---------- */

function findSheet_(ss, wanted) {
  var target = norm_(wanted);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (norm_(sheets[i].getName()) === target) return sheets[i];
  }
  // fallback: starts-with match (handles "Dealer Upholstery1" vs variants)
  for (var j = 0; j < sheets.length; j++) {
    if (norm_(sheets[j].getName()).indexOf(target) === 0) return sheets[j];
  }
  throw new Error('Tab not found: ' + wanted);
}

function getRows_(sheet) {
  if (!sheet) return null;
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  return { header: values[0], data: values.slice(1) };
}

// find a column index by header name, ignoring case + surrounding spaces
function col_(header, name) {
  var want = norm_(name);
  for (var i = 0; i < header.length; i++) {
    if (norm_(header[i]) === want) return i;
  }
  return -1;
}

function norm_(v) {
  return String(v == null ? '' : v).trim().toLowerCase();
}

function str_(v) {
  if (v == null) return '';
  return String(v).trim();
}

function num_(v) {
  if (v === '' || v == null) return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

// Return a YYYY-MM-DD string or '' — handles real Dates and typo text like "6/17/0025".
function toISODate_(v, tz) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  // m/d/yyyy  or  m/d/yy  (0025 -> 2025)
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{1,4})$/);
  if (m) {
    var mo = +m[1], da = +m[2], yr = +m[3];
    if (yr < 100) yr += 2000;          // 25 -> 2025
    else if (yr < 1000) yr += 2000;    // 0025 -> 2025
    return pad_(yr, 4) + '-' + pad_(mo, 2) + '-' + pad_(da, 2);
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return '';
}

function pad_(n, len) {
  var s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}
