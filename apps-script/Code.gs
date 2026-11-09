/**
 * TEAM ATTENDANCE – Google Sheet backend
 * --------------------------------------
 * 1. Create a NEW blank Google Sheet > Extensions > Apps Script
 * 2. Paste this file into Code.gs, add an HTML file named "Index" and paste Index.html
 * 3. Choose the function "setup" at the top and press Run (only once – builds all sheets)
 * 4. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
 *
 * First admin login:  Login ID "admin"  password "admin123"  (reset it in the admin panel!)
 *
 * SHEETS CREATED
 *  Employees         – one row per employee (master list)
 *  Attendance        – one row per employee per day (login + logout side by side)
 *  Employee Summary  – auto-calculated monthly report per employee (pick the month in B1)
 *  Branch Summary    – auto-calculated monthly report per branch (pick the month in B1)
 *  Branch Transfers  – history of every branch change
 *  Branches / SMs / Locations – dropdown lists used by the website
 *  Settings          – reminder time, late time, weekly off
 */

const TZ = 'Asia/Kolkata';

// ID of the Google Sheet that stores all data (from its link: /spreadsheets/d/<ID>/edit)
const SHEET_ID = '13NcJsjTp8ECfZ9iocGIQaAXtrkzfF5Ai_STyYir4ADU';
function ss_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID); }

const COLS = {
  'Employees': ['Employee ID', 'First Name', 'Last Name', 'Full Name', 'Phone', 'Branch', 'Joining Date',
                'Role', 'Status', 'Created On', 'Password Hash'],
  'Attendance': ['Date', 'Day', 'Employee ID', 'Employee Name', 'Branch', 'SM', 'CP Firm', 'Location',
                 'Login Time', 'Late', 'Logout Time', 'Hours Worked', 'Total Calls', 'Connect Calls', 'Connect %',
                 'Lead of the Day', 'Site Visit Done', 'Tomorrow Site Visit', 'Booking', 'Site Visit Confirmed WTD', 'Status',
                 'Login Posted At', 'Logout Posted At', 'Admin Note', 'Login Message', 'Logout Message'],
  'Branch Transfers': ['Changed On', 'Employee ID', 'Employee Name', 'From Branch', 'To Branch', 'Effective Date', 'Changed By'],
  'Branches': ['Branch'],
  'SMs': ['SM Name'],
  'Locations': ['Location'],
  'Settings': ['Setting', 'Value', 'Description'],
};
const DATE_COLS = ['Date', 'Joining Date', 'Effective Date'];
const NUM_COLS = ['Total Calls', 'Connect Calls', 'Connect %', 'Site Visit Done', 'Tomorrow Site Visit', 'Site Visit Confirmed WTD', 'Booking'];
const SUMMARY_ROWS = 300;     // employee rows prepared in "Employee Summary"

/* ================= Web page ================= */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Team Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= One-time setup ================= */
function setup() {
  const ss = ss_();
  Object.keys(COLS).forEach(function (name) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    s.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]);
    s.setFrozenRows(1);
  });
  if (users_().length === 0)
    appendObj_('Employees', { 'Employee ID': 'admin', 'First Name': 'Admin', 'Last Name': '', 'Full Name': 'Admin',
      'Role': 'Admin', 'Status': 'Active', 'Created On': now_('yyyy-MM-dd HH:mm'), 'Password Hash': hash_('admin', 'admin123') });
  seed_('Branches', ['Elements Callers', 'Airoli Commercial Calling', 'Eden Garden Airoli Caller',
                     'Upvan Callers', 'Nerul Commercial Callers', 'Juinagar Commercial Calling']);
  seed_('SMs', ['sohail sir', 'Aman sir', 'anand sir', 'Saurabh Khandagale']);
  seed_('Locations', ['on site', 'office', 'field visit']);
  [['Logout reminder time', '19:00', 'After this time (24-hour), callers who have not posted the logout see "Kindly post the logout"'],
   ['Late after', '10:00', 'Login after this time (24-hour) is marked Late'],
   ['Weekly off', 'Monday', 'This day is not counted as absent (write None for no weekly off)']].forEach(function (r) {
    if (!setting_(r[0], '')) append_('Settings', r);
  });
  buildSummaries_(ss);
  try { format_(ss); } catch (e) { Logger.log('Formatting skipped: ' + e); }
  try { installNightly_(); } catch (e) { Logger.log('Trigger skipped: ' + e); }
  try {
    const s1 = ss.getSheetByName('Sheet1');
    if (s1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  } catch (e) {}
  return 'Setup done';
}

/* ================= Summary sheets (formulas – update automatically) ================= */
function buildSummaries_(ss) {
  const A = function (c) { return "Attendance!$" + colLetter_('Attendance', c) + ":$" + colLetter_('Attendance', c); };
  const E = function (c) { return "Employees!" + colLetter_('Employees', c) + "2:" + colLetter_('Employees', c); };
  const S = 'DATE(YEAR($B$1),MONTH($B$1),1)', END = 'EOMONTH($B$1,0)', PAST = 'MIN(EOMONTH($B$1,0),TODAY()-1)';
  const first = new Date(); first.setDate(1);

  /* ---- Employee Summary ---- */
  let s = ss.getSheetByName('Employee Summary') || ss.insertSheet('Employee Summary');
  s.clear();
  s.getRange(1, 1, 3, 2).setValues([['Month', first], ['Weekly off', ''], ['Weekend code', '']]);
  s.getRange(2, 2).setFormula('=IFERROR(VLOOKUP("Weekly off",Settings!A:B,2,FALSE),"Monday")');
  // 7-character weekend code for NETWORKDAYS.INTL (Mon..Sun, 1 = off day)
  s.getRange(3, 2).setFormula('=IFERROR(REPT("0",MATCH(B2,{"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"},0)-1)&"1"&REPT("0",7-MATCH(B2,{"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"},0)),"0000000")');
  s.getRange(1, 4).setValue('Change the month in B1 (any date in that month). Everything below updates automatically.');
  s.getRange(3, 4).setValue('(used by the formulas, do not edit)');
  const eh = ['Employee ID', 'Employee Name', 'Branch', 'Joining Date', 'Working Days', 'Present', 'Absent', 'Attendance %',
              'Late Logins', 'Logout Missed', 'Total Calls', 'Connect Calls', 'Connect %', 'Leads', 'Site Visits', 'Bookings'];
  s.getRange(5, 1, 1, eh.length).setValues([eh]);
  s.getRange(6, 1).setFormula('=IFERROR(SORT(FILTER({' + E('Employee ID') + ',' + E('Full Name') + ',' + E('Branch') + ',' + E('Joining Date') +
    '},' + E('Role') + '="Caller",' + E('Status') + '="Active"),3,TRUE,2,TRUE),"")');
  const id = A('Employee ID'), dt = A('Date');
  const f = [];
  for (let r = 6; r < 6 + SUMMARY_ROWS; r++) {
    const ID = '$A' + r, J = 'IF($D' + r + '="",0,$D' + r + ')';
    const inMonth = id + ',' + ID + ',' + dt + ',">="&' + S + ',' + dt + ',"<="&' + END;
    const gone = id + ',' + ID + ',' + dt + ',">="&MAX(' + S + ',' + J + '),' + dt + ',"<="&' + PAST;
    const W = (x) => '=IF(' + ID + '="","",' + x + ')';
    f.push([
      W('MAX(0,NETWORKDAYS.INTL(MAX(' + S + ',' + J + '),' + PAST + ',$B$3))'),
      W('COUNTIFS(' + inMonth + ')'),
      W('MAX(0,$E' + r + '-COUNTIFS(' + gone + ',' + A('Day') + ',"<>"&LEFT($B$2,3)))'),
      '=IF(OR(' + ID + '="",$F' + r + '+$G' + r + '=0),"",$F' + r + '/($F' + r + '+$G' + r + '))',
      W('COUNTIFS(' + inMonth + ',' + A('Late') + ',"Yes")'),
      W('COUNTIFS(' + gone + ',' + A('Status') + ',"Logout pending")+COUNTIFS(' + gone + ',' + A('Status') + ',"Logout missed")'),
      W('SUMIFS(' + A('Total Calls') + ',' + inMonth + ')'),
      W('SUMIFS(' + A('Connect Calls') + ',' + inMonth + ')'),
      '=IF(OR(' + ID + '="",N($K' + r + ')=0),"",$L' + r + '/$K' + r + ')',
      W('SUMIFS(' + A('Lead of the Day') + ',' + inMonth + ')'),
      W('SUMIFS(' + A('Site Visit Done') + ',' + inMonth + ')'),
      W('SUMIFS(' + A('Booking') + ',' + inMonth + ')'),
    ]);
  }
  s.getRange(6, 5, SUMMARY_ROWS, 12).setFormulas(f);
  s.setFrozenRows(5);

  /* ---- Branch Summary ---- */
  s = ss.getSheetByName('Branch Summary') || ss.insertSheet('Branch Summary');
  s.clear();
  s.getRange(1, 1, 1, 2).setValues([['Month', first]]);
  s.getRange(1, 4).setValue('Change the month in B1. Present days and calls are counted under the branch the caller was in on that day.');
  const bh = ['Branch', 'Active Callers', 'Present Days', 'Late Logins', 'Logout Missed', 'Total Calls', 'Connect Calls',
              'Connect %', 'Leads', 'Site Visits', 'Bookings'];
  s.getRange(3, 1, 1, bh.length).setValues([bh]);
  s.getRange(4, 1).setFormula('=IFERROR(FILTER(Branches!A2:A,Branches!A2:A<>""),"")');
  const br = A('Branch'), g = [];
  for (let r = 4; r < 64; r++) {
    const B = '$A' + r;
    const inMonth = br + ',' + B + ',' + dt + ',">="&' + S + ',' + dt + ',"<="&' + END;
    const past = br + ',' + B + ',' + dt + ',">="&' + S + ',' + dt + ',"<="&' + PAST;
    const W = (x) => '=IF(' + B + '="","",' + x + ')';
    g.push([
      W('COUNTIFS(Employees!$' + colLetter_('Employees', 'Branch') + ':$' + colLetter_('Employees', 'Branch') + ',' + B +
        ',Employees!$' + colLetter_('Employees', 'Role') + ':$' + colLetter_('Employees', 'Role') + ',"Caller",Employees!$' +
        colLetter_('Employees', 'Status') + ':$' + colLetter_('Employees', 'Status') + ',"Active")'),
      W('COUNTIFS(' + inMonth + ')'),
      W('COUNTIFS(' + inMonth + ',' + A('Late') + ',"Yes")'),
      W('COUNTIFS(' + past + ',' + A('Status') + ',"Logout pending")+COUNTIFS(' + past + ',' + A('Status') + ',"Logout missed")'),
      W('SUMIFS(' + A('Total Calls') + ',' + inMonth + ')'),
      W('SUMIFS(' + A('Connect Calls') + ',' + inMonth + ')'),
      '=IF(OR(' + B + '="",N($F' + r + ')=0),"",$G' + r + '/$F' + r + ')',
      W('SUMIFS(' + A('Lead of the Day') + ',' + inMonth + ')'),
      W('SUMIFS(' + A('Site Visit Done') + ',' + inMonth + ')'),
      W('SUMIFS(' + A('Booking') + ',' + inMonth + ')'),
    ]);
  }
  s.getRange(4, 2, 60, 10).setFormulas(g);
  s.setFrozenRows(3);
}

/* ================= Look & feel of the sheets ================= */
function format_(ss) {
  const GREEN = '#128c7e';
  const order = ['Employees', 'Attendance', 'Employee Summary', 'Branch Summary', 'Branch Transfers', 'Branches', 'SMs', 'Locations', 'Settings'];
  order.forEach(function (n, i) {
    const s = ss.getSheetByName(n); if (!s) return;
    ss.setActiveSheet(s); ss.moveActiveSheet(i + 1);
    const lastCol = Math.max(s.getLastColumn(), 1);
    const hr = n === 'Employee Summary' ? 5 : n === 'Branch Summary' ? 3 : 1;
    s.getRange(hr, 1, 1, lastCol).setBackground(GREEN).setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
    s.setRowHeight(hr, 34);
    s.setTabColor(i < 2 ? GREEN : i < 4 ? '#f0b45a' : '#9aa5a0');
  });
  // Employees
  let s = ss.getSheetByName('Employees');
  s.getRange('G2:G').setNumberFormat('dd-mmm-yyyy');
  s.getRange('F2:F').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName('Branches').getRange('A2:A'), true).setAllowInvalid(true).build());
  s.getRange('H2:H').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Caller', 'Admin'], true).build());
  s.getRange('I2:I').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Active', 'Inactive'], true).build());
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Inactive').setBackground('#fdecea').setFontColor('#8a1f14').setRanges([s.getRange('I2:I')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Active').setBackground('#e7f6ef').setFontColor('#0b6b45').setRanges([s.getRange('I2:I')]).build(),
  ]);
  [130, 110, 110, 160, 130, 190, 110, 80, 80, 130, 120].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  s.hideColumns(11);                                          // Password Hash – never needed by admin
  s.getRange('K1').setNote('Encrypted password. Do not edit. Use "Reset password" in the admin panel.');
  // Attendance
  s = ss.getSheetByName('Attendance');
  s.getRange('A2:A').setNumberFormat('dd-mmm-yyyy');
  s.getRange('O2:O').setNumberFormat('0%');
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Complete').setBackground('#e7f6ef').setFontColor('#0b6b45').setRanges([s.getRange('U2:U')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Logout pending').setBackground('#fff4e0').setFontColor('#7a4a00').setRanges([s.getRange('U2:U')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Logout missed').setBackground('#fdecea').setFontColor('#8a1f14').setRanges([s.getRange('U2:U')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes').setBackground('#fff4e0').setFontColor('#7a4a00').setRanges([s.getRange('J2:J')]).build(),
  ]);
  [100, 50, 130, 160, 190, 140, 90, 100, 85, 50, 85, 80, 70, 75, 70, 75, 75, 85, 90, 75, 120, 160, 160, 200, 60, 60]
    .forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  s.hideColumns(25, 2);                                       // full WhatsApp messages – kept for records
  s.getRange('X1').setNote('Admin can write a note here, e.g. "Logout missed – phone battery dead".');
  // Transfers
  s = ss.getSheetByName('Branch Transfers');
  s.getRange('F2:F').setNumberFormat('dd-mmm-yyyy');
  [140, 150, 170, 190, 190, 110, 110].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  // Lists & settings
  ['Branches', 'SMs', 'Locations'].forEach(function (n) { ss.getSheetByName(n).setColumnWidth(1, 240); });
  s = ss.getSheetByName('Settings'); s.setColumnWidth(1, 170); s.setColumnWidth(2, 90); s.setColumnWidth(3, 520);
  // Summaries
  s = ss.getSheetByName('Employee Summary');
  s.getRange('B1').setNumberFormat('mmmm yyyy').setBackground('#fff4e0').setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  s.getRange('A1:A3').setFontWeight('bold'); s.getRange('A3:B3').setFontColor('#9aa5a0');
  s.getRange('D6:D').setNumberFormat('dd-mmm-yyyy');
  s.getRange('H6:H').setNumberFormat('0%'); s.getRange('M6:M').setNumberFormat('0%');
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.75).setBackground('#fdecea').setFontColor('#8a1f14').setRanges([s.getRange('H6:H')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor('#c0392b').setRanges([s.getRange('J6:J')]).build(),
  ]);
  [130, 170, 190, 100, 75, 70, 70, 85, 65, 70, 75, 75, 75, 60, 65, 75].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  s = ss.getSheetByName('Branch Summary');
  s.getRange('B1').setNumberFormat('mmmm yyyy').setBackground('#fff4e0').setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().build());
  s.getRange('A1').setFontWeight('bold');
  s.getRange('H4:H').setNumberFormat('0%');
  [220, 80, 80, 75, 75, 80, 80, 75, 60, 70, 75].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  ss.setActiveSheet(ss.getSheetByName('Employees'));
}

/* ================= Nightly check (runs every night at 11 PM) ================= */
function installNightly_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'nightlyCheck') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('nightlyCheck').timeBased().atHour(23).everyDays(1).inTimezone(TZ).create();
}
// Marks every past day that still says "Logout pending" as "Logout missed"
function nightlyCheck() {
  const today = todayKey_(), t = table_('Attendance'), col = t.head.indexOf('Status') + 1;
  if (!col) return;
  t.rows.forEach(function (o) {
    if (o['Status'] === 'Logout pending' && (o['Date'] < today || nowMinutes_() >= 23 * 60))
      t.sheet.getRange(o._row, col, 1, 1).setValues([["'Logout missed"]]);
  });
}

/* ================= Sheet helpers ================= */
// Read a sheet as objects keyed by the header row (columns can be moved or added safely)
function table_(name) {
  const s = sheet_(name), range = s.getDataRange();
  const disp = range.getDisplayValues(), raw = range.getValues();
  const head = (disp[0] || []).map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < disp.length; i++) {
    const o = { _row: i + 1 }; let any = false;
    head.forEach(function (h, j) {
      if (!h) return;
      let v = disp[i][j];
      const r = raw[i][j];
      if (r instanceof Date) v = Utilities.formatDate(r, ssTz_(), 'yyyy-MM-dd');
      else if (DATE_COLS.indexOf(h) >= 0) v = normDate_(v);
      else if (typeof r === 'number' && NUM_COLS.indexOf(h) >= 0) v = String(r);
      o[h] = String(v == null ? '' : v).trim();
      if (o[h]) any = true;
    });
    if (any) rows.push(o);
  }
  return { sheet: s, head: head, rows: rows };
}
function ssTz_() { try { return ss_().getSpreadsheetTimeZone() || TZ; } catch (e) { return TZ; } }
function normDate_(v) {            // accepts 2026-11-09, 09-11-2026, 09/11/2026
  v = String(v || '').trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(v);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/.exec(v);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return v;
}
// value for a cell: dates as real dates, numbers as numbers, everything else as plain text
function cell_(col, v) {
  if (v === null || v === undefined || v === '') return '';
  if (DATE_COLS.indexOf(col) >= 0) return normDate_(v);
  if ((NUM_COLS.indexOf(col) >= 0 || col === 'Lead of the Day') && /^\d+(\.\d+)?$/.test(String(v))) return Number(v);
  if (col === 'Connect %' && typeof v === 'number') return v;
  return "'" + String(v);
}
function headOf_(name) { return sheet_(name).getDataRange().getDisplayValues()[0].map(function (h) { return String(h).trim(); }); }
function rowArr_(head, obj) { return head.map(function (h) { return h ? cell_(h, obj[h]) : ''; }); }
function appendObj_(name, obj) { sheet_(name).appendRow(rowArr_(headOf_(name), obj)); }
function updateObj_(name, rowNum, obj) {
  const head = headOf_(name);
  sheet_(name).getRange(rowNum, 1, 1, head.length).setValues([rowArr_(head, obj)]);
}
function colLetter_(name, col) {   // column letter of a header, e.g. ('Attendance','Status') -> 'T'
  let n = COLS[name].indexOf(col) + 1, out = '';
  while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); }
  return out;
}

function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" missing – please run setup once.');
  return s;
}
function rows_(name) { return sheet_(name).getDataRange().getDisplayValues().slice(1); }
// leading ' keeps dates/times as plain text in the sheet
function txt_(v) { return "'" + (v === null || v === undefined ? '' : String(v)); }
function append_(name, arr) { sheet_(name).appendRow(arr.map(txt_)); }
function seed_(name, list) { if (rows_(name).length === 0) list.forEach(function (x) { append_(name, [x]); }); }
function list_(name) { return rows_(name).map(function (r) { return r[0].trim(); }).filter(String); }
function replaceList_(name, values) {
  const s = sheet_(name), last = s.getLastRow();
  if (last > 1) s.deleteRows(2, last - 1);
  values.forEach(function (v) { append_(name, [v]); });
}
function setting_(key, def) {
  const r = rows_('Settings').filter(function (r) { return r[0] === key; })[0];
  return r && r[1] ? r[1] : def;
}
function setSetting_(key, value) {
  const s = sheet_('Settings'), all = rows_('Settings');
  for (let i = 0; i < all.length; i++)
    if (all[i][0] === key) { s.getRange(i + 2, 2, 1, 1).setValues([[txt_(value)]]); return; }
  append_('Settings', [key, value]);
}

/* ================= Time helpers (server time, India) ================= */
function now_(fmt) { return Utilities.formatDate(new Date(), TZ, fmt); }
function todayKey_() { return now_('yyyy-MM-dd'); }
function nowMinutes_() { const p = now_('HH:mm').split(':'); return Number(p[0]) * 60 + Number(p[1]); }
function fmt12_(mins) { let h = Math.floor(mins / 60), m = mins % 60; const ap = h >= 12 ? 'PM' : 'AM'; h = h % 12 || 12; return h + ':' + (m < 10 ? '0' : '') + m + ' ' + ap; }
function msgDate_(key) { const p = key.split('-'); return Number(p[2]) + '-' + Number(p[1]) + '-' + p[0].slice(2); }

function addDays_(key, n) { const d = new Date(key + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
const DAYS_ = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
function dow_(key) { return new Date(key + 'T00:00:00Z').getUTCDay(); }
function parse12_(t) {            // "9:40 AM" -> minutes after midnight
  const m = /^(\d{1,2}):(\d{2})\s*(AM|PM)$/i.exec(String(t || '').trim());
  if (!m) return null;
  return (Number(m[1]) % 12 + (m[3].toUpperCase() === 'PM' ? 12 : 0)) * 60 + Number(m[2]);
}
function num_(v) { const n = parseFloat(String(v || '').replace(/[^\d.]/g, '')); return isNaN(n) ? 0 : n; }

/* ================= Users & sessions ================= */
function hash_(username, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(username).toLowerCase() + ':' + password);
  return bytes.map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); }).join('');
}
function users_() {
  return table_('Employees').rows.map(function (o) {
    const name = o['Full Name'] || [o['First Name'], o['Last Name']].filter(String).join(' ');
    return { username: String(o['Employee ID'] || '').trim().toLowerCase(), hash: o['Password Hash'], name: name,
             first: o['First Name'], last: o['Last Name'], phone: o['Phone'] || '', branch: o['Branch'], joining: o['Joining Date'],
             role: String(o['Role'] || 'Caller').toLowerCase() === 'admin' ? 'admin' : 'caller',
             active: String(o['Status'] || 'Active').toLowerCase() === 'inactive' ? 'no' : 'yes',
             created: o['Created On'], row: o._row, raw: o };
  }).filter(function (u) { return u.username; });
}
function findUser_(username) {
  username = String(username || '').trim().toLowerCase();
  return users_().filter(function (u) { return u.username === username; })[0] || null;
}
function publicUser_(u) {
  return { username: u.username, name: u.name, phone: u.phone, branch: u.branch, joining: u.joining, role: u.role, active: u.active, created: u.created };
}
function newToken_(username) {
  const props = PropertiesService.getScriptProperties(), today = todayKey_(), all = props.getProperties();
  Object.keys(all).forEach(function (k) {                // remove tokens from earlier days
    if (k.indexOf('t_') === 0 && all[k].split('|')[1] !== today) props.deleteProperty(k);
  });
  const token = Utilities.getUuid();
  props.setProperty('t_' + token, username + '|' + today);
  return token;
}
function auth_(token, needAdmin) {
  const v = token ? PropertiesService.getScriptProperties().getProperty('t_' + token) : null;
  if (!v || v.split('|')[1] !== todayKey_()) throw new Error('SESSION_EXPIRED');
  const u = findUser_(v.split('|')[0]);
  if (!u || u.active !== 'yes') throw new Error('SESSION_EXPIRED');
  if (needAdmin && u.role !== 'admin') throw new Error('Only admin can do this.');
  return u;
}

// Login ID = first name + last name + last 4 digits of phone, e.g. vivekjadhao3210
function makeUsername_(name, digits) {
  const parts = name.toLowerCase().split(' ').map(function (w) { return w.replace(/[^a-z]/g, ''); }).filter(String);
  let base = parts.length > 1 ? parts[0] + parts[parts.length - 1] : (parts[0] || 'caller');
  base = base.slice(0, 20) + (digits ? digits.slice(-4) : '');
  let u = base, n = 2;
  while (findUser_(u)) u = base + (n++);
  return u;
}
// 6-character password without confusing letters (no 0/o, 1/l/i)
function makePassword_() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let out = '';
  for (let i = 0; i < 6; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}
function appUrl_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

/* ================= Attendance ================= */
function attendance_() {
  const out = [];
  table_('Attendance').rows.forEach(function (o) {
    const base = { date: o['Date'], username: String(o['Employee ID'] || '').toLowerCase(), name: o['Employee Name'],
                   branch: o['Branch'], sm: o['SM'], row: o._row, raw: o };
    if (!base.date || !base.username) return;
    if (o['Login Time']) out.push(Object.assign({}, base, { type: 'login', time: o['Login Time'], cp: o['CP Firm'],
      loc: o['Location'], late: o['Late'], message: o['Login Message'] }));
    if (o['Logout Time']) out.push(Object.assign({}, base, { type: 'logout', time: o['Logout Time'], total: o['Total Calls'],
      connect: o['Connect Calls'], lead: o['Lead of the Day'], visitToday: o['Site Visit Done'],
      visitTom: o['Tomorrow Site Visit'], wtd: o['Site Visit Confirmed WTD'], booking: o['Booking'], message: o['Logout Message'] }));
  });
  return out;
}

function buildMessage_(type, dateKey, name, d) {
  if (type === 'login') return [
    'Date : ' + msgDate_(dateKey),
    'Caller name : ' + name,
    'Cp firm : ' + d.cp,
    'SM :  ' + d.sm,
    'Location:  ' + d.loc,
    'Login time :' + d.time,
  ].join('\n');
  return [
    'Date: ' + msgDate_(dateKey),
    'Caller Name :- ' + name,
    'SM Name - ' + d.sm,
    'Total Calls:- ' + d.total,
    'Connect Calls: ' + d.connect,
    'Log out time:- ' + d.time,
    'Lead of the day:- ' + d.lead,
    'Site visit done today- ' + d.visitToday,
    'Tomorrow site visit: ' + d.visitTom,
    'Booking: ' + d.booking,
    'Site visit confirmed WTD: ' + d.wtd,
  ].join('\n');
}

function callerState_(u) {
  const today = todayKey_();
  const mine = attendance_().filter(function (a) { return a.username === u.username; });
  const byDay = {};
  mine.forEach(function (a) { (byDay[a.date] = byDay[a.date] || {})[a.type] = a; });
  const days = Object.keys(byDay).sort().reverse();
  return {
    user: publicUser_(u),
    isAdmin: false,
    today: { date: today, login: (byDay[today] || {}).login || null, logout: (byDay[today] || {}).logout || null },
    missed: days.filter(function (d) { return d < today && byDay[d].login && !byDay[d].logout; }).slice(0, 5),
    history: days.slice(0, 10).map(function (d) {
      return { date: d, login: byDay[d].login ? byDay[d].login.time : '', logout: byDay[d].logout ? byDay[d].logout.time : '' };
    }),
    config: { sms: list_('SMs'), locations: list_('Locations'), reminderAt: setting_('Logout reminder time', '19:00') },
    serverMinutes: nowMinutes_(),
  };
}
function stateFor_(u) { return u.role === 'admin' ? { user: publicUser_(u), isAdmin: true } : callerState_(u); }

/* ================= API (called from the web page) ================= */
function api(action, p) {
  p = p || {};
  try {
    const fn = API[action];
    if (!fn) throw new Error('Unknown action');
    return JSON.parse(JSON.stringify({ ok: true, data: fn(p) }));
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

const API = {
  login: function (p) {
    const u = findUser_(p.username);
    if (!u || u.hash !== hash_(u.username, p.password || '')) throw new Error('Wrong Login ID or password.');
    if (u.active !== 'yes') throw new Error('Your account is inactive. Please contact admin.');
    const st = stateFor_(u);
    st.token = newToken_(u.username);
    return st;
  },

  me: function (p) { return stateFor_(auth_(p.token)); },

  post: function (p) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const u = auth_(p.token);
      if (u.role === 'admin') throw new Error('Admin account cannot post attendance.');
      const type = p.type, d = p.data || {}, today = todayKey_();
      if (type !== 'login' && type !== 'logout') throw new Error('Wrong entry type.');
      const st = callerState_(u);
      if (st.today[type]) throw new Error((type === 'login' ? 'Login' : 'Logout') + ' already posted today at ' + st.today[type].time + '. Only one ' + type + ' is allowed per day.');
      const rec = {};
      if (type === 'login') {
        const h = Number(d.hour), m = Number(d.min);
        if (!(h >= 1 && h <= 12) || !(m >= 0 && m <= 59) || (d.ap !== 'AM' && d.ap !== 'PM')) throw new Error('Please choose the login time.');
        const mins = (h % 12 + (d.ap === 'PM' ? 12 : 0)) * 60 + m;
        if (mins > nowMinutes_()) throw new Error('Login time cannot be in the future.');
        rec.time = fmt12_(mins);
        rec.sm = String(d.sm || '').trim();
        rec.cp = String(d.cp || '').trim();
        rec.loc = String(d.loc || '').trim();
        if (!rec.sm) throw new Error('Please choose the SM.');
        if (!rec.cp) throw new Error('Please enter the CP firm.');
        if (!rec.loc) throw new Error('Please enter the location.');
      } else {
        if (!st.today.login) throw new Error('Please post your Login first.');
        rec.time = fmt12_(nowMinutes_());           // logout time = server time, cannot be changed
        rec.sm = st.today.login.sm;                 // SM fixed from today's login
        ['total', 'connect', 'lead', 'visitToday', 'visitTom', 'booking', 'wtd'].forEach(function (k) {
          rec[k] = String(d[k] === undefined ? '' : d[k]).trim();
          if (!rec[k]) throw new Error('Please fill all logout fields (put 0 if none).');
        });
        if (Number(rec.connect) > Number(rec.total)) throw new Error('Connect calls cannot be more than Total calls.');
      }
      const message = buildMessage_(type, today, u.name, rec);
      const posted = now_('dd-MM-yyyy HH:mm:ss');
      if (type === 'login') {
        const lp = setting_('Late after', '10:00').split(':');
        appendObj_('Attendance', { 'Date': today, 'Day': DAYS_[dow_(today)].slice(0, 3), 'Employee ID': u.username,
          'Employee Name': u.name, 'Branch': u.branch, 'SM': rec.sm, 'CP Firm': rec.cp, 'Location': rec.loc,
          'Login Time': rec.time, 'Late': parse12_(rec.time) > Number(lp[0]) * 60 + Number(lp[1] || 0) ? 'Yes' : 'No',
          'Status': 'Logout pending', 'Login Posted At': posted, 'Login Message': message });
      } else {
        const row = st.today.login.raw;
        const lm = parse12_(st.today.login.time), om = parse12_(rec.time), tot = Number(rec.total), con = Number(rec.connect);
        updateObj_('Attendance', row._row, Object.assign({}, row, {
          'Logout Time': rec.time,
          'Hours Worked': om > lm ? Math.floor((om - lm) / 60) + 'h ' + ((om - lm) % 60) + 'm' : '',
          'Total Calls': rec.total, 'Connect Calls': rec.connect, 'Connect %': tot ? Math.round(con * 1000 / tot) / 1000 : '',
          'Lead of the Day': rec.lead, 'Site Visit Done': rec.visitToday, 'Tomorrow Site Visit': rec.visitTom,
          'Site Visit Confirmed WTD': rec.wtd, 'Booking': rec.booking, 'Status': 'Complete', 'Logout Posted At': posted, 'Logout Message': message }));
      }
      return { message: message, state: callerState_(u) };
    } finally {
      lock.releaseLock();
    }
  },

  /* ---------- Admin ---------- */
  adminData: function (p) {
    auth_(p.token, true);
    return {
      users: users_().map(publicUser_),
      branches: list_('Branches'), sms: list_('SMs'), locations: list_('Locations'),
      reminderAt: setting_('Logout reminder time', '19:00'),
      lateAfter: setting_('Late after', '10:00'),
      weeklyOff: setting_('Weekly off', 'Monday'),
      transfers: table_('Branch Transfers').rows.slice(-100).reverse().map(function (o) {
        return { changedOn: o['Changed On'], username: o['Employee ID'], name: o['Employee Name'], from: o['From Branch'],
                 to: o['To Branch'], effective: o['Effective Date'], by: o['Changed By'] };
      }),
      today: todayKey_(),
      appUrl: appUrl_(),
    };
  },

  saveUser: function (p) {
    const admin = auth_(p.token, true), x = p.user || {};
    const first = String(x.firstName || '').trim().replace(/\s+/g, ' ');
    const last = String(x.lastName || '').trim().replace(/\s+/g, ' ');
    if (!first) throw new Error('Please enter the first name.');
    if (!last) throw new Error('Please enter the last name.');
    const name = first + ' ' + last;
    const role = x.role === 'admin' ? 'admin' : 'caller';
    const active = x.active === 'no' ? 'no' : 'yes';
    // phone: 10-digit mobile number, saved as +91 9876543210
    let phone = String(x.phone || '').replace(/\D/g, '');
    if (phone.length === 12 && phone.indexOf('91') === 0) phone = phone.slice(2);
    if (phone.length === 11 && phone[0] === '0') phone = phone.slice(1);
    if (phone && !/^[6-9]\d{9}$/.test(phone)) throw new Error('Please enter a valid 10-digit mobile number.');
    if (!phone && role === 'caller') throw new Error('Please enter the phone number.');
    const digits = phone;
    if (phone) phone = '+91 ' + phone;

    if (x.isNew) {
      if (role === 'caller' && !x.branch) throw new Error('Please choose the branch.');
      if (digits && users_().some(function (u) { return u.phone.replace(/\D/g, '').slice(-10) === digits; }))
        throw new Error('This phone number is already added for another caller.');
      const username = makeUsername_(name, digits);
      const password = makePassword_();
      appendObj_('Employees', { 'Employee ID': username, 'First Name': first, 'Last Name': last, 'Full Name': name,
        'Phone': phone, 'Branch': x.branch || '', 'Joining Date': x.joining || '', 'Role': role === 'admin' ? 'Admin' : 'Caller',
        'Status': active === 'no' ? 'Inactive' : 'Active', 'Created On': now_('yyyy-MM-dd HH:mm'), 'Password Hash': hash_(username, password) });
      const out = API.adminData(p);
      out.credentials = { name: name, username: username, password: password, phone: phone, isNew: true };
      return out;
    }

    const existing = findUser_(x.username);
    if (!existing) throw new Error('User not found.');
    if (existing.username === admin.username && (active === 'no' || role !== 'admin'))
      throw new Error('You cannot deactivate or demote your own admin account.');
    if (digits && users_().some(function (u) { return u.username !== existing.username && u.phone.replace(/\D/g, '').slice(-10) === digits; }))
      throw new Error('This phone number is already added for another caller.');
    updateObj_('Employees', existing.row, Object.assign({}, existing.raw, { 'First Name': first, 'Last Name': last,
      'Full Name': name, 'Phone': phone, 'Joining Date': x.joining || '', 'Role': role === 'admin' ? 'Admin' : 'Caller',
      'Status': active === 'no' ? 'Inactive' : 'Active' }));
    return API.adminData(p);
  },

  resetPassword: function (p) {
    auth_(p.token, true);
    const u = findUser_(p.username);
    if (!u) throw new Error('User not found.');
    const password = makePassword_();
    updateObj_('Employees', u.row, Object.assign({}, u.raw, { 'Password Hash': hash_(u.username, password) }));
    const out = API.adminData(p);
    out.credentials = { name: u.name, username: u.username, password: password, phone: u.phone, isNew: false };
    return out;
  },

  transfer: function (p) {
    const admin = auth_(p.token, true);
    const u = findUser_(p.username);
    if (!u) throw new Error('User not found.');
    const to = String(p.branch || '').trim();
    if (!to) throw new Error('Please choose the new branch.');
    if (to === u.branch) throw new Error(u.name + ' is already in ' + to + '.');
    updateObj_('Employees', u.row, Object.assign({}, u.raw, { 'Branch': to }));
    appendObj_('Branch Transfers', { 'Changed On': now_('yyyy-MM-dd HH:mm'), 'Employee ID': u.username, 'Employee Name': u.name,
      'From Branch': u.branch, 'To Branch': to, 'Effective Date': p.effective || todayKey_(), 'Changed By': admin.name });
    return API.adminData(p);
  },

  saveLists: function (p) {
    auth_(p.token, true);
    const clean = function (a) {
      const seen = {};
      return (a || []).map(function (s) { return String(s).trim(); })
        .filter(function (s) { if (!s || seen[s.toLowerCase()]) return false; seen[s.toLowerCase()] = 1; return true; });
    };
    const b = clean(p.branches), s = clean(p.sms), l = clean(p.locations);
    if (!b.length || !s.length || !l.length) throw new Error('Branches, SMs and Locations each need at least one entry.');
    if (!/^\d{1,2}:\d{2}$/.test(p.reminderAt || '')) throw new Error('Reminder time must look like 19:00');
    if (!/^\d{1,2}:\d{2}$/.test(p.lateAfter || '')) throw new Error('Late-after time must look like 10:00');
    if (p.weeklyOff !== 'None' && DAYS_.indexOf(p.weeklyOff) < 0) throw new Error('Please choose the weekly off day.');
    replaceList_('Branches', b); replaceList_('SMs', s); replaceList_('Locations', l);
    setSetting_('Logout reminder time', p.reminderAt);
    setSetting_('Late after', p.lateAfter);
    setSetting_('Weekly off', p.weeklyOff);
    return API.adminData(p);
  },

  /* Period report: branch summary + every employee's attendance & calls between two dates */
  periodReport: function (p) {
    auth_(p.token, true);
    const today = todayKey_(), re = /^\d{4}-\d{2}-\d{2}$/;
    let from = p.from, to = p.to;
    if (!re.test(from || '') || !re.test(to || '')) throw new Error('Please choose From and To dates.');
    if (to > today) to = today;
    if (from > to) throw new Error('From date must be before To date.');
    if ((new Date(to) - new Date(from)) / 864e5 > 92) throw new Error('Please choose at most 3 months at a time.');
    const lateAfter = setting_('Late after', '10:00'), lp = lateAfter.split(':'), lateMin = Number(lp[0]) * 60 + Number(lp[1]);
    const weeklyOff = setting_('Weekly off', 'Monday'), offDay = DAYS_.indexOf(weeklyOff);
    const days = []; for (let d = from; d <= to; d = addDays_(d, 1)) days.push(d);

    const by = {};
    attendance_().forEach(function (a) {
      if (a.date < from || a.date > to) return;
      const u = (by[a.username] = by[a.username] || {});
      (u[a.date] = u[a.date] || {})[a.type] = a;
    });

    const employees = users_().filter(function (u) {
      return u.role === 'caller' && (u.active === 'yes' || by[u.username]);
    }).map(function (u) {
      const mine = by[u.username] || {}, start = u.joining && u.joining > from ? u.joining : from;
      const e = { username: u.username, name: u.name, phone: u.phone, branch: u.branch, joining: u.joining, active: u.active,
                  present: 0, absent: 0, late: 0, missed: 0, offWorked: 0, total: 0, connect: 0, lead: 0, visits: 0, booking: 0, daily: [] };
      let loginSum = 0, workMins = 0, workN = 0;
      days.forEach(function (d) {
        if (d < start) return;
        const li = (mine[d] || {}).login, lo = (mine[d] || {}).logout, isOff = offDay >= 0 && dow_(d) === offDay;
        const row = { date: d, day: DAYS_[dow_(d)].slice(0, 3), login: li ? li.time : '', logout: lo ? lo.time : '',
                      sm: li ? li.sm : '', hours: '', total: '', connect: '', lead: '', visits: '', booking: '', status: '' };
        if (li) {
          e.present++; if (isOff) e.offWorked++;
          const lm = parse12_(li.time), isLate = lm !== null && lm > lateMin;
          if (lm !== null) loginSum += lm;
          if (isLate) e.late++;
          if (lo) {
            const om = parse12_(lo.time);
            if (lm !== null && om !== null && om > lm) { workMins += om - lm; workN++; row.hours = Math.floor((om - lm) / 60) + 'h ' + ((om - lm) % 60) + 'm'; }
            row.total = lo.total; row.connect = lo.connect; row.lead = lo.lead; row.visits = lo.visitToday; row.booking = lo.booking || '';
            e.total += num_(lo.total); e.connect += num_(lo.connect); e.lead += num_(lo.lead); e.visits += num_(lo.visitToday); e.booking += num_(lo.booking);
            row.status = isLate ? 'Late' : 'Present';
          } else if (d < today) { e.missed++; row.status = 'Logout missed'; }
          else row.status = 'Working';
        } else if (isOff) row.status = 'Weekly off';
        else if (d === today) row.status = 'Not logged in';
        else { e.absent++; row.status = 'Absent'; }
        e.daily.push(row);
      });
      e.daily.reverse();
      e.attPct = e.present + e.absent ? Math.round(e.present * 100 / (e.present + e.absent)) : null;
      e.connectPct = e.total ? Math.round(e.connect * 100 / e.total) : null;
      e.avgLogin = e.present ? fmt12_(Math.round(loginSum / e.present)) : '';
      e.avgHours = workN ? Math.floor(workMins / workN / 60) + 'h ' + Math.round(workMins / workN % 60) + 'm' : '';
      return e;
    });

    const names = list_('Branches');
    employees.forEach(function (e) { if (e.branch && names.indexOf(e.branch) < 0) names.push(e.branch); });
    const branches = names.map(function (b) {
      const es = employees.filter(function (e) { return e.branch === b; });
      const sum = function (k) { return es.reduce(function (t, e) { return t + e[k]; }, 0); };
      const pr = sum('present'), ab = sum('absent'), tot = sum('total'), con = sum('connect');
      return { branch: b, callers: es.filter(function (e) { return e.active === 'yes'; }).length,
               present: pr, absent: ab, late: sum('late'), missed: sum('missed'),
               attPct: pr + ab ? Math.round(pr * 100 / (pr + ab)) : null,
               total: tot, connect: con, connectPct: tot ? Math.round(con * 100 / tot) : null, lead: sum('lead'), visits: sum('visits'), booking: sum('booking') };
    });
    return { from: from, to: to, today: today, lateAfter: lateAfter, weeklyOff: weeklyOff, branches: branches, employees: employees };
  },

  report: function (p) {
    auth_(p.token, true);
    const date = p.date || todayKey_();
    const att = attendance_().filter(function (a) { return a.date === date; });
    const callers = users_().filter(function (u) { return u.role === 'caller' && u.active === 'yes'; });
    return {
      date: date,
      isToday: date === todayKey_(),
      isOff: DAYS_.indexOf(setting_('Weekly off', 'Monday')) === dow_(date),
      reminderAt: setting_('Logout reminder time', '19:00'),
      serverMinutes: nowMinutes_(),
      rows: callers.map(function (u) {
        const li = att.filter(function (a) { return a.username === u.username && a.type === 'login'; })[0];
        const lo = att.filter(function (a) { return a.username === u.username && a.type === 'logout'; })[0];
        return { name: u.name, username: u.username, branch: li ? li.branch : u.branch,
                 login: li ? li.time : '', logout: lo ? lo.time : '', sm: li ? li.sm : '',
                 total: lo ? lo.total : '', connect: lo ? lo.connect : '', booking: lo ? (lo.booking || '') : '' };
      }),
    };
  },
};
