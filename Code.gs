/** Install with Index.html as a Google Apps Script Web app. */
const ROSTER_TAB = 'Attendance_Roster';
const LOG_TAB = 'Attendance_Log';
const ROSTER_HEADER = ['School', 'Student_ID', 'Student_Name'];
const LOG_HEADER = ['Date', 'School', 'Student_ID', 'Student_Name', 'Status', 'Recorded_At'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Παρουσιολόγιο')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function lock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function text_(value, label) {
  if (typeof value !== 'string' || !value.trim() || value.trim().length > 200 || /[\r\n\t]/.test(value))
    throw new Error('Μη έγκυρη τιμή: ' + label);
  const result = value.trim();
  if (/^[=+@-]/.test(result)) throw new Error(label + ': δεν επιτρέπεται αρχικό =, +, -, @.');
  return result;
}
function tab_(ss, name, header) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (!sheet.getLastRow()) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    const actual = sheet.getRange(1, 1, 1, header.length).getDisplayValues()[0];
    if (JSON.stringify(actual) !== JSON.stringify(header))
      throw new Error('Το φύλλο ' + name + ' έχει διαφορετικές στήλες. Μετονομάστε το υπάρχον φύλλο πρώτα.');
  }
  return sheet;
}
function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('ATTENDANCE_SHEET_ID');
  if (!id) throw new Error('Συνδέστε πρώτα το Google Sheet από τις Ρυθμίσεις.');
  return SpreadsheetApp.openById(id);
}
function capacity_(sheet, endRow) {
  if (endRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), endRow - sheet.getMaxRows());
}
function roster_(ss) {
  const sheet = tab_(ss, ROSTER_TAB, ROSTER_HEADER);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues()
    .filter(row => row.some(Boolean)).map(row => ({school: row[0], id: row[1], name: row[2]}));
}
function state_(ss) {
  return {connected:true, sheetUrl:ss.getUrl(), roster:roster_(ss)};
}
function getState() {
  return lock_(() => {
    if (!PropertiesService.getScriptProperties().getProperty('ATTENDANCE_SHEET_ID'))
      return {connected:false, sheetUrl:'', roster:[]};
    return state_(spreadsheet_());
  });
}
function configureSpreadsheet(url) {
  return lock_(() => {
    const match = String(url).trim().match(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)(?:\/|$)/);
    if (!match) throw new Error('Δώστε την πλήρη διεύθυνση του Google Sheet.');
    const ss = SpreadsheetApp.openById(match[1]);
    tab_(ss, ROSTER_TAB, ROSTER_HEADER);
    tab_(ss, LOG_TAB, LOG_HEADER);
    PropertiesService.getScriptProperties().setProperty('ATTENDANCE_SHEET_ID', match[1]);
    return state_(ss);
  });
}
function saveRoster(rows) {
  return lock_(() => {
    if (!Array.isArray(rows) || !rows.length || rows.length > 5000) throw new Error('Απαιτούνται 1–5000 παιδιά.');
    const ids = new Set();
    const values = rows.map(row => {
      const school = text_(row.school, 'Σχολείο');
      const id = text_(row.id, 'Κωδικός παιδιού');
      const name = text_(row.name, 'Όνομα παιδιού');
      if (ids.has(id)) throw new Error('Διπλός κωδικός παιδιού: ' + id);
      ids.add(id);
      return [school, id, name];
    });
    const ss = spreadsheet_();
    const sheet = tab_(ss, ROSTER_TAB, ROSTER_HEADER);
    const oldLast = sheet.getLastRow();
    capacity_(sheet, values.length + 1);
    sheet.getRange(2, 1, values.length, 3).setNumberFormat('@').setValues(values);
    if (oldLast > values.length + 1) sheet.getRange(values.length + 2, 1, oldLast - values.length - 1, 3).clearContent();
    SpreadsheetApp.flush();
    return state_(ss);
  });
}
function saveAttendance(payload) {
  return lock_(() => {
    if (!payload || !/^\d{4}-\d{2}-\d{2}$/.test(payload.date)) throw new Error('Μη έγκυρη ημερομηνία.');
    const date = new Date(payload.date + 'T12:00:00Z');
    if (isNaN(date.getTime()) || date.toISOString().slice(0,10) !== payload.date) throw new Error('Μη έγκυρη ημερομηνία.');
    const school = text_(payload.school, 'Σχολείο');
    if (!Array.isArray(payload.ids) || !payload.ids.length || payload.ids.length > 5000) throw new Error('Επιλέξτε παιδιά.');
    const ss = spreadsheet_();
    const roster = roster_(ss);
    const ids = [...new Set(payload.ids.map(id => text_(id, 'Κωδικός παιδιού')))];
    const students = ids.map(id => {
      const matches = roster.filter(s => s.school === school && s.id === id);
      if (matches.length !== 1) throw new Error('Ο κωδικός ' + id + ' δεν αντιστοιχεί σε ένα παιδί του σχολείου. Ανανεώστε τη λίστα.');
      text_(matches[0].name, 'Όνομα παιδιού');
      return matches[0];
    });
    const sheet = tab_(ss, LOG_TAB, LOG_HEADER);
    const existing = sheet.getLastRow() < 2 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,3).getDisplayValues();
    const key = (d,s,id) => JSON.stringify([d,s,id]);
    const seen = new Set(existing.map(row => key(...row)));
    const stamp = new Date().toISOString();
    const rows = students.filter(s => !seen.has(key(payload.date, school, s.id)))
      .map(s => [payload.date, school, s.id, s.name, 'Παρών', stamp]);
    if (rows.length) {
      capacity_(sheet, sheet.getLastRow() + rows.length);
      sheet.getRange(sheet.getLastRow()+1,1,rows.length,6).setNumberFormat('@').setValues(rows);
    }
    SpreadsheetApp.flush();
    return {inserted:rows.length, skipped:ids.length-rows.length};
  });
}
