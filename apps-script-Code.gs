// ZBT Dispatch Tools — shared rate con log in the "Log" tab of this Google Sheet.
// The website (GitHub Pages) talks to this script with POST requests that must carry the passcode.

// >>> Put your passcode between the quotes. Use 8+ characters. Everyone types this once per device. <<<
const PASSCODE = 'CHANGE-ME';

const LOG_TAB = 'Log';

function doGet() {
  return ContentService.createTextOutput('ZBT Dispatch Tools log. Nothing to see here.');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (PASSCODE === 'CHANGE-ME' || String(body.key || '') !== PASSCODE) {
      Utilities.sleep(800);                      // slows down anyone guessing
      return json_({ ok: false, error: 'locked' });
    }
    if (body.action === 'log') return json_({ ok: true, log: getLog_() });
    if (body.action === 'add') return json_({ ok: true, log: addEntry_(body.rec || {}) });
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function logSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(LOG_TAB);
    sh.appendRow(['Number', 'Week', 'Load', 'Date', 'Consignee', 'PO#', 'Saved at']);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('D:D').setNumberFormat('@');
    sh.getRange('F:F').setNumberFormat('@');
    // starting row: the last real rate con before this log existed (your copy in Apps Script has the real one)
    sh.appendRow(['1.1', 1, 1, '2026-01-01', 'Sample Warehouse 1234', '123456789', new Date(2026, 0, 1)]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function asDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

function getLog_() {
  const rows = logSheet_().getDataRange().getValues().slice(1).filter(r => r[0] !== '');
  return rows.map(r => ({
    number: String(r[0]), week: Number(r[1]), load: Number(r[2]), date: asDate_(r[3]),
    cons: String(r[4]), po: String(r[5]), ts: r[6] instanceof Date ? r[6].getTime() : 0,
  })).reverse().slice(0, 200);
}

function addEntry_(rec) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    logSheet_().appendRow([String(rec.number || '').slice(0, 20), Number(rec.week) || 0, Number(rec.load) || 0,
      String(rec.date || '').slice(0, 10), String(rec.cons || '').slice(0, 300), String(rec.po || '').slice(0, 60), new Date()]);
  } finally {
    lock.releaseLock();
  }
  return getLog_();
}
