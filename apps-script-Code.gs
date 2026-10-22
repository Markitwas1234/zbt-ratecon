// ZBT Rate Con Maker — shared log in the "Log" tab of this Google Sheet.
// The website lives on GitHub Pages and talks to this script:
//   GET  ...exec?action=log  -> recent rate cons (newest first)
//   POST ...exec  {action:"add", rec:{...}} -> adds one, returns the log
// Opening ...exec with no action still shows the old built-in page.
const LOG_TAB = 'Log';

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'log') return json_({ ok: true, log: getLog() });
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ZBT Rate Con Maker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'add') return json_({ ok: true, log: addEntry(body.rec) });
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
    sh.appendRow(['85.3', 85, 3, '2026-10-20', 'WCD Palmdale 5027', '156662821', new Date(2026, 9, 20)]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function asDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

function getLog() {
  const rows = logSheet_().getDataRange().getValues().slice(1).filter(r => r[0] !== '');
  return rows.map(r => ({
    number: String(r[0]), week: Number(r[1]), load: Number(r[2]), date: asDate_(r[3]),
    cons: String(r[4]), po: String(r[5]), ts: r[6] instanceof Date ? r[6].getTime() : 0,
  })).reverse().slice(0, 200);
}

function addEntry(rec) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    logSheet_().appendRow([String(rec.number), Number(rec.week), Number(rec.load), String(rec.date),
      String(rec.cons || '').slice(0, 300), String(rec.po || '').slice(0, 60), new Date()]);
  } finally {
    lock.releaseLock();
  }
  return getLog();
}
