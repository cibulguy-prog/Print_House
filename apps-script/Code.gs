/**
 * ניהול כספי הדפוס - צד שרת (Google Apps Script)
 * הנתונים נשמרים בלשוניות של הגיליון הזה. כל טבלה היא לשונית, כל רשומה היא שורה.
 * מותר לערוך נתונים ישירות בגיליון, אבל לא לשנות את שמות הכותרות בשורה 1
 * ולא את עמודת "מזהה".
 */

var SCHEMA = {
  customers: { sheet: 'לקוחות', cols: [
    ['id','מזהה','t'],['name','שם לקוח','t'],['taxId','ח.פ / ע.מ','t'],['contact','איש קשר','t'],['phone','טלפון','t'],
    ['email','אימייל','t'],['address','כתובת','t'],['terms','תנאי תשלום (ימים)','n'],['notes','הערות','t'],['sample','דוגמה','b']]},
  suppliers: { sheet: 'ספקים', cols: [
    ['id','מזהה','t'],['name','שם ספק','t'],['taxId','ח.פ','t'],['category','קטגוריה','t'],['terms','תנאי תשלום (ימים)','n'],
    ['contact','איש קשר','t'],['phone','טלפון','t'],['email','אימייל','t'],['sample','דוגמה','b']]},
  jobs: { sheet: 'עבודות', cols: [
    ['id','מזהה','t'],['number','מס׳ עבודה','n'],['date','תאריך','d'],['customerId','מזהה לקוח','t'],['customerName','לקוח','t'],
    ['title','עבודה','t'],['specs','מפרט','t'],['qty','כמות','n'],['price','מחיר לפני מע״מ','n'],['estCost','עלות משוערת','n'],
    ['status','סטטוס','t'],['dueDate','תאריך יעד','d'],['invoiceId','מזהה חשבונית','t'],['notes','הערות','t'],['calc','נתוני מחשבון','j'],['sample','דוגמה','b']]},
  docs: { sheet: 'מסמכים', cols: [
    ['id','מזהה','t'],['type','קוד סוג','t'],['typeName','סוג מסמך','t'],['number','מספר','n'],['date','תאריך','d'],['dueDate','לתשלום עד','d'],
    ['customerId','מזהה לקוח','t'],['customerName','לקוח','t'],['customerTaxId','ח.פ לקוח','t'],['customerAddress','כתובת לקוח','t'],
    ['net','לפני מע״מ','n'],['vat','מע״מ','n'],['total','סה״כ','n'],['vatRate','שיעור מע״מ','n'],['discount','הנחה','n'],
    ['allocation','מספר הקצאה','t'],['refDocId','מסמך מקור','t'],['notes','הערות','t'],
    ['lines','שורות','j'],['payments','תקבולים','j'],['appliesTo','שיוך לחשבוניות','j'],['jobIds','עבודות','j'],
    ['createdBy','הופק על ידי','t'],['createdAt','נוצר','t'],['sample','דוגמה','b']]},
  expenses: { sheet: 'הוצאות', cols: [
    ['id','מזהה','t'],['date','תאריך','d'],['supplierId','מזהה ספק','t'],['supplierName','ספק','t'],['category','קטגוריה','t'],
    ['desc','תיאור','t'],['docNo','אסמכתא','t'],['total','סה״כ כולל מע״מ','n'],['hasVat','כולל מע״מ','b'],['isAsset','רכוש קבוע','b'],
    ['vatRate','שיעור מע״מ','n'],['paid','שולם','b'],['paidDate','תאריך תשלום','d'],['method','אמצעי תשלום','t'],['sample','דוגמה','b']]}
};
var SETTINGS_SHEET = 'הגדרות';
var DOC_NAMES = {'305':'חשבונית מס','320':'חשבונית מס/קבלה','330':'חשבונית זיכוי','400':'קבלה'};

/* ---------- web app & menu ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ניהול כספי הדפוס')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('מערכת כספים')
    .addItem('פתיחת המערכת', 'showAppLink')
    .addItem('הכנת הגיליונות (פעם ראשונה)', 'setup')
    .addToUi();
}

function showAppLink() {
  var url = ScriptApp.getService().getUrl();
  var html = url
    ? '<div dir="rtl" style="font-family:Arial;font-size:14px"><p>המערכת פתוחה בכתובת:</p><p><a href="' + url + '" target="_blank">פתיחת ניהול כספי הדפוס</a></p><p style="color:#666">כדאי לשמור את הקישור במועדפים.</p></div>'
    : '<div dir="rtl" style="font-family:Arial;font-size:14px"><p>המערכת עוד לא פורסמה כאפליקציית אינטרנט.</p><p>בעורך Apps Script: פריסה ← פריסה חדשה ← סוג: אפליקציית אינטרנט.</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(160), 'ניהול כספי הדפוס');
}

/** יוצר את הלשוניות והכותרות. בטוח להריץ שוב: לא מוחק נתונים. */
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(function (k) { ensureSheet_(ss, SCHEMA[k]); });
  var st = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
  if (!st.getRange('A1').getValue()) {
    st.getRange('A1:B1').setValues([['settings', '']]);
    st.getRange('A3').setValue('הגדרות המערכת נשמרות כאן אוטומטית. עורכים אותן ממסך ההגדרות במערכת.');
  }
  st.getRange('B1').setNumberFormat('@');
  var first = ss.getSheets()[0];
  if (first.getName() === 'Sheet1' || first.getName() === 'גיליון1') {
    if (first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  }
  return true;
}

function ensureSheet_(ss, def) {
  var sh = ss.getSheetByName(def.sheet);
  if (!sh) sh = ss.insertSheet(def.sheet);
  var headers = def.cols.map(function (c) { return c[1]; });
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var current = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (!current.join('')) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    // add any missing columns at the end (keeps existing order and data)
    var missing = headers.filter(function (h) { return current.indexOf(h) < 0; });
    if (missing.length) sh.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#EDF1F4');
  sh.setRightToLeft(true);
  // text columns stay text (keeps leading zeros in ח.פ / phone numbers)
  var map = headerMap_(sh);
  def.cols.forEach(function (c) {
    var col = map[c[1]];
    if (!col) return;
    var rng = sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1);
    if (c[2] === 't' || c[2] === 'j') rng.setNumberFormat('@');
    if (c[2] === 'd') rng.setNumberFormat('dd/mm/yyyy');
    if (c[2] === 'n') rng.setNumberFormat('#,##0.##');
  });
  return sh;
}

function headerMap_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var row = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var m = {};
  row.forEach(function (h, i) { if (h !== '' && !m[h]) m[String(h)] = i + 1; });
  return m;
}

function sheetFor_(col) {
  var def = SCHEMA[col];
  if (!def) throw new Error('טבלה לא מוכרת: ' + col);
  var ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(def.sheet) || ensureSheet_(ss, def);
}

/* ---------- conversion ---------- */
var TZ_ = null;
function tz_() { return TZ_ || (TZ_ = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone()); }

function toCell_(v, type) {
  if (v === undefined || v === null) return '';
  if (type === 'j') return (v === '' ? '' : JSON.stringify(v));
  if (type === 'b') return v === true || v === 'true' || v === 'TRUE';
  if (type === 'n') return (v === '' ? '' : Number(v));
  if (type === 'd') {
    if (!v) return '';
    var p = String(v).split('-');
    if (p.length === 3) return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    return v;
  }
  return String(v);
}

function fromCell_(v, type) {
  if (type === 'j') { if (v === '' || v === null) return undefined; try { return JSON.parse(v); } catch (e) { return undefined; } }
  if (type === 'b') return v === true || v === 'TRUE' || v === 'true' || v === 'כן';
  if (type === 'n') return (v === '' || v === null) ? undefined : Number(v);
  if (type === 'd') {
    if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
    return v ? String(v) : '';
  }
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v === null ? '' : String(v);
}

function readAll_(col) {
  var def = SCHEMA[col], sh = sheetFor_(col);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var map = headerMap_(sh);
  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var out = [];
  values.forEach(function (row) {
    var o = {};
    def.cols.forEach(function (c) {
      var idx = map[c[1]];
      if (!idx) return;
      var v = fromCell_(row[idx - 1], c[2]);
      if (v !== undefined) o[c[0]] = v;
    });
    if (o.id) out.push(o);
  });
  return out;
}

function findRow_(sh, map, id) {
  var idCol = map['מזהה'];
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, idCol, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function writeRow_(col, obj) {
  var def = SCHEMA[col], sh = sheetFor_(col), map = headerMap_(sh);
  var width = sh.getLastColumn();
  var r = findRow_(sh, map, obj.id);
  var row = r ? sh.getRange(r, 1, 1, width).getValues()[0] : new Array(width).fill('');
  def.cols.forEach(function (c) {
    var idx = map[c[1]];
    if (!idx || !(c[0] in obj)) return;
    row[idx - 1] = toCell_(obj[c[0]], c[2]);
  });
  if (!r) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, width).setValues([row]);
}

function withLock_(fn) {
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- API for the web app ---------- */
function getAll() {
  var out = {};
  Object.keys(SCHEMA).forEach(function (k) { out[k] = readAll_(k); });
  var st = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  var raw = st ? st.getRange('B1').getValue() : '';
  try { out.settings = raw ? JSON.parse(raw) : null; } catch (e) { out.settings = null; }
  out.user = API_WHO_ || Session.getActiveUser().getEmail() || '';
  out.sheetUrl = SpreadsheetApp.getActive().getUrl();
  out.serverTime = new Date().toISOString();
  return out;
}

function saveRecord(col, obj) {
  if (!obj || !obj.id) throw new Error('רשומה ללא מזהה');
  return withLock_(function () { writeRow_(col, obj); return obj; });
}

function saveRecords(col, list) {
  return withLock_(function () { list.forEach(function (o) { writeRow_(col, o); }); return list.length; });
}

function deleteRecord(col, id) {
  return withLock_(function () {
    var sh = sheetFor_(col), map = headerMap_(sh);
    var r = findRow_(sh, map, id);
    if (r) sh.deleteRow(r);
    return !!r;
  });
}

/** Issues a document with the next sequential number for its type. Numbering is done here, under a lock, so two users never get the same number. */
function createDoc(doc, startNumbers) {
  return withLock_(function () {
    var docs = readAll_('docs');
    var start = Number((startNumbers || {})[doc.type]) || 1;
    var mx = start - 1;
    docs.forEach(function (d) { if (d.type === doc.type && Number(d.number) > mx) mx = Number(d.number); });
    doc.number = mx + 1;
    doc.typeName = DOC_NAMES[doc.type] || '';
    doc.createdBy = API_WHO_ || Session.getActiveUser().getEmail() || '';
    writeRow_('docs', doc);
    return doc;
  });
}

function saveSettings(obj) {
  return withLock_(function () {
    var ss = SpreadsheetApp.getActive();
    var st = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
    st.getRange('A1').setValue('settings');
    st.getRange('B1').setNumberFormat('@').setValue(JSON.stringify(obj));
    return true;
  });
}

/** Writes a report table to its own tab ("דוח · name"), replacing the previous version. Returns a link to the tab. */
function writeReport(name, rows) {
  var ss = SpreadsheetApp.getActive();
  var title = ('דוח · ' + name).replace(/[\[\]\*\?\/\\:]/g, '-').slice(0, 95);
  var sh = ss.getSheetByName(title) || ss.insertSheet(title);
  sh.clear();
  sh.setRightToLeft(true);
  if (rows && rows.length) {
    var w = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var data = rows.map(function (r) { var x = r.slice(); while (x.length < w) x.push(''); return x; });
    sh.getRange(1, 1, data.length, w).setValues(data);
    sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#EDF1F4');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, w);
  }
  sh.getRange(rows.length + 2, 1).setValue('עודכן: ' + Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm')).setFontColor('#5A6471');
  return ss.getUrl() + '#gid=' + sh.getSheetId();
}

/* ---------- JSON API for the website (GitHub Pages) ----------
 * The site sends POST (Content-Type text/plain) with {"action": "...", "args": [...]}.
 * Deploy the web app as: Execute as "Me", Who has access "Anyone".
 */
var API_ACTIONS = {
  getAll: getAll, saveRecord: saveRecord, saveRecords: saveRecords, deleteRecord: deleteRecord,
  createDoc: createDoc, saveSettings: saveSettings, writeReport: writeReport, ping: function () { return 'ok'; }
};
var API_WHO_ = '';

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'בקשה לא תקינה' }); }
  var fn = API_ACTIONS[req && req.action];
  if (!fn) return json_({ ok: false, error: 'פעולה לא מוכרת' });
  try {
    return json_({ ok: true, result: fn.apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
