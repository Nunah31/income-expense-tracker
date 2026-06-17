// Google Apps Script - הדבק קוד זה ב-Google Apps Script ופרוס כ-Web App
// פרויקט: הכנסות והוצאות 2025

const SECRET_KEY    = 'esther-udi-2025';
const INCOME_SHEET  = 'הכנסות';
const EXPENSE_SHEET = 'הוצאות';

const INCOME_HEADERS  = ['שם', 'ת.ביצוע', 'ת.פרעון', 'סה"כ', 'הערות', 'סטטוס', 'ת.תשלום'];
const EXPENSE_HEADERS = ['שם', 'ת.ביצוע', 'ת.פרעון', 'סה"כ', 'הערות', 'סטטוס', 'ת.תשלום'];

const GREEN_DARK  = '#1e7e34';
const GREEN_LIGHT = '#c6efce';
const GREEN_ROW   = '#f0fff4';
const RED_DARK    = '#c0392b';
const RED_LIGHT   = '#ffc7ce';
const RED_ROW     = '#fff0f0';
const WHITE       = '#ffffff';

// -------------------------------------------------------
function doGet(e) {
  const cb = e.parameter.callback || null;
  if (e.parameter.key !== SECRET_KEY)
    return buildResponse({ success: false, error: 'unauthorized' }, cb);

  try {
    const action = e.parameter.action || '';

    if (action === 'upsert' && e.parameter.entry) {
      const entry = JSON.parse(e.parameter.entry);
      PropertiesService.getScriptProperties()
        .setProperty('entry_' + entry.id, JSON.stringify(entry));
      rebuildSheets();
      return buildResponse({ success: true }, cb);
    }

    if (action === 'delete' && e.parameter.id) {
      PropertiesService.getScriptProperties()
        .deleteProperty('entry_' + e.parameter.id);
      rebuildSheets();
      return buildResponse({ success: true }, cb);
    }

    if (action === 'write' && e.parameter.data) {
      const entries = JSON.parse(e.parameter.data);
      const props = PropertiesService.getScriptProperties();
      Object.keys(props.getProperties())
        .filter(k => k.startsWith('entry_'))
        .forEach(k => props.deleteProperty(k));
      entries.forEach(e => props.setProperty('entry_' + e.id, JSON.stringify(e)));
      rebuildSheets();
      return buildResponse({ success: true, count: entries.length }, cb);
    }

    if (action === 'read_sheet') {
      return buildResponse({ success: true, data: readFromSheets() }, cb);
    }

    // ברירת מחדל: קריאה
    return buildResponse({ success: true, data: readAllEntries() }, cb);

  } catch (err) {
    return buildResponse({ success: false, error: err.message }, cb);
  }
}

// -------------------------------------------------------
function readAllEntries() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const entries = Object.keys(props)
    .filter(k => k.startsWith('entry_'))
    .map(k => JSON.parse(props[k]));
  return entries.length > 0 ? entries : readFromSheets();
}

function readFromSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entries = [];
  const fmtDate = v => v instanceof Date
    ? Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd') : '';

  [INCOME_SHEET, EXPENSE_SHEET].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    data.forEach((row, i) => {
      if (!row[0]) return;
      entries.push({
        id: name === INCOME_SHEET ? 'inc_' + i : 'exp_' + i,
        type: name === INCOME_SHEET ? 'income' : 'expense',
        name:     row[0],
        execDate: fmtDate(row[1]),
        payDate:  fmtDate(row[2]),
        amount:   row[3] || 0,
        notes:    row[4] || '',
        status:   statusKey(row[5]),
        paidDate: fmtDate(row[6]),
      });
    });
  });
  return entries;
}

// -------------------------------------------------------
function rebuildSheets() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const all = Object.keys(props)
    .filter(k => k.startsWith('entry_'))
    .map(k => JSON.parse(props[k]));

  writeTab(all.filter(e => e.type === 'income'),  INCOME_SHEET,
           GREEN_DARK, GREEN_LIGHT, GREEN_ROW);
  writeTab(all.filter(e => e.type === 'expense'), EXPENSE_SHEET,
           RED_DARK,   RED_LIGHT,   RED_ROW);
}

function writeTab(entries, sheetName, darkColor, lightColor, rowColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  const headers = sheetName === INCOME_SHEET ? INCOME_HEADERS : EXPENSE_HEADERS;

  // כותרת ראשית (שורה 1)
  // אין שורה ראשית נפרדת — הכותרות ישירות בשורה 1

  // כותרות עמודות — שורה 1
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setFontWeight('bold')
    .setBackground(darkColor)
    .setFontColor(WHITE)
    .setHorizontalAlignment('center');

  // נתונים
  if (entries.length > 0) {
    const rows = entries.map(e => [
      e.name    || '',
      e.execDate  ? new Date(e.execDate)  : '',
      e.payDate   ? new Date(e.payDate)   : '',
      e.amount  || 0,
      e.notes   || '',
      statusLabel(e.status),
      e.paidDate  ? new Date(e.paidDate)  : '',
    ]);
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);

    // עיצוב שורות
    for (let i = 0; i < rows.length; i++) {
      sheet.getRange(i + 2, 1, 1, 7).setBackground(i % 2 === 0 ? rowColor : WHITE);
    }

    // פורמטים
    const n = entries.length;
    sheet.getRange(2, 2, n, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 3, n, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 7, n, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(2, 4, n, 1).setNumberFormat('₪#,##0');
  }

  // רוחב עמודות
  sheet.setColumnWidth(1, 160); // שם
  sheet.setColumnWidth(2, 110); // ת.ביצוע
  sheet.setColumnWidth(3, 110); // ת.פרעון
  sheet.setColumnWidth(4, 90);  // סה"כ
  sheet.setColumnWidth(5, 220); // הערות
  sheet.setColumnWidth(6, 120); // סטטוס
  sheet.setColumnWidth(7, 110); // ת.תשלום

  // הקפא שורת כותרת + פילטר
  sheet.setFrozenRows(1);
  try { const f = sheet.getFilter(); if (f) f.remove(); } catch(e) {}
  if (entries.length > 0) {
    sheet.getRange(1, 1, entries.length + 1, 7).createFilter();
  }
}

// -------------------------------------------------------
function statusLabel(key) {
  const map = {
    paid:          'שולם',
    paid_bit:      'שולם בביט',
    paid_paybox:   'שולם בפייבוקס',
    paid_transfer: 'שולם בהעברה',
    unpaid:        'לא שולם',
    partial:       'שולם חלקית',
  };
  return map[key] || key || '';
}

function statusKey(label) {
  const map = {
    'שולם':           'paid',
    'שולם בביט':      'paid_bit',
    'שולם בפייבוקס':  'paid_paybox',
    'שולם בהעברה':    'paid_transfer',
    'לא שולם':        'unpaid',
    'שולם חלקית':     'partial',
  };
  return map[label] || 'unpaid';
}

function buildResponse(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
