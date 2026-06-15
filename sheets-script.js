// Google Apps Script - הדבק קוד זה ב-Google Apps Script ופרוס כ-Web App
// פרויקט: הכנסות והוצאות 2025
// גיליון: זכות

const SHEET_NAME = 'זכות';
const SECRET_KEY = 'esther-udi-2025';

function doGet(e) {
  const cb = e.parameter.callback || null;
  if (e.parameter.key !== SECRET_KEY) {
    return buildResponse({ success: false, error: 'unauthorized' }, cb);
  }
  try {
    const action = e.parameter.action || '';

    // upsert: שמירת רשומה אחת (הוספה או עדכון)
    if (action === 'upsert' && e.parameter.entry) {
      const entry = JSON.parse(e.parameter.entry);
      const props = PropertiesService.getScriptProperties();
      props.setProperty('entry_' + entry.id, JSON.stringify(entry));
      rebuildSheet();
      return buildResponse({ success: true }, cb);
    }

    // delete: מחיקת רשומה לפי ID
    if (action === 'delete' && e.parameter.id) {
      const props = PropertiesService.getScriptProperties();
      props.deleteProperty('entry_' + e.parameter.id);
      rebuildSheet();
      return buildResponse({ success: true }, cb);
    }

    // write: שמירת כל הרשומות בבת אחת (גיבוי מלא)
    if (action === 'write' && e.parameter.data) {
      const entries = JSON.parse(e.parameter.data);
      const props = PropertiesService.getScriptProperties();
      // מחק רשומות קיימות
      const existing = props.getProperties();
      Object.keys(existing).filter(k => k.startsWith('entry_')).forEach(k => props.deleteProperty(k));
      // שמור חדשות
      entries.forEach(entry => props.setProperty('entry_' + entry.id, JSON.stringify(entry)));
      rebuildSheet();
      return buildResponse({ success: true, count: entries.length }, cb);
    }

    // read (ברירת מחדל): קריאת כל הרשומות
    const data = readAllEntries();
    return buildResponse({ success: true, data: data }, cb);

  } catch (err) {
    return buildResponse({ success: false, error: err.message }, cb);
  }
}

function doPost(e) {
  if (e.parameter && e.parameter.key !== SECRET_KEY) {
    return buildResponse({ success: false, error: 'unauthorized' });
  }
  try {
    let entries;
    if (e.parameter && e.parameter.data) {
      entries = JSON.parse(e.parameter.data);
    } else if (e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      entries = body.data || body;
    } else {
      return buildResponse({ success: false, error: 'no data' });
    }
    const props = PropertiesService.getScriptProperties();
    const existing = props.getProperties();
    Object.keys(existing).filter(k => k.startsWith('entry_')).forEach(k => props.deleteProperty(k));
    entries.forEach(entry => props.setProperty('entry_' + entry.id, JSON.stringify(entry)));
    rebuildSheet();
    return buildResponse({ success: true, count: entries.length });
  } catch (err) {
    return buildResponse({ success: false, error: err.message });
  }
}

// קריאת כל הרשומות מה-Script Properties
function readAllEntries() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const entries = Object.keys(props)
    .filter(k => k.startsWith('entry_'))
    .map(k => JSON.parse(props[k]));

  // אם אין רשומות ב-Properties, נסה לקרוא מהגיליון (מצב ראשוני)
  if (entries.length === 0) {
    return readFromSheet();
  }
  return entries;
}

// קריאה ישירה מהגיליון (רק לצורך מיגרציה ראשונית)
function readFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return [];

  const data = sheet.getRange(5, 1, lastRow - 4, 14).getValues();
  const entries = [];

  data.forEach((row, i) => {
    if (row[0]) {
      entries.push({
        id: row[6] || ('inc_' + i),
        type: 'income',
        name: row[0],
        execDate: row[1] ? Utilities.formatDate(new Date(row[1]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
        payDate: row[2] ? Utilities.formatDate(new Date(row[2]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
        amount: row[3] || 0,
        notes: row[4] || '',
        status: statusKey(row[5]),
        paidDate: row[6] && row[6] instanceof Date ? Utilities.formatDate(new Date(row[6]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
      });
    }
    if (row[7]) {
      entries.push({
        id: row[13] || ('exp_' + i),
        type: 'expense',
        name: row[7],
        execDate: row[8] ? Utilities.formatDate(new Date(row[8]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
        payDate: row[9] ? Utilities.formatDate(new Date(row[9]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
        amount: row[10] || 0,
        notes: row[11] || '',
        status: statusKey(row[12]),
        paidDate: row[13] && row[13] instanceof Date ? Utilities.formatDate(new Date(row[13]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
      });
    }
  });

  return entries;
}

// בניה מחדש של הגיליון מהרשומות שב-Properties
function rebuildSheet() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const allEntries = Object.keys(props)
    .filter(k => k.startsWith('entry_'))
    .map(k => JSON.parse(props[k]));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const income = allEntries.filter(e => e.type === 'income');
  const expense = allEntries.filter(e => e.type === 'expense');
  const rows = Math.max(income.length, expense.length);

  sheet.clearContents();
  sheet.getRange('A3').setValue('הכנסות').setFontWeight('bold').setFontSize(13);
  sheet.getRange('H3').setValue('הוצאות').setFontWeight('bold').setFontSize(13);

  const headers = ['שם', 'ת.ביצוע', 'ת.פרעון', 'סה"כ', 'הערות', 'סטטוס', 'ת.תשלום'];
  headers.forEach((h, i) => sheet.getRange(4, i + 1).setValue(h).setFontWeight('bold'));
  headers.forEach((h, i) => sheet.getRange(4, i + 8).setValue(h).setFontWeight('bold'));

  for (let i = 0; i < rows; i++) {
    const row = i + 5;
    if (income[i]) {
      const e = income[i];
      sheet.getRange(row, 1).setValue(e.name || '');
      sheet.getRange(row, 2).setValue(e.execDate ? new Date(e.execDate) : '');
      sheet.getRange(row, 3).setValue(e.payDate ? new Date(e.payDate) : '');
      sheet.getRange(row, 4).setValue(e.amount || 0);
      sheet.getRange(row, 5).setValue(e.notes || '');
      sheet.getRange(row, 6).setValue(statusLabel(e.status));
      sheet.getRange(row, 7).setValue(e.paidDate ? new Date(e.paidDate) : '');
    }
    if (expense[i]) {
      const e = expense[i];
      sheet.getRange(row, 8).setValue(e.name || '');
      sheet.getRange(row, 9).setValue(e.execDate ? new Date(e.execDate) : '');
      sheet.getRange(row, 10).setValue(e.payDate ? new Date(e.payDate) : '');
      sheet.getRange(row, 11).setValue(e.amount || 0);
      sheet.getRange(row, 12).setValue(e.notes || '');
      sheet.getRange(row, 13).setValue(statusLabel(e.status));
      sheet.getRange(row, 14).setValue(e.paidDate ? new Date(e.paidDate) : '');
    }
  }

  if (rows > 0) {
    sheet.getRange(5, 2, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 3, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 7, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 9, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 10, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 14, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 4, rows, 1).setNumberFormat('₪#,##0');
    sheet.getRange(5, 11, rows, 1).setNumberFormat('₪#,##0');
  }
}

function statusLabel(key) {
  const map = {
    paid: 'שולם',
    paid_bit: 'שולם בביט',
    paid_paybox: 'שולם בפייבוקס',
    paid_transfer: 'שולם בהעברה',
    unpaid: 'לא שולם',
    partial: 'שולם חלקית',
  };
  return map[key] || key || '';
}

function statusKey(label) {
  const map = {
    'שולם': 'paid',
    'שולם בביט': 'paid_bit',
    'שולם בפייבוקס': 'paid_paybox',
    'שולם בהעברה': 'paid_transfer',
    'לא שולם': 'unpaid',
    'שולם חלקית': 'partial',
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
