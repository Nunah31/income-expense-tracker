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

  // INC_START=1 (A), EXP_START=10 (J) → צריך 16 עמודות (A-P)
  const data = sheet.getRange(5, 1, lastRow - 4, 16).getValues();
  const entries = [];
  const fmtDate = v => v instanceof Date ? Utilities.formatDate(v, 'Asia/Jerusalem', 'yyyy-MM-dd') : '';

  data.forEach((row, i) => {
    // הכנסות: עמודות A-G = indices 0-6
    if (row[0]) {
      entries.push({
        id: 'inc_' + i,
        type: 'income',
        name: row[0],
        execDate: fmtDate(row[1]),
        payDate:  fmtDate(row[2]),
        amount:   row[3] || 0,
        notes:    row[4] || '',
        status:   statusKey(row[5]),
        paidDate: fmtDate(row[6]),
      });
    }
    // הוצאות: עמודות J-P = indices 9-15
    if (row[9]) {
      entries.push({
        id: 'exp_' + i,
        type: 'expense',
        name: row[9],
        execDate: fmtDate(row[10]),
        payDate:  fmtDate(row[11]),
        amount:   row[12] || 0,
        notes:    row[13] || '',
        status:   statusKey(row[14]),
        paidDate: fmtDate(row[15]),
      });
    }
  });

  return entries;
}

// עמודות: הכנסות A-G, רווח H-I, הוצאות J-P
const INC_START = 1;  // A
const EXP_START = 10; // J

// בניה מחדש של הגיליון מהרשומות שב-Properties
function rebuildSheet() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const allEntries = Object.keys(props)
    .filter(k => k.startsWith('entry_'))
    .map(k => JSON.parse(props[k]));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const income  = allEntries.filter(e => e.type === 'income');
  const expense = allEntries.filter(e => e.type === 'expense');
  const rows = Math.max(income.length, expense.length);

  const GREEN_DARK  = '#1e7e34';
  const GREEN_LIGHT = '#c6efce';
  const RED_DARK    = '#c0392b';
  const RED_LIGHT   = '#ffc7ce';
  const WHITE       = '#ffffff';

  sheet.clear(); // מנקה תוכן + עיצוב

  // ---- כותרות טבלה (שורה 3) ----
  const incTitle = sheet.getRange(3, INC_START, 1, 7);
  incTitle.merge();
  incTitle.setValue('הכנסות')
    .setFontWeight('bold').setFontSize(13).setFontColor(WHITE)
    .setBackground(GREEN_DARK).setHorizontalAlignment('center');

  const expTitle = sheet.getRange(3, EXP_START, 1, 7);
  expTitle.merge();
  expTitle.setValue('הוצאות')
    .setFontWeight('bold').setFontSize(13).setFontColor(WHITE)
    .setBackground(RED_DARK).setHorizontalAlignment('center');

  // ---- כותרות עמודות (שורה 4) ----
  const headers = ['שם', 'ת.ביצוע', 'ת.פרעון', 'סה"כ', 'הערות', 'סטטוס', 'ת.תשלום'];

  const incHeaders = sheet.getRange(4, INC_START, 1, 7);
  incHeaders.setValues([headers])
    .setFontWeight('bold').setBackground(GREEN_LIGHT)
    .setHorizontalAlignment('center');

  const expHeaders = sheet.getRange(4, EXP_START, 1, 7);
  expHeaders.setValues([headers])
    .setFontWeight('bold').setBackground(RED_LIGHT)
    .setHorizontalAlignment('center');

  // ---- נתונים ----
  for (let i = 0; i < rows; i++) {
    const row = i + 5;
    if (income[i]) {
      const e = income[i];
      sheet.getRange(row, INC_START,     1, 1).setValue(e.name || '');
      sheet.getRange(row, INC_START + 1, 1, 1).setValue(e.execDate  ? new Date(e.execDate)  : '');
      sheet.getRange(row, INC_START + 2, 1, 1).setValue(e.payDate   ? new Date(e.payDate)   : '');
      sheet.getRange(row, INC_START + 3, 1, 1).setValue(e.amount || 0);
      sheet.getRange(row, INC_START + 4, 1, 1).setValue(e.notes || '');
      sheet.getRange(row, INC_START + 5, 1, 1).setValue(statusLabel(e.status));
      sheet.getRange(row, INC_START + 6, 1, 1).setValue(e.paidDate  ? new Date(e.paidDate)  : '');
      sheet.getRange(row, INC_START, 1, 7).setBackground('#f0fff4'); // ירוק בהיר מאוד לשורות
    }
    if (expense[i]) {
      const e = expense[i];
      sheet.getRange(row, EXP_START,     1, 1).setValue(e.name || '');
      sheet.getRange(row, EXP_START + 1, 1, 1).setValue(e.execDate  ? new Date(e.execDate)  : '');
      sheet.getRange(row, EXP_START + 2, 1, 1).setValue(e.payDate   ? new Date(e.payDate)   : '');
      sheet.getRange(row, EXP_START + 3, 1, 1).setValue(e.amount || 0);
      sheet.getRange(row, EXP_START + 4, 1, 1).setValue(e.notes || '');
      sheet.getRange(row, EXP_START + 5, 1, 1).setValue(statusLabel(e.status));
      sheet.getRange(row, EXP_START + 6, 1, 1).setValue(e.paidDate  ? new Date(e.paidDate)  : '');
      sheet.getRange(row, EXP_START, 1, 7).setBackground('#fff0f0'); // ורוד בהיר מאוד לשורות
    }
  }

  // ---- פורמט תאריכים וסכומים ----
  if (rows > 0) {
    sheet.getRange(5, INC_START + 1, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, INC_START + 2, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, INC_START + 6, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, EXP_START + 1, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, EXP_START + 2, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, EXP_START + 6, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, INC_START + 3, rows, 1).setNumberFormat('₪#,##0');
    sheet.getRange(5, EXP_START + 3, rows, 1).setNumberFormat('₪#,##0');
  }

  // ---- רוחב עמודות ----
  sheet.setColumnWidth(INC_START,     160);
  sheet.setColumnWidth(INC_START + 4, 200);
  sheet.setColumnWidth(EXP_START,     140);
  sheet.setColumnWidth(EXP_START + 4, 200);

  // ---- הקפאת שורות כותרת + פילטרים ----
  sheet.setFrozenRows(4);
  // מחק פילטר קיים (אם יש) ואז צור חדש על כל הטווח
  try { const f = sheet.getFilter(); if (f) f.remove(); } catch(e) {}
  if (rows > 0) {
    sheet.getRange(4, INC_START, rows + 1, EXP_START + 6).createFilter();
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
