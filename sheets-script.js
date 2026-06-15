// Google Apps Script - הדבק קוד זה ב-Google Apps Script ופרוס כ-Web App
// פרויקט: הכנסות והוצאות 2025
// גיליון: זכות

const SHEET_NAME = 'זכות';
const SECRET_KEY = 'habari25'; // ← חייב להיות זהה לקוד באפליקציה

function doGet(e) {
  const cb = e.parameter.callback || null;
  if (e.parameter.key !== SECRET_KEY) {
    return buildResponse({ success: false, error: 'unauthorized' }, cb);
  }
  try {
    if (e.parameter && e.parameter.data) {
      const entries = JSON.parse(e.parameter.data);
      syncSheet(entries);
      return buildResponse({ success: true, count: entries.length }, cb);
    }
    const data = readSheet();
    return buildResponse({ success: true, data }, cb);
  } catch (err) {
    return buildResponse({ success: false, error: err.message }, cb);
  }
}

function doPost(e) {
  try {
    let entries;
    if (e.postData && e.postData.contents) {
      entries = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      entries = JSON.parse(e.parameter.data);
    } else {
      return buildResponse({ success: false, error: 'no data' });
    }
    syncSheet(entries);
    return buildResponse({ success: true, count: entries.length });
  } catch (err) {
    return buildResponse({ success: false, error: err.message });
  }
}

function syncSheet(entries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const income = entries.filter(e => e.type === 'income');
  const expense = entries.filter(e => e.type === 'expense');
  const rows = Math.max(income.length, expense.length);

  // כותרות
  sheet.clearContents();
  sheet.getRange('A3').setValue('הכנסות').setFontWeight('bold').setFontSize(13);
  sheet.getRange('H3').setValue('הוצאות').setFontWeight('bold').setFontSize(13);

  const headers = ['שם', 'ת.ביצוע', 'ת.פרעון', 'סה"כ', 'סטטוס', 'הערות'];
  // הכנסות - עמודות A-F
  headers.forEach((h, i) => sheet.getRange(4, i + 1).setValue(h).setFontWeight('bold'));
  // הוצאות - עמודות H-M
  headers.forEach((h, i) => sheet.getRange(4, i + 8).setValue(h).setFontWeight('bold'));

  for (let i = 0; i < rows; i++) {
    const row = i + 5;
    if (income[i]) {
      const e = income[i];
      sheet.getRange(row, 1).setValue(e.name || '');
      sheet.getRange(row, 2).setValue(e.execDate ? new Date(e.execDate) : '');
      sheet.getRange(row, 3).setValue(e.payDate ? new Date(e.payDate) : '');
      sheet.getRange(row, 4).setValue(e.amount || 0);
      sheet.getRange(row, 5).setValue(statusLabel(e.status));
      sheet.getRange(row, 6).setValue(e.notes || '');
    }
    if (expense[i]) {
      const e = expense[i];
      sheet.getRange(row, 8).setValue(e.name || '');
      sheet.getRange(row, 9).setValue(e.execDate ? new Date(e.execDate) : '');
      sheet.getRange(row, 10).setValue(e.payDate ? new Date(e.payDate) : '');
      sheet.getRange(row, 11).setValue(e.amount || 0);
      sheet.getRange(row, 12).setValue(statusLabel(e.status));
      sheet.getRange(row, 13).setValue(e.notes || '');
    }
  }

  // עיצוב תאריכים
  if (rows > 0) {
    sheet.getRange(5, 2, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 3, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 9, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 10, rows, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(5, 4, rows, 1).setNumberFormat('₪#,##0');
    sheet.getRange(5, 11, rows, 1).setNumberFormat('₪#,##0');
  }
}

function readSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return [];

  const data = sheet.getRange(5, 1, lastRow - 4, 13).getValues();
  const entries = [];

  data.forEach((row, i) => {
    if (row[0]) {
      entries.push({
        id: 'inc_' + i,
        type: 'income',
        name: row[0],
        execDate: row[1] ? Utilities.formatDate(new Date(row[1]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
        payDate: row[2] ? Utilities.formatDate(new Date(row[2]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
        amount: row[3] || 0,
        status: statusKey(row[4]),
        notes: row[5] || '',
      });
    }
    if (row[7]) {
      entries.push({
        id: 'exp_' + i,
        type: 'expense',
        name: row[7],
        execDate: row[8] ? Utilities.formatDate(new Date(row[8]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
        payDate: row[9] ? Utilities.formatDate(new Date(row[9]), 'Asia/Jerusalem', 'yyyy-MM-dd') : '',
        amount: row[10] || 0,
        status: statusKey(row[11]),
        notes: row[12] || '',
      });
    }
  });

  return entries;
}

function statusLabel(key) {
  const map = { paid: 'שולם', unpaid: 'לא שולם', partial: 'שולם חלקית' };
  return map[key] || key || '';
}

function statusKey(label) {
  const map = { 'שולם': 'paid', 'לא שולם': 'unpaid', 'שולם חלקית': 'partial' };
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
