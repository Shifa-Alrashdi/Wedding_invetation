/**
 * يستقبل ردود تأكيد الحضور ويحفظها في Google Sheets.
 * - ورقة "الردود": الاسم، الحالة، العدد، آخر تحديث
 * - ورقة "الملخص": المجاميع تلقائياً
 * إذا نفس الاسم أرسل مرة ثانية، يتحدث رده بدل ما يتكرر.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheets_(ss);

    const p = e.parameter || {};
    const name = String(p.name || '').trim().slice(0, 100);
    const status = ['حاضر', 'ربما', 'معتذر'].indexOf(p.status) >= 0 ? p.status : 'معتذر';
    const count = status !== 'معتذر' ? Math.max(1, Math.min(20, parseInt(p.count, 10) || 1)) : 0;
    if (!name) return ContentService.createTextOutput('missing name');

    const row = [name, status, count, new Date()];
    const key = normalize_(name);
    const last = sheet.getLastRow();
    let target = 0;
    if (last > 1) {
      const names = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < names.length; i++) {
        if (normalize_(names[i][0]) === key) { target = i + 2; break; }
      }
    }
    if (target) sheet.getRange(target, 1, 1, 4).setValues([row]);
    else sheet.appendRow(row);

    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function normalize_(s) {
  return String(s).trim().replace(/\s+/g, ' ')
    .replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي');
}

function getOrCreateSheets_(ss) {
  let sheet = ss.getSheetByName('الردود');
  if (!sheet) {
    sheet = ss.insertSheet('الردود');
    sheet.setRightToLeft(true);
    sheet.appendRow(['الاسم', 'الحالة', 'عدد الحضور', 'آخر تحديث']);
    sheet.getRange('A1:D1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  if (!ss.getSheetByName('الملخص')) {
    const s = ss.insertSheet('الملخص', 0);
    s.setRightToLeft(true);
    s.getRange('A1:B6').setValues([
      ['مجموع الحاضرين المؤكدين (أشخاص)', "=SUMIF('الردود'!B2:B,\"حاضر\",'الردود'!C2:C)"],
      ['مجموع \"ربما\" (أشخاص)', "=SUMIF('الردود'!B2:B,\"ربما\",'الردود'!C2:C)"],
      ['عدد الردود بالحضور', "=COUNTIF('الردود'!B2:B,\"حاضر\")"],
      ['عدد ردود ربما', "=COUNTIF('الردود'!B2:B,\"ربما\")"],
      ['عدد المعتذرين', "=COUNTIF('الردود'!B2:B,\"معتذر\")"],
      ['إجمالي الردود', "=COUNTA('الردود'!A2:A)"]
    ]);
    s.getRange('A1:A6').setFontWeight('bold');
    s.getRange('B1').setFontSize(18).setFontWeight('bold');
    s.autoResizeColumn(1);
  }
  return sheet;
}
