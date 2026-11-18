/**
 * يستقبل ردود تأكيد الحضور ويحفظها في Google Sheets.
 * - ورقة "الردود": الاسم، الحالة، العدد، آخر تحديث
 * - ورقة "الملخص": المجاميع تلقائياً
 * إذا نفس الجوال أرسل مرة ثانية، يتحدث رده بدل ما يتكرر.
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

    const id = String(p.id || '').slice(0, 40);
    const row = [name, status, count, new Date(), id];
    const last = sheet.getLastRow();
    let target = 0;
    if (id && last > 1) {
      const ids = sheet.getRange(2, 5, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === id) { target = i + 2; break; }
      }
    }
    if (target) sheet.getRange(target, 1, 1, 5).setValues([row]);
    else sheet.appendRow(row);

    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheets_(ss) {
  let sheet = ss.getSheetByName('الردود');
  if (!sheet) {
    sheet = ss.insertSheet('الردود');
    sheet.setRightToLeft(true);
    sheet.appendRow(['الاسم', 'الحالة', 'عدد الحضور', 'آخر تحديث', 'رقم الجهاز']);
    sheet.getRange('A1:E1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideColumns(5);
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
