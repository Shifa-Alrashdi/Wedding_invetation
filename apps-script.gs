/**
 * يستقبل ردود تأكيد الحضور ويحفظها في Google Sheets.
 * - ورقة "الردود": الاسم، الحالة، العدد، الجهة، آخر تحديث
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
    const side = String(p.side || '').slice(0, 30);
    const row = [name, status, count, side, new Date(), id];
    const last = sheet.getLastRow();
    let target = 0;
    if (id && last > 1) {
      const ids = sheet.getRange(2, 6, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === id) { target = i + 2; break; }
      }
    }
    if (target) sheet.getRange(target, 1, 1, 6).setValues([row]);
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
    sheet.appendRow(['الاسم', 'الحالة', 'عدد الحضور', 'الجهة', 'آخر تحديث', 'رقم الجهاز']);
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideColumns(6);
  }
  if (!ss.getSheetByName('الملخص')) {
    const s = ss.insertSheet('الملخص', 0);
    s.setRightToLeft(true);
    const R = "'الردود'!";
    s.getRange('A1:D1').setValues([['', 'أهل العروس', 'أهل العريس', 'المجموع']]);
    const rows = [
      ['الحاضرين المؤكدين (أشخاص)', 'SUMIFS(' + R + 'C2:C,' + R + 'B2:B,"حاضر",' + R + 'D2:D,"%s")', 'SUMIF(' + R + 'B2:B,"حاضر",' + R + 'C2:C)'],
      ['"ربما" (أشخاص)', 'SUMIFS(' + R + 'C2:C,' + R + 'B2:B,"ربما",' + R + 'D2:D,"%s")', 'SUMIF(' + R + 'B2:B,"ربما",' + R + 'C2:C)'],
      ['عدد ردود الحضور', 'COUNTIFS(' + R + 'B2:B,"حاضر",' + R + 'D2:D,"%s")', 'COUNTIF(' + R + 'B2:B,"حاضر")'],
      ['عدد ردود ربما', 'COUNTIFS(' + R + 'B2:B,"ربما",' + R + 'D2:D,"%s")', 'COUNTIF(' + R + 'B2:B,"ربما")'],
      ['عدد المعتذرين', 'COUNTIFS(' + R + 'B2:B,"معتذر",' + R + 'D2:D,"%s")', 'COUNTIF(' + R + 'B2:B,"معتذر")'],
      ['إجمالي الردود', 'COUNTIFS(' + R + 'A2:A,"<>",' + R + 'D2:D,"%s")', 'COUNTA(' + R + 'A2:A)']
    ];
    s.getRange(2, 1, rows.length, 4).setValues(rows.map(r => [
      r[0], '=' + r[1].replace('%s', 'أهل العروس'), '=' + r[1].replace('%s', 'أهل العريس'), '=' + r[2]
    ]));
    s.getRange('A1:D1').setFontWeight('bold');
    s.getRange('A2:A7').setFontWeight('bold');
    s.getRange('B2:D2').setFontSize(16).setFontWeight('bold');
    s.autoResizeColumn(1);
  }
  return sheet;
}
