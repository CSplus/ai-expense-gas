/*************************************************
 * Duplicate.gs
 * 重複検出
 *************************************************/

function detectDuplicateExpenses() {
  const sheet = getExpenseSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastColumn = Math.max(sheet.getLastColumn(), COL.INVOICE_NOTE);
  const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const map = {};
  let dupNo = 1;

  // 既存の重複判定を初期化。ただし集計対象は空・保留・対象のみ対象へ戻す。
  for (let i = 0; i < data.length; i++) {
    const rowNo = i + 2;
    sheet.getRange(rowNo, COL.DUPLICATE).setValue('');
    sheet.getRange(rowNo, COL.DUPLICATE_ID).setValue('');

    const current = sheet.getRange(rowNo, COL.SUMMARY_TARGET).getValue();
    if (current === '' || current === '保留' || current === '対象') {
      sheet.getRange(rowNo, COL.SUMMARY_TARGET).setValue('対象');
    }
  }

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const date = row[COL.DATE - 1];
    const vendor = row[COL.VENDOR_NORMALIZED - 1] || row[COL.VENDOR - 1];
    const amount = row[COL.AMOUNT - 1];
    const invoiceNumber = normalizeInvoiceNumber(row[COL.INVOICE_NUMBER - 1]);

    if (!date || !vendor || !amount) continue;

    const key = formatDuplicateDate(date) + '|' + vendor + '|' + amount + (invoiceNumber ? '|' + invoiceNumber : '');
    if (!map[key]) map[key] = [];
    map[key].push(i + 2);
  }

  Object.keys(map).forEach(key => {
    const rows = map[key];
    if (rows.length < 2) return;

    const dupId = 'DUP' + Utilities.formatString('%04d', dupNo++);

    rows.forEach(rowNo => {
      sheet.getRange(rowNo, COL.DUPLICATE).setValue('重複候補');
      sheet.getRange(rowNo, COL.DUPLICATE_ID).setValue(dupId);
      sheet.getRange(rowNo, COL.SUMMARY_TARGET).setValue('保留');
    });
  });
}

function formatDuplicateDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd');
  }

  return String(value).replace(/\//g, '-').trim();
}

