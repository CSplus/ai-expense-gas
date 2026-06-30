/*************************************************
 * Summary.gs
 * 月次集計
 *************************************************/

function createMonthlySummary() {
  const today = new Date();
  const target = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  createMonthlySummaryFor(target.getFullYear(), target.getMonth() + 1);
}

function createMonthlySummaryFor(year, month) {
  const expenseSheet = getExpenseSheet();
  const summarySheet = getSummarySheet();
  summarySheet.clear();

  ensureExpenseInvoiceColumns(expenseSheet);
  const col = getExpenseColumnsByName(['date', 'vendor', 'amount', 'accountName', 'status', 'confirm', 'duplicate', 'duplicateId', 'summaryTarget']);
  const values = expenseSheet.getDataRange().getValues();
  const summary = {};
  const excludedRows = [];
  const detailRows = [];
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 1);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const date = toDate(row[col.date - 1]);
    if (!date) continue;
    if (date < start || date >= end) continue;

    const vendor = row[col.vendor - 1];
    const amount = Number(row[col.amount - 1]);
    const category = row[col.accountName - 1] || 'その他';
    const status = row[col.status - 1];
    const confirm = row[col.confirm - 1] || '未確認';
    const duplicate = row[col.duplicate - 1] || '';
    const duplicateId = row[col.duplicateId - 1] || '';
    const summaryTarget = row[col.summaryTarget - 1] || '対象';

    if (!amount) continue;

    const isIncluded = confirm === '確認済' && summaryTarget === '対象';

    detailRows.push([
      formatDate(date),
      vendor,
      amount,
      category,
      status,
      confirm,
      duplicate,
      duplicateId,
      summaryTarget
    ]);

    if (isIncluded) {
      if (!summary[category]) summary[category] = 0;
      summary[category] += amount;
    } else {
      excludedRows.push([
        formatDate(date),
        vendor,
        amount,
        category,
        confirm,
        duplicate,
        duplicateId,
        summaryTarget
      ]);
    }
  }

  let r = 1;
  summarySheet.getRange(r, 1).setValue(`${year}年${month}月 経費集計`);
  summarySheet.getRange(r, 1).setFontWeight('bold');
  r += 2;

  summarySheet.getRange(r, 1, 1, 2).setValues([['勘定科目', '金額']]);
  summarySheet.getRange(r, 1, 1, 2).setFontWeight('bold');
  r++;

  let total = 0;
  const categories = ['接待交際費', '旅費交通費', '車両費', '通信費', '消耗品費', 'その他'];

  categories.forEach(category => {
    const amount = summary[category] || 0;
    summarySheet.getRange(r, 1, 1, 2).setValues([[category, amount]]);
    total += amount;
    r++;
  });

  summarySheet.getRange(r, 1, 1, 2).setValues([['合計', total]]);
  summarySheet.getRange(r, 1, 1, 2).setFontWeight('bold');
  r += 2;

  summarySheet.getRange(r, 1).setValue('集計対象外・確認が必要な明細');
  summarySheet.getRange(r, 1).setFontWeight('bold');
  r++;

  summarySheet.getRange(r, 1, 1, 8)
    .setValues([['取引日', '店舗名', '金額', '勘定科目', '確認', '重複判定', '重複候補ID', '集計対象']]);
  summarySheet.getRange(r, 1, 1, 8).setFontWeight('bold');
  r++;

  if (excludedRows.length > 0) {
    summarySheet.getRange(r, 1, excludedRows.length, 8).setValues(excludedRows);
    r += excludedRows.length;
  } else {
    summarySheet.getRange(r, 1).setValue('なし');
    r++;
  }

  r += 2;
  summarySheet.getRange(r, 1).setValue('明細一覧');
  summarySheet.getRange(r, 1).setFontWeight('bold');
  r++;

  summarySheet.getRange(r, 1, 1, 9)
    .setValues([['取引日', '店舗名', '金額', '勘定科目', '処理状態', '確認', '重複判定', '重複候補ID', '集計対象']]);
  summarySheet.getRange(r, 1, 1, 9).setFontWeight('bold');
  r++;

  if (detailRows.length > 0) {
    summarySheet.getRange(r, 1, detailRows.length, 9).setValues(detailRows);
  }

  summarySheet.autoResizeColumns(1, 9);
}

function toDate(value) {
  if (value instanceof Date) return value;

  if (typeof value === 'string') {
    const normalized = value.replace(/\//g, '-');
    const date = new Date(normalized);
    if (!isNaN(date.getTime())) return date;
  }

  return null;
}

function formatDate(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}
