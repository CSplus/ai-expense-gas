/*************************************************
 * Export.gs
 * メール送信・CSV出力
 *************************************************/

function sendMonthlyExpenseReport() {
  const summarySheet = getSummarySheet();
  const values = summarySheet.getDataRange().getValues();
  const title = values[0][0];
  const report = {};

  for (let i = 3; i <= 8; i++) {
    const category = values[i][0];
    const amount = values[i][1];
    report[category] = amount;
  }

  const total = values[9][1];

  let body = '';
  body += title + '\n\n';
  body += '■接待交際費\n' + Number(report['接待交際費'] || 0).toLocaleString() + '円\n\n';
  body += '■旅費交通費\n' + Number(report['旅費交通費'] || 0).toLocaleString() + '円\n\n';
  body += '■車両費\n' + Number(report['車両費'] || 0).toLocaleString() + '円\n\n';
  body += '■通信費\n' + Number(report['通信費'] || 0).toLocaleString() + '円\n\n';
  body += '■消耗品費\n' + Number(report['消耗品費'] || 0).toLocaleString() + '円\n\n';
  body += '■その他\n' + Number(report['その他'] || 0).toLocaleString() + '円\n\n';
  body += '------------------------\n';
  body += '合計\n' + Number(total || 0).toLocaleString() + '円\n\n';
  body += '詳細CSVを添付します。\n';

  const csvBlob = createExpenseCsvBlobFromSummary();

  GmailApp.sendEmail(
    getNoticeEmail(),
    '【経費報告】' + title,
    body,
    {
      attachments: [csvBlob]
    }
  );
}

function createExpenseCsvBlobFromSummary() {
  const summarySheet = getSummarySheet();
  const values = summarySheet.getDataRange().getValues();
  const title = values[0][0] || '経費集計';
  const detailStartRow = findRowByText(summarySheet, '明細一覧');

  if (!detailStartRow) {
    throw new Error('月次集計シートに明細一覧が見つかりません。');
  }

  const lastRow = summarySheet.getLastRow();
  const dataRange = summarySheet.getRange(detailStartRow + 1, 1, lastRow - detailStartRow, 9);
  const data = dataRange.getValues();

  const csv = data.map(row => {
    return row.map(value => {
      const text = value instanceof Date
        ? Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd')
        : String(value ?? '');

      return '"' + text.replace(/"/g, '""') + '"';
    }).join(',');
  }).join('\r\n');

  const fileName = title.replace(/\s/g, '_') + '_明細.csv';
  return Utilities.newBlob('\uFEFF' + csv, 'text/csv', fileName);
}

function findRowByText(sheet, text) {
  const values = sheet.getDataRange().getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === text) return i + 1;
  }

  return null;
}

/**
 * TKC FX2 仕訳明細のAU列（47列目）に登録番号を設定するためのヘルパー。
 * 既存のFX2出力実装へ組み込む際は、仕訳明細配列を47列以上にし、
 * この関数で T + 13桁形式の登録番号のみを AU 列へ出力する。
 * TODO: FX2仕訳明細テキスト出力本体の実装時に、税区分マスタ参照とあわせて呼び出す。
 */
function applyFx2InvoiceNumberToJournalRow(rowValues, invoiceNumber) {
  const normalized = normalizeInvoiceNumber(invoiceNumber);
  while (rowValues.length < 47) rowValues.push('');
  rowValues[46] = normalized || '';
  return rowValues;
}
