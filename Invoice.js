/*************************************************
 * Invoice.gs
 * インボイス登録番号処理（API未使用版）
 *
 * 当面は国税庁Web-APIを呼び出さず、Geminiが抽出した登録番号を保存する。
 * 将来API利用可能になったら、このファイルの getInvoiceInfo() だけを差し替える。
 *************************************************/
/**
 * 登録番号を正規化する
 * T + 数字13桁のみ有効
 */
function normalizeInvoiceNumber(value) {
  if (!value) return '';

  const text = String(value)
    .replace(/[Ｔｔ]/g, 'T')
    .replace(/[０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[^Tt0-9]/g, '')
    .toUpperCase();

  if (/^\d{13}$/.test(text)) return 'T' + text;

  const match = text.match(/T\d{13}/);
  if (!match) return '';

  return match[0];
}

/**
 * API未使用版のインボイス判定
 * ※正式な適格判定は国税庁API利用開始後に行う
 */
function getInvoiceInfo(invoiceNumber, invoiceJudgement) {
  const normalized = normalizeInvoiceNumber(invoiceNumber);

  let judgement = '不明';

  if (normalized) {
    judgement = '登録番号あり（未照会）';
  } else if (invoiceJudgement === '登録番号なし') {
    judgement = '登録番号なし';
  }

  return {
    registrationNumber: normalized,
    invoiceJudgement: judgement,
    invoiceStatus: 'API未使用',
    officialName: '',
    checkedAt: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm'),
    note: normalized
      ? 'Gemini判定（API未使用・正式名称未確認）'
      : '登録番号を確認できませんでした'
  };
}