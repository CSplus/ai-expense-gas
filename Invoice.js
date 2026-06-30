/*************************************************
 * Invoice.gs
 * 登録番号処理（API未使用版）
 *
 * 当面は国税庁Web-APIを呼び出さず、Geminiが抽出した登録番号を保存する。
 * 将来API利用可能になったら、このファイルの getInvoiceInfo() だけを差し替える。
 *************************************************/
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
    registrationNumber: normalized || '',
    invoiceJudgement: judgement,
    invoiceStatus: 'API未使用',
    officialName: '',
    checkedAt: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm'),
    note: normalized
      ? 'Gemini判定（API未使用・正式名称未確認）'
      : '登録番号を確認できませんでした'
  };
}