/*************************************************
 * Invoice.gs
 * インボイス登録番号処理（API未使用版）
 *
 * 当面は国税庁Web-APIを呼び出さず、Geminiが抽出した登録番号を保存する。
 * 将来API利用可能になったら、このファイルの getInvoiceInfo() だけを差し替える。
 *************************************************/
/**
 * API未使用版のインボイス判定
 * ※正式な適格判定は国税庁API利用開始後に行う
 */
function getInvoiceInfo(invoiceNumber, invoiceJudgement) {
  const hasInput = invoiceNumber !== null && invoiceNumber !== undefined && String(invoiceNumber).trim() !== '';
  const normalized = normalizeInvoiceNumber(invoiceNumber);

  let judgement = 'API未確認';
  let status = '未確認';
  let note = 'API未使用';

  if (normalized) {
    judgement = '登録番号あり';
    note = 'Gemini判定（API未使用・正式名称未確認）';
  } else if (hasInput) {
    judgement = '番号不正';
    status = '番号不正';
    note = '登録番号が T + 13桁形式ではありません';
  } else if (invoiceJudgement === '登録番号なし') {
    judgement = '登録番号なし';
    note = '登録番号を確認できませんでした';
  }

  return {
    registrationNumber: normalized || '',
    invoiceJudgement: judgement,
    invoiceStatus: status,
    officialName: '',
    address: '',
    checkedAt: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm'),
    note: note
  };
}