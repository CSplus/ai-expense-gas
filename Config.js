/*************************************************
 * Config.gs
 * 共通設定・シート名・列番号・設定取得
 *************************************************/

const TIMEZONE = 'Asia/Tokyo';

const SHEET_EXPENSE = '経費台帳';
const SHEET_SUMMARY = '月次集計';
const SHEET_CONFIG  = 'システム設定';
const SHEET_RULE    = '仕分けルール';
const SHEET_ACCOUNT = '勘定科目マスタ';
const SHEET_TAX_RATE = '税率マスタ';

// 経費台帳の列番号（1始まり）
const COL = {
  TIMESTAMP: 1,        // A タイムスタンプ
  RECEIPT_URL: 2,      // B 領収書画像アップロード
  MEMO: 3,             // C 内容のメモ
  DATE: 4,             // D 取引日
  VENDOR: 5,           // E 店舗名
  AMOUNT: 6,           // F 金額
  ACCOUNT_CODE: 7,     // G 勘定科目コード
  ACCOUNT_NAME: 8,     // H 勘定科目名
  STATUS: 9,           // I 処理状態
  FILE_ID: 10,         // J ファイルID
  ERROR: 11,           // K エラー内容
  VENDOR_NORMALIZED: 12,// L 取引先正規名
  PAYMENT_METHOD: 13,  // M 支払方法
  EVIDENCE_TYPE: 14,   // N 証憑種別
  SOURCE_FILE: 15,     // O 元ファイル名
  CONFIRM: 16,         // P 確認
  INPUT_CATEGORY: 17,  // Q 入力区分
  DUPLICATE: 18,       // R 重複判定
  DUPLICATE_ID: 19,    // S 重複候補ID
  SUMMARY_TARGET: 20,  // T 集計対象
  INVOICE_NUMBER: 21,  // U 登録番号
  INVOICE_JUDGEMENT: 22,// V インボイス判定
  INVOICE_STATUS: 23,  // W インボイス登録状態
  INVOICE_CHECKED_AT: 24,// X インボイス確認日
  TAX_RATE: 25,        // Y 税率
  TAX_AMOUNT: 26,      // Z 消費税額
  INVOICE_NOTE: 27,    // AA インボイス備考
  INVOICE_REGISTERED_NAME: 28, // AB インボイス正式名称
  INVOICE_REGISTRATION_DATE: 29, // AC インボイス登録年月日
  INVOICE_EXPIRE_DATE: 30, // AD インボイス失効年月日
  INVOICE_API_ERROR: 31 // AE インボイスAPIエラー
};

function getConfig(key) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_CONFIG);

  if (!sheet) {
    throw new Error('システム設定シートが見つかりません。');
  }

  const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();

  for (const row of values) {
    if (String(row[0]).trim() === key) {
      return String(row[1]).trim();
    }
  }

  throw new Error('設定項目が見つかりません: ' + key);
}

function getExpenseSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EXPENSE);
  if (!sheet) throw new Error('経費台帳シートが見つかりません。');
  return sheet;
}

function getSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) sheet = ss.insertSheet(SHEET_SUMMARY);
  return sheet;
}

function getNoticeEmail() {
  return getConfig('通知先メール');
}

function getAccountantEmail() {
  return getConfig('会計事務所メール');
}

function getGeminiModel() {
  return getConfig('Geminiモデル');
}

function getInboxFolderId() {
  return getConfig('受信箱フォルダID');
}

function getReceiptFolderId() {
  return getConfig('領収書フォルダID');
}

function getCardUnprocessedFolderId() {
  return getConfig('カード明細・未処理フォルダID');
}

function getCardProcessedFolderId() {
  return getConfig('カード明細・処理済フォルダID');
}

function formatNow() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm');
}

function formatDateForSheet(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}
