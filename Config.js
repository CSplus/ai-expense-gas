/*************************************************
 * Config.gs
 * 共通設定・シート名・経費台帳v1.0列定義・設定取得
 *************************************************/

const TIMEZONE = 'Asia/Tokyo';

const SHEET_EXPENSE = '経費台帳';
const SHEET_SUMMARY = '月次集計';
const SHEET_CONFIG  = 'システム設定';
const SHEET_RULE    = 'カード明細仕訳ルール';
const SHEET_TAX_RATE = '税率マスタ';
const SHEET_ACCOUNTING_EXPORT_HISTORY = '会計出力履歴';

// 経費台帳 v1.0 の正式列番号（1始まり）。読み書き時はヘッダー名から列位置を取得する。
const COL = {
  TIMESTAMP: 1,        // A タイムスタンプ
  RECEIPT_URL: 2,      // B 領収書画像アップロード
  MEMO: 3,             // C 内容メモ
  DATE: 4,             // D 取引日
  VENDOR: 5,           // E 店舗名
  AMOUNT: 6,           // F 金額
  ACCOUNT_CODE: 7,     // G 勘定科目コード
  ACCOUNT_NAME: 8,     // H 勘定科目
  STATUS: 9,           // I 処理状態
  FILE_ID: 10,         // J ファイルID
  ERROR: 11,           // K エラー内容
  PAYMENT_METHOD: 12,  // L 支払方法
  EVIDENCE_TYPE: 13,   // M 証憑種別
  SOURCE_FILE: 14,     // N 元ファイル名
  CONFIRM: 15,         // O 確認
  INPUT_CATEGORY: 16,  // P 入力区分
  DUPLICATE: 17,       // Q 重複判定
  DUPLICATE_ID: 18,    // R 重複候補ID
  SUMMARY_TARGET: 19,  // S 集計対象
  INVOICE_NUMBER: 20,  // T 登録番号
  INVOICE_JUDGEMENT: 21,// U インボイス判定
  INVOICE_STATUS: 22,  // V インボイス登録状態
  INVOICE_REGISTERED_NAME: 23, // W インボイス正式名称
  INVOICE_ADDRESS: 24, // X インボイス住所
  INVOICE_REGISTRATION_DATE: 25, // Y インボイス登録日
  INVOICE_EXPIRE_DATE: 26, // Z インボイス失効日
  INVOICE_CHECKED_AT: 27,// AA インボイス確認日
  TAX_RATE: 28,        // AB 税率
  TAX_AMOUNT: 29,      // AC 消費税額
  INVOICE_NOTE: 30,    // AD インボイス備考
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
