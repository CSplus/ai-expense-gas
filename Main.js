/*************************************************
 * Main.gs
 * Webアプリ入口・領収書アップロード処理
 *************************************************/

/**
 * Webアプリ表示
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('CS+ 経費領収書アップロード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Webアプリから領収書画像を受け取り、Drive保存・Gemini解析・経費台帳登録を行う
 */
function uploadReceiptFromWebApp(data) {
  const sheet = getExpenseSheet();
  const folder = DriveApp.getFolderById(getReceiptFolderId());

  const categoryInput = data.categoryInput;
  const memo = data.memo || '';
  const base64 = data.imageBase64;
  const mimeType = data.mimeType || 'image/jpeg';

  if (!categoryInput) throw new Error('経費区分が選択されていません。');
  if (!base64) throw new Error('領収書画像がありません。');

  const bytes = Utilities.base64Decode(base64);
  const now = new Date();
  const fileName =
    Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd_HHmmss') + '_receipt.jpg';

  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);

  const fileId = file.getId();
  const fileUrl = file.getUrl();

  const inputRule = getAccountingRuleFromInput(categoryInput);

  // まず受信済みとして1行追加
  sheet.appendRow([
    now,                    // A タイムスタンプ
    fileUrl,                // B 領収書画像アップロード
    memo,                   // C 内容のメモ
    '',                     // D 取引日
    '',                     // E 店舗名
    '',                     // F 金額
    inputRule.accountCode,  // G 勘定科目コード
    inputRule.accountName,  // H 勘定科目名
    '受信済',               // I 処理状態
    fileId,                 // J ファイルID
    '',                     // K エラー内容
    '',                     // L 取引先正規名
    '現金',                 // M 支払方法
    '領収書画像',           // N 証憑種別
    '',                     // O 元ファイル名
    '未確認',               // P 確認
    categoryInput,          // Q 入力区分
    '',                     // R 重複判定
    '',                     // S 重複候補ID
    '対象',                 // T 集計対象
    '',                     // U 登録番号
    '',                     // V インボイス判定
    '',                     // W インボイス登録状態
    '',                     // X インボイス確認日
    '',                     // Y 税率
    '',                     // Z 消費税額
    ''                      // AA インボイス備考
  ]);

  const row = sheet.getLastRow();

  try {
    // Geminiで領収書を解析
    const result = analyzeReceiptWithGemini(fileId);

    Logger.log(JSON.stringify(result, null, 2));

    // 仕訳ルール取得
    const rule = getAccountingRule(result.vendor || '');

    // API未使用版インボイス情報
    const invoiceInfo = getInvoiceInfo(
      result.invoiceNumber || result.invoice_number || '',
      result.invoiceJudgement || result.invoice_judgement || ''
    );

    // 取引先正規名は、現時点では仕訳ルール優先
    // 将来API利用時は invoiceInfo.officialName を優先する
    const vendorOfficialName =
      invoiceInfo.officialName ||
      rule.vendorName ||
      result.vendor ||
      '';

    // 税率・消費税額の補完
    const taxInfo = normalizeTaxInfo(result);

    // 基本情報を書き込み
    sheet.getRange(row, COL.DATE).setValue(result.date || '');
    sheet.getRange(row, COL.VENDOR).setValue(result.vendor || '');
    sheet.getRange(row, COL.AMOUNT).setValue(result.amount || '');
    sheet.getRange(row, COL.ACCOUNT_CODE).setValue(inputRule.accountCode);
    sheet.getRange(row, COL.ACCOUNT_NAME).setValue(inputRule.accountName);
    sheet.getRange(row, COL.STATUS).setValue('読取済');
    sheet.getRange(row, COL.VENDOR_NORMALIZED).setValue(vendorOfficialName);
    sheet.getRange(row, COL.PAYMENT_METHOD).setValue(result.paymentMethod || '現金');

    // インボイス情報を書き込み
    sheet.getRange(row, COL.INVOICE_NUMBER).setValue(invoiceInfo.registrationNumber || '');
    sheet.getRange(row, COL.INVOICE_JUDGEMENT).setValue(invoiceInfo.invoiceJudgement || '');
    sheet.getRange(row, COL.INVOICE_STATUS).setValue(invoiceInfo.invoiceStatus || '');
    sheet.getRange(row, COL.INVOICE_CHECKED_AT).setValue(invoiceInfo.checkedAt || '');

    // 税率・消費税額を書き込み
    sheet.getRange(row, COL.TAX_RATE).setValue(taxInfo.taxRate || '');
    sheet.getRange(row, COL.TAX_AMOUNT).setValue(taxInfo.taxAmount || '');

    // 備考
    const invoiceNote = [
      invoiceInfo.note || '',
      taxInfo.taxNote || ''
    ]
      .filter(function(v) { return v; })
      .join(' / ');

    sheet.getRange(row, COL.INVOICE_NOTE).setValue(invoiceNote);

    return {
      success: true,
      date: result.date || '',
      vendor: result.vendor || '',
      amount: result.amount || '',
      category: inputRule.accountName,
      invoiceNumber: invoiceInfo.registrationNumber || '',
      invoiceJudgement: invoiceInfo.invoiceJudgement || '',
      taxRate: taxInfo.taxRate || '',
      taxAmount: taxInfo.taxAmount || '',
      officialName: vendorOfficialName
    };

  } catch (error) {
    sheet.getRange(row, COL.STATUS).setValue('エラー：読取失敗');
    sheet.getRange(row, COL.ERROR).setValue(error.message);

    return {
      success: false,
      message: error.message
    };
  }
}