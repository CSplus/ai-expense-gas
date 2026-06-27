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

  const row = appendInitialReceiptRow(sheet, {
    now: now,
    fileUrl: fileUrl,
    memo: memo,
    accountCode: inputRule.accountCode,
    accountName: inputRule.accountName,
    fileId: fileId,
    categoryInput: categoryInput
  });

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

    const invoiceNote = [
      invoiceInfo.note || '',
      taxInfo.taxNote || ''
    ]
      .filter(function(v) { return v; })
      .join(' / ');

    updateReceiptAnalysisResult(sheet, row, {
      result: result,
      inputRule: inputRule,
      invoiceInfo: invoiceInfo,
      vendorOfficialName: vendorOfficialName,
      taxInfo: taxInfo,
      invoiceNote: invoiceNote
    });

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
    markReceiptAnalysisError(sheet, row, error.message);

    return {
      success: false,
      message: error.message
    };
  }
}


/**
 * ユーザーが確認・修正した内容で、送信時にだけDrive保存・経費台帳登録を行う
 */
function submitConfirmedReceiptFromWebApp(data) {
  const sheet = getExpenseSheet();
  const folder = DriveApp.getFolderById(getReceiptFolderId());

  const categoryInput = data.categoryInput;
  const memo = data.memo || '';
  const base64 = data.imageBase64;
  const mimeType = data.mimeType || 'image/jpeg';
  const confirmed = data.confirmed || {};

  if (!categoryInput) throw new Error('経費区分が選択されていません。');
  if (!base64) throw new Error('領収書画像がありません。');
  if (!confirmed.date) throw new Error('取引日を入力してください。');
  if (!confirmed.vendor) throw new Error('店舗名・取引先を入力してください。');
  if (!confirmed.amount) throw new Error('金額を入力してください。');

  const bytes = Utilities.base64Decode(base64);
  const now = new Date();
  const fileName =
    Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd_HHmmss') + '_receipt.jpg';

  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);

  const fileId = file.getId();
  const fileUrl = file.getUrl();
  const inputRule = getAccountingRuleFromInput(categoryInput);

  const row = appendInitialReceiptRow(sheet, {
    now: now,
    fileUrl: fileUrl,
    memo: memo,
    accountCode: inputRule.accountCode,
    accountName: inputRule.accountName,
    fileId: fileId,
    categoryInput: categoryInput
  });

  const result = {
    date: confirmed.date || '',
    vendor: confirmed.vendor || '',
    amount: confirmed.amount || '',
    category: confirmed.category || '',
    paymentMethod: confirmed.paymentMethod || '現金',
    invoiceNumber: confirmed.invoiceNumber || '',
    invoiceJudgement: confirmed.invoiceJudgement || '',
    taxRate: confirmed.taxRate || '',
    taxAmount: confirmed.taxAmount || '',
    taxNote: confirmed.taxNote || '',
    memo: memo
  };

  const invoiceInfo = getInvoiceInfo(
    result.invoiceNumber,
    result.invoiceJudgement
  );
  invoiceInfo.invoiceJudgement = result.invoiceJudgement || invoiceInfo.invoiceJudgement;
  invoiceInfo.note = invoiceInfo.registrationNumber
    ? 'ユーザー確認済（API未使用・正式名称未確認）'
    : 'ユーザー確認済';

  const rule = getAccountingRule(result.vendor || '');
  const vendorOfficialName =
    invoiceInfo.officialName ||
    rule.vendorName ||
    result.vendor ||
    '';
  const taxInfo = {
    taxRate: normalizeTaxRateText(result.taxRate),
    taxAmount: result.taxAmount === '' ? '' : Number(result.taxAmount),
    taxNote: result.taxNote || 'ユーザー確認済'
  };
  const invoiceNote = [
    'ユーザー確認・修正後の内容で登録',
    invoiceInfo.note || '',
    taxInfo.taxNote || ''
  ]
    .filter(function(v) { return v; })
    .join(' / ');

  updateReceiptAnalysisResult(sheet, row, {
    result: result,
    inputRule: inputRule,
    invoiceInfo: invoiceInfo,
    vendorOfficialName: vendorOfficialName,
    taxInfo: taxInfo,
    invoiceNote: invoiceNote
  });

  return {
    success: true,
    message: '登録が完了しました。',
    row: row,
    fileUrl: fileUrl
  };
}

/**
 * Webアプリから領収書画像を受け取り、Drive/Sheetsへ保存せずGemini解析だけを行う
 */
function analyzeReceiptFromWebApp(data) {
  const categoryInput = data.categoryInput;
  const base64 = data.imageBase64;
  const mimeType = data.mimeType || 'image/jpeg';

  if (!categoryInput) throw new Error('経費区分が選択されていません。');
  if (!base64) throw new Error('領収書画像がありません。');

  const result = analyzeReceiptImageWithGemini(base64, mimeType);
  const taxInfo = normalizeTaxInfo(result);

  return {
    success: true,
    date: result.date || '',
    vendor: result.vendor || '',
    amount: result.amount || '',
    inputCategory: categoryInput,
    category: result.category || '',
    paymentMethod: result.paymentMethod || '現金',
    invoiceNumber: result.invoiceNumber || '',
    invoiceJudgement: result.invoiceJudgement || '',
    taxRate: taxInfo.taxRate || result.taxRate || '',
    taxAmount: taxInfo.taxAmount || result.taxAmount || '',
    taxEstimated: taxInfo.taxEstimated || result.taxEstimated || false,
    taxNote: taxInfo.taxNote || result.taxNote || '',
    invoiceNote: result.invoiceNote || ''
  };
}
