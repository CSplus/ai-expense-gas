/*************************************************
 * Sheet.gs
 * 経費台帳への書き込み処理
 *************************************************/

function appendInitialReceiptRow(sheet, params) {
  const rowValues = new Array(COL.INVOICE_NOTE).fill('');

  rowValues[COL.TIMESTAMP - 1] = params.now;
  rowValues[COL.RECEIPT_URL - 1] = params.fileUrl;
  rowValues[COL.MEMO - 1] = params.memo;
  rowValues[COL.ACCOUNT_CODE - 1] = params.accountCode;
  rowValues[COL.ACCOUNT_NAME - 1] = params.accountName;
  rowValues[COL.STATUS - 1] = '受信済';
  rowValues[COL.FILE_ID - 1] = params.fileId;
  rowValues[COL.PAYMENT_METHOD - 1] = '現金';
  rowValues[COL.EVIDENCE_TYPE - 1] = '領収書画像';
  rowValues[COL.CONFIRM - 1] = '未確認';
  rowValues[COL.INPUT_CATEGORY - 1] = params.categoryInput;
  rowValues[COL.SUMMARY_TARGET - 1] = '対象';

  sheet.appendRow(rowValues);
  return sheet.getLastRow();
}

function updateReceiptAnalysisResult(sheet, row, params) {
  const result = params.result;
  const inputRule = params.inputRule;
  const invoiceInfo = params.invoiceInfo;
  const vendorOfficialName = params.vendorOfficialName;

  sheet.getRange(row, COL.DATE).setValue(result.date || '');
  sheet.getRange(row, COL.VENDOR).setValue(result.vendor || '');
  sheet.getRange(row, COL.AMOUNT).setValue(result.amount || '');
  sheet.getRange(row, COL.ACCOUNT_CODE).setValue(inputRule.accountCode);
  sheet.getRange(row, COL.ACCOUNT_NAME).setValue(inputRule.accountName);
  sheet.getRange(row, COL.STATUS).setValue('読取済');
  sheet.getRange(row, COL.VENDOR_NORMALIZED).setValue(vendorOfficialName || '');
  sheet.getRange(row, COL.PAYMENT_METHOD).setValue(result.paymentMethod || '現金');

  sheet.getRange(row, COL.INVOICE_NUMBER).setValue(invoiceInfo.registrationNumber || '');
  sheet.getRange(row, COL.INVOICE_JUDGEMENT).setValue(invoiceInfo.invoiceJudgement || '');
  sheet.getRange(row, COL.INVOICE_STATUS).setValue(invoiceInfo.invoiceStatus || '');
  sheet.getRange(row, COL.INVOICE_CHECKED_AT).setValue(invoiceInfo.checkedAt || '');
  sheet.getRange(row, COL.TAX_RATE).setValue(result.taxRate || '');
  sheet.getRange(row, COL.TAX_AMOUNT).setValue(result.taxAmount || '');
  sheet.getRange(row, COL.INVOICE_NOTE).setValue(invoiceInfo.note || result.invoiceNote || '');
}

function appendCardExpenseRow(sheet, date, vendor, amount, paymentMethod, fileName) {
  const rule = getAccountingRule(vendor);
  const rowValues = new Array(COL.INVOICE_NOTE).fill('');

  rowValues[COL.TIMESTAMP - 1] = new Date();
  rowValues[COL.DATE - 1] = date;
  rowValues[COL.VENDOR - 1] = vendor;
  rowValues[COL.AMOUNT - 1] = amount;
  rowValues[COL.ACCOUNT_CODE - 1] = rule.accountCode;
  rowValues[COL.ACCOUNT_NAME - 1] = rule.accountName;
  rowValues[COL.STATUS - 1] = 'カード明細取込済';
  rowValues[COL.VENDOR_NORMALIZED - 1] = rule.vendorName;
  rowValues[COL.PAYMENT_METHOD - 1] = paymentMethod;
  rowValues[COL.EVIDENCE_TYPE - 1] = 'カード明細';
  rowValues[COL.SOURCE_FILE - 1] = fileName;
  rowValues[COL.CONFIRM - 1] = '未確認';
  rowValues[COL.SUMMARY_TARGET - 1] = '対象';

  sheet.appendRow(rowValues);
}

/**
 * 税率・消費税額を整える
 * 領収書に税額記載がない場合は、税込金額から推定する
 */
function normalizeTaxInfo(result) {
  const amount = Number(result.amount || 0);

  let taxRate = result.taxRate || result.tax_rate || '';
  let taxAmount = result.taxAmount || result.tax_amount || '';
  let taxEstimated = result.taxEstimated === true || result.tax_estimated === true;
  let taxNote = result.taxNote || result.tax_note || '';

  // 表記ゆれ補正
  taxRate = String(taxRate || '').trim();

  if (taxRate === '１０％') taxRate = '10%';
  if (taxRate === '８％') taxRate = '8%';

  // Geminiが税額を返している場合
  if (taxAmount !== '' && taxAmount !== null && taxAmount !== undefined) {
    return {
      taxRate: taxRate || '不明',
      taxAmount: Number(taxAmount),
      taxNote: taxEstimated ? '消費税額はAI推定' : '消費税額は領収書記載'
    };
  }

  // 金額が無い場合は計算不可
  if (!amount) {
    return {
      taxRate: taxRate || '不明',
      taxAmount: '',
      taxNote: '税込金額が不明のため消費税額を計算できません'
    };
  }

  // 税率が10%なら税込金額から計算
  if (taxRate === '10%') {
    return {
      taxRate: '10%',
      taxAmount: Math.round(amount / 1.1 * 0.1),
      taxNote: '消費税額は税込金額からAI推定'
    };
  }

  // 税率が8%なら税込金額から計算
  if (taxRate === '8%') {
    return {
      taxRate: '8%',
      taxAmount: Math.round(amount / 1.08 * 0.08),
      taxNote: '消費税額は税込金額からAI推定'
    };
  }

  // 税率が不明でも、飲食・サービス系は10%推定
  const vendor = String(result.vendor || '').trim();
  const category = String(result.category || '').trim();

  if (
    category === '接待交際費' ||
    vendor.match(/居酒屋|食堂|レストラン|バル|カフェ|喫茶|駐車場|パーキング|ホテル|ガソリン|ENEOS|出光|コスモ/)
  ) {
    return {
      taxRate: '10%',
      taxAmount: Math.round(amount / 1.1 * 0.1),
      taxNote: '税率・消費税額は業種からAI推定'
    };
  }

  return {
    taxRate: '不明',
    taxAmount: '',
    taxNote: '税率を判定できませんでした'
  };
}
