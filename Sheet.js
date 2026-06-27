/*************************************************
 * Sheet.gs
 * 経費台帳への書き込み処理
 *************************************************/

function appendInitialReceiptRow(sheet, params) {
  const rowValues = createExpenseRowValues({
    timestamp: params.now,
    receiptUrl: params.fileUrl,
    memo: params.memo,
    accountCode: params.accountCode,
    accountName: params.accountName,
    status: '受信済',
    fileId: params.fileId,
    paymentMethod: '現金',
    evidenceType: '領収書画像',
    confirm: '未確認',
    inputCategory: params.categoryInput,
    summaryTarget: '対象'
  });

  sheet.appendRow(rowValues);
  return sheet.getLastRow();
}

function updateReceiptAnalysisResult(sheet, row, params) {
  const result = params.result;
  const inputRule = params.inputRule;
  const invoiceInfo = params.invoiceInfo;
  const vendorOfficialName = params.vendorOfficialName;
  const taxInfo = params.taxInfo || {};
  const invoiceNote = params.invoiceNote || invoiceInfo.note || result.invoiceNote || '';

  setRowValues(sheet, row, {
    date: result.date || '',
    vendor: result.vendor || '',
    amount: result.amount || '',
    accountCode: inputRule.accountCode,
    accountName: inputRule.accountName,
    status: '読取済',
    vendorNormalized: vendorOfficialName || '',
    paymentMethod: result.paymentMethod || '現金',
    invoiceNumber: invoiceInfo.registrationNumber || '',
    invoiceJudgement: invoiceInfo.invoiceJudgement || '',
    invoiceStatus: invoiceInfo.invoiceStatus || '',
    invoiceCheckedAt: invoiceInfo.checkedAt || '',
    taxRate: taxInfo.taxRate || '',
    taxAmount: taxInfo.taxAmount || '',
    invoiceNote: invoiceNote
  });
}

function markReceiptAnalysisError(sheet, row, message) {
  setRowValues(sheet, row, {
    status: 'エラー：読取失敗',
    error: message
  });
}

function appendCardExpenseRow(sheet, date, vendor, amount, paymentMethod, fileName) {
  const rule = getAccountingRule(vendor);
  const rowValues = createExpenseRowValues({
    timestamp: new Date(),
    date: date,
    vendor: vendor,
    amount: amount,
    accountCode: rule.accountCode,
    accountName: rule.accountName,
    status: 'カード明細取込済',
    vendorNormalized: rule.vendorName,
    paymentMethod: paymentMethod,
    evidenceType: 'カード明細',
    sourceFile: fileName,
    confirm: '未確認',
    summaryTarget: '対象'
  });

  sheet.appendRow(rowValues);
}


function createExpenseRowValues(valuesByName) {
  const rowValues = new Array(COL.INVOICE_NOTE).fill('');
  applyNamedValuesToRow(rowValues, valuesByName);
  return rowValues;
}

function setRowValues(sheet, row, valuesByName) {
  Object.keys(valuesByName).forEach(function(name) {
    sheet.getRange(row, getExpenseColumnByName(name)).setValue(valuesByName[name]);
  });
}

function applyNamedValuesToRow(rowValues, valuesByName) {
  Object.keys(valuesByName).forEach(function(name) {
    rowValues[getExpenseColumnByName(name) - 1] = valuesByName[name];
  });
}

function getExpenseColumnByName(name) {
  const columns = {
    timestamp: COL.TIMESTAMP,
    receiptUrl: COL.RECEIPT_URL,
    memo: COL.MEMO,
    date: COL.DATE,
    vendor: COL.VENDOR,
    amount: COL.AMOUNT,
    accountCode: COL.ACCOUNT_CODE,
    accountName: COL.ACCOUNT_NAME,
    status: COL.STATUS,
    fileId: COL.FILE_ID,
    error: COL.ERROR,
    vendorNormalized: COL.VENDOR_NORMALIZED,
    paymentMethod: COL.PAYMENT_METHOD,
    evidenceType: COL.EVIDENCE_TYPE,
    sourceFile: COL.SOURCE_FILE,
    confirm: COL.CONFIRM,
    inputCategory: COL.INPUT_CATEGORY,
    duplicate: COL.DUPLICATE,
    duplicateId: COL.DUPLICATE_ID,
    summaryTarget: COL.SUMMARY_TARGET,
    invoiceNumber: COL.INVOICE_NUMBER,
    invoiceJudgement: COL.INVOICE_JUDGEMENT,
    invoiceStatus: COL.INVOICE_STATUS,
    invoiceCheckedAt: COL.INVOICE_CHECKED_AT,
    taxRate: COL.TAX_RATE,
    taxAmount: COL.TAX_AMOUNT,
    invoiceNote: COL.INVOICE_NOTE
  };

  if (!columns[name]) {
    throw new Error('未定義の経費台帳列です: ' + name);
  }

  return columns[name];
}
