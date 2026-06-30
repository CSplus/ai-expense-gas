/*************************************************
 * Sheet.gs
 * 経費台帳への書き込み処理
 *************************************************/

function appendInitialReceiptRow(sheet, params) {
  ensureExpenseInvoiceColumns(sheet);
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
  ensureExpenseInvoiceColumns(sheet);
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
    invoiceRegisteredName: invoiceInfo.officialName || invoiceInfo.registeredName || '',
    invoiceAddress: invoiceInfo.address || '',
    invoiceJudgement: invoiceInfo.invoiceJudgement || '',
    invoiceStatus: invoiceInfo.invoiceStatus || '',
    invoiceCheckedAt: invoiceInfo.checkedAt || '',
    invoiceRegistrationDate: invoiceInfo.registrationDate || '',
    invoiceExpireDate: invoiceInfo.expireDate || '',
    invoiceApiError: invoiceInfo.apiError || '',
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
  const rowValues = new Array(getExpenseLastColumn()).fill('');
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
    invoiceRegisteredName: COL.INVOICE_REGISTERED_NAME,
    invoiceAddress: COL.INVOICE_ADDRESS,
    invoiceJudgement: COL.INVOICE_JUDGEMENT,
    invoiceStatus: COL.INVOICE_STATUS,
    invoiceCheckedAt: COL.INVOICE_CHECKED_AT,
    invoiceRegistrationDate: COL.INVOICE_REGISTRATION_DATE,
    invoiceExpireDate: COL.INVOICE_EXPIRE_DATE,
    invoiceApiError: COL.INVOICE_API_ERROR,
    taxRate: COL.TAX_RATE,
    taxAmount: COL.TAX_AMOUNT,
    invoiceNote: COL.INVOICE_NOTE
  };

  const headerLabels = {
    invoiceNumber: ['登録番号', 'インボイス登録番号'],
    invoiceRegisteredName: 'インボイス正式名称',
    invoiceAddress: 'インボイス住所',
    invoiceJudgement: 'インボイス判定',
    invoiceStatus: 'インボイス登録状態',
    invoiceRegistrationDate: ['インボイス登録日', 'インボイス登録年月日'],
    invoiceExpireDate: ['インボイス失効日', 'インボイス失効年月日'],
    invoiceCheckedAt: ['インボイス確認日', 'インボイスAPI確認日時'],
    invoiceApiError: 'インボイスAPIエラー',
    taxRate: '税率',
    taxAmount: '消費税額',
    invoiceNote: 'インボイス備考'
  };

  if (headerLabels[name]) {
    const labels = Array.isArray(headerLabels[name]) ? headerLabels[name] : [headerLabels[name]];
    for (let i = 0; i < labels.length; i++) {
      const headerColumn = findExpenseHeaderColumn(labels[i]);
      if (headerColumn) return headerColumn;
    }
  }

  if (!columns[name]) {
    throw new Error('未定義の経費台帳列です: ' + name);
  }

  return columns[name];
}


function getExpenseLastColumn() {
  return Math.max(COL.INVOICE_API_ERROR, getExpenseSheet().getLastColumn());
}

function ensureExpenseInvoiceColumns(sheet) {
  sheet = sheet || getExpenseSheet();
  const requiredHeaderGroups = [
    ['登録番号', 'インボイス登録番号'],
    ['インボイス判定'],
    ['インボイス登録状態'],
    ['インボイス正式名称'],
    ['インボイス住所'],
    ['インボイス登録日', 'インボイス登録年月日'],
    ['インボイス失効日', 'インボイス失効年月日'],
    ['インボイス確認日', 'インボイスAPI確認日時'],
    ['税率'],
    ['消費税額'],
    ['インボイス備考'],
    ['インボイスAPIエラー']
  ];
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) {
    return String(v || '').trim();
  });

  requiredHeaderGroups.forEach(function(group) {
    const exists = group.some(function(header) {
      return headers.indexOf(header) !== -1;
    });
    if (!exists) {
      sheet.getRange(1, headers.length + 1).setValue(group[0]);
      headers.push(group[0]);
    }
  });
}

function findExpenseHeaderColumn(headerName) {
  const sheet = getExpenseSheet();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === headerName) return i + 1;
  }
  return null;
}
