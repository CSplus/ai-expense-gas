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
    invoiceAddress: invoiceInfo.address || invoiceInfo.invoiceAddress || '',
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

const EXPENSE_INVOICE_CANONICAL_HEADERS = {
  invoiceNumber: '登録番号',
  invoiceJudgement: 'インボイス判定',
  invoiceStatus: 'インボイス登録状態',
  invoiceRegisteredName: 'インボイス正式名称',
  invoiceAddress: 'インボイス住所',
  invoiceRegistrationDate: 'インボイス登録日',
  invoiceExpireDate: 'インボイス失効日',
  invoiceCheckedAt: 'インボイス確認日',
  taxRate: '税率',
  taxAmount: '消費税額',
  invoiceNote: 'インボイス備考',
  invoiceApiError: 'インボイスAPIエラー'
};

const EXPENSE_INVOICE_LEGACY_HEADER_ALIASES = {
  invoiceNumber: ['インボイス登録番号'],
  invoiceRegistrationDate: ['インボイス登録年月日'],
  invoiceExpireDate: ['インボイス失効年月日'],
  invoiceCheckedAt: ['インボイスAPI確認日', 'インボイスAPI確認日時']
};

const EXPENSE_INVOICE_MIGRATIONS = [
  { from: 'インボイス登録番号', to: '登録番号' },
  { from: 'インボイス登録年月日', to: 'インボイス登録日' },
  { from: 'インボイス失効年月日', to: 'インボイス失効日' },
  { from: 'インボイスAPI確認日', to: 'インボイス確認日' },
  { from: 'インボイスAPI確認日時', to: 'インボイス確認日' }
];

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
    invoiceAddress: null,
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

  if (EXPENSE_INVOICE_CANONICAL_HEADERS[name]) {
    const headerColumn = findExpenseHeaderColumn(EXPENSE_INVOICE_CANONICAL_HEADERS[name]);
    if (headerColumn) return headerColumn;
  }

  if (EXPENSE_INVOICE_CANONICAL_HEADERS[name]) {
    throw new Error('経費台帳に正規列がありません: ' + EXPENSE_INVOICE_CANONICAL_HEADERS[name]);
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
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) {
    return String(v || '').trim();
  });

  Object.keys(EXPENSE_INVOICE_CANONICAL_HEADERS).forEach(function(name) {
    const header = EXPENSE_INVOICE_CANONICAL_HEADERS[name];
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(header);
      headers.push(header);
    }
  });

  migrateExpenseInvoiceLegacyColumns(sheet);
}

function migrateExpenseInvoiceLegacyColumns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) {
    return String(v || '').trim();
  });

  EXPENSE_INVOICE_MIGRATIONS.forEach(function(migration) {
    const fromCol = headers.indexOf(migration.from) + 1;
    const toCol = headers.indexOf(migration.to) + 1;
    if (!fromCol || !toCol || fromCol === toCol) return;

    const fromValues = sheet.getRange(2, fromCol, lastRow - 1, 1).getValues();
    const toRange = sheet.getRange(2, toCol, lastRow - 1, 1);
    const toValues = toRange.getValues();
    let changed = false;

    for (let i = 0; i < toValues.length; i++) {
      if ((toValues[i][0] === '' || toValues[i][0] === null || toValues[i][0] === undefined)
          && fromValues[i][0] !== '' && fromValues[i][0] !== null && fromValues[i][0] !== undefined) {
        toValues[i][0] = fromValues[i][0];
        changed = true;
      }
    }

    if (changed) toRange.setValues(toValues);
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
