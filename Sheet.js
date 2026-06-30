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

const EXPENSE_CANONICAL_HEADERS = {
  timestamp: 'タイムスタンプ',
  receiptUrl: '領収書画像アップロード',
  memo: '内容メモ',
  date: '取引日',
  vendor: '店舗名',
  amount: '金額',
  accountCode: '勘定科目コード',
  accountName: '勘定科目',
  status: '処理状態',
  fileId: 'ファイルID',
  error: 'エラー内容',
  vendorNormalized: '取引先正規名',
  paymentMethod: '支払方法',
  evidenceType: '証憑種別',
  sourceFile: '元ファイル名',
  confirm: '確認',
  inputCategory: '入力区分',
  duplicate: '重複判定',
  duplicateId: '重複候補ID',
  summaryTarget: '集計対象',
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

const EXPENSE_HEADER_ORDER = [
  'timestamp', 'receiptUrl', 'memo', 'date', 'vendor', 'amount',
  'accountCode', 'accountName', 'status', 'fileId', 'error', 'vendorNormalized',
  'paymentMethod', 'evidenceType', 'sourceFile', 'confirm', 'inputCategory',
  'duplicate', 'duplicateId', 'summaryTarget', 'invoiceNumber', 'invoiceJudgement',
  'invoiceStatus', 'invoiceRegisteredName', 'invoiceAddress', 'invoiceRegistrationDate',
  'invoiceExpireDate', 'invoiceCheckedAt', 'taxRate', 'taxAmount', 'invoiceNote',
  'invoiceApiError'
];

function getExpenseColumnByName(name) {
  const header = EXPENSE_CANONICAL_HEADERS[name];
  if (!header) {
    throw new Error('未定義の経費台帳列です: ' + name);
  }

  const headerColumn = findExpenseHeaderColumn(header);
  if (!headerColumn) {
    throw new Error('経費台帳に正規列がありません: ' + header);
  }
  return headerColumn;
}



function getExpenseColumnsByName(names) {
  const columns = {};
  names.forEach(function(name) {
    columns[name] = getExpenseColumnByName(name);
  });
  return columns;
}

function getExpenseLastColumn() {
  return Math.max(EXPENSE_HEADER_ORDER.length, getExpenseSheet().getLastColumn());
}

function ensureExpenseInvoiceColumns(sheet) {
  sheet = sheet || getExpenseSheet();
  const requiredHeaders = EXPENSE_HEADER_ORDER.map(function(name) {
    return EXPENSE_CANONICAL_HEADERS[name];
  });
  const lastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length, 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) {
    return String(v || '').trim();
  });

  requiredHeaders.forEach(function(header, index) {
    if (headers[index] === '') {
      sheet.getRange(1, index + 1).setValue(header);
      headers[index] = header;
    }
  });

  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(header);
      headers.push(header);
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
