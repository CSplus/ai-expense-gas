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
  const taxInfo = params.taxInfo || {};
  const invoiceNote = params.invoiceNote || invoiceInfo.note || result.invoiceNote || '';

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
  sheet.getRange(row, COL.TAX_RATE).setValue(taxInfo.taxRate || '');
  sheet.getRange(row, COL.TAX_AMOUNT).setValue(taxInfo.taxAmount || '');
  sheet.getRange(row, COL.INVOICE_NOTE).setValue(invoiceNote);
}

function markReceiptAnalysisError(sheet, row, message) {
  sheet.getRange(row, COL.STATUS).setValue('エラー：読取失敗');
  sheet.getRange(row, COL.ERROR).setValue(message);
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
