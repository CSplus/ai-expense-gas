const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

function createSheet(headers) {
  const rows = [headers.slice(), []];
  function ensureCell(row, col) {
    while (rows.length < row) rows.push([]);
    while (rows[row - 1].length < col) rows[row - 1].push('');
  }
  return {
    rows,
    getLastColumn() { return rows[0].length; },
    getLastRow() { return rows.length; },
    getRange(row, col, numRows = 1, numCols = 1) {
      return {
        getValues() {
          return Array.from({ length: numRows }, (_, r) => Array.from({ length: numCols }, (_, c) => (rows[row - 1 + r] || [])[col - 1 + c] ?? ''));
        },
        setValue(value) { ensureCell(row, col); rows[row - 1][col - 1] = value; },
        setValues(values) { values.forEach((valueRow, r) => valueRow.forEach((value, c) => { ensureCell(row + r, col + c); rows[row - 1 + r][col - 1 + c] = value; })); }
      };
    },
    appendRow(row) { rows.push(row.slice()); }
  };
}

const existingHeaders = [
  'タイムスタンプ', '領収書画像アップロード', '内容のメモ', '取引日', '店舗名', '金額',
  '勘定科目コード', '勘定科目名', '処理状態', 'ファイルID', 'エラー内容', '取引先正規名',
  '支払方法', '証憑種別', '元ファイル名', '確認', '入力区分', '重複判定', '重複候補ID', '集計対象',
  '登録番号', 'インボイス判定', 'インボイス登録状態', 'インボイス確認日', '税率', '消費税額', 'インボイス備考'
];

const sheet = createSheet(existingHeaders);
const ctx = {
  COL: {
    TIMESTAMP: 1, RECEIPT_URL: 2, MEMO: 3, DATE: 4, VENDOR: 5, AMOUNT: 6,
    ACCOUNT_CODE: 7, ACCOUNT_NAME: 8, STATUS: 9, FILE_ID: 10, ERROR: 11,
    VENDOR_NORMALIZED: 12, PAYMENT_METHOD: 13, EVIDENCE_TYPE: 14, SOURCE_FILE: 15,
    CONFIRM: 16, INPUT_CATEGORY: 17, DUPLICATE: 18, DUPLICATE_ID: 19, SUMMARY_TARGET: 20,
    INVOICE_NUMBER: 21, INVOICE_JUDGEMENT: 22, INVOICE_STATUS: 23, INVOICE_CHECKED_AT: 24,
    TAX_RATE: 25, TAX_AMOUNT: 26, INVOICE_NOTE: 27, INVOICE_REGISTERED_NAME: 28,
    INVOICE_ADDRESS: 29, INVOICE_REGISTRATION_DATE: 30, INVOICE_EXPIRE_DATE: 31, INVOICE_API_ERROR: 32
  },
  getExpenseSheet() { return sheet; },
  console
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync('Sheet.js', 'utf8'), ctx);

ctx.ensureExpenseInvoiceColumns(sheet);
const once = sheet.rows[0].slice();
ctx.ensureExpenseInvoiceColumns(sheet);
assert.deepStrictEqual(sheet.rows[0], once, 'ensureExpenseInvoiceColumns should not duplicate existing columns');
['インボイス正式名称', 'インボイス住所', 'インボイス登録日', 'インボイス失効日'].forEach((header) => {
  assert.ok(sheet.rows[0].includes(header), `${header} should be added`);
});

ctx.updateReceiptAnalysisResult(sheet, 2, {
  result: { date: '2026-06-30', vendor: 'セブンイレブン津駅前店', amount: 1100, paymentMethod: '現金' },
  inputRule: { accountCode: 6231, accountName: '雑費' },
  invoiceInfo: {
    registrationNumber: 'T1234567890123',
    invoiceJudgement: 'API確認済み',
    invoiceStatus: '有効',
    officialName: '株式会社セブン‐イレブン・ジャパン',
    address: '東京都千代田区',
    registrationDate: '2023-10-01',
    expireDate: '',
    checkedAt: '2026/06/30 12:00',
    apiError: ''
  },
  vendorOfficialName: 'セブンイレブン津駅前店',
  taxInfo: { taxRate: '10%', taxAmount: 100 },
  invoiceNote: 'API確認済み'
});

function value(header) {
  return sheet.rows[1][sheet.rows[0].indexOf(header)];
}
assert.strictEqual(value('店舗名'), 'セブンイレブン津駅前店');
assert.strictEqual(value('インボイス正式名称'), '株式会社セブン‐イレブン・ジャパン');
assert.strictEqual(value('インボイス住所'), '東京都千代田区');
assert.strictEqual(value('インボイス登録日'), '2023-10-01');
assert.strictEqual(value('インボイス登録状態'), '有効');
assert.strictEqual(value('税率'), '10%');
assert.strictEqual(value('消費税額'), 100);

console.log('Sheet invoice column tests passed');
