const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

function formatCacheDate(date) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}/${pad(date.getMonth() + 1)}/${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function daysAgo(days) {
  const date = new Date();
  date.setDate(date.getDate() - days);
  return formatCacheDate(date);
}

function createSheet(initialRows) {
  const rows = initialRows.map((row) => row.slice());

  function ensureCell(row, col) {
    while (rows.length < row) rows.push([]);
    while (rows[row - 1].length < col) rows[row - 1].push('');
  }

  return {
    rows,
    getLastColumn() {
      return rows.reduce((max, row) => Math.max(max, row.length), 0);
    },
    getLastRow() {
      return rows.length;
    },
    getRange(row, col, numRows = 1, numCols = 1) {
      return {
        getValues() {
          const values = [];
          for (let r = 0; r < numRows; r++) {
            const current = [];
            for (let c = 0; c < numCols; c++) {
              current.push((rows[row - 1 + r] || [])[col - 1 + c] ?? '');
            }
            values.push(current);
          }
          return values;
        },
        setValue(value) {
          ensureCell(row, col);
          rows[row - 1][col - 1] = value;
        },
        setValues(values) {
          values.forEach((valueRow, r) => {
            valueRow.forEach((value, c) => {
              ensureCell(row + r, col + c);
              rows[row - 1 + r][col - 1 + c] = value;
            });
          });
        }
      };
    },
    appendRow(row) {
      rows.push(row.slice());
    }
  };
}

function createContext({ rows, appId = 'dummy-app-id', fetchImpl }) {
  const sheet = createSheet(rows);
  let fetchCalls = 0;
  const ctx = {
    TIMEZONE: 'Asia/Tokyo',
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === 'InvoiceCache' ? sheet : null;
          },
          insertSheet() {
            return sheet;
          }
        };
      }
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty() {
            return appId;
          }
        };
      }
    },
    UrlFetchApp: {
      fetch(url, options) {
        fetchCalls += 1;
        return fetchImpl(url, options);
      }
    },
    Utilities: {
      formatDate(date) {
        return formatCacheDate(date);
      }
    },
    console
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync('InvoiceApi.gs', 'utf8'), ctx);
  return { ctx, sheet, getFetchCalls: () => fetchCalls };
}

function successResponse() {
  return {
    getResponseCode: () => 200,
    getContentText: () => JSON.stringify({
      announcement: [{
        registratedNumber: 'T1234567890123',
        name: '株式会社API結果',
        address: '東京都API区',
        registrationDate: '2023-10-01',
        expireDate: '',
        process: '公表'
      }]
    })
  };
}

function errorResponse() {
  return {
    getResponseCode: () => 500,
    getContentText: () => 'server error'
  };
}

const headers = ['インボイス登録番号', '正式名称', '住所', '登録状態', '登録日', '失効日', '最終確認日', 'Version'];

(function freshVersionOneCacheSkipsApi() {
  const { ctx, getFetchCalls } = createContext({
    rows: [headers, ['T1234567890123', '株式会社キャッシュ', '東京都', '公表', '2023-10-01', '', daysAgo(3), '1']],
    fetchImpl: () => { throw new Error('API should not be called'); }
  });

  const result = ctx.fetchInvoiceBusinessInfo('1234567890123');
  assert.strictEqual(getFetchCalls(), 0);
  assert.strictEqual(result.fromCache, true);
  assert.strictEqual(result.cacheStale, false);
  assert.strictEqual(result.registeredName, '株式会社キャッシュ');
})();

(function blankVersionCallsApiAndUpdatesVersionOne() {
  const { ctx, sheet, getFetchCalls } = createContext({
    rows: [headers, ['T1234567890123', '株式会社旧', '東京都', '公表', '2023-10-01', '', daysAgo(3), '']],
    fetchImpl: () => successResponse()
  });

  const result = ctx.fetchInvoiceBusinessInfo('T1234567890123');
  assert.strictEqual(getFetchCalls(), 1);
  assert.strictEqual(result.registeredName, '株式会社API結果');
  assert.strictEqual(sheet.rows[1][7], '1');
})();

(function differentVersionCallsApi() {
  const { ctx, getFetchCalls } = createContext({
    rows: [headers, ['T1234567890123', '株式会社旧', '東京都', '公表', '2023-10-01', '', daysAgo(3), '2']],
    fetchImpl: () => successResponse()
  });

  const result = ctx.fetchInvoiceBusinessInfo('T1234567890123');
  assert.strictEqual(getFetchCalls(), 1);
  assert.strictEqual(result.registeredName, '株式会社API結果');
})();

(function apiSuccessAppendsVersionOne() {
  const { ctx, sheet } = createContext({
    rows: [headers],
    fetchImpl: () => successResponse()
  });

  ctx.fetchInvoiceBusinessInfo('T1234567890123');
  assert.strictEqual(sheet.rows.length, 2);
  assert.strictEqual(sheet.rows[1][0], 'T1234567890123');
  assert.strictEqual(sheet.rows[1][7], '1');
})();

(function apiFailureUsesOldCacheAsStale() {
  const { ctx, getFetchCalls } = createContext({
    rows: [headers, ['T1234567890123', '株式会社古い', '東京都', '公表', '2023-10-01', '', daysAgo(45), '1']],
    fetchImpl: () => errorResponse()
  });

  const result = ctx.fetchInvoiceBusinessInfo('T1234567890123');
  assert.strictEqual(getFetchCalls(), 1);
  assert.strictEqual(result.fromCache, true);
  assert.strictEqual(result.cacheStale, true);
  assert.strictEqual(result.registeredName, '株式会社古い');
  assert.match(result.apiErrorMessage, /HTTPエラー: 500/);
})();

(function enrichDoesNotThrowOnApiFailureWithOldCache() {
  const { ctx } = createContext({
    rows: [headers, ['T1234567890123', '株式会社古い', '東京都', '公表', '2023-10-01', '', daysAgo(45), '1']],
    fetchImpl: () => errorResponse()
  });

  const expense = ctx.enrichWithInvoiceInfo({ invoiceNumber: 'T1234567890123', vendor: 'OCR店舗名' });
  assert.strictEqual(expense.vendor, 'OCR店舗名');
  assert.strictEqual(expense.invoiceRegisteredName, '株式会社古い');
  assert.strictEqual(expense.invoiceStatus, '有効');
  assert.match(expense.invoiceApiError, /HTTPエラー: 500/);
})();

function createExpenseContext(rows) {
  const sheet = createSheet(rows);
  const ctx = {
    TIMEZONE: 'Asia/Tokyo',
    SHEET_EXPENSE: '経費台帳',
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return {
          getSheetByName(name) {
            return name === '経費台帳' ? sheet : null;
          }
        };
      }
    },
    console
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync('Config.js', 'utf8'), ctx);
  vm.runInContext(fs.readFileSync('Sheet.js', 'utf8'), ctx);
  return { ctx, sheet };
}

(function expenseColumnsPreferCanonicalAndMigrateLegacyValues() {
  const baseHeaders = [
    'タイムスタンプ', '領収書画像アップロード', '内容のメモ', '取引日', '店舗名', '金額',
    '勘定科目コード', '勘定科目名', '処理状態', 'ファイルID', 'エラー内容', '取引先正規名',
    '支払方法', '証憑種別', '元ファイル名', '確認', '入力区分', '重複判定', '重複候補ID', '集計対象',
    '登録番号', 'インボイス判定', 'インボイス登録状態', 'インボイス確認日', '税率', '消費税額',
    'インボイス備考', 'インボイス正式名称', 'インボイス登録日', 'インボイス失効日', 'インボイスAPIエラー',
    'インボイス登録番号', 'インボイス登録年月日', 'インボイス失効年月日', 'インボイスAPI確認日'
  ];
  const row = new Array(baseHeaders.length).fill('');
  row[31] = 'T9999999999999';
  row[32] = '2023/10/01';
  row[33] = '2024/10/01';
  row[34] = '2026/06/30 12:00';
  const { ctx, sheet } = createExpenseContext([baseHeaders, row]);

  ctx.ensureExpenseInvoiceColumns(sheet);
  assert.strictEqual(sheet.rows[1][20], 'T9999999999999');
  assert.strictEqual(sheet.rows[1][28], '2023/10/01');
  assert.strictEqual(sheet.rows[1][29], '2024/10/01');
  assert.strictEqual(sheet.rows[1][23], '2026/06/30 12:00');

  ctx.setRowValues(sheet, 2, {
    vendor: 'OCR店舗名',
    vendorNormalized: '取引先名',
    invoiceRegisteredName: 'API正式名称',
    invoiceAddress: '東京都API区',
    invoiceNumber: 'T1234567890123',
    invoiceStatus: '有効',
    taxRate: '10%',
    taxAmount: 100,
    duplicate: '重複なし'
  });

  assert.strictEqual(sheet.rows[1][4], 'OCR店舗名');
  assert.strictEqual(sheet.rows[1][11], '取引先名');
  assert.strictEqual(sheet.rows[1][27], 'API正式名称');
  assert.notStrictEqual(sheet.rows[1][27], 'OCR店舗名');
  assert.strictEqual(sheet.rows[1][35], '東京都API区');
  assert.strictEqual(sheet.rows[1][20], 'T1234567890123');
  assert.strictEqual(sheet.rows[1][31], 'T9999999999999');
  assert.strictEqual(sheet.rows[1][22], '有効');
  assert.strictEqual(sheet.rows[1][24], '10%');
  assert.strictEqual(sheet.rows[1][25], 100);
  assert.strictEqual(sheet.rows[1][17], '重複なし');
})();

(function enrichMapsCacheAndApiNamesAndDisplayStatus() {
  const cacheCtx = createContext({
    rows: [headers, ['T1234567890123', '株式会社キャッシュ正式名', '東京都', '01', '2023-10-01', '', daysAgo(1), '1']],
    fetchImpl: () => { throw new Error('API should not be called'); }
  }).ctx;
  const cachedExpense = cacheCtx.enrichWithInvoiceInfo({ invoiceNumber: 'T1234567890123', vendor: 'OCR店舗名' });
  assert.strictEqual(cachedExpense.invoiceRegisteredName, '株式会社キャッシュ正式名');
  assert.strictEqual(cachedExpense.invoiceAddress, '東京都');
  assert.strictEqual(cachedExpense.invoiceRegistrationDate, '2023-10-01');
  assert.strictEqual(cachedExpense.invoiceExpireDate, '');
  assert.strictEqual(cachedExpense.invoiceStatus, '有効');
  assert.notStrictEqual(cachedExpense.invoiceRegisteredName, cachedExpense.vendor);

  const apiCtx = createContext({ rows: [headers], fetchImpl: () => successResponse() }).ctx;
  const first = apiCtx.enrichWithInvoiceInfo({ invoiceNumber: 'T1234567890123', vendor: 'OCR店舗名' });
  const second = apiCtx.enrichWithInvoiceInfo({ invoiceNumber: 'T1234567890123', vendor: 'OCR店舗名' });
  assert.strictEqual(first.invoiceRegisteredName, '株式会社API結果');
  assert.strictEqual(first.invoiceAddress, '東京都API区');
  assert.strictEqual(first.invoiceRegistrationDate, '2023-10-01');
  assert.strictEqual(first.invoiceExpireDate, '');
  assert.strictEqual(second.invoiceRegisteredName, '株式会社API結果');
  assert.strictEqual(second.invoiceAddress, '東京都API区');
  assert.strictEqual(second.invoiceStatus, '有効');
  assert.notStrictEqual(second.invoiceRegisteredName, second.vendor);
})();

console.log('InvoiceApi cache and expense sheet mapping tests passed');
