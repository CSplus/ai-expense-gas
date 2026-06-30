/*************************************************
 * InvoiceApi.gs
 * 国税庁「適格請求書発行事業者公表サイト Web-API」連携
 *************************************************/

const NTA_INVOICE_APP_ID_PROPERTY = 'NTA_INVOICE_APP_ID';
const NTA_INVOICE_API_URL = 'https://web-api.invoice-kohyo.nta.go.jp/1/num';
const SHEET_INVOICE_CACHE = 'InvoiceCache';
const INVOICE_CACHE_MAX_AGE_DAYS = 30;
const INVOICE_CACHE_HEADERS = [
  'インボイス登録番号',
  '正式名称',
  '住所',
  '登録状態',
  '登録日',
  '失効日',
  '最終確認日',
  'Version'
];

/**
 * 登録番号を T + 数字13桁に正規化する。
 */
function normalizeInvoiceNumber(value) {
  if (value === null || value === undefined || value === '') return null;

  const text = String(value)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .replace(/[‐‑‒–—―ー－ｰ−]/g, '')
    .replace(/\s+/g, '')
    .replace(/[^A-Za-z0-9]/g, '')
    .toUpperCase();

  if (/^\d{13}$/.test(text)) return 'T' + text;
  if (/^T\d{13}$/.test(text)) return text;

  const match = text.match(/T\d{13}/);
  return match ? match[0] : null;
}

/**
 * 国税庁Web-APIから登録事業者情報を取得する。
 */
function fetchInvoiceBusinessInfo(invoiceNumber) {
  const normalized = normalizeInvoiceNumber(invoiceNumber);
  if (!normalized) {
    return {
      invoiceNumber: '',
      status: 'invalid_number',
      errorMessage: 'インボイス登録番号が T + 13桁形式ではありません。'
    };
  }

  const cached = getCachedInvoiceBusinessInfo(normalized);
  if (cached && isFreshInvoiceCache(cached)) {
    return cachedToInvoiceBusinessInfo(cached, false);
  }

  const appId = PropertiesService.getScriptProperties().getProperty(NTA_INVOICE_APP_ID_PROPERTY);
  if (!appId) {
    if (cached) return cachedToInvoiceBusinessInfo(cached, true);
    return {
      invoiceNumber: normalized,
      status: 'not_configured',
      errorMessage: NTA_INVOICE_APP_ID_PROPERTY + ' が Script Properties に設定されていません。'
    };
  }

  try {
    const url = NTA_INVOICE_API_URL
      + '?id=' + encodeURIComponent(appId)
      + '&type=21&history=0&number=' + encodeURIComponent(normalized);

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    const body = response.getContentText('UTF-8');

    if (code < 200 || code >= 300) {
      if (cached) {
        const cachedResult = cachedToInvoiceBusinessInfo(cached, true);
        cachedResult.apiErrorMessage = '国税庁Web-API HTTPエラー: ' + code;
        cachedResult.raw = body;
        return cachedResult;
      }
      return {
        invoiceNumber: normalized,
        status: 'api_error',
        errorMessage: '国税庁Web-API HTTPエラー: ' + code,
        raw: body
      };
    }

    const json = JSON.parse(body);
    const item = extractInvoiceApiItem(json);
    if (!item) {
      return {
        invoiceNumber: normalized,
        status: 'not_found',
        raw: json
      };
    }

    const result = {
      invoiceNumber: item.registratedNumber || item.registrationNumber || item.invoiceNumber || normalized,
      registeredName: item.name || item.registeredName || '',
      address: buildInvoiceBusinessAddress(item),
      registrationDate: item.registrationDate || item.registratedDate || '',
      updateDate: item.updateDate || '',
      disposalDate: item.disposalDate || '',
      expireDate: item.expireDate || '',
      status: item.process || item.status || 'found',
      raw: json
    };
    updateInvoiceBusinessInfoCache(result);
    return result;
  } catch (error) {
    if (cached) {
      const cachedResult = cachedToInvoiceBusinessInfo(cached, true);
      cachedResult.apiErrorMessage = error.message;
      return cachedResult;
    }
    return {
      invoiceNumber: normalized,
      status: 'api_error',
      errorMessage: error.message
    };
  }
}

/**
 * 経費オブジェクトにインボイスAPI取得結果を別項目として付与する。
 */
function enrichWithInvoiceInfo(expense) {
  expense = expense || {};
  const normalized = normalizeInvoiceNumber(expense.invoiceNumber);
  expense.invoiceNumber = normalized || '';

  if (!normalized) return expense;

  const info = fetchInvoiceBusinessInfo(normalized);
  expense.invoiceRegisteredName = info.registeredName || '';
  expense.invoiceAddress = info.address || '';
  expense.invoiceStatus = getInvoiceStatusDisplayName(info.status);
  expense.invoiceRegistrationDate = info.registrationDate || '';
  expense.invoiceExpireDate = info.expireDate || '';
  expense.invoiceCheckedAt = info.checkedAt || Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm');
  expense.invoiceApiCheckedAt = expense.invoiceCheckedAt;
  expense.invoiceApiError = info.errorMessage || info.apiErrorMessage || '';
  expense.invoiceApiRaw = info.raw || '';

  return expense;
}


function getInvoiceStatusDisplayName(status) {
  const value = String(status || '').trim();
  if (!value) return '未確認';

  const normalized = value.toLowerCase();
  const statusMap = {
    '1': '有効',
    '01': '有効',
    'found': '有効',
    '公表': '有効',
    '有効': '有効',
    '2': '失効',
    '02': '失効',
    '失効': '失効',
    '3': '取消',
    '03': '取消',
    '取消': '取消',
    'not_found': '登録なし',
    '登録なし': '登録なし',
    'invalid_number': '番号不正',
    '番号不正': '番号不正',
    'not_configured': '未確認',
    'api_error': '未確認',
    'cached': '有効'
  };

  return statusMap[value] || statusMap[normalized] || value;
}

function extractInvoiceApiItem(json) {
  if (!json) return null;
  if (json.announcement && json.announcement.length) return json.announcement[0];
  if (json.data && json.data.length) return json.data[0];
  if (json.items && json.items.length) return json.items[0];
  if (json.registratedNumber || json.name) return json;
  return null;
}

function buildInvoiceBusinessAddress(item) {
  return item.address || [item.prefecture, item.city, item.street].filter(function(v) { return v; }).join('');
}

function setNtaInvoiceAppId(appId) {
  PropertiesService.getScriptProperties().setProperty(NTA_INVOICE_APP_ID_PROPERTY, String(appId || '').trim());
}

function testNtaInvoiceApi(invoiceNumber) {
  const result = fetchInvoiceBusinessInfo(invoiceNumber);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


function getInvoiceCacheSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_INVOICE_CACHE);
  if (!sheet) sheet = ss.insertSheet(SHEET_INVOICE_CACHE);
  ensureInvoiceCacheHeaders(sheet);
  return sheet;
}

function ensureInvoiceCacheHeaders(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) {
    return String(v || '').trim();
  });

  INVOICE_CACHE_HEADERS.forEach(function(header, index) {
    if (headers[index] !== header) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });

}


function backfillInvoiceCacheVersion(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const versionColumn = INVOICE_CACHE_HEADERS.length;
  const versionRange = sheet.getRange(2, versionColumn, lastRow - 1, 1);
  const values = versionRange.getValues();
  let changed = false;

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null || values[i][0] === undefined) {
      values[i][0] = '1';
      changed = true;
    }
  }

  if (changed) versionRange.setValues(values);
}

function getCachedInvoiceBusinessInfo(invoiceNumber) {
  const sheet = getInvoiceCacheSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, INVOICE_CACHE_HEADERS.length).getValues();
  backfillInvoiceCacheVersion(sheet);
  for (let i = 0; i < values.length; i++) {
    if (normalizeInvoiceNumber(values[i][0]) === invoiceNumber) {
      return {
        row: i + 2,
        invoiceNumber: invoiceNumber,
        registeredName: values[i][1] || '',
        address: values[i][2] || '',
        status: values[i][3] || '',
        registrationDate: values[i][4] || '',
        expireDate: values[i][5] || '',
        checkedAt: values[i][6] || '',
        version: values[i][7] || ''
      };
    }
  }

  return null;
}

function isFreshInvoiceCache(cached) {
  if (!cached || String(cached.version || '') !== '1') return false;

  const checkedDate = toInvoiceCacheDate(cached.checkedAt);
  if (!checkedDate) return false;

  const ageMillis = new Date().getTime() - checkedDate.getTime();
  return ageMillis >= 0 && ageMillis <= INVOICE_CACHE_MAX_AGE_DAYS * 24 * 60 * 60 * 1000;
}

function cachedToInvoiceBusinessInfo(cached, stale) {
  return {
    invoiceNumber: cached.invoiceNumber,
    registeredName: cached.registeredName || '',
    address: cached.address || '',
    registrationDate: cached.registrationDate || '',
    updateDate: '',
    disposalDate: '',
    expireDate: cached.expireDate || '',
    status: cached.status || 'cached',
    checkedAt: formatInvoiceCacheDate(cached.checkedAt),
    fromCache: true,
    cacheStale: Boolean(stale),
    cacheVersion: cached.version || '',
    raw: cached
  };
}

function updateInvoiceBusinessInfoCache(info) {
  const normalized = normalizeInvoiceNumber(info.invoiceNumber);
  if (!normalized) return;

  const sheet = getInvoiceCacheSheet();
  const cached = getCachedInvoiceBusinessInfo(normalized);
  const rowValues = [
    normalized,
    info.registeredName || '',
    info.address || '',
    info.status || '',
    info.registrationDate || '',
    info.expireDate || '',
    Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd HH:mm'),
    '1'
  ];

  if (cached) {
    sheet.getRange(cached.row, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}

function toInvoiceCacheDate(value) {
  if (value instanceof Date) return value;
  if (!value) return null;

  const date = new Date(String(value).replace(/\//g, '-'));
  return isNaN(date.getTime()) ? null : date;
}

function formatInvoiceCacheDate(value) {
  const date = toInvoiceCacheDate(value);
  if (!date) return String(value || '');
  return Utilities.formatDate(date, TIMEZONE, 'yyyy/MM/dd HH:mm');
}
