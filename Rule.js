/*************************************************
 * Rule.gs
 * カード明細仕訳ルール・入力区分別勘定科目設定
 *************************************************/

function getAccountingRule(vendor, invoiceNumber, registeredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RULE);

  if (!sheet) {
    throw new Error('カード明細仕訳ルールシートが見つかりません。');
  }

  ensureAccountingRuleInvoiceColumns(sheet);

  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const invoiceNumberCol = findHeaderIndex(headers, '登録番号');
  const registeredNameCol = findHeaderIndex(headers, '正式事業者名');
  const normalizedInvoiceNumber = normalizeInvoiceNumber(invoiceNumber);
  const targetRegisteredName = String(registeredName || '').toUpperCase().trim();
  const target = String(vendor || '').toUpperCase().trim();

  if (normalizedInvoiceNumber && invoiceNumberCol >= 0) {
    for (let i = 1; i < values.length; i++) {
      if (normalizeInvoiceNumber(values[i][invoiceNumberCol]) === normalizedInvoiceNumber) {
        return {
          accountCode: values[i][1],
          accountName: values[i][2],
          vendorName: registeredName || vendor
        };
      }
    }
  }

  if (targetRegisteredName && registeredNameCol >= 0) {
    for (let i = 1; i < values.length; i++) {
      const ruleName = String(values[i][registeredNameCol] || '').toUpperCase().trim();
      if (ruleName && targetRegisteredName === ruleName) {
        return {
          accountCode: values[i][1],
          accountName: values[i][2],
          vendorName: registeredName || vendor
        };
      }
    }
  }

  for (let i = 1; i < values.length; i++) {
    const keyword = String(values[i][0] || '').toUpperCase().trim();
    if (!keyword) continue;

    if (target.includes(keyword)) {
      return {
        accountCode: values[i][1],
        accountName: values[i][2],
        vendorName: vendor
      };
    }
  }

  return {
    accountCode: 6231,
    accountName: '雑費',
    vendorName: vendor
  };
}

function getAccountingRuleFromInput(categoryInput) {
  const normalizedCategoryInput = String(categoryInput || '').trim();
  // 入力区分名そのものをシステム設定の「項目」として検索する。
  const account = getAccountFromSystemConfig(normalizedCategoryInput) ||
    getAccountFromSystemConfig('デフォルト');

  if (!account) {
    throw new Error('システム設定に入力区分別の勘定科目設定が見つかりません。');
  }

  return account;
}

function getAccountFromSystemConfig(itemName) {
  const target = String(itemName || '').trim();
  if (!target) return null;

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_CONFIG);

  if (!sheet) {
    throw new Error('システム設定シートが見つかりません。');
  }

  const values = sheet.getDataRange().getValues();
  if (!values.length) return null;

  const headers = values[0].map(function(v) { return String(v || '').trim(); });
  let itemCol = findHeaderIndex(headers, '項目');
  let codeCol = findHeaderIndex(headers, 'コード');
  let accountNameCol = findHeaderIndex(headers, '勘定科目名');

  if (itemCol < 0 || codeCol < 0 || accountNameCol < 0) {
    itemCol = 0;
    codeCol = 1;
    accountNameCol = 2;
  }

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][itemCol] || '').trim() === target) {
      return {
        accountCode: values[i][codeCol],
        accountName: values[i][accountNameCol]
      };
    }
  }

  return null;
}


function ensureAccountingRuleInvoiceColumns(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RULE);
  if (!sheet) return;

  const requiredHeaders = ['登録番号', '正式事業者名'];
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) {
    return String(v || '').trim();
  });

  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(header);
      headers.push(header);
    }
  });
}

function findHeaderIndex(headers, headerName) {
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === headerName) return i;
  }
  return -1;
}
