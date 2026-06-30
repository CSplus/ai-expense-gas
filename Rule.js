/*************************************************
 * Rule.gs
 * 勘定科目・取引先正規名ルール
 *************************************************/

function getAccountingRule(vendor, invoiceNumber, registeredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RULE);

  if (!sheet) {
    throw new Error('仕分けルールシートが見つかりません。');
  }

  ensureAccountingRuleInvoiceColumns(sheet);

  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const invoiceNumberCol = findHeaderIndex(headers, 'インボイス登録番号');
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
          vendorName: values[i][3] || registeredName || vendor
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
          vendorName: values[i][3] || registeredName || vendor
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
        vendorName: values[i][3] || vendor
      };
    }
  }

  return {
    accountCode: 6231,
    accountName: '雑費',
    vendorName: vendor
  };
}

function getAccountByCode(code) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_ACCOUNT);

  if (!sheet) {
    throw new Error('勘定科目マスタシートが見つかりません。');
  }

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(code)) {
      return {
        accountCode: values[i][0],
        accountName: values[i][1]
      };
    }
  }

  return {
    accountCode: code,
    accountName: '未設定'
  };
}

function getAccountingRuleFromInput(categoryInput) {
  let code;

  switch (categoryInput) {
    case '駐車場・交通費':
      code = getConfig('駐車場・交通費コード');
      break;

    case '飲食費・会食':
      code = getConfig('飲食費・会食コード');
      break;

    case 'ガソリン':
      code = getConfig('ガソリンコード');
      break;

    case 'その他経費':
      code = getConfig('その他経費コード');
      break;

    default:
      code = getConfig('デフォルトコード');
  }

  return getAccountByCode(code);
}


function ensureAccountingRuleInvoiceColumns(sheet) {
  sheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RULE);
  if (!sheet) return;

  const requiredHeaders = ['インボイス登録番号', '正式事業者名'];
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
