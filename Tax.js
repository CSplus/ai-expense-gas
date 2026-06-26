/*************************************************
 * Tax.gs
 * 税率マスタ・消費税額計算
 *************************************************/

/**
 * 取引日に対応する税率設定を取得する
 */
function getTaxRateMaster(transactionDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TAX_RATE);

  if (!sheet) {
    throw new Error('税率マスタシートが見つかりません。');
  }

  const values = sheet.getDataRange().getValues();
  const targetDate = toDate(transactionDate) || new Date();

  for (let i = 1; i < values.length; i++) {
    const startDate = toDate(values[i][0]);
    const endDate = toDate(values[i][1]);

    if (!startDate || !endDate) continue;

    if (targetDate >= startDate && targetDate <= endDate) {
      return {
        standardRate: Number(values[i][2]),
        reducedRate: Number(values[i][3])
      };
    }
  }

  return {
    standardRate: 10,
    reducedRate: 8
  };
}

/**
 * 税込金額から消費税額を計算する
 */
function calculateTaxIncluded(amount, rate) {
  amount = Number(amount || 0);
  rate = Number(rate || 0);

  if (!amount) return '';
  if (rate === 0) return 0;

  return Math.round(amount / (100 + rate) * rate);
}

/**
 * Gemini結果と税率マスタから税率・消費税額を整える
 */
function normalizeTaxInfo(result) {
  const amount = Number(result.amount || 0);
  const master = getTaxRateMaster(result.date);

  let taxRate = result.taxRate || result.tax_rate || '';
  let taxAmount = result.taxAmount || result.tax_amount || '';
  let taxNote = result.taxNote || result.tax_note || '';

  taxRate = normalizeTaxRateText(taxRate);

  // 1. 領収書に税額が明記されている場合は優先
  if (taxAmount !== '' && taxAmount !== null && taxAmount !== undefined) {
    return {
      taxRate: taxRate || '不明',
      taxAmount: Number(taxAmount),
      taxNote: taxNote || '消費税額は領収書記載'
    };
  }

  if (!amount) {
    return {
      taxRate: taxRate || '不明',
      taxAmount: '',
      taxNote: '税込金額が不明のため消費税額を計算できません'
    };
  }

  // 2. 領収書に税率が明記されている場合
  if (taxRate === '10%' || taxRate === '8%' || taxRate === '0%') {
    const rateNumber = Number(taxRate.replace('%', ''));
    return {
      taxRate: taxRate,
      taxAmount: calculateTaxIncluded(amount, rateNumber),
      taxNote: '消費税額は領収書記載税率から計算'
    };
  }

  // 3. 軽減税率っぽい場合
  if (isReducedTaxCandidate(result)) {
    return {
      taxRate: master.reducedRate + '%',
      taxAmount: calculateTaxIncluded(amount, master.reducedRate),
      taxNote: '税率は税率マスタの軽減税率を適用'
    };
  }

  // 4. 通常は標準税率
  return {
    taxRate: master.standardRate + '%',
    taxAmount: calculateTaxIncluded(amount, master.standardRate),
    taxNote: '税率は税率マスタの標準税率を適用'
  };
}

/**
 * 税率表記を正規化する
 */
function normalizeTaxRateText(value) {
  if (!value) return '';

  let text = String(value)
    .replace(/[０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .replace('％', '%')
    .trim();

  if (text === '10') text = '10%';
  if (text === '8') text = '8%';
  if (text === '0') text = '0%';

  return text;
}

/**
 * 軽減税率候補か判定する
 * 必要に応じて後で強化
 */
function isReducedTaxCandidate(result) {
  const vendor = String(result.vendor || '');
  const category = String(result.category || '');
  const memo = String(result.memo || '');

  const text = vendor + ' ' + category + ' ' + memo;

  return /スーパー|食品|弁当|惣菜|パン|青果|精肉|鮮魚|米|飲料|テイクアウト|持ち帰り/.test(text);
}