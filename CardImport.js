/*************************************************
 * CardImport.gs
 * クレジットカードCSV取込
 *************************************************/

function processCardCsvFiles() {
  const sheet = getExpenseSheet();
  const unprocessedFolder = DriveApp.getFolderById(getCardUnprocessedFolderId());
  const processedFolder = DriveApp.getFolderById(getCardProcessedFolderId());
  const files = unprocessedFolder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (!fileName.toLowerCase().endsWith('.csv')) continue;

    if (fileName.startsWith('enavi')) {
      importRakutenCardCsv(sheet, file);
    } else if (fileName.startsWith('IDEMITSU')) {
      importIdemitsuCardCsv(sheet, file);
    } else {
      throw new Error('未対応のCSVファイル名です: ' + fileName);
    }

    file.moveTo(processedFolder);
  }
}

function importRakutenCardCsv(sheet, file) {
  const csv = readJapaneseCsvText(file);
  const rows = Utilities.parseCsv(csv);
  const fileName = file.getName();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const date = row[0];
    const vendor = row[1];
    const amount = Number(String(row[4]).replace(/,/g, ''));

    if (!date || !vendor || !amount) continue;
    if (vendor.includes('現地利用額')) continue;

    appendCardExpenseRow(sheet, date, vendor, amount, '楽天カード', fileName);
  }
}

function importIdemitsuCardCsv(sheet, file) {
  const csv = readJapaneseCsvText(file);
  const rows = Utilities.parseCsv(csv);
  const fileName = file.getName();

  for (let i = 6; i < rows.length; i++) {
    const row = rows[i];
    const date = row[0];
    const vendor = row[1];
    const amount = Number(String(row[5]).replace(/,/g, ''));

    if (!date || !vendor || !amount) continue;

    appendCardExpenseRow(sheet, date, vendor, amount, '出光カード', fileName);
  }
}

function readJapaneseCsvText(file) {
  const blob = file.getBlob();
  const encodings = ['Windows-31J', 'MS932', 'Shift_JIS', 'UTF-8'];

  for (const enc of encodings) {
    try {
      const text = blob.getDataAsString(enc);

      if (
        !text.includes('�') &&
        !text.includes('譁') &&
        !text.includes('縺') &&
        !text.includes('繧') &&
        !text.includes('驥')
      ) {
        return text;
      }
    } catch (e) {}
  }

  return blob.getDataAsString('Windows-31J');
}
