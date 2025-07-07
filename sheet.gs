// Formatos
const formatContabilidadeBR = '_([$R$ -416]* #,##0.00_);_([$R$ -416]* \\(#,##0.00\\);_([$R$ -416]* "-"??_);_(@_)';

function initializeSheet(sheetName, append, headersConfig, options = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  const headers = headersConfig.map(h => h.name);

  if (!append) {
    if (sheet) {
      if (sheet.getFilter()) sheet.getFilter().remove();
      sheet.clear();
    } else {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.appendRow(headers);
  } else {
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
    }
  }

  headersConfig.forEach((config, i) => {
    const col = i + 1;
    const colLetter = getColumnLetter(col);

    if (config.width) sheet.setColumnWidth(col, config.width);

    if (config.align) {
      sheet.getRange(`${colLetter}:${colLetter}`).setHorizontalAlignment(config.align);
    }

    if (config.format) {
      sheet.getRange(`${colLetter}:${colLetter}`).setNumberFormat(config.format);
    }
  });

  if (options.autoFilter && !sheet.getFilter()) {
    sheet.getRange(1, 1, 1, headers.length).createFilter();
  }

  if (options.frozenRows && options.frozenRows > 0) {
    sheet.setFrozenRows(options.frozenRows);
  }

  return sheet;
}

function getColumnLetter(col) {
  let temp = "", letter = "";
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }
  return letter;
}
