/**
 * Возвращает список букв колонок и их заголовков
 * @param {number} headerRow - Номер строки заголовков
 * @return {Object[]} Массив объектов [{ column: string, header: string }]
 */
function getSelectedColumns(headerRow) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const selection = sheet.getActiveRange();
    const startColumn = selection.getColumn();
    const numColumns = selection.getNumColumns();
    const columns = [];

    const headers = sheet
      .getRange(headerRow, startColumn, 1, numColumns)
      .getValues()[0];

    for (let i = 0; i < numColumns; i++) {
      const columnIndex = startColumn + i;
      const columnLetter = columnToLetter(columnIndex);
      const header = headers[i] ? String(headers[i]).trim() : columnLetter;
      columns.push({ column: columnLetter, header: header });
    }

    Logger.log("Выделенные колонки и заголовки: " + JSON.stringify(columns));
    return columns;
  } catch (error) {
    Logger.log("Ошибка в getSelectedColumns: " + error.message);
    throw new Error(
      "Не удалось получить заголовки. Проверьте номер строки и выделение."
    );
  }
}

/**
 * Записывает результаты парсинга по артикулам
 */
function writeResultsToSheet(
  sheet,
  articleColumn,
  finalPriceColumn,
  priceWithSppColumn,
  priceInLkColumn,
  sppPercentColumn,
  priceDifferenceColumn,
  totalQuantityColumn,
  vendorCodeColumn,
  processedData,
  articles,
  headerRow
) {
  const lastRow = sheet.getLastRow();
  const columnValues = sheet
    .getRange(headerRow + 1, articleColumn, lastRow - headerRow, 1)
    .getValues();
  const dataMap = new Map();

  processedData.forEach((item) => {
    dataMap.set(item.id, item);
  });

  const results = [];

  for (let i = 0; i < columnValues.length; i++) {
    const row = columnValues[i];
    const article = extractArticle(row[0]);

    if (!article || !articles.includes(article)) continue;

    const data = dataMap.get(article);
    if (!data) continue;

    const rowData = {};

    rowData[finalPriceColumn] = formatValue(
      data.client_price ? `${data.client_price} ₽` : null
    );
    rowData[priceWithSppColumn] = formatValue(
      data.price_spp ? `${data.price_spp} ₽` : null
    );
    rowData[priceInLkColumn] = formatValue(
      data.seller_price ? `${data.seller_price} ₽` : null
    );
    rowData[sppPercentColumn] = formatValue(data.spp);

    if (priceDifferenceColumn) {
      rowData[priceDifferenceColumn] = formatValue(
        data.diffPrice ? `${data.diffPrice} ₽` : null
      );
    }
    if (totalQuantityColumn) {
      rowData[totalQuantityColumn] = formatValue(data.totalQuantity);
    }
    if (vendorCodeColumn) {
      rowData[vendorCodeColumn] = formatValue(data.vendorCode);
    }

    results.push({
      row: headerRow + i + 1,
      data: rowData,
    });
  }

  results.forEach((item) => {
    Object.entries(item.data).forEach(([column, value]) => {
      sheet.getRange(item.row, parseInt(column)).setValue(value);
    });
  });
}

/**
 * Записывает результаты парсинга всех товаров из ЛК
 */
function writeLkResultsToSheet(
  sheet,
  articleColumn,
  finalPriceColumn,
  priceWithSppColumn,
  priceInLkColumn,
  sppPercentColumn,
  priceDifferenceColumn,
  totalQuantityColumn,
  vendorCodeColumn,
  products,
  headerRow
) {
  products.forEach((product, index) => {
    const row = headerRow + 1 + index;
    const rowData = {};

    rowData[articleColumn] = formatValue(product.id);
    rowData[finalPriceColumn] = formatValue(
      product.client_price ? `${product.client_price} ₽` : null
    );
    rowData[priceWithSppColumn] = formatValue(
      product.price_spp ? `${product.price_spp} ₽` : null
    );
    rowData[priceInLkColumn] = formatValue(
      product.seller_price ? `${product.seller_price} ₽` : null
    );
    rowData[sppPercentColumn] = formatValue(product.spp);

    if (priceDifferenceColumn) {
      rowData[priceDifferenceColumn] = formatValue(
        product.diffPrice ? `${product.diffPrice} ₽` : null
      );
    }
    if (totalQuantityColumn) {
      rowData[totalQuantityColumn] = formatValue(product.totalQuantity);
    }
    if (vendorCodeColumn) {
      rowData[vendorCodeColumn] = formatValue(product.vendorCode);
    }

    Object.entries(rowData).forEach(([column, value]) => {
      sheet.getRange(row, parseInt(column)).setValue(value);
    });
  });
}
