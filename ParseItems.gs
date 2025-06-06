/**
 * Парсит товары по артикулам из таблицы
 * @param {Object} columnConfig - Конфигурация колонок
 * @param {number} headerRow - Номер строки заголовков
 * @param {number} walletPercent - Процент WB-кошелька
 * @return {Object} Результат обработки
 */
function parseArticles(columnConfig, headerRow, walletPercent) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const articleColumn = letterToColumn(columnConfig.article);
  const finalPriceColumn = letterToColumn(columnConfig.finalPrice);
  const priceWithSppColumn = letterToColumn(columnConfig.priceWithSpp);
  const priceInLkColumn = letterToColumn(columnConfig.priceInLk);
  const sppPercentColumn = letterToColumn(columnConfig.sppPercent);
  
  const priceDifferenceColumn = columnConfig.priceDifference ? letterToColumn(columnConfig.priceDifference) : null;
  const totalQuantityColumn = columnConfig.totalQuantity ? letterToColumn(columnConfig.totalQuantity) : null;
  const vendorCodeColumn = columnConfig.vendorCode ? letterToColumn(columnConfig.vendorCode) : null;
  
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(headerRow + 1, articleColumn, lastRow - headerRow, 1);
  const columnValues = range.getValues();
  
  const articles = extractArticles(columnValues);
  
  if (articles.length === 0) {
    const columnLetter = columnToLetter(articleColumn);
    return {
      title: '❌ Артикулы не найдены',
      message: `В столбце ${columnLetter} не обнаружено валидных артикулов!\n\n` +
               'Проверьте:\n' +
               '1. Содержит ли столбец числовые значения\n' +
               '2. Нет ли скрытых символов в ячейках'
    };
  }
  
  const processedData = processArticlesDirectly(articles, walletPercent);
  
  writeResultsToSheet(
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
  );
  
  return formatResult(articles, columnToLetter(articleColumn), processedData);
}

/**
 * Прямая обработка артикулов без получения базовых товаров
 * @param {number[]} articles - Массив артикулов
 * @param {number} walletPercent - Процент WB-кошелька
 * @return {Object[]} Обработанные данные
 */
function processArticlesDirectly(articles, walletPercent) {
  const wbApi = new WBPrivateAPI({ destination: CONFIG.DESTINATIONS.KRASNODAR });
  
  const uniqueArticles = [...new Set(articles)]; // Убираем дубликаты
  const chunkSize = 100;
  const chunks = chunkArray(uniqueArticles, chunkSize);
  let detailedProducts = [];
  
  chunks.forEach((chunk, index) => {
    try {
      const products = wbApi.getListOfProducts(chunk);
      detailedProducts = [...detailedProducts, ...products];
    } catch (err) {
      Logger.log(`Чанк ${index + 1} не обработан: ${err.message}`);
    }
  });
  
  return processDetailedProducts(detailedProducts, walletPercent);
}

/**
 * Обработка только детализированных товаров
 * @param {Object[]} detailedProducts - Детализированные продукты
 * @param {number} walletPercent - Процент WB-кошелька
 * @return {Object[]} Обработанные данные
 */
function processDetailedProducts(detailedProducts, walletPercent) {
  const walletFraction = walletPercent / 100;
  
  return detailedProducts.map(product => {
    const sppPercentage = calculateSPP(product.priceU, product.salePriceU);
    const priceSpp = product.salePriceU || 0;
    const clientPrice = calculateClientPrice(priceSpp, walletFraction);
    
    return {
      id: product.id,
      vendorCode: product.vendorCode || 'N/A',
      seller_price: convertToRubles(product.priceU || 0),
      price_spp: convertToRubles(priceSpp),
      client_price: convertToRubles(clientPrice),
      diffPrice: convertToRubles((product.priceU || 0) - (product.salePriceU || 0)),
      spp: `${sppPercentage}%`,
      totalQuantity: product.totalQuantity || 0
    };
  });
}

/**
 * Парсит все товары из личного кабинета
 * @param {Object} columnConfig - Конфигурация колонок
 * @param {number} headerRow - Номер строки заголовков
 * @param {number} walletPercent - Процент WB-кошелька
 * @return {Object} Результат обработки
 */
function parseAllFromLk(columnConfig, headerRow, walletPercent) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const articleColumn = letterToColumn(columnConfig.article);
  const finalPriceColumn = letterToColumn(columnConfig.finalPrice);
  const priceWithSppColumn = letterToColumn(columnConfig.priceWithSpp);
  const priceInLkColumn = letterToColumn(columnConfig.priceInLk);
  const sppPercentColumn = letterToColumn(columnConfig.sppPercent);
  
  const priceDifferenceColumn = columnConfig.priceDifference ? letterToColumn(columnConfig.priceDifference) : null;
  const totalQuantityColumn = columnConfig.totalQuantity ? letterToColumn(columnConfig.totalQuantity) : null;
  const vendorCodeColumn = columnConfig.vendorCode ? letterToColumn(columnConfig.vendorCode) : null;
  
  const products = fetchAllProducts(walletPercent);
  
  if (products.length === 0) {
    return {
      title: '❌ Товары не найдены',
      message: 'В личном кабинете не найдено товаров. Проверьте токен и настройки API.'
    };
  }
  
  writeLkResultsToSheet(
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
  );
  
  return {
    title: '✅ Результаты парсинга всех товаров',
    message: `Обработано товаров: ${products.length}\n\n` +
             `Первый артикул: ${products[0].id}\n` +
             `Последний артикул: ${products[products.length - 1].id}\n` +
             `Данные записаны в таблицу, начиная со строки ${headerRow + 1}`
  };
}

/**
 * Форматирует результат обработки артикулов
 * @param {number[]} articles - Массив артикулов
 * @param {string} columnLetter - Буква столбца
 * @param {Object[]} processedData - Обработанные данные
 * @return {Object} Объект с данными для отображения
 */
function formatResult(articles, columnLetter, processedData) {
  const count = articles.length;
  const processedCount = processedData.length;
  
  return {
    title: `✅ Результаты обработки столбца ${columnLetter}`,
    message: `Найдено артикулов: ${count}\n` +
             `Обработано продуктов: ${processedCount}\n\n` +
             `Примеры значений:\n${articles.slice(0, 10).join('\n')}${count > 10 ? '\n...' : ''}\n\n` +
             `Первый артикул: ${articles[0]}\n` +
             `Последний артикул: ${articles[count - 1]}\n` +
             `Уникальных: ${[...new Set(articles)].length}\n\n` +
             '⚠️ Данные записаны в таблицу'
  };
}