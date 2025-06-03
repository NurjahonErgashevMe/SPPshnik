/**
 * Парсит товары по артикулам из таблицы
 * @param {Object} columnConfig - Конфигурация колонок
 * @param {number} headerRow - Номер строки заголовков
 * @return {Object} Результат обработки
 */
function parseArticles(columnConfig, headerRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const articleColumn = letterToColumn(columnConfig.article);
  const finalPriceColumn = letterToColumn(columnConfig.finalPrice);
  const priceWithSppColumn = letterToColumn(columnConfig.priceWithSpp);
  const priceInLkColumn = letterToColumn(columnConfig.priceInLk);
  const sppPercentColumn = letterToColumn(columnConfig.sppPercent);

  const priceDifferenceColumn = columnConfig.priceDifference
    ? letterToColumn(columnConfig.priceDifference)
    : null;
  const totalQuantityColumn = columnConfig.totalQuantity
    ? letterToColumn(columnConfig.totalQuantity)
    : null;
  const vendorCodeColumn = columnConfig.vendorCode
    ? letterToColumn(columnConfig.vendorCode)
    : null;

  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(
    headerRow + 1,
    articleColumn,
    lastRow - headerRow,
    1
  );
  const columnValues = range.getValues();

  const articles = extractArticles(columnValues);

  if (articles.length === 0) {
    const columnLetter = columnToLetter(articleColumn);
    return {
      title: "❌ Артикулы не найдены",
      message:
        `В столбце ${columnLetter} не обнаружено валидных артикулов!\n\n` +
        "Проверьте:\n" +
        "1. Содержит ли столбец числовые значения\n" +
        "2. Нет ли скрытых символов в ячейках",
    };
  }

  const processedData = processProducts(articles);

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
 * Парсит все товары из личного кабинета
 * @param {Object} columnConfig - Конфигурация колонок
 * @param {number} headerRow - Номер строки заголовков
 * @param {string} token - WB API токен
 * @return {Object} Результат обработки
 */
function parseAllFromLk(columnConfig, headerRow, token) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const articleColumn = letterToColumn(columnConfig.article);
  const finalPriceColumn = letterToColumn(columnConfig.finalPrice);
  const priceWithSppColumn = letterToColumn(columnConfig.priceWithSpp);
  const priceInLkColumn = letterToColumn(columnConfig.priceInLk);
  const sppPercentColumn = letterToColumn(columnConfig.sppPercent);

  const priceDifferenceColumn = columnConfig.priceDifference
    ? letterToColumn(columnConfig.priceDifference)
    : null;
  const totalQuantityColumn = columnConfig.totalQuantity
    ? letterToColumn(columnConfig.totalQuantity)
    : null;
  const vendorCodeColumn = columnConfig.vendorCode
    ? letterToColumn(columnConfig.vendorCode)
    : null;

  const products = fetchAllProducts(token);

  if (products.length === 0) {
    return {
      title: "❌ Товары не найдены",
      message:
        "В личном кабинете не найдено товаров. Проверьте токен и настройки API.",
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
    title: "✅ Результаты парсинга всех товаров",
    message:
      `Обработано товаров: ${products.length}\n\n` +
      `Первый артикул: ${products[0].id}\n` +
      `Последний артикул: ${products[products.length - 1].id}\n` +
      `Данные записаны в таблицу, начиная со строки ${headerRow + 1}`,
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
    message:
      `Найдено артикулов: ${count}\n` +
      `Обработано продуктов: ${processedCount}\n\n` +
      `Примеры значений:\n${articles.slice(0, 10).join("\n")}${
        count > 10 ? "\n..." : ""
      }\n\n` +
      `Первый артикул: ${articles[0]}\n` +
      `Последний артикул: ${articles[count - 1]}\n` +
      `Уникальных: ${[...new Set(articles)].length}\n\n` +
      "⚠️ Данные записаны в таблицу",
  };
}

/**
 * Обрабатывает продукты через API
 * @param {number[]} articles - Массив артикулов
 * @return {Object[]} Обработанные данные продуктов
 */
function processProducts(articles) {
  const WB_API_TOKEN = JSON.parse(
    PropertiesService.getScriptProperties().getProperty("WB_CONFIG")
  )?.token;
  if (!WB_API_TOKEN) {
    throw new Error("WB_API_TOKEN не установлен.");
  }

  const wbApi = new WBPrivateAPI({
    destination: CONFIG.DESTINATIONS.KRASNODAR,
  });

  const baseProducts = fetchBaseProducts(WB_API_TOKEN);
  const filteredBaseProducts = baseProducts.filter((product) =>
    articles.includes(product.nmID)
  );
  const nmIds = filteredBaseProducts
    .map((p) => p.nmID)
    .filter((id) => id && !isNaN(id));

  if (nmIds.length === 0) {
    throw new Error(
      "Не найдено соответствующих продуктов для указанных артикулов"
    );
  }

  const chunkSize = 100;
  const chunks = chunkArray(nmIds, chunkSize);
  let detailedProducts = [];

  chunks.forEach((chunk, index) => {
    try {
      const products = wbApi.getListOfProducts(chunk);
      detailedProducts = [...detailedProducts, ...products];
    } catch (err) {
      Logger.log(`Чанк ${index + 1} не обработан: ${err.message}`);
    }
  });

  return processData(filteredBaseProducts, detailedProducts);
}

/**
 * Получает базовые продукты через API
 * @param {string} token - WB API токен
 * @return {Object[]} Массив базовых продуктов
 */
function fetchBaseProducts(token) {
  const options = {
    method: "get",
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  };
  const url = `${CONFIG.API_ENDPOINTS.PRODUCTS_FILTER}?limit=${CONFIG.REQUEST_CONFIG.LIMIT}&offset=${CONFIG.REQUEST_CONFIG.OFFSET}`;
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(`Ошибка запроса к API: ${responseCode}`);
  }

  const data = JSON.parse(response.getContentText());
  const listGoods = data?.data?.listGoods || [];
  Logger.log("Base products count: " + listGoods.length);
  return listGoods;
}
