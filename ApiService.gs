/**
 * Создает конфигурацию HTTP-запроса
 * @return {Object} Конфигурация сессии
 */
function createSession() {
  return {
    headers: {
      'User-Agent': CONFIG.USERAGENT,
      'Accept-Encoding': 'gzip, deflate, br',
    },
    muteHttpExceptions: true,
  };
}

/**
 * Класс для работы с API Wildberries
 */
class WBPrivateAPI {
  constructor({ destination }) {
    this.sessionConfig = createSession();
    this.destination = destination;
  }

  /**
   * Получает список продуктов по их nmIDs
   * @param {number[]} productIds - Массив артикулов
   * @returns {Object[]} Массив продуктов
   */
  getListOfProducts(productIds) {
    const queryParams = {
      appType: CONFIG.APPTYPES.DESKTOP,
      curr: 'rub',
      dest: CONFIG.DESTINATIONS.DEFAULT.id,
      hide_dtype: 13,
      ab_testing: false,
      lang: 'ru',
      nm: productIds.join(';'),
    };

    const queryString = Object.keys(queryParams)
      .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(queryParams[key])}`)
      .join('&');

    const url = `${CONFIG.URLS.SEARCH.PRODUCTS}?${queryString}`;

    const options = {
      ...this.sessionConfig,
      method: 'get',
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      throw new Error(`Ошибка запроса к API: статус ${responseCode}`);
    }

    const data = JSON.parse(response.getContentText());
    const rawProducts = data?.products || [];

    return rawProducts.map(p => {
      const size = p.sizes?.[0] || {};
      const price = size.price || {};
      const stocks = size.stocks || [];
      const totalQuantity = stocks.reduce((acc, s) => acc + (s.qty || 0), 0);

      return {
        id: p.id,
        vendorCode: p.vendorCode || 'N/A',
        priceU: price.basic || 0,
        salePriceU: price.product || 0,
        totalQuantity,
      };
    });
  }

}

/**
 * Тестирует токен с помощью пробного API запроса
 * @return {Object} Результат проверки
 */
function testToken() {
  const token = getTokenFromSheet();
  if (!token) {
    return {
      success: false,
      message: 'Токен не найден. Укажите токен в ячейке C1.'
    };
  }

  const options = {
    method: 'get',
    headers: { 'Authorization': `Bearer ${token}` },
    muteHttpExceptions: true,
  };
  const url = `${CONFIG.API_ENDPOINTS.PRODUCTS_FILTER}?limit=1&offset=0`;

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      return { success: true, message: 'Токен валиден.' };
    } else if (responseCode === 401) {
      return { success: false, message: 'Недействительный токен. Проверьте правильность токена.' };
    } else {
      return { success: false, message: `Ошибка API: статус ${responseCode}.` };
    }
  } catch (error) {
    return { success: false, message: `Ошибка проверки токена: ${error.message}` };
  }
}

/**
 * Получает все товары из личного кабинета
 * @param {number} walletPercent - Процент WB-кошелька
 * @return {Object[]} Массив товаров
 */
function fetchAllProducts(walletPercent) {
  const token = getTokenFromSheet();
  if (!token) {
    throw new Error('Токен не найден. Укажите токен в ячейке C1.');
  }

  const options = {
    method: 'get',
    headers: { 'Authorization': `Bearer ${token}` },
    muteHttpExceptions: true,
  };
  const url = `${CONFIG.API_ENDPOINTS.PRODUCTS_FILTER}?limit=${CONFIG.REQUEST_CONFIG.LIMIT}&offset=${CONFIG.REQUEST_CONFIG.OFFSET}`;

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(`Ошибка запроса к API: ${responseCode}`);
  }

  const data = JSON.parse(response.getContentText());
  const baseProducts = data?.data?.listGoods || [];
  Logger.log('Всего товаров: ' + baseProducts.length);

  const wbApi = new WBPrivateAPI({ destination: CONFIG.DESTINATIONS.KRASNODAR });
  const nmIds = baseProducts.map(p => p.nmID).filter(id => id && !isNaN(id));

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

  return processDataForLk(baseProducts, detailedProducts, walletPercent);
}

/**
 * Обработка данных для режима ЛК
 * @param {Object[]} baseProducts - Базовые продукты
 * @param {Object[]} detailedProducts - Детализированные продукты
 * @param {number} walletPercent - Процент WB-кошелька
 * @return {Object[]} Обработанные данные
 */
function processDataForLk(baseProducts, detailedProducts, walletPercent) {
  const walletFraction = walletPercent / 100;
  const detailsMap = new Map();

  detailedProducts.forEach(product => {
    detailsMap.set(product.id, product);
  });

  return baseProducts.map(baseProduct => {
    const details = detailsMap.get(baseProduct.nmID) || {};
    const sppPercentage = calculateSPP(details.priceU, details.salePriceU);
    const priceSpp = details.salePriceU || 0;
    const clientPrice = calculateClientPrice(priceSpp, walletFraction);

    return {
      id: baseProduct.nmID,
      vendorCode: baseProduct?.vendorCode,
      seller_price: convertToRubles(details.priceU || 0),
      price_spp: convertToRubles(priceSpp),
      client_price: convertToRubles(clientPrice),
      diffPrice: convertToRubles((details.priceU || 0) - (details.salePriceU || 0)),
      spp: `${sppPercentage}%`,
      totalQuantity: details.totalQuantity || 0
    };
  });
}

/**
 * Конвертирует цену из копеек в рубли
 * @param {number} priceInCents - Цена в копейках
 * @return {number} Цена в рублях
 */
function convertToRubles(priceInCents) {
  return Math.floor(priceInCents / 100);
}

/**
 * Вычисляет процент СПП
 * @param {number} priceU - Цена до СПП
 * @param {number} salePriceU - Цена с СПП
 * @return {number} Процент СПП
 */
function calculateSPP(priceU, salePriceU) {
  if (!priceU) return 0;
  const priceRub = convertToRubles(priceU);
  const salePriceRub = convertToRubles(salePriceU);
  return Math.floor(((priceRub - salePriceRub) / priceRub) * 100);
}

/**
 * Вычисляет цену для клиента
 * @param {number} priceSpp - Цена с СПП
 * @param {number} walletFraction - Доля WB-кошелька (например, 0.03 для 3%)
 * @return {number} Цена для клиента
 */
function calculateClientPrice(priceSpp, walletFraction) {
  return priceSpp * (1 - walletFraction);
}
