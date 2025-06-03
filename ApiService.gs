/**
 * Создает конфигурацию HTTP-запроса
 * @return {Object} Конфигурация сессии
 */
function createSession() {
  return {
    headers: {
      "User-Agent": CONFIG.USERAGENT,
      "Accept-Encoding": "gzip, deflate, br",
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
    const options = {
      ...this.sessionConfig,
      method: "get",
      params: {
        nm: productIds.join(";"),
        appType: CONFIG.APPTYPES.DESKTOP,
        dest: this.destination.ids[0],
      },
    };

    const queryString = Object.keys(options.params)
      .map(
        (key) =>
          `${encodeURIComponent(key)}=${encodeURIComponent(
            options.params[key]
          )}`
      )
      .join("&");
    const url = `${CONFIG.URLS.SEARCH.PRODUCTS}?${queryString}`;

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      throw new Error(`Ошибка запроса к API: статус ${responseCode}`);
    }

    const data = JSON.parse(response.getContentText());
    Logger.log(
      "Detailed products: " + JSON.stringify(data?.data?.products?.slice(0, 2))
    );
    return data?.data?.products || [];
  }
}

/**
 * Тестирует токен с помощью пробного API запроса
 * @param {string} token - Токен для проверки
 * @return {Object} Результат проверки
 */
function testToken(token) {
  const options = {
    method: "get",
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  };
  const url = `${CONFIG.API_ENDPOINTS.PRODUCTS_FILTER}?limit=1&offset=0`;

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      return { success: true, message: "Токен валиден." };
    } else if (responseCode === 401) {
      return {
        success: false,
        message: "Недействительный токен. Проверьте правильность токена.",
      };
    } else {
      return { success: false, message: `Ошибка API: статус ${responseCode}.` };
    }
  } catch (error) {
    return {
      success: false,
      message: `Ошибка проверки токена: ${error.message}`,
    };
  }
}

/**
 * Получает все товары из личного кабинета
 * @param {string} token - WB API токен
 * @return {Object[]} Массив товаров
 */
function fetchAllProducts(token) {
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
  const baseProducts = data?.data?.listGoods || [];
  Logger.log("Всего товаров: " + baseProducts.length);

  const wbApi = new WBPrivateAPI({
    destination: CONFIG.DESTINATIONS.KRASNODAR,
  });
  const nmIds = baseProducts
    .map((p) => p.nmID)
    .filter((id) => id && !isNaN(id));

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

  return processData(baseProducts, detailedProducts);
}

/**
 * Обрабатывает данные продуктов
 * @param {Object[]} baseProducts - Базовые продукты
 * @param {Object[]} detailedProducts - Детализированные продукты
 * @return {Object[]} Обработанные данные
 */
function processData(baseProducts, detailedProducts) {
  const detailsMap = new Map();
  detailedProducts.forEach((product) => {
    detailsMap.set(product.id, product);
  });

  return baseProducts.map((baseProduct) => {
    const details = detailsMap.get(baseProduct.nmID) || {};
    const sppPercentage = calculateSPP(details.priceU, details.salePriceU);
    const priceSpp = details.salePriceU || 0;
    const clientPrice = calculateClientPrice(priceSpp);

    return {
      id: baseProduct.nmID,
      vendorCode: baseProduct.vendorCode || "N/A",
      seller_price: convertToRubles(details.priceU || 0),
      price_spp: convertToRubles(priceSpp),
      client_price: convertToRubles(clientPrice),
      diffPrice: convertToRubles(
        (details.priceU || 0) - (details.salePriceU || 0)
      ),
      spp: `${sppPercentage}%`,
      totalQuantity: details.totalQuantity || 0,
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
 * @return {number} Цена для клиента
 */
function calculateClientPrice(priceSpp) {
  return priceSpp * (1 - CONFIG.WALLET_PERCENTAGE);
}
