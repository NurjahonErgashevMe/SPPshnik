/**
 * Конфигурация и константы API
 */
const CONFIG = {
  URLS: {
    SEARCH: {
      PRODUCTS: 'https://card.wb.ru/cards/v4/detail',
    },
  },
  APPTYPES: {
    DESKTOP: 1,
  },
  DESTINATIONS: {
    KRASNODAR: {
      ids: [1059500],
      regions: [64, 65, 83, 7, 8, 80, 33, 70, 82, 86, 30, 69, 22, 66, 31, 40, 1, 48],
    },
    DEFAULT : {
      id : 12358062
    }
  },
  USERAGENT: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
  API_ENDPOINTS: {
    PRODUCTS_FILTER: 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter',
  },
  REQUEST_CONFIG: {
    LIMIT: 1000,
    OFFSET: 0,
  },
};
