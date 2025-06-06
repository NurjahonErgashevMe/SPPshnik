/**
 * Создает пользовательское меню
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WB Артикулы')
    .addItem('Парсить по артикулам', 'showConfigModalForArticles')
    .addItem('Спарсить все в ЛК', 'showConfigModalForLk')
    .addToUi();
}

/**
 * Показывает модалку для парсинга по артикулам
 */
function showConfigModalForArticles() {
  const html = HtmlService.createHtmlOutputFromFile('modal')
    .setWidth(600)
    .setHeight(600);
  html.append('<script>window.parseMode = "parseArticles";</script>');
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка: Парсить по артикулам');
}

/**
 * Показывает модалку для парсинга всех товаров из ЛК
 */
function showConfigModalForLk() {
  const html = HtmlService.createHtmlOutputFromFile('modal')
    .setWidth(600)
    .setHeight(600);
  html.append('<script>window.parseMode = "parseAllLk";</script>');
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройка: Спарсить все в ЛК');
}

/**
 * Обрабатывает конфигурацию из формы
 * @param {Object} data - Данные конфигурации
 * @return {Object} Результат обработки
 */
function processConfig(data) {
  try {
    if (!data.headerRow || isNaN(data.headerRow) || data.headerRow < 1) {
      return { success: false, message: 'Укажите корректный номер строки заголовков (число больше 0).' };
    }
    if (!data.walletPercent || data.walletPercent < 0 || data.walletPercent > 100) {
      return { success: false, message: 'Процент WB-кошелька должен быть числом от 0 до 100.' };
    }

    saveConfig(data);
    const testResult = testToken();
    if (!testResult.success) {
      return testResult;
    }

    let result;
    if (data.mode === 'parseAllLk') {
      result = parseAllFromLk(data.columns, data.headerRow, data.walletPercent);
    } else {
      result = parseArticles(data.columns, data.headerRow, data.walletPercent);
    }
    return { success: true, message: result.message };
  } catch (error) {
    return { success: false, message: `Ошибка: ${error.message}` };
  }
}

/**
 * Сохраняет конфигурацию
 * @param {Object} config - Конфигурация
 */
function saveConfig(config) {
  PropertiesService.getScriptProperties().setProperty('WB_CONFIG', JSON.stringify(config));
}

/**
 * Получает сохраненную конфигурацию
 * @return {Object|null} Конфигурация
 */
function getSavedConfig() {
  const configJson = PropertiesService.getScriptProperties().getProperty('WB_CONFIG');
  return configJson ? JSON.parse(configJson) : null;
}