function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("WB Артикулы")
    .addItem("Парсить по артикулам", "showConfigModalForArticles")
    .addItem("Спарсить все в ЛК", "showConfigModalForLk")
    .addToUi();
}

/**
 * Показывает модалку для парсинга по артикулам
 */
function showConfigModalForArticles() {
  const html = HtmlService.createHtmlOutputFromFile("modal")
    .setWidth(600)
    .setHeight(600);
  html.append('<script>window.parseMode = "parseArticles";</script>');
  SpreadsheetApp.getUi().showModalDialog(
    html,
    "Настройка: Парсить по артикулам"
  );
}

/**
 * Показывает модалку для парсинга всех товаров из ЛК
 */
function showConfigModalForLk() {
  const html = HtmlService.createHtmlOutputFromFile("modal")
    .setWidth(600)
    .setHeight(600);
  html.append('<script>window.parseMode = "parseAllLk";</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Настройка: Спарсить все в ЛК");
}

/**
 * Обрабатывает конфигурацию из формы
 * @param {Object} data - Данные конфигурации
 * @return {Object} Результат обработки
 */
function processConfig(data) {
  try {
    if (!data.token || data.token.trim() === "") {
      return { success: false, message: "Токен не может быть пустым." };
    }
    if (!data.headerRow || isNaN(data.headerRow) || data.headerRow < 1) {
      return {
        success: false,
        message: "Укажите корректный номер строки заголовков (число больше 0).",
      };
    }

    saveConfig(data);
    const testResult = testToken(data.token.trim());
    if (!testResult.success) {
      return { success: false, message: testResult.message };
    }

    let result;
    if (data.mode === "parseAllLk") {
      result = parseAllFromLk(data.columns, data.headerRow, data.token);
    } else {
      result = parseArticles(data.columns, data.headerRow);
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
  PropertiesService.getScriptProperties().setProperty(
    "WB_CONFIG",
    JSON.stringify(config)
  );
}

/**
 * Получает сохраненную конфигурацию
 * @return {Object|null} Конфигурация
 */
function getSavedConfig() {
  const configJson =
    PropertiesService.getScriptProperties().getProperty("WB_CONFIG");
  return configJson ? JSON.parse(configJson) : null;
}
