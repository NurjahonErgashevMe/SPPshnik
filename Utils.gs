/**
 * Преобразует индекс столбца в буквенное обозначение
 * @param {number} columnIndex - Индекс столбца (начиная с 1)
 * @return {string} Буквенное обозначение (A, B, C, ...)
 */
function columnToLetter(columnIndex) {
  let temp,
    letter = "";
  while (columnIndex > 0) {
    temp = (columnIndex - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnIndex = (columnIndex - temp - 1) / 26;
  }
  return letter || "A";
}

/**
 * Преобразует букву колонки в индекс
 * @param {string} letter - Буква колонки (A, B, ..., AA, AB)
 * @return {number} Индекс колонки (начиная с 1)
 */
function letterToColumn(letter) {
  let column = 0;
  const letters = letter.toUpperCase();

  for (let i = 0; i < letters.length; i++) {
    const charCode = letters.charCodeAt(i) - 64;
    column = column * 26 + charCode;
  }

  return column;
}

/**
 * Извлекает валидные артикулы из массива значений
 * @param {Array} values - Массив значений столбца
 * @return {number[]} Отфильтрованные артикулы
 */
function extractArticles(values) {
  return values
    .map((row, index) => {
      const value = row[0];
      return extractArticle(value);
    })
    .filter((article) => article !== null);
}

/**
 * Извлекает артикул из одного значения
 * @param {*} value - Значение ячейки
 * @return {number|null} Артикул или null
 */
function extractArticle(value) {
  if (typeof value === "string") {
    const cleaned = value.replace(/\D/g, "");
    return cleaned ? parseInt(cleaned, 10) : null;
  }
  return typeof value === "number" ? value : null;
}

/**
 * Разбивает массив на чанки
 * @param {Array} array - Массив для разделения
 * @param {number} chunkSize - Размер чанка
 * @return {Array[]} Массив чанков
 */
function chunkArray(array, chunkSize) {
  const chunks = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    chunks.push(array.slice(i, i + chunkSize));
  }
  return chunks;
}

/**
 * Форматирует значение для записи в таблицу
 * @param {*} value - Значение
 * @return {string|number} Форматированное значение
 */
function formatValue(value) {
  return value === null || value === undefined || value === 0 ? "-" : value;
}
