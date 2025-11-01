/**
 * ============================================
 * МОДУЛЬ: РУЧНОЕ СОЗДАНИЕ ПЛИТОК ТЕГОВ
 * ============================================
 *
 * Функционал:
 * 1. Инициализация таблицы ключевых слов
 * 2. Валидация и проверка категорий
 * 3. Создание новых категорий для тегов
 * 4. Генерация HTML плиток на основе ручных данных
 * 5. Динамический расчет позиций секций
 */

// ============================================
// ИНИЦИАЛИЗАЦИЯ ТАБЛИЦЫ КЛЮЧЕВЫХ СЛОВ
// ============================================

/**
 * Создает таблицу для ручного ввода ключевых слов в детальном листе категории
 * Вызывается автоматически при открытии детального листа или вручную из меню
 */
function initializeTagKeywordsTable() {
  const sheet = SpreadsheetApp.getActiveSheet();

  // Проверяем, что это детальный лист категории
  const categoryId = sheet.getRange(DETAIL_SHEET_SECTIONS.CATEGORY_ID_CELL).getValue();
  if (!categoryId) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Это не детальный лист категории', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  console.log('[INFO] Инициализация таблицы ключевых слов для категории:', categoryId);

  const startRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_START;
  const headerRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_HEADER_ROW;

  // Заголовок секции
  sheet.getRange(startRow, 1, 1, 7)
    .merge()
    .setValue('📝 КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПЛИТОК ТЕГОВ (ручной ввод)')
    .setBackground('#E8F0FE')
    .setFontWeight('bold')
    .setFontSize(12);

  // Заголовки столбцов
  const headers = [
    '☑️',                    // A: Чекбокс
    'Ключевое слово',        // B
    'Тип плитки',            // C
    'Текст анкора',          // D
    'URL/ID категории',      // E
    'Статус категории',      // F
    'ID родителя (новая)'    // G
  ];

  const headerRange = sheet.getRange(headerRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#D0E1F9');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // Форматирование столбцов
  sheet.setColumnWidth(1, 40);   // Чекбокс
  sheet.setColumnWidth(2, 200);  // Ключевое слово
  sheet.setColumnWidth(3, 120);  // Тип плитки
  sheet.setColumnWidth(4, 250);  // Текст анкора
  sheet.setColumnWidth(5, 150);  // URL/ID категории
  sheet.setColumnWidth(6, 150);  // Статус категории
  sheet.setColumnWidth(7, 120);  // ID родителя

  // Добавляем валидацию для столбца "Тип плитки"
  const dataStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;
  const tileTypeColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.TILE_TYPE;

  const tileTypeRange = sheet.getRange(dataStartRow, tileTypeColumn, 50, 1); // 50 строк для ввода
  const tileTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Верхняя', 'Нижняя'], true)
    .setAllowInvalid(false)
    .build();
  tileTypeRange.setDataValidation(tileTypeRule);

  // Добавляем чекбоксы в первый столбец
  const checkboxColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.CHECKBOX;
  const checkboxRange = sheet.getRange(dataStartRow, checkboxColumn, 50, 1);
  checkboxRange.insertCheckboxes();

  // Добавляем формулы для автоматической проверки статуса категории (столбец F)
  const statusColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.CATEGORY_STATUS;
  const categoryLinkColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.CATEGORY_LINK;

  for (let i = 0; i < 50; i++) {
    const row = dataStartRow + i;
    const linkCell = columnToLetter(categoryLinkColumn) + row;

    // Формула: если ячейка пустая - "Не указана", иначе будет обновлено скриптом
    const formula = `=IF(ISBLANK(${linkCell}), "Не указана", "Проверить")`;
    sheet.getRange(row, statusColumn).setFormula(formula);
  }

  console.log('[SUCCESS] ✅ Таблица ключевых слов инициализирована');
  SpreadsheetApp.getActiveSpreadsheet().toast('Таблица готова к заполнению', '✅ Успех', 3);
}

/**
 * Вспомогательная функция: конвертирует номер столбца в букву (1 → A, 2 → B, и т.д.)
 */
function columnToLetter(column) {
  let temp;
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// ============================================
// РАСЧЕТ ДИНАМИЧЕСКИХ ПОЗИЦИЙ
// ============================================

/**
 * Рассчитывает актуальные позиции всех секций в детальном листе
 * @param {Sheet} sheet - Лист Google Sheets
 * @returns {Object} Объект с позициями секций
 */
function calculateSheetSections(sheet) {
  // НОВАЯ НАДЁЖНАЯ СИСТЕМА: Ищем секции по заголовкам вместо расчёта отступов
  const maxRow = Math.min(sheet.getLastRow(), 1000);

  let keywordsStart = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;
  let keywordsEnd = keywordsStart - 1;
  let upperTileStart = null;
  let lowerTileStart = null;
  let statsStart = null;
  let productsStart = null;

  // Сканируем лист и ищем заголовки секций
  console.log(`[DEBUG calculateSections] Начинаем сканирование ${maxRow} строк...`);

  for (let row = 1; row <= maxRow; row++) {
    const cellValue = sheet.getRange(row, 1).getValue();
    if (!cellValue) continue;

    const text = cellValue.toString();

    // Ищем заголовок "КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПЛИТОК"
    if (text.includes('КЛЮЧЕВЫЕ СЛОВА') && text.includes('ПЛИТК')) {
      keywordsStart = row + 2; // +2 = заголовок + пустая строка + заголовки столбцов = данные
      console.log(`[DEBUG calculateSections] Строка ${row}: Найдены КЛЮЧЕВЫЕ СЛОВА`);
    }

    // Ищем заголовок "ВЕРХНЯЯ ПЛИТКА"
    if (text.includes('ВЕРХН') && text.includes('ПЛИТК')) {
      upperTileStart = row;
      console.log(`[DEBUG calculateSections] Строка ${row}: Найдена ВЕРХНЯЯ ПЛИТКА`);
    }

    // Ищем заголовок "НИЖНЯЯ ПЛИТКА"
    if (text.includes('НИЖН') && text.includes('ПЛИТК')) {
      lowerTileStart = row;
      console.log(`[DEBUG calculateSections] Строка ${row}: Найдена НИЖНЯЯ ПЛИТКА`);
    }

    // Ищем заголовок "СТАТИСТИКА ТОВАРОВ"
    if (text.includes('СТАТИСТИКА') && text.includes('ТОВАР')) {
      statsStart = row;
      console.log(`[DEBUG calculateSections] Строка ${row}: Найдена СТАТИСТИКА ТОВАРОВ`);
    }

    // Ищем заголовок "ТОВАРЫ" или "ТЕКУЩИЕ ТОВАРЫ"
    if ((text.includes('ТОВАР') || text.includes('PRODUCT')) && !text.includes('ПЛИТК') && !text.includes('СТАТИСТИКА')) {
      productsStart = row + 2; // +2 = заголовок + заголовки столбцов = данные
      console.log(`[DEBUG calculateSections] Строка ${row}: Найден заголовок ТОВАРЫ (текст: "${text}"), данные с ${productsStart}`);
      break; // Товары - последняя секция
    }
  }

  // Находим последнюю заполненную строку в таблице ключевых слов
  const keywordColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.KEYWORD;
  for (let row = keywordsStart; row < keywordsStart + 100; row++) {
    const value = sheet.getRange(row, keywordColumn).getValue();
    if (value && value.toString().trim() !== '') {
      keywordsEnd = row;
    } else if (row > keywordsEnd + 5) {
      break;
    }
  }

  const keywordsCount = keywordsEnd - keywordsStart + 1;

  // Fallback значения, если заголовки не найдены
  if (upperTileStart === null) {
    upperTileStart = keywordsEnd + 3;
  }
  if (lowerTileStart === null) {
    lowerTileStart = upperTileStart + 12;
  }
  if (productsStart === null) {
    productsStart = lowerTileStart + 37;
  }

  console.log(`[DEBUG] Рассчитанные позиции секций:`, {
    keywordsStart,
    keywordsEnd,
    keywordsCount,
    upperTileStart,
    lowerTileStart,
    statsStart,
    productsStart
  });

  return {
    keywordsStart: keywordsStart,
    keywordsEnd: keywordsEnd,
    keywordsCount: Math.max(0, keywordsCount),
    upperTileStart: upperTileStart,
    lowerTileStart: lowerTileStart,
    statsStart: statsStart,
    productsStart: productsStart
  };
}

/**
 * Получает начальную строку списка товаров (динамически)
 */
function getProductsStartRow(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSheet();
  }
  const sections = calculateSheetSections(sheet);
  return sections.productsStart;
}

// ============================================
// ВАЛИДАЦИЯ И ПРОВЕРКА КАТЕГОРИЙ
// ============================================

/**
 * Проверяет все ключевые слова и обновляет статусы категорий
 */
function validateTagKeywords() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const categoryId = sheet.getRange(DETAIL_SHEET_SECTIONS.CATEGORY_ID_CELL).getValue();

  if (!categoryId) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Это не детальный лист категории', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Проверка категорий...', '⏳ Обработка', -1);
  console.log('[INFO] Начало валидации ключевых слов');

  const sections = calculateSheetSections(sheet);
  const dataStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;
  const dataEndRow = sections.keywordsEnd;

  if (dataEndRow < dataStartRow) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Нет данных для проверки', '⚠️ Внимание', 3);
    return;
  }

  const cols = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS;
  const rowCount = dataEndRow - dataStartRow + 1;

  // Читаем данные
  const dataRange = sheet.getRange(dataStartRow, 1, rowCount, 7);
  const data = dataRange.getValues();

  let checkedCount = 0;
  let existingCount = 0;
  let toCreateCount = 0;
  let emptyCount = 0;

  // Проверяем каждую строку
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const keyword = row[cols.KEYWORD - 1];
    const categoryLink = row[cols.CATEGORY_LINK - 1];

    if (!keyword || keyword.toString().trim() === '') {
      continue; // Пропускаем пустые строки
    }

    checkedCount++;
    let status = 'Не указана';

    if (categoryLink && categoryLink.toString().trim() !== '') {
      const linkStr = categoryLink.toString().trim();

      // Проверяем, это ID или URL
      if (/^\d+$/.test(linkStr)) {
        // Это ID категории - проверяем существование
        const categoryExists = checkCategoryExists(parseInt(linkStr));
        status = categoryExists ? '✅ Существует' : '❌ Не найдена';
        if (categoryExists) existingCount++;
      } else if (linkStr.includes('http')) {
        // Это URL - извлекаем handle и проверяем
        status = '🔗 URL указан';
        existingCount++;
      } else {
        status = '⚠️ Неверный формат';
      }
    } else {
      status = '➕ Создать новую';
      toCreateCount++;
    }

    // Обновляем статус в таблице
    sheet.getRange(dataStartRow + i, cols.CATEGORY_STATUS).setValue(status);
  }

  console.log(`[SUCCESS] Проверено строк: ${checkedCount}, Существующих: ${existingCount}, Создать: ${toCreateCount}`);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Проверено: ${checkedCount} | Существуют: ${existingCount} | Создать: ${toCreateCount}`,
    '✅ Валидация завершена',
    5
  );
}

/**
 * Проверяет, существует ли категория с данным ID
 */
function checkCategoryExists(categoryId) {
  try {
    const config = getInsalesConfig();
    const url = `${config.baseUrl}/admin/collections/${categoryId}.json`;

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(config.apiKey + ':' + config.password),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    return response.getResponseCode() === 200;
  } catch (error) {
    console.log(`[WARNING] Ошибка проверки категории ${categoryId}:`, error.message);
    return false;
  }
}

// ============================================
// СОЗДАНИЕ НОВЫХ КАТЕГОРИЙ ДЛЯ ТЕГОВ
// ============================================

/**
 * Создает новые категории для всех тегов со статусом "Создать новую"
 */
function createCategoriesForTags() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const parentCategoryId = sheet.getRange(DETAIL_SHEET_SECTIONS.CATEGORY_ID_CELL).getValue();

  if (!parentCategoryId) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Не удалось получить ID текущей категории', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Создание категорий...', '⏳ Обработка', -1);
  console.log('[INFO] Начало создания категорий для тегов');

  const sections = calculateSheetSections(sheet);
  const dataStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;
  const dataEndRow = sections.keywordsEnd;

  if (dataEndRow < dataStartRow) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Нет данных для создания', '⚠️ Внимание', 3);
    return;
  }

  const cols = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS;
  const rowCount = dataEndRow - dataStartRow + 1;
  const dataRange = sheet.getRange(dataStartRow, 1, rowCount, 7);
  const data = dataRange.getValues();

  let createdCount = 0;
  let errorCount = 0;

  // Создаем категории для каждой строки со статусом "Создать новую"
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const keyword = row[cols.KEYWORD - 1];
    const anchorText = row[cols.ANCHOR_TEXT - 1];
    const categoryLink = row[cols.CATEGORY_LINK - 1];
    const status = row[cols.CATEGORY_STATUS - 1];
    const customParentId = row[cols.PARENT_CATEGORY - 1];

    if (!keyword || keyword.toString().trim() === '') {
      continue;
    }

    // Создаем только для строк со статусом "Создать новую" или пустым URL
    if (status && status.toString().includes('Создать') || !categoryLink) {
      const categoryTitle = anchorText || keyword;
      const parentId = customParentId || parentCategoryId;

      console.log(`[INFO] Создаю категорию: "${categoryTitle}" с parent_id: ${parentId}`);

      try {
        const newCategory = createCategoryForTag(categoryTitle, parentId);

        if (newCategory && newCategory.id) {
          // Обновляем таблицу
          sheet.getRange(dataStartRow + i, cols.CATEGORY_LINK).setValue(newCategory.id);
          sheet.getRange(dataStartRow + i, cols.CATEGORY_STATUS).setValue('✅ Создана');
          createdCount++;

          console.log(`[SUCCESS] ✅ Категория создана: ID ${newCategory.id}`);
        } else {
          sheet.getRange(dataStartRow + i, cols.CATEGORY_STATUS).setValue('❌ Ошибка');
          errorCount++;
        }
      } catch (error) {
        console.log(`[ERROR] Ошибка создания категории:`, error.message);
        sheet.getRange(dataStartRow + i, cols.CATEGORY_STATUS).setValue('❌ ' + error.message);
        errorCount++;
      }

      // Задержка между запросами
      Utilities.sleep(500);
    }
  }

  console.log(`[SUCCESS] Создано категорий: ${createdCount}, Ошибок: ${errorCount}`);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Создано: ${createdCount} | Ошибок: ${errorCount}`,
    '✅ Создание завершено',
    5
  );
}

/**
 * Создает новую категорию через InSales API
 * @param {string} title - Название категории
 * @param {number} parentId - ID родительской категории
 * @returns {Object} Данные созданной категории
 */
function createCategoryForTag(title, parentId) {
  const config = getInsalesConfig();
  const url = `${config.baseUrl}/admin/collections.json`;

  // Генерируем handle из названия
  const handle = transliterate(title)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');

  const payload = {
    collection: {
      title: title,
      parent_id: parseInt(parentId) || 9069711, // По умолчанию корневая категория
      handle: handle,
      is_hidden: false,
      position: 999
    }
  };

  console.log('[DEBUG] Payload для создания категории:', JSON.stringify(payload, null, 2));

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(config.apiKey + ':' + config.password),
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode === 201 || responseCode === 200) {
    return JSON.parse(responseText);
  } else {
    throw new Error(`HTTP ${responseCode}: ${responseText}`);
  }
}

// ============================================
// ГЕНЕРАЦИЯ ПЛИТОК НА ОСНОВЕ РУЧНЫХ ДАННЫХ
// ============================================

/**
 * ГЛАВНАЯ ФУНКЦИЯ: Генерирует плитки тегов на основе ручных данных
 */
function generateTilesFromManualData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const categoryId = sheet.getRange(DETAIL_SHEET_SECTIONS.CATEGORY_ID_CELL).getValue();
  const categoryTitle = sheet.getRange(DETAIL_SHEET_SECTIONS.TITLE_CELL).getValue();

  if (!categoryId) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Это не детальный лист категории', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Генерация плиток...', '⏳ Обработка', -1);
  console.log('[INFO] Начало генерации плиток из ручных данных');

  // 1. Читаем данные из таблицы
  const tilesData = readManualTilesData(sheet);

  if (tilesData.upper.length === 0 && tilesData.lower.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Отметьте чекбоксами ключевые слова для генерации',
      '⚠️ Нет данных',
      5
    );
    return;
  }

  console.log(`[INFO] Найдено: ${tilesData.upper.length} для верхней плитки, ${tilesData.lower.length} для нижней`);

  // 2. Генерируем HTML
  const htmlResult = generateManualTilesHTML(tilesData, categoryTitle);

  // 3. Сохраняем в лист
  saveGeneratedTilesToSheet(sheet, htmlResult, tilesData);

  console.log('[SUCCESS] ✅ Плитки успешно сгенерированы и сохранены');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Верхняя: ${tilesData.upper.length} анкоров | Нижняя: ${tilesData.lower.length} анкоров`,
    '✅ Плитки созданы',
    5
  );
}

/**
 * Читает данные из таблицы ключевых слов
 */
function readManualTilesData(sheet) {
  const sections = calculateSheetSections(sheet);
  const dataStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;
  const dataEndRow = sections.keywordsEnd;

  const cols = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS;
  const rowCount = dataEndRow - dataStartRow + 1;

  if (rowCount <= 0) {
    return { upper: [], lower: [] };
  }

  const dataRange = sheet.getRange(dataStartRow, 1, rowCount, 7);
  const data = dataRange.getValues();

  const upperTile = [];
  const lowerTile = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const checked = row[cols.CHECKBOX - 1];
    const keyword = row[cols.KEYWORD - 1];
    const tileType = row[cols.TILE_TYPE - 1];
    const anchorText = row[cols.ANCHOR_TEXT - 1];
    const categoryLink = row[cols.CATEGORY_LINK - 1];
    const status = row[cols.CATEGORY_STATUS - 1];

    // Пропускаем неотмеченные или пустые строки
    if (!checked || !keyword || keyword.toString().trim() === '') {
      continue;
    }

    // Пропускаем строки без категории
    if (!categoryLink || status.includes('Не указана') || status.includes('Ошибка')) {
      console.log(`[WARNING] Пропущена строка ${i + 1}: категория не указана или ошибка`);
      continue;
    }

    const anchor = {
      keyword: keyword.toString().trim(),
      anchor: anchorText ? anchorText.toString().trim() : keyword.toString().trim(),
      link: buildCategoryLink(categoryLink),
      category_id: extractCategoryId(categoryLink)
    };

    // Распределяем по плиткам
    if (tileType && tileType.toString().includes('Верхняя')) {
      upperTile.push(anchor);
    } else if (tileType && tileType.toString().includes('Нижняя')) {
      lowerTile.push(anchor);
    }
  }

  return { upper: upperTile, lower: lowerTile };
}

/**
 * Извлекает ID категории из URL или возвращает ID напрямую
 */
function extractCategoryId(categoryLink) {
  const linkStr = categoryLink.toString().trim();

  if (/^\d+$/.test(linkStr)) {
    return parseInt(linkStr);
  }

  // Пытаемся извлечь ID из URL
  const match = linkStr.match(/\/collection\/([^\/]+)/);
  if (match) {
    return match[1]; // Возвращаем handle
  }

  return linkStr;
}

/**
 * Создает URL категории из ID или handle
 */
function buildCategoryLink(categoryLink) {
  const linkStr = categoryLink.toString().trim();

  // Если это уже полный URL
  if (linkStr.includes('http')) {
    return linkStr;
  }

  // Если это ID - получаем handle из API
  if (/^\d+$/.test(linkStr)) {
    const handle = getCategoryHandle(parseInt(linkStr));
    return handle ? `/collection/${handle}` : '#';
  }

  // Если это handle
  if (linkStr.startsWith('/')) {
    return linkStr;
  }

  return `/collection/${linkStr}`;
}

/**
 * Получает handle категории по ID
 */
function getCategoryHandle(categoryId) {
  try {
    const config = getInsalesConfig();
    const url = `${config.baseUrl}/admin/collections/${categoryId}.json`;

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(config.apiKey + ':' + config.password),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return data.handle || data.url || data.id;
    }
  } catch (error) {
    console.log(`[WARNING] Не удалось получить handle для категории ${categoryId}`);
  }

  return null;
}

/**
 * Генерирует HTML для плиток на основе ручных данных
 */
function generateManualTilesHTML(tilesData, categoryTitle) {
  // Используем существующую функцию генерации HTML
  const result = {
    categoryName: categoryTitle,
    topTile: tilesData.upper,
    bottomTile: tilesData.lower
  };

  return generateTilesHTML(result);
}

/**
 * Показывает предпросмотр плитки на основе ручных данных
 */
function showTilesPreviewManual() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const categoryId = sheet.getRange(DETAIL_SHEET_SECTIONS.CATEGORY_ID_CELL).getValue();
  const categoryTitle = sheet.getRange(DETAIL_SHEET_SECTIONS.TITLE_CELL).getValue();

  if (!categoryId) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Это не детальный лист категории', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  console.log('[INFO] Предпросмотр плиток для категории:', categoryTitle);

  // Читаем данные
  const tilesData = readManualTilesData(sheet);

  if (tilesData.upper.length === 0 && tilesData.lower.length === 0) {
    SpreadsheetApp.getUi().alert(
      'Нет данных',
      'Отметьте чекбоксами ключевые слова для предпросмотра',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Генерируем HTML
  const htmlResult = generateManualTilesHTML(tilesData, categoryTitle);

  // Создаем HTML для предпросмотра
  const cssStyles = generateTilesCSS();

  const previewHTML = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background: #f5f5f5;
          }
          .preview-container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          }
          .preview-section {
            margin-bottom: 40px;
          }
          .preview-title {
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 15px;
            color: #333;
          }
          .stats {
            background: #e3f2fd;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
          }
          ${cssStyles}
        </style>
      </head>
      <body>
        <div class="preview-container">
          <h1>Предпросмотр плиток тегов</h1>
          <p><strong>Категория:</strong> ${categoryTitle}</p>

          <div class="stats">
            <strong>Статистика:</strong><br>
            Верхняя плитка: ${tilesData.upper.length} анкоров<br>
            Нижняя плитка: ${tilesData.lower.length} анкоров
          </div>

          <div class="preview-section">
            <div class="preview-title">🏷️ ВЕРХНЯЯ ПЛИТКА (Навигация)</div>
            ${htmlResult.topHTML || '<p style="color: #999;">Нет данных для верхней плитки</p>'}
          </div>

          <div class="preview-section">
            <div class="preview-title">🏷️ НИЖНЯЯ ПЛИТКА (SEO)</div>
            ${htmlResult.bottomHTML || '<p style="color: #999;">Нет данных для нижней плитки</p>'}
          </div>

          <button onclick="google.script.host.close()" style="
            padding: 10px 20px;
            background: #4285f4;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-size: 14px;
          ">Закрыть</button>
        </div>
      </body>
    </html>
  `;

  // Показываем диалог
  const htmlOutput = HtmlService.createHtmlOutput(previewHTML)
    .setWidth(1000)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '👁️ Предпросмотр плитки тегов');
}

/**
 * Сохраняет сгенерированные плитки в лист (в колонки СТАЛО: E-H)
 * Находит существующие блоки плиток и заполняет правую часть таблицы
 * Читает чекбоксы и объединяет старые + новые теги в финальный HTML
 */
function saveGeneratedTilesToSheet(sheet, htmlResult, tilesData) {
  console.log('[INFO] Сохранение сгенерированных плиток в формате БЫЛО/СТАЛО');

  // Ищем блоки плиток в листе по заголовкам
  const upperTileRow = findRowByText(sheet, '🏷️ ПЛИТКА ТЕГОВ - ВЕРХНЯЯ');
  const lowerTileRow = findRowByText(sheet, '🏷️ ПЛИТКА ТЕГОВ - НИЖНЯЯ');

  if (!upperTileRow || !lowerTileRow) {
    throw new Error('Не найдены блоки плиток в детальном листе. Пересоздайте лист категории.');
  }

  console.log(`[INFO] Найдены блоки: Верхняя плитка - строка ${upperTileRow}, Нижняя плитка - строка ${lowerTileRow}`);

  // === ВЕРХНЯЯ ПЛИТКА: сохраняем в колонки E-H (СТАЛО) ===
  const upperDataStartRow = upperTileRow + 4; // +4 = заголовок + инструкция + пустая + заголовки столбцов

  // ДИНАМИЧЕСКИЙ РАСЧЕТ: находим сколько строк уже есть в таблице
  const upperExistingRows = sheet.getRange(upperDataStartRow, 1, 50, 1).getValues().filter(row => row[0] !== '').length;
  const upperTotalRows = Math.max(upperExistingRows, tilesData.upper.length, 3); // Минимум 3 строки

  // Очищаем старые данные в колонках E-H (СТАЛО)
  sheet.getRange(upperDataStartRow, 5, upperTotalRows, 4).clearContent();

  // Записываем новые данные
  for (let i = 0; i < Math.min(tilesData.upper.length, upperTotalRows); i++) {
    const anchor = tilesData.upper[i];
    const row = upperDataStartRow + i;

    sheet.getRange(row, 5, 1, 4).setValues([[
      anchor.anchor,                // E: СТАЛО - Текст анкора
      anchor.link,                  // F: СТАЛО - URL
      anchor.category_id || '',     // G: СТАЛО - ID категории
      'Сгенерировано'               // H: СТАЛО - Примечание
    ]]);
  }

  // Читаем старые теги с чекбоксами (колонки A, B, C) - ДИНАМИЧЕСКИ
  const upperOldData = sheet.getRange(upperDataStartRow, 1, upperTotalRows, 3).getValues();
  const upperOldTagsChecked = upperOldData
    .filter(row => row[2] === true && row[0] && row[1]) // Чекбокс включен и есть данные
    .map(row => ({
      text: row[0],
      url: row[1],
      anchor: row[0],
      link: row[1]
    }));

  console.log(`[INFO] Верхняя плитка: ${upperOldTagsChecked.length} старых тегов включено, ${tilesData.upper.length} новых`);

  // === НИЖНЯЯ ПЛИТКА: сохраняем в колонки E-H (СТАЛО) ===
  const lowerDataStartRow = lowerTileRow + 4;

  // ДИНАМИЧЕСКИЙ РАСЧЕТ: находим сколько строк уже есть в таблице
  const lowerExistingRows = sheet.getRange(lowerDataStartRow, 1, 100, 1).getValues().filter(row => row[0] !== '').length;
  const lowerTotalRows = Math.max(lowerExistingRows, tilesData.lower.length, 5); // Минимум 5 строк

  // Очищаем старые данные в колонках E-H (СТАЛО)
  sheet.getRange(lowerDataStartRow, 5, lowerTotalRows, 4).clearContent();

  // Записываем новые данные
  for (let i = 0; i < Math.min(tilesData.lower.length, lowerTotalRows); i++) {
    const anchor = tilesData.lower[i];
    const row = lowerDataStartRow + i;

    sheet.getRange(row, 5, 1, 4).setValues([[
      anchor.anchor,                // E: СТАЛО - Текст анкора
      anchor.link,                  // F: СТАЛО - URL
      anchor.category_id || '',     // G: СТАЛО - ID категории
      'Сгенерировано'               // H: СТАЛО - Примечание
    ]]);
  }

  // Читаем старые теги с чекбоксами (колонки A, B, C) - ДИНАМИЧЕСКИ
  const lowerOldData = sheet.getRange(lowerDataStartRow, 1, lowerTotalRows, 3).getValues();
  const lowerOldTagsChecked = lowerOldData
    .filter(row => row[2] === true && row[0] && row[1]) // Чекбокс включен и есть данные
    .map(row => ({
      text: row[0],
      url: row[1],
      anchor: row[0],
      link: row[1]
    }));

  console.log(`[INFO] Нижняя плитка: ${lowerOldTagsChecked.length} старых тегов включено, ${tilesData.lower.length} новых`);

  // === ГЕНЕРИРУЕМ ФИНАЛЬНЫЙ HTML: Старые (отмеченные) + Новые ===
  const finalUpperAnchors = [...upperOldTagsChecked, ...tilesData.upper];
  const finalLowerAnchors = [...lowerOldTagsChecked, ...tilesData.lower];

  const finalUpperHTML = generateTopTileHTML(finalUpperAnchors);
  const finalLowerHTML = generateBottomTileHTML(finalLowerAnchors);

  console.log(`[INFO] Финальная верхняя плитка: ${finalUpperAnchors.length} тегов`);
  console.log(`[INFO] Финальная нижняя плитка: ${finalLowerAnchors.length} тегов`);

  // Добавляем финальный HTML код в конец каждого блока - ДИНАМИЧЕСКИ
  // Верхняя плитка - HTML в строке после данных
  const upperHTMLRow = upperDataStartRow + upperTotalRows + 1;
  console.log(`[DEBUG] Записываем HTML верхней плитки в строку ${upperHTMLRow} (после ${upperTotalRows} строк данных)`);

  // Обновляем существующее поле (оно было создано в setupDetailedCategorySheet)
  sheet.getRange(upperHTMLRow, 2, 1, 7)
    .merge()
    .setValue(finalUpperHTML)
    .setWrap(true)
    .setBackground('#c8e6c9');

  // Нижняя плитка - HTML в строке после данных
  const lowerHTMLRow = lowerDataStartRow + lowerTotalRows + 1;
  console.log(`[DEBUG] Записываем HTML нижней плитки в строку ${lowerHTMLRow} (после ${lowerTotalRows} строк данных)`);

  // Обновляем существующее поле (оно было создано в setupDetailedCategorySheet)
  sheet.getRange(lowerHTMLRow, 2, 1, 7)
    .merge()
    .setValue(finalLowerHTML)
    .setWrap(true)
    .setBackground('#c8e6c9');

  console.log('[SUCCESS] ✅ Финальные HTML коды добавлены (старые + новые теги)');
}

/**
 * Находит строку по тексту в первой колонке
 * @param {Sheet} sheet - Лист Google Sheets
 * @param {string} searchText - Текст для поиска
 * @returns {number|null} Номер строки или null
 */
function findRowByText(sheet, searchText) {
  const maxRows = sheet.getMaxRows();
  const searchRange = sheet.getRange(1, 1, maxRows, 1);
  const values = searchRange.getValues();

  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i][0];
    if (cellValue && cellValue.toString().includes(searchText)) {
      return i + 1; // +1 потому что массив начинается с 0, а строки с 1
    }
  }

  return null;
}
