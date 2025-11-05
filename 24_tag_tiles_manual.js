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
 * УСТАРЕЛО: НЕ ИСПОЛЬЗУЙТЕ ЭТУ ФУНКЦИЮ!
 * Используйте addCategoryPickerColumn() вместо этого.
 *
 * Эта функция создавала 50 строк и затирала нижние блоки.
 * Оставлена для обратной совместимости.
 */
function initializeTagKeywordsTable() {
  SpreadsheetApp.getUi().alert(
    'Эта функция устарела',
    'Используйте "Добавить кнопки выбора категорий" из меню',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Добавляет колонку H "Выбор" с кнопками подбора категорий
 * к существующей таблице ключевых слов
 *
 * ВАЖНО: Таблица ключевых слов должна уже существовать!
 * (создается автоматически при создании детального листа)
 */
function addCategoryPickerColumn() {
  const sheet = SpreadsheetApp.getActiveSheet();

  // Проверяем, что это детальный лист категории
  const categoryId = sheet.getRange(DETAIL_SHEET_SECTIONS.CATEGORY_ID_CELL).getValue();
  if (!categoryId) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Это не детальный лист категории', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  console.log('[INFO] Добавление колонки выбора категорий для категории:', categoryId);

  const startRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_START;
  const headerRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_HEADER_ROW;
  const dataStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;

  // Проверяем, есть ли уже колонка H
  const existingHeader = sheet.getRange(headerRow, 8).getValue();
  if (existingHeader && existingHeader.includes('Выбор')) {
    SpreadsheetApp.getUi().alert(
      'Колонка уже существует',
      'Колонка "Выбор" уже добавлена к таблице',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Обновляем заголовок секции - расширяем до 8 колонок
  const sectionHeader = sheet.getRange(startRow, 1).getValue();
  sheet.getRange(startRow, 1, 1, 8)
    .merge()
    .setValue(sectionHeader || '📝 КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПЛИТОК ТЕГОВ (ручной ввод)')
    .setBackground('#E8F0FE')
    .setFontWeight('bold')
    .setFontSize(12);

  // Добавляем заголовок колонки H
  sheet.getRange(headerRow, 8)
    .setValue('Выбор')
    .setBackground('#D0E1F9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Форматирование колонки H
  sheet.setColumnWidth(8, 100);  // Выбор категории

  // Определяем количество строк в таблице
  // Ищем первую пустую строку или строку со следующей секцией
  let rowCount = 0;
  for (let i = 0; i < 100; i++) { // Максимум 100 строк
    const row = dataStartRow + i;
    const cellValue = sheet.getRange(row, 1).getValue();

    // Если встретили заголовок следующей секции - останавливаемся
    if (cellValue && cellValue.toString().includes('📊')) {
      break;
    }

    rowCount++;
  }

  console.log(`[INFO] Найдено ${rowCount} строк в таблице ключевых слов`);

  // Добавляем rich text ссылки в колонку H
  addCategoryPickerLinksToRange(sheet, dataStartRow, rowCount);

  console.log('[SUCCESS] ✅ Колонка выбора категорий добавлена');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Добавлена колонка "Выбор" с кнопками для ${rowCount} строк`,
    '✅ Успех',
    3
  );
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

/**
 * Добавляет rich text ссылки в колонку H для вызова диалога подбора категорий
 *
 * @param {Sheet} sheet - Лист для добавления ссылок
 * @param {number} dataStartRow - Строка начала данных
 * @param {number} rowCount - Количество строк для добавления ссылок
 */
function addCategoryPickerLinksToRange(sheet, dataStartRow, rowCount) {
  const pickerColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.PICKER_LINK;

  // Создаем ссылки для указанного количества строк
  for (let i = 0; i < rowCount; i++) {
    const row = dataStartRow + i;

    // Создаем rich text ссылку
    // ВАЖНО: URL "#gid=0" - это фиктивный URL, реальное действие обрабатывается через onSelectionChange
    const richText = SpreadsheetApp.newRichTextValue()
      .setText('🔍 Выбрать')
      .setLinkUrl('#gid=0')  // Фиктивная ссылка для визуального эффекта
      .build();

    const cell = sheet.getRange(row, pickerColumn);
    cell.setRichTextValue(richText);
    cell.setHorizontalAlignment('center');
    cell.setBackground('#F8F9FA');  // Светло-серый фон для визуального отличия
  }

  console.log(`[INFO] ✅ Добавлено ${rowCount} ссылок для подбора категорий в колонку H`);
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
  console.log('[calculateSheetSections] ========== НАЧАЛО РАСЧЁТА ПОЗИЦИЙ ==========');

  const markers = DETAIL_SHEET_SECTION_MARKERS;

  // === ШАГ 1: Находим все заголовки секций по emoji-маркерам ===
  console.log('[calculateSheetSections] Шаг 1: Поиск заголовков секций...');

  const keywordsRow = findSectionByMarker(sheet, markers.KEYWORDS_HEADER) || markers.KEYWORDS_TABLE_START;
  const statsRow = findSectionByMarker(sheet, markers.STATS_HEADER);
  const productsRow = findSectionByMarker(sheet, markers.PRODUCTS_HEADER);
  const fieldsRow = findSectionByMarker(sheet, markers.FIELDS_HEADER);
  const upperTileRow = findSectionByMarker(sheet, markers.UPPER_TILE_HEADER);
  const lowerTileRow = findSectionByMarker(sheet, markers.LOWER_TILE_HEADER);

  // === ШАГ 2: Рассчитываем начало данных (заголовок + offset) ===
  console.log('[calculateSheetSections] Шаг 2: Расчёт начала данных...');

  const keywordsDataStart = keywordsRow + markers.HEADER_TO_DATA_OFFSET;
  const productsDataStart = productsRow ? productsRow + markers.HEADER_TO_DATA_OFFSET : null;

  // === ШАГ 3: Находим конец таблицы ключевых слов (динамический поиск) ===
  console.log('[calculateSheetSections] Шаг 3: Поиск конца таблицы ключевых слов...');

  let keywordsEnd = keywordsDataStart - 1;
  const keywordColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.KEYWORD;

  // Сканируем до statsRow (если найдена) или до 100 строк
  const searchLimit = statsRow ? Math.min(statsRow, keywordsDataStart + 100) : keywordsDataStart + 100;

  for (let row = keywordsDataStart; row < searchLimit; row++) {
    const value = sheet.getRange(row, keywordColumn).getValue();
    if (value && value.toString().trim() !== '') {
      keywordsEnd = row;
    } else if (row > keywordsEnd + 3) {
      // 3+ пустые строки подряд = конец секции
      break;
    }
  }

  const keywordsCount = Math.max(0, keywordsEnd - keywordsDataStart + 1);
  console.log(`[calculateSheetSections] Таблица ключевых слов: строки ${keywordsDataStart}-${keywordsEnd} (${keywordsCount} записей)`);

  // === ШАГ 4: Читаем количество товаров из ARCHITECTURAL CONSTANT ===
  console.log('[calculateSheetSections] Шаг 4: Чтение количества товаров...');

  let productsCount = 0;
  if (statsRow) {
    try {
      // Количество товаров находится в ячейке B{statsRow + 1}
      const productsCountCell = `B${statsRow + 1}`;
      const productsValue = sheet.getRange(productsCountCell).getValue();
      productsCount = parseInt(productsValue) || 0;
      console.log(`[calculateSheetSections] Количество товаров из ${productsCountCell}: ${productsCount}`);
    } catch (error) {
      console.warn(`[calculateSheetSections] ⚠️ Не удалось прочитать количество товаров:`, error.message);
    }
  }

  // === ШАГ 5: Находим HTML-поля плиток по под-маркеру ===
  console.log('[calculateSheetSections] Шаг 5: Поиск HTML-полей плиток...');

  let upperHTMLRow = null;
  let lowerHTMLRow = null;

  // Ищем HTML верхней плитки ТОЛЬКО между upperTileRow и lowerTileRow
  if (upperTileRow && lowerTileRow) {
    upperHTMLRow = findSectionByMarker(sheet, markers.HTML_CODE_MARKER, upperTileRow, lowerTileRow);
    console.log(`[calculateSheetSections] Верхняя HTML: строка ${upperHTMLRow || 'не найдена'}`);
  }

  // Ищем HTML нижней плитки ТОЛЬКО после lowerTileRow
  if (lowerTileRow) {
    lowerHTMLRow = findSectionByMarker(sheet, markers.HTML_CODE_MARKER, lowerTileRow, lowerTileRow + 100);
    console.log(`[calculateSheetSections] Нижняя HTML: строка ${lowerHTMLRow || 'не найдена'}`);
  }

  // === РЕЗУЛЬТАТ: Объект с позициями всех секций ===
  const sections = {
    // Заголовки секций
    keywordsRow,
    statsRow,
    productsRow,
    fieldsRow,
    upperTileRow,
    lowerTileRow,

    // Данные секций
    keywordsStart: keywordsDataStart,
    keywordsEnd,
    keywordsCount,

    productsStart: productsDataStart,
    productsCount,

    // HTML-поля
    upperHTMLRow,
    lowerHTMLRow,

    // Метаданные
    calculatedAt: new Date().toISOString(),
    method: 'marker-based-v2'
  };

  console.log('[calculateSheetSections] ========== РЕЗУЛЬТАТ ==========');
  console.log(JSON.stringify(sections, null, 2));
  console.log('[calculateSheetSections] ========== КОНЕЦ РАСЧЁТА ==========');

  return sections;
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

  // ДИНАМИЧЕСКИЙ РАСЧЕТ: находим сколько строк данных в блоке верхней плитки
  // Читаем до тех пор, пока не встретим "HTML код (финальный)" или начало нижней плитки
  let upperExistingRows = 0;
  for (let row = upperDataStartRow; row < lowerTileRow; row++) {
    const cellValue = sheet.getRange(row, 1).getValue();
    const cellText = cellValue ? cellValue.toString() : '';

    // Если встретили "HTML код" - это конец данных
    if (cellText.includes('HTML код')) {
      break;
    }

    // Если ячейка не пустая - увеличиваем счётчик
    if (cellText.trim() !== '') {
      upperExistingRows = row - upperDataStartRow + 1;
    }
  }

  const upperTotalRows = Math.max(upperExistingRows, tilesData.upper.length, 3); // Минимум 3 строки
  console.log(`[DEBUG] Верхняя плитка: найдено ${upperExistingRows} заполненных строк, итого ${upperTotalRows}`);

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

  // ДИНАМИЧЕСКИЙ РАСЧЕТ: находим сколько строк данных в блоке нижней плитки
  // Читаем до тех пор, пока не встретим "HTML код (финальный)"
  let lowerExistingRows = 0;
  for (let row = lowerDataStartRow; row < lowerDataStartRow + 100; row++) {
    const cellValue = sheet.getRange(row, 1).getValue();
    const cellText = cellValue ? cellValue.toString() : '';

    // Если встретили "HTML код" - это конец данных
    if (cellText.includes('HTML код')) {
      break;
    }

    // Если ячейка не пустая - увеличиваем счётчик
    if (cellText.trim() !== '') {
      lowerExistingRows = row - lowerDataStartRow + 1;
    }
  }

  const lowerTotalRows = Math.max(lowerExistingRows, tilesData.lower.length, 5); // Минимум 5 строк
  console.log(`[DEBUG] Нижняя плитка: найдено ${lowerExistingRows} заполненных строк, итого ${lowerTotalRows}`);

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

  // Добавляем финальный HTML код - ИЩЕМ СУЩЕСТВУЮЩИЕ ПОЛЯ ПО ТЕКСТУ
  // Верхняя плитка - ищем ПЕРВОЕ "HTML код (финальный)" после заголовка верхней плитки, но ДО нижней
  let upperHTMLRow = null;
  for (let row = upperTileRow; row < lowerTileRow; row++) {
    const cellValue = sheet.getRange(row, 1).getValue();
    if (cellValue && cellValue.toString().includes('HTML код') && cellValue.toString().includes('финальный')) {
      upperHTMLRow = row;
      break;
    }
  }

  if (upperHTMLRow) {
    console.log(`[DEBUG] Записываем HTML верхней плитки в строку ${upperHTMLRow} (найдено поле по тексту)`);
    sheet.getRange(upperHTMLRow, 2, 1, 7)
      .merge()
      .setValue(finalUpperHTML)
      .setWrap(true)
      .setBackground('#c8e6c9');
  } else {
    console.log('[WARNING] Поле для HTML верхней плитки не найдено');
  }

  // Нижняя плитка - ищем "HTML код (финальный)" после заголовка нижней плитки
  let lowerHTMLRow = null;
  for (let row = lowerTileRow; row < lowerTileRow + 100; row++) {
    const cellValue = sheet.getRange(row, 1).getValue();
    if (cellValue && cellValue.toString().includes('HTML код') && cellValue.toString().includes('финальный')) {
      lowerHTMLRow = row;
      break;
    }
  }

  if (lowerHTMLRow) {
    console.log(`[DEBUG] Записываем HTML нижней плитки в строку ${lowerHTMLRow} (найдено поле по тексту)`);
    sheet.getRange(lowerHTMLRow, 2, 1, 7)
      .merge()
      .setValue(finalLowerHTML)
      .setWrap(true)
      .setBackground('#c8e6c9');
  } else {
    console.log('[WARNING] Поле для HTML нижней плитки не найдено');
  }

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
