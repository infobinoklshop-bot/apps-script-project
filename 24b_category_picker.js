/**
 * ============================================
 * МОДУЛЬ: ПОДБОР КАТЕГОРИЙ ДЛЯ ПЛИТОК ТЕГОВ
 * ============================================
 *
 * Диалог для удобного выбора категорий при создании плиток тегов.
 * Используется в таблице ключевых слов (строка 20+) для заполнения колонки E.
 *
 * Функционал:
 * - Поиск категорий в реальном времени
 * - Отображение иерархии с визуальными отступами
 * - Метаданные категорий (количество товаров, наличие)
 * - Возврат ID выбранной категории в ячейку таблицы
 */

// ============================================
// ПОКАЗ ДИАЛОГА ПОДБОРА КАТЕГОРИЙ
// ============================================

/**
 * Показывает диалог подбора категории для текущей выбранной ячейки
 *
 * Вызывается из меню "Плитки тегов → 🔍 Выбрать категорию для текущей ячейки"
 */
function showCategoryPickerForCurrentCell() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const row = range.getRow();
  const sheetName = sheet.getName();

  // Проверяем, что это детальный лист категории
  if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
    SpreadsheetApp.getUi().alert('Ошибка',
      'Эту функцию можно использовать только в детальном листе категории',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Проверяем, что выбрана строка в диапазоне таблицы ключевых слов
  const dataStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;
  if (row < dataStartRow) {
    SpreadsheetApp.getUi().alert('Выберите строку',
      `Выберите ячейку в строке ${dataStartRow} или ниже (в таблице ключевых слов)`,
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Открываем диалог для выбранной строки
  showCategoryPickerForTagTiles(row, sheetName);
}

/**
 * Показывает диалог подбора категории для плитки тега
 *
 * Вызывается из:
 * - onSelectionChange триггера (клик на ссылку в колонке H) - НЕ РАБОТАЕТ из-за ограничений
 * - showCategoryPickerForCurrentCell() (через меню)
 *
 * @param {number} targetRow - Номер строки в таблице ключевых слов
 * @param {string} sheetName - Название листа категории
 */
function showCategoryPickerForTagTiles(targetRow, sheetName) {
  const context = 'showCategoryPickerForTagTiles';

  try {
    logInfo(`📌 Открытие диалога подбора категории для строки ${targetRow}`, null, context);

    // Проверяем, что sheetName передан, иначе берем активный лист
    if (!sheetName) {
      const activeSheet = SpreadsheetApp.getActiveSheet();
      sheetName = activeSheet.getName();
    }

    // Проверяем, что это детальный лист категории
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      SpreadsheetApp.getUi().alert('Ошибка',
        'Диалог можно открыть только из детального листа категории',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Создаем HTML диалог из файла
    const htmlTemplate = HtmlService.createTemplateFromFile('24b_category_picker_dialog');

    // Передаем параметры в HTML шаблон
    htmlTemplate.targetRow = targetRow;
    htmlTemplate.sheetName = sheetName;

    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(600)
      .setHeight(650);

    // Показываем модальный диалог
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔍 Выбрать категорию для плитки');

    logInfo(`✅ Диалог подбора категории открыт для строки ${targetRow}`, null, context);

  } catch (error) {
    logError(`❌ Ошибка открытия диалога подбора категории`, error, context);

    // Показываем ошибку пользователю
    SpreadsheetApp.getUi().alert('Ошибка',
      `Не удалось открыть диалог подбора: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================
// СЕРВЕРНЫЕ ФУНКЦИИ ДЛЯ ДИАЛОГА
// ============================================

/**
 * Возвращает список категорий для диалога
 *
 * Переиспользует существующую функцию getCategoriesForSearch()
 * из модуля 13_categories_search.js
 *
 * @returns {Array} Массив объектов категорий
 */
function getCategoriesForPicker() {
  const context = 'getCategoriesForPicker';

  try {
    logInfo('📋 Загрузка категорий для диалога подбора', null, context);

    // Используем существующую функцию из 13_categories_search.js
    const categories = getCategoriesForSearch();

    logInfo(`✅ Загружено ${categories.length} категорий для подбора`, null, context);

    return categories;

  } catch (error) {
    logError('❌ Ошибка загрузки категорий для подбора', error, context);
    throw new Error(`Не удалось загрузить категории: ${error.message}`);
  }
}

/**
 * Обрабатывает выбор категории и записывает ID в ячейку
 *
 * @param {Object} categoryData - Данные выбранной категории
 * @param {number} targetRow - Номер строки для записи
 * @param {string} sheetName - Название листа
 * @returns {Object} Результат операции
 */
function selectCategoryForTag(categoryData, targetRow, sheetName) {
  const context = 'selectCategoryForTag';

  try {
    logInfo(`📌 Выбор категории ${categoryData.id} для строки ${targetRow}`, null, context);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Лист "${sheetName}" не найден`);
    }

    // Получаем номер колонки для URL/ID категории (колонка E)
    const categoryLinkColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.CATEGORY_LINK;

    // Записываем ID категории в ячейку
    const targetCell = sheet.getRange(targetRow, categoryLinkColumn);
    targetCell.setValue(categoryData.id);

    logInfo(`✅ ID категории ${categoryData.id} записан в ${sheetName}!${targetCell.getA1Notation()}`,
      null, context);

    // Обновляем статус категории в колонке F
    // Устанавливаем статус "✅ Существует" сразу, так как категория выбрана из списка
    const statusColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.CATEGORY_STATUS;
    const statusCell = sheet.getRange(targetRow, statusColumn);
    statusCell.setValue('✅ Существует');

    // Показываем уведомление пользователю
    ss.toast(
      `Категория "${categoryData.title}" выбрана`,
      '✅ Успех',
      3
    );

    logInfo('✅ Категория успешно выбрана и записана', null, context);

    return {
      success: true,
      categoryId: categoryData.id,
      categoryTitle: categoryData.title,
      targetRow: targetRow
    };

  } catch (error) {
    logError(`❌ Ошибка записи выбранной категории`, error, context);

    // Показываем ошибку пользователю
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Ошибка: ${error.message}`,
      '❌ Ошибка',
      5
    );

    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ============================================

/**
 * Включает CSS и JavaScript файлы в HTML шаблон
 * Используется в HTML для подключения внешних файлов
 *
 * @param {string} filename - Имя файла для включения
 * @returns {string} Содержимое файла
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
