/**
 * ========================================
 * МОДУЛЬ: SEO-ТЕГИ (TITLE & DESCRIPTION)
 * ========================================
 *
 * Генерация SEO Title и Description для страниц категорий
 * с использованием Gemini API.
 *
 * ДВА РЕЖИМА РАБОТЫ:
 * 1. Детальный лист - генерация для текущей открытой категории
 * 2. Массовый режим - обработка нескольких категорий по чекбоксам
 *
 * ЭТАПЫ ПОЭТАПНОЙ ГЕНЕРАЦИИ:
 * - Этап 1 (Аналитик): анализ конкурентов и семантики
 * - Этап 2 (Копирайтер): создание черновиков
 * - Этап 3 (Редактор): финальная доработка
 * Температуры задаются в SEO_TAGS_CONFIG.TEMPERATURES
 */

// ========================================
// РЕЖИМ 1: ДЕТАЛЬНЫЙ ЛИСТ КАТЕГОРИИ
// ========================================

/**
 * Генерация SEO-тегов для текущей открытой категории (режим "Сразу")
 * Вызывается из меню: AI Генерация → 🏷️ Генерировать SEO-теги
 */
function generateSeoTagsForActiveCategory() {
  const context = 'generateSeoTagsForActiveCategory';
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  try {
    Logger.log('=== ГЕНЕРАЦИЯ SEO-ТЕГОВ (режим: Сразу) ===');
    logInfo('🏷️ Запуск генерации SEO-тегов для активной категории', null, context);

    // Проверяем, что мы на детальном листе категории
    const sheetName = activeSheet.getName();
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      ui.alert('Ошибка',
        `Эта функция работает только на листе категории.\n` +
        `Текущий лист: "${sheetName}"\n\n` +
        `Откройте категорию через меню: Категории → Найти и открыть категорию`,
        ui.ButtonSet.OK);
      return;
    }

    // Получаем ID категории
    const cells = SEO_TAGS_CONFIG.DETAIL_SHEET_CELLS;
    const categoryId = activeSheet.getRange(cells.CATEGORY_ID).getValue();
    const categoryName = activeSheet.getRange(cells.CATEGORY_NAME).getValue();

    if (!categoryId) {
      ui.alert('Ошибка', 'Не удалось получить ID категории из ячейки B2', ui.ButtonSet.OK);
      return;
    }

    Logger.log(`📋 Категория: "${categoryName}" (ID: ${categoryId})`);
    logInfo(`📋 Категория: "${categoryName}" (ID: ${categoryId})`, null, context);

    // Ищем данные на листе "SEO-теги"
    const rowData = findCategoryInSeoTagsSheet_(categoryId, categoryName);

    if (!rowData) {
      ui.alert('Категория не найдена',
        `Категория "${categoryName}" (ID: ${categoryId}) не найдена на листе "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}".\n\n` +
        `Добавьте её через меню:\nКатегории → SEO-теги (массово) → Загрузить категории`,
        ui.ButtonSet.OK);
      return;
    }

    // Получаем feedback из детального листа
    const feedback = activeSheet.getRange(cells.GEMINI_FEEDBACK).getValue() || '';
    rowData.feedback = feedback;

    // Показываем подтверждение
    const confirm = ui.alert('Генерация SEO-тегов',
      `Категория: ${categoryName}\n` +
      `ID: ${categoryId}\n` +
      (feedback ? `Замечания: "${feedback.substring(0, 50)}..."\n` : '') +
      `\nРежим: Сразу (один запрос)\n\n` +
      `Продолжить?`,
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    ss.toast('Генерация SEO-тегов...', '⏳ Пожалуйста, подождите', -1);

    // Получаем промпт из ячейки (столбец K)
    const prompt = rowData.promptSingle && rowData.promptSingle.trim() !== ''
      ? rowData.promptSingle
      : null;

    if (!prompt) {
      throw new Error('Промпт не задан. Заполните ячейку промпта (столбец K) на листе "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '"');
    }

    // Заменяем плейсхолдеры
    const preparedPrompt = replaceSeoTagsPlaceholders_(prompt, rowData);

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('📤 ЗАПРОС К GEMINI (temp: ' + SEO_TAGS_CONFIG.TEMPERATURES.SINGLE + ')');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('ПРОМПТ (полный):');
    Logger.log(preparedPrompt);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    // Вызываем Gemini API
    const response = callGeminiWithTemperature_(
      preparedPrompt,
      SEO_TAGS_PROMPTS.SYSTEM,
      SEO_TAGS_CONFIG.TEMPERATURES.SINGLE
    );

    Logger.log('📥 ОТВЕТ ОТ GEMINI:');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log(response);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    // Парсим результат
    const parsed = parseSeoTagsResult_(response);

    if (!parsed) {
      throw new Error('Не удалось распарсить ответ Gemini');
    }

    // Записываем результат в ОБА места
    // 1. Детальный лист
    activeSheet.getRange(cells.SEO_TITLE).setValue(parsed.title);
    activeSheet.getRange(cells.META_DESC).setValue(parsed.description);
    activeSheet.getRange(cells.SEO_TITLE).setBackground('#e8f5e9');
    activeSheet.getRange(cells.META_DESC).setBackground('#e8f5e9');

    // 2. Лист SEO-теги
    writeSeoTagsResult_(rowData.rowNumber, parsed);

    ss.toast(`Готово! Title: ${parsed.title.length} симв., Description: ${parsed.description.length} симв.`, '✅ SEO-теги', 5);

    Logger.log(`✅ SEO-теги сгенерированы:`);
    Logger.log(`   Title (${parsed.title.length} симв.):`);
    Logger.log(parsed.title);
    Logger.log(`   Description (${parsed.description.length} симв.):`);
    Logger.log(parsed.description);
    logInfo(`✅ SEO-теги сгенерированы для "${categoryName}"`, {
      titleLength: parsed.title.length,
      descLength: parsed.description.length
    }, context);

    ui.alert('Генерация завершена',
      `Title (${parsed.title.length} симв.):\n${parsed.title}\n\n` +
      `Description (${parsed.description.length} симв.):\n${parsed.description}`,
      ui.ButtonSet.OK);

  } catch (error) {
    logError('❌ Ошибка генерации SEO-тегов', error, context);
    ss.toast('Ошибка генерации', '❌', 3);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Поэтапная генерация SEO-тегов для текущей категории
 * Этап 1 (Аналитик) → Этап 2 (Копирайтер) → Этап 3 (Редактор)
 */
function generateSeoTagsStagedForActiveCategory() {
  const context = 'generateSeoTagsStagedForActiveCategory';
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  try {
    Logger.log('=== ПОЭТАПНАЯ ГЕНЕРАЦИЯ SEO-ТЕГОВ (3 этапа) ===');
    logInfo('🏷️ Запуск ПОЭТАПНОЙ генерации SEO-тегов', null, context);

    // Проверяем, что мы на детальном листе категории
    const sheetName = activeSheet.getName();
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      ui.alert('Ошибка',
        `Эта функция работает только на листе категории.\n` +
        `Откройте категорию через меню: Категории → Найти и открыть категорию`,
        ui.ButtonSet.OK);
      return;
    }

    // Получаем ID категории
    const cells = SEO_TAGS_CONFIG.DETAIL_SHEET_CELLS;
    const categoryId = activeSheet.getRange(cells.CATEGORY_ID).getValue();
    const categoryName = activeSheet.getRange(cells.CATEGORY_NAME).getValue();

    if (!categoryId) {
      ui.alert('Ошибка', 'Не удалось получить ID категории', ui.ButtonSet.OK);
      return;
    }

    // Ищем данные на листе "SEO-теги"
    const rowData = findCategoryInSeoTagsSheet_(categoryId, categoryName);

    if (!rowData) {
      ui.alert('Категория не найдена',
        `Категория "${categoryName}" (ID: ${categoryId}) не найдена на листе "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}".\n\n` +
        `Добавьте её через меню: SEO-теги (массово) → Загрузить категории`,
        ui.ButtonSet.OK);
      return;
    }

    // Получаем feedback
    const feedback = activeSheet.getRange(cells.GEMINI_FEEDBACK).getValue() || '';
    rowData.feedback = feedback;

    // Подтверждение
    const confirm = ui.alert('Поэтапная генерация SEO-тегов',
      `Категория: ${categoryName}\n\n` +
      `Будет выполнено 3 этапа:\n` +
      `1. Аналитик (temp ${SEO_TAGS_CONFIG.TEMPERATURES.STAGE_1})\n` +
      `2. Копирайтер (temp ${SEO_TAGS_CONFIG.TEMPERATURES.STAGE_2})\n` +
      `3. Редактор (temp ${SEO_TAGS_CONFIG.TEMPERATURES.STAGE_3})\n\n` +
      `Продолжить?`,
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    const stageResults = {};

    // ========== ЭТАП 1: АНАЛИТИК ==========
    ss.toast('Этап 1: Аналитик...', '📊', -1);
    logInfo('📊 Этап 1 (Аналитик)', null, context);

    if (!rowData.promptStage1 || rowData.promptStage1.trim() === '') {
      throw new Error('Промпт Этапа 1 (Аналитик) не задан. Заполните столбец L на листе "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '"');
    }
    const prompt1 = rowData.promptStage1;

    const preparedPrompt1 = replaceSeoTagsPlaceholders_(prompt1, rowData, stageResults);

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('📤 ЭТАП 1: ЗАПРОС (Аналитик, temp: 0.3)');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('ПРОМПТ (полный):');
    Logger.log(preparedPrompt1);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    const response1 = callGeminiWithTemperature_(
      preparedPrompt1,
      SEO_TAGS_PROMPTS.STAGE_1_SYSTEM,
      SEO_TAGS_CONFIG.TEMPERATURES.STAGE_1
    );

    stageResults.stage1 = response1;
    Logger.log('📥 ЭТАП 1: ОТВЕТ (Аналитик):');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log(response1);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    logInfo(`✅ Этап 1 завершён (${response1.length} симв.)`, null, context);

    // Записываем результат Этапа 1 в столбец L
    if (rowData.rowNumber && rowData.rowNumber !== 2) {
      const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
      if (seoSheet) seoSheet.getRange(rowData.rowNumber, SEO_TAGS_CONFIG.MASS_COLUMNS.PROMPT_STAGE_1).setValue(response1);
    }

    Utilities.sleep(SEO_TAGS_CONFIG.DELAY_BETWEEN_API_CALLS);

    // ========== ЭТАП 2: КОПИРАЙТЕР ==========
    ss.toast('Этап 2: Копирайтер...', '✍️', -1);
    logInfo('✍️ Этап 2 (Копирайтер)', null, context);

    if (!rowData.promptStage2 || rowData.promptStage2.trim() === '') {
      throw new Error('Промпт Этапа 2 (Копирайтер) не задан. Заполните столбец M на листе "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '"');
    }
    const prompt2 = rowData.promptStage2;

    const preparedPrompt2 = replaceSeoTagsPlaceholders_(prompt2, rowData, stageResults);

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('📤 ЭТАП 2: ЗАПРОС (Копирайтер, temp: ' + SEO_TAGS_CONFIG.TEMPERATURES.STAGE_2 + ')');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('ПРОМПТ (полный):');
    Logger.log(preparedPrompt2);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    const response2 = callGeminiWithTemperature_(
      preparedPrompt2,
      SEO_TAGS_PROMPTS.STAGE_2_SYSTEM,
      SEO_TAGS_CONFIG.TEMPERATURES.STAGE_2
    );

    stageResults.stage2 = response2;
    Logger.log('📥 ЭТАП 2: ОТВЕТ (Копирайтер):');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log(response2);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    logInfo(`✅ Этап 2 завершён (${response2.length} симв.)`, null, context);

    // Записываем результат Этапа 2 в столбец M (с комментариями — для наглядности)
    if (rowData.rowNumber && rowData.rowNumber !== 2) {
      const seoSheet2 = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
      if (seoSheet2) seoSheet2.getRange(rowData.rowNumber, SEO_TAGS_CONFIG.MASS_COLUMNS.PROMPT_STAGE_2).setValue(response2);
    }

    Utilities.sleep(SEO_TAGS_CONFIG.DELAY_BETWEEN_API_CALLS);

    // ========== ЭТАП 3: РЕДАКТОР ==========
    ss.toast('Этап 3: Редактор...', '📝', -1);
    logInfo('📝 Этап 3 (Редактор)', null, context);

    if (!rowData.promptStage3 || rowData.promptStage3.trim() === '') {
      throw new Error('Промпт Этапа 3 (Редактор) не задан. Заполните столбец N на листе "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '"');
    }
    const prompt3 = rowData.promptStage3;

    const preparedPrompt3 = replaceSeoTagsPlaceholders_(prompt3, rowData, stageResults);

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('📤 ЭТАП 3: ЗАПРОС (Редактор, temp: 0.1)');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('ПРОМПТ:');
    Logger.log(preparedPrompt3);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    const response3 = callGeminiWithTemperature_(
      preparedPrompt3,
      SEO_TAGS_PROMPTS.STAGE_3_SYSTEM,
      SEO_TAGS_CONFIG.TEMPERATURES.STAGE_3
    );

    Logger.log('📥 ЭТАП 3: ОТВЕТ (Редактор) - ФИНАЛЬНЫЙ:');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log(response3);
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    logInfo(`✅ Этап 3 завершён`, null, context);

    // Записываем результат Этапа 3 в столбец N (с комментариями — для наглядности)
    if (rowData.rowNumber && rowData.rowNumber !== 2) {
      const seoSheet3 = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
      if (seoSheet3) seoSheet3.getRange(rowData.rowNumber, SEO_TAGS_CONFIG.MASS_COLUMNS.PROMPT_STAGE_3).setValue(response3);
    }

    // Парсим финальный результат
    const parsed = parseSeoTagsResult_(response3);

    if (!parsed) {
      throw new Error('Не удалось распарсить финальный ответ (Этап 3)');
    }

    // Записываем результат в ОБА места
    activeSheet.getRange(cells.SEO_TITLE).setValue(parsed.title);
    activeSheet.getRange(cells.META_DESC).setValue(parsed.description);
    activeSheet.getRange(cells.SEO_TITLE).setBackground('#e8f5e9');
    activeSheet.getRange(cells.META_DESC).setBackground('#e8f5e9');

    writeSeoTagsResult_(rowData.rowNumber, parsed);

    ss.toast('Поэтапная генерация завершена!', '✅', 5);

    Logger.log(`✅ ПОЭТАПНАЯ ГЕНЕРАЦИЯ ЗАВЕРШЕНА`);
    Logger.log(`   Title (${parsed.title.length} симв.):`);
    Logger.log(parsed.title);
    Logger.log(`   Description (${parsed.description.length} симв.):`);
    Logger.log(parsed.description);
    logInfo(`✅ Поэтапная генерация завершена для "${categoryName}"`, null, context);

    ui.alert('Поэтапная генерация завершена',
      `Title (${parsed.title.length} симв.):\n${parsed.title}\n\n` +
      `Description (${parsed.description.length} симв.):\n${parsed.description}`,
      ui.ButtonSet.OK);

  } catch (error) {
    logError('❌ Ошибка поэтапной генерации', error, context);
    ss.toast('Ошибка генерации', '❌', 3);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

// ========================================
// РЕЖИМ 2: МАССОВАЯ ОБРАБОТКА
// ========================================

/**
 * Массовая генерация SEO-тегов (режим "Сразу")
 * Обрабатывает строки с отмеченными чекбоксами на листе "SEO-теги"
 */
function generateSeoTagsSingleMass() {
  const context = 'generateSeoTagsSingleMass';
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    Logger.log('=== МАССОВАЯ ГЕНЕРАЦИЯ SEO-ТЕГОВ (режим: Сразу) ===');
    logInfo('🚀 Запуск массовой генерации SEO-тегов (режим: Сразу)', null, context);

    const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) {
      ui.alert('Ошибка',
        `Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" не найден.\n\n` +
        `Создайте его через меню: SEO-теги (массово) → Создать лист`,
        ui.ButtonSet.OK);
      return;
    }

    // Получаем отмеченные строки
    const checkedRows = getCheckedRowsFromSeoTagsSheet_(sheet);

    if (checkedRows.length === 0) {
      ui.alert('Ошибка', 'Не выбрано ни одной строки (чекбокс в столбце A)', ui.ButtonSet.OK);
      return;
    }

    // Подтверждение
    const confirm = ui.alert('Массовая генерация',
      `Будет обработано строк: ${checkedRows.length}\n` +
      `Режим: Сразу (один запрос на строку)\n\n` +
      `Продолжить?`,
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    Logger.log(`📋 Выбрано строк: ${checkedRows.length}`);
    let successCount = 0;
    let errorCount = 0;
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

    for (let i = 0; i < checkedRows.length; i++) {
      const row = checkedRows[i];

      Logger.log(`\n[${i + 1}/${checkedRows.length}] Обработка: ${row.pageName} (ID: ${row.id})`);
      ss.toast(`Обработка ${i + 1}/${checkedRows.length}: ${row.pageName}`, '⏳', -1);

      try {
        // Получаем промпт из ячейки
        if (!row.promptSingle || row.promptSingle.trim() === '') {
          throw new Error('Промпт не задан для "' + row.pageName + '". Заполните столбец K.');
        }
        const prompt = row.promptSingle;

        // Заменяем плейсхолдеры
        const preparedPrompt = replaceSeoTagsPlaceholders_(prompt, row);

        Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
        Logger.log('📤 ЗАПРОС (temp: 0.5):');
        Logger.log('ПРОМПТ (полный):');
        Logger.log(preparedPrompt);
        Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

        // Вызываем API
        const response = callGeminiWithTemperature_(
          preparedPrompt,
          SEO_TAGS_PROMPTS.SYSTEM,
          SEO_TAGS_CONFIG.TEMPERATURES.SINGLE
        );

        Logger.log('📥 ОТВЕТ:');
        Logger.log(response);
        Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

        // Парсим результат
        const parsed = parseSeoTagsResult_(response);

        if (!parsed) {
          throw new Error('Не удалось распарсить ответ');
        }

        // Записываем результат
        sheet.getRange(row.rowNumber, cols.RESULT_TITLE).setValue(parsed.title);
        sheet.getRange(row.rowNumber, cols.RESULT_DESC).setValue(parsed.description);
        sheet.getRange(row.rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#e8f5e9');

        Logger.log(`✅ РЕЗУЛЬТАТ: Title (${parsed.title.length}), Description (${parsed.description.length})`);
        logInfo(`✅ Строка ${row.rowNumber}: "${row.pageName}"`, null, context);
        successCount++;

        // Пауза между строками
        if (i < checkedRows.length - 1) {
          Utilities.sleep(SEO_TAGS_CONFIG.DELAY_BETWEEN_ROWS);
        }

      } catch (error) {
        Logger.log(`   ❌ ОШИБКА: ${error.message}`);
        logError(`❌ Ошибка строки ${row.rowNumber}`, error, context);

        sheet.getRange(row.rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#ffebee');
        sheet.getRange(row.rowNumber, cols.RESULT_TITLE).setValue('ОШИБКА: ' + error.message);

        errorCount++;
      }
    }

    Logger.log(`\n=== МАССОВАЯ ГЕНЕРАЦИЯ ЗАВЕРШЕНА ===`);
    Logger.log(`   Успешно: ${successCount}, Ошибок: ${errorCount}`);
    ss.toast(`Готово! Успешно: ${successCount}, Ошибок: ${errorCount}`, '✅', 10);

    ui.alert('Массовая генерация завершена',
      `Успешно: ${successCount}\nОшибок: ${errorCount}`,
      ui.ButtonSet.OK);

    logInfo(`🏁 Массовая генерация завершена. Успешно: ${successCount}, Ошибок: ${errorCount}`, null, context);

  } catch (error) {
    logError('❌ Критическая ошибка массовой генерации', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Массовая генерация краткого описания категории (бывший "SEO-теги (1 запрос)").
 * Читает промпт из колонки M (PROMPT_SINGLE), отправляет в AI,
 * записывает результат в колонку S (SHORT_DESC).
 */
function generateCategoryDescriptionMass() {
  const context = 'generateCategoryDescriptionMass';
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    Logger.log('=== МАССОВАЯ ГЕНЕРАЦИЯ ОПИСАНИЯ КАТЕГОРИИ ===');
    logInfo('🚀 Запуск массовой генерации описания категории', null, context);

    const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) {
      ui.alert('Ошибка',
        `Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" не найден.\n\n` +
        `Создайте его через меню: SEO-теги (массово) → Создать лист`,
        ui.ButtonSet.OK);
      return;
    }

    const checkedRows = getCheckedRowsFromSeoTagsSheet_(sheet);

    if (checkedRows.length === 0) {
      ui.alert('Ошибка', 'Не выбрано ни одной строки (чекбокс в столбце A)', ui.ButtonSet.OK);
      return;
    }

    const confirm = ui.alert('Описание категории',
      `Будет обработано строк: ${checkedRows.length}\n` +
      `Режим: генерация краткого описания\n` +
      `Результат → столбец S (Краткое описание)\n\n` +
      `Продолжить?`,
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    let successCount = 0;
    let errorCount = 0;
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

    for (let i = 0; i < checkedRows.length; i++) {
      const row = checkedRows[i];

      Logger.log(`\n[${i + 1}/${checkedRows.length}] Описание: ${row.pageName} (ID: ${row.id})`);
      ss.toast(`Описание ${i + 1}/${checkedRows.length}: ${row.pageName}`, '⏳', -1);

      try {
        if (!row.promptSingle || row.promptSingle.trim() === '') {
          throw new Error('Промпт не задан для "' + row.pageName + '". Заполните столбец M.');
        }
        const prompt = row.promptSingle;
        const preparedPrompt = replaceSeoTagsPlaceholders_(prompt, row);

        const response = callGeminiWithTemperature_(
          preparedPrompt,
          SEO_TAGS_PROMPTS.SYSTEM,
          SEO_TAGS_CONFIG.TEMPERATURES.SINGLE
        );

        // Записываем ответ напрямую в колонку S (Краткое описание)
        const description = (response || '').trim();
        if (!description) {
          throw new Error('Пустой ответ от AI');
        }

        sheet.getRange(row.rowNumber, cols.SHORT_DESC).setValue(description);
        sheet.getRange(row.rowNumber, cols.SHORT_DESC).setBackground('#e8f5e9');

        Logger.log(`✅ Описание (${description.length} символов)`);
        logInfo(`✅ Строка ${row.rowNumber}: "${row.pageName}"`, null, context);
        successCount++;

        if (i < checkedRows.length - 1) {
          Utilities.sleep(SEO_TAGS_CONFIG.DELAY_BETWEEN_ROWS);
        }

      } catch (error) {
        Logger.log(`   ❌ ОШИБКА: ${error.message}`);
        logError(`❌ Ошибка строки ${row.rowNumber}`, error, context);

        sheet.getRange(row.rowNumber, cols.SHORT_DESC).setBackground('#ffebee');
        sheet.getRange(row.rowNumber, cols.SHORT_DESC).setValue('ОШИБКА: ' + error.message);

        errorCount++;
      }
    }

    Logger.log(`\n=== ГЕНЕРАЦИЯ ОПИСАНИЙ ЗАВЕРШЕНА ===`);
    Logger.log(`   Успешно: ${successCount}, Ошибок: ${errorCount}`);
    ss.toast(`Готово! Успешно: ${successCount}, Ошибок: ${errorCount}`, '✅', 10);

    ui.alert('Генерация описаний завершена',
      `Успешно: ${successCount}\nОшибок: ${errorCount}`,
      ui.ButtonSet.OK);

    logInfo(`🏁 Генерация описаний завершена. Успешно: ${successCount}, Ошибок: ${errorCount}`, null, context);

  } catch (error) {
    logError('❌ Критическая ошибка генерации описаний', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Массовая поэтапная генерация SEO-тегов
 */
function generateSeoTagsStagedMass(silent) {
  const context = 'generateSeoTagsStagedMass';
  const ui = silent ? null : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    Logger.log('=== МАССОВАЯ ПОЭТАПНАЯ ГЕНЕРАЦИЯ SEO-ТЕГОВ (3 этапа) ===');
    logInfo('🚀 Запуск массовой ПОЭТАПНОЙ генерации', null, context);

    const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) {
      if (!silent) ui.alert('Ошибка', `Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" не найден.`, ui.ButtonSet.OK);
      return { complete: false, pendingCount: 0 };
    }

    const checkedRows = getCheckedRowsFromSeoTagsSheet_(sheet);

    if (checkedRows.length === 0) {
      if (!silent) ui.alert('Ошибка', 'Не выбрано ни одной строки', ui.ButtonSet.OK);
      return { complete: true, pendingCount: 0 };
    }

    if (!silent) {
      const confirm = ui.alert('Поэтапная массовая генерация',
        `Будет обработано строк: ${checkedRows.length}\n` +
        `Режим: Поэтапно (3 запроса на строку)\n\n` +
        `Этап 1: Аналитик (temp ${SEO_TAGS_CONFIG.TEMPERATURES.STAGE_1})\n` +
        `Этап 2: Копирайтер (temp ${SEO_TAGS_CONFIG.TEMPERATURES.STAGE_2})\n` +
        `Этап 3: Редактор (temp ${SEO_TAGS_CONFIG.TEMPERATURES.STAGE_3})\n\n` +
        `Продолжить?`,
        ui.ButtonSet.YES_NO);

      if (confirm !== ui.Button.YES) return { complete: false, pendingCount: 0 };
    }

    Logger.log(`📋 Выбрано строк: ${checkedRows.length}`);
    let successCount = 0;
    let errorCount = 0;
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

    for (let i = 0; i < checkedRows.length; i++) {
      const row = checkedRows[i];

      Logger.log(`\n[${i + 1}/${checkedRows.length}] Обработка: ${row.pageName}`);
      ss.toast(`Обработка ${i + 1}/${checkedRows.length}: ${row.pageName}`, '⏳', -1);

      try {
        const stageResults = {};

        // Этап 1
        ss.toast(`${row.pageName}: Этап 1 - Аналитик...`, '📊', -1);
        if (!row.promptStage1 || row.promptStage1.trim() === '') {
          throw new Error('Промпт Этапа 1 не задан для "' + row.pageName + '". Заполните столбец L.');
        }
        const prompt1 = row.promptStage1;
        const prepPrompt1 = replaceSeoTagsPlaceholders_(prompt1, row, stageResults);

        Logger.log('   ━━━ ЭТАП 1: Аналитик (temp: 0.3) ━━━');
        Logger.log('   📤 Промпт Этап 1 (полный):');
        Logger.log(prepPrompt1);

        const response1 = callGeminiWithTemperature_(
          prepPrompt1,
          SEO_TAGS_PROMPTS.STAGE_1_SYSTEM,
          SEO_TAGS_CONFIG.TEMPERATURES.STAGE_1
        );
        stageResults.stage1 = response1;
        Logger.log('   📥 Ответ Этап 1 (полный):');
        Logger.log(response1);

        // Записываем результат Этапа 1 в столбец L (кроме строки 2 — там промпт-шаблон)
        if (row.rowNumber !== 2) {
          sheet.getRange(row.rowNumber, cols.PROMPT_STAGE_1).setValue(response1);
        }

        Utilities.sleep(SEO_TAGS_CONFIG.DELAY_BETWEEN_API_CALLS);

        // Этап 2
        ss.toast(`${row.pageName}: Этап 2 - Копирайтер...`, '✍️', -1);
        if (!row.promptStage2 || row.promptStage2.trim() === '') {
          throw new Error('Промпт Этапа 2 не задан для "' + row.pageName + '". Заполните столбец M.');
        }
        const prompt2 = row.promptStage2;
        const prepPrompt2 = replaceSeoTagsPlaceholders_(prompt2, row, stageResults);

        Logger.log('   ━━━ ЭТАП 2: Копирайтер (temp: ' + SEO_TAGS_CONFIG.TEMPERATURES.STAGE_2 + ') ━━━');
        Logger.log('   📤 Промпт Этап 2 (полный):');
        Logger.log(prepPrompt2);

        const response2 = callGeminiWithTemperature_(
          prepPrompt2,
          SEO_TAGS_PROMPTS.STAGE_2_SYSTEM,
          SEO_TAGS_CONFIG.TEMPERATURES.STAGE_2
        );
        stageResults.stage2 = response2;
        Logger.log('   📥 Ответ Этап 2 (полный):');
        Logger.log(response2);

        // Записываем результат Этапа 2 в столбец M (с комментариями — для наглядности)
        if (row.rowNumber !== 2) {
          sheet.getRange(row.rowNumber, cols.PROMPT_STAGE_2).setValue(response2);
        }

        Utilities.sleep(SEO_TAGS_CONFIG.DELAY_BETWEEN_API_CALLS);

        // Этап 3
        ss.toast(`${row.pageName}: Этап 3 - Редактор...`, '📝', -1);
        if (!row.promptStage3 || row.promptStage3.trim() === '') {
          throw new Error('Промпт Этапа 3 не задан для "' + row.pageName + '". Заполните столбец N.');
        }
        const prompt3 = row.promptStage3;
        const prepPrompt3 = replaceSeoTagsPlaceholders_(prompt3, row, stageResults);

        Logger.log('   ━━━ ЭТАП 3: Редактор (temp: 0.1) ━━━');
        Logger.log('   📤 Промпт Этап 3 (полный):');
        Logger.log(prepPrompt3);

        const response3 = callGeminiWithTemperature_(
          prepPrompt3,
          SEO_TAGS_PROMPTS.STAGE_3_SYSTEM,
          SEO_TAGS_CONFIG.TEMPERATURES.STAGE_3
        );
        Logger.log('   📥 ФИНАЛЬНЫЙ ОТВЕТ: ' + response3);

        // Записываем результат Этапа 3 в столбец N (с комментариями — для наглядности)
        if (row.rowNumber !== 2) {
          sheet.getRange(row.rowNumber, cols.PROMPT_STAGE_3).setValue(response3);
        }

        // Парсим результат
        const parsed = parseSeoTagsResult_(response3);

        if (!parsed) {
          throw new Error('Не удалось распарсить финальный ответ');
        }

        // Записываем
        sheet.getRange(row.rowNumber, cols.RESULT_TITLE).setValue(parsed.title);
        sheet.getRange(row.rowNumber, cols.RESULT_DESC).setValue(parsed.description);
        sheet.getRange(row.rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#e8f5e9');

        Logger.log(`   ✅ Готово! Title (${parsed.title.length} симв.):`);
        Logger.log(parsed.title);
        Logger.log(`   Description (${parsed.description.length} симв.):`);
        Logger.log(parsed.description);
        successCount++;

        if (i < checkedRows.length - 1) {
          Utilities.sleep(SEO_TAGS_CONFIG.DELAY_BETWEEN_ROWS);
        }

      } catch (error) {
        Logger.log(`   ❌ ОШИБКА: ${error.message}`);
        logError(`❌ Ошибка строки ${row.rowNumber}`, error, context);

        sheet.getRange(row.rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#ffebee');
        sheet.getRange(row.rowNumber, cols.RESULT_TITLE).setValue('ОШИБКА: ' + error.message);

        errorCount++;
      }
    }

    Logger.log(`\n=== ПОЭТАПНАЯ МАССОВАЯ ГЕНЕРАЦИЯ ЗАВЕРШЕНА ===`);
    Logger.log(`   Успешно: ${successCount}, Ошибок: ${errorCount}`);
    ss.toast(`Готово! Успешно: ${successCount}, Ошибок: ${errorCount}`, '✅', 10);

    if (!silent) {
      ui.alert('Поэтапная генерация завершена',
        `Успешно: ${successCount}\nОшибок: ${errorCount}`,
        ui.ButtonSet.OK);
    }

    return { complete: true, pendingCount: 0 };

  } catch (error) {
    logError('❌ Критическая ошибка', error, context);
    if (!silent) ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
    return { complete: false, pendingCount: 0 };
  }
}

// ========================================
// ОТПРАВКА SEO-ТЕГОВ В INSALES
// ========================================

/**
 * Массовая отправка сгенерированных SEO-тегов из листа "SEO-теги" в InSales.
 * Берёт Title и Description из столбцов результата (O, P) для выбранных строк (☑️)
 * и отправляет их через InSales API как seo_title и seo_description.
 */
function pushSeoTagsToInSales() {
  const context = 'pushSeoTagsToInSales';
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) {
      ui.alert('Ошибка', `Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" не найден.`, ui.ButtonSet.OK);
      return;
    }

    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    const data = sheet.getDataRange().getValues();

    // Собираем выбранные строки с результатами
    const rowsToSend = [];
    for (let i = SEO_TAGS_CONFIG.MASS_DATA_START_ROW - 1; i < data.length; i++) {
      if (data[i][cols.CHECKBOX - 1] === true) {
        const id = data[i][cols.ID - 1];
        const title = data[i][cols.RESULT_TITLE - 1] || '';
        const desc = data[i][cols.RESULT_DESC - 1] || '';
        const name = data[i][cols.PAGE_NAME - 1] || '';

        if (!id) {
          Logger.log(`⚠️ Строка ${i + 1}: нет ID категории, пропускаем`);
          continue;
        }
        if (!title && !desc) {
          Logger.log(`⚠️ Строка ${i + 1} (${name}): нет Title и Description, пропускаем`);
          continue;
        }

        rowsToSend.push({ rowNumber: i + 1, id: id, title: title, desc: desc, name: name });
      }
    }

    if (rowsToSend.length === 0) {
      ui.alert('Ошибка', 'Не выбрано строк с результатами для отправки.\nУбедитесь, что ☑️ отмечены и Title/Description заполнены.', ui.ButtonSet.OK);
      return;
    }

    const confirm = ui.alert('Отправка SEO-тегов в InSales',
      `Будет отправлено: ${rowsToSend.length} категорий\n\n` +
      `Обновляемые поля:\n` +
      `• seo_title (Title)\n` +
      `• seo_description (Description)\n\n` +
      `Продолжить?`,
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    Logger.log(`=== ОТПРАВКА SEO-ТЕГОВ В INSALES (${rowsToSend.length} категорий) ===`);
    logInfo(`🚀 Начинаем отправку ${rowsToSend.length} категорий`, null, context);

    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < rowsToSend.length; i++) {
      const row = rowsToSend[i];
      ss.toast(`Отправка ${i + 1}/${rowsToSend.length}: ${row.name}`, '📤', -1);

      try {
        // InSales API: html_title = SEO Title, meta_description = SEO Description
        const changes = {
          collection: {
            html_title: row.title || null,
            meta_description: row.desc || null
          }
        };

        Logger.log(`[${i + 1}/${rowsToSend.length}] ID: ${row.id}, Title: ${(row.title || '').substring(0, 50)}...`);

        const result = sendCategoryUpdateToInSalesAPI(row.id, changes);

        if (result && result.success) {
          successCount++;
          // Отмечаем зелёным фон строки результата
          sheet.getRange(row.rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#c8e6c9');
        } else {
          errorCount++;
          sheet.getRange(row.rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#ffcdd2');
        }

      } catch (error) {
        Logger.log(`   ❌ Ошибка: ${error.message}`);
        logError(`❌ Ошибка отправки категории ${row.id}`, error, context);
        errorCount++;
        sheet.getRange(row.rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#ffcdd2');
      }

      // Пауза между запросами
      if (i < rowsToSend.length - 1) {
        Utilities.sleep(APP_CONFIG.apiDelay || 500);
      }
    }

    Logger.log(`=== ОТПРАВКА ЗАВЕРШЕНА: Успешно: ${successCount}, Ошибок: ${errorCount} ===`);
    logInfo(`🏁 Отправка завершена. Успешно: ${successCount}, Ошибок: ${errorCount}`, null, context);

    ss.toast(`Готово! Успешно: ${successCount}, Ошибок: ${errorCount}`, '✅', 10);
    ui.alert('Отправка SEO-тегов завершена',
      `Успешно: ${successCount}\nОшибок: ${errorCount}`,
      ui.ButtonSet.OK);

  } catch (error) {
    logError('❌ Критическая ошибка отправки', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

// ========================================
// ОБЩИЕ ФУНКЦИИ
// ========================================

/**
 * Вызывает Gemini API с кастомной температурой и retry-логикой
 *
 * @param {string} prompt - Текст промпта
 * @param {string} systemInstruction - Системная инструкция
 * @param {number} temperature - Температура генерации (0.0 - 1.0)
 * @returns {string} - Ответ от API
 */
function callGeminiWithTemperature_(prompt, systemInstruction, temperature) {
  const context = 'callGeminiWithTemperature';

  const apiKey = GOOGLE_GEMINI_V2_CONFIG.apiKey;
  const model = GOOGLE_GEMINI_V2_CONFIG.model;
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    systemInstruction: {
      parts: [{ text: systemInstruction }]
    },
    generationConfig: {
      temperature: temperature,
      maxOutputTokens: GOOGLE_GEMINI_V2_CONFIG.maxOutputTokens
    }
  };

  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const maxRetries = SEO_TAGS_CONFIG.MAX_RETRIES;
  const retryDelays = SEO_TAGS_CONFIG.RETRY_DELAYS;

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      if (attempt > 0) {
        Logger.log(`   🔄 Gemini API повтор (попытка ${attempt + 1}/${maxRetries})`);
      }
      logInfo(`🔄 API запрос (попытка ${attempt + 1}/${maxRetries}, temp=${temperature})`, null, context);

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      // Успех
      if (responseCode === 200) {
        const data = JSON.parse(responseText);

        if (!data.candidates || data.candidates.length === 0) {
          throw new Error('Gemini вернул пустой ответ (no candidates)');
        }

        var candidate = data.candidates[0];
        if (!candidate.content || !candidate.content.parts || candidate.content.parts.length === 0) {
          var finishReason = candidate.finishReason || 'unknown';
          throw new Error('Gemini кандидат без контента (finishReason: ' + finishReason + ')');
        }

        const result = candidate.content.parts[0].text;
        logInfo(`✅ API ответ получен (${result.length} символов)`, null, context);
        return result;
      }

      // 503 - Overloaded, ретраим
      if (responseCode === 503) {
        Logger.log(`   ⚠️ API 503 Overloaded - повтор через ${retryDelays[attempt]}ms...`);
        logWarning(`⚠️ API 503 Overloaded (попытка ${attempt + 1}/${maxRetries})`, null, context);

        if (attempt < maxRetries - 1) {
          const delay = retryDelays[attempt] || retryDelays[retryDelays.length - 1];
          logInfo(`⏳ Ожидание ${delay}ms перед повтором...`, null, context);
          Utilities.sleep(delay);
          continue;
        }
      }

      // 429 - Rate limit
      if (responseCode === 429) {
        Logger.log(`   ⚠️ API 429 Rate Limit - повтор...`);
        logWarning(`⚠️ API 429 Rate Limit (попытка ${attempt + 1}/${maxRetries})`, null, context);

        if (attempt < maxRetries - 1) {
          const delay = (retryDelays[attempt] || retryDelays[retryDelays.length - 1]) * 2;
          logInfo(`⏳ Ожидание ${delay}ms перед повтором...`, null, context);
          Utilities.sleep(delay);
          continue;
        }
      }

      // Другие ошибки
      throw new Error(`Gemini API Error (${responseCode}): ${responseText.substring(0, 300)}`);

    } catch (e) {
      if (attempt >= maxRetries - 1) {
        logError(`❌ API запрос неуспешен после ${attempt + 1} попыток`, e, context);
        throw e;
      }

      // Для сетевых ошибок тоже ретраим
      if (e.message.includes('timeout') || e.message.includes('connect')) {
        const delay = retryDelays[attempt] || retryDelays[retryDelays.length - 1];
        Utilities.sleep(delay);
        continue;
      }

      throw e;
    }
  }

  throw new Error('Превышено количество попыток API запроса');
}

/**
 * Очищает результат этапа от ролевых артефактов (приветствия, заключения персон)
 * Работает с структурированным форматом (### СЕКЦИЯ) и свободным текстом.
 *
 * Убирает:
 * - Вступления: "Приветствую!", "Отличный запрос!", "Вот мой анализ:" и т.п.
 * - Заключения: "Надеюсь, это поможет!", "Жду указаний", "Готов к этапу 2" и т.п.
 * - Обёртки markdown code blocks (```...```)
 *
 * @param {string} text - Сырой результат этапа
 * @returns {string} - Очищенный текст
 */
function cleanStageOutput_(text) {
  if (!text) return '';
  let cleaned = text;

  // Убираем обёртку markdown code blocks
  cleaned = cleaned.replace(/^```[\w]*\n?/m, '').replace(/\n?```\s*$/m, '');

  // Убираем ролевые приветствия — всё до первого ### заголовка
  const sectionStart = cleaned.search(/^###\s/m);
  if (sectionStart > 0) {
    cleaned = cleaned.substring(sectionStart);
  }

  // Убираем заключительные ролевые фразы после последнего раздела
  // Паттерн: пустая строка + текст без ### маркера, содержащий типичные sign-off слова
  cleaned = cleaned.replace(/\n---\n[\s\S]*$/i, '');
  cleaned = cleaned.replace(/\n{2,}(?:готов|жду|надеюсь|ожидаю|если нужн|буду рад|обращайтесь|удачи|успехов|с уважением|рад помочь|пишите|не стесняйтесь|фундамент заложен|на связи|вот мой|это всё|это мой|подытож|резюмир|в заключени)[\s\S]*$/i, '');

  return cleaned.trim();
}

/**
 * Удаляет комментарии (<!-- КОММЕНТАРИЙ: ... -->) из ответа этапа.
 * Используется перед передачей результата в следующий этап цепочки,
 * чтобы комментарии не засоряли рабочий контекст.
 * @param {string} text - Ответ этапа с комментариями
 * @returns {string} - Ответ без комментариев
 * @private
 */
function stripStageComments_(text) {
  if (!text) return '';
  return text.replace(/<!--\s*КОММЕНТАРИЙ[\s\S]*?-->/gi, '').trim();
}

/**
 * Заменяет плейсхолдеры в промпте на реальные данные
 * Поддерживает форматы: [C2], [F], {CATEGORY_NAME}, {ANALYSIS_REPORT}
 *
 * @param {string} prompt - Исходный промпт
 * @param {Object} rowData - Данные строки
 * @param {Object} stageResults - Результаты предыдущих этапов
 * @returns {string} - Промпт с заменёнными плейсхолдерами
 */
function replaceSeoTagsPlaceholders_(prompt, rowData, stageResults = {}) {
  if (!prompt) return '';

  let result = prompt;

  // === ОТЛАДКА: Проверяем наличие плейсхолдеров в исходном промпте ===
  Logger.log('━━━ DEBUG: Замена плейсхолдеров ━━━');

  const foundPlaceholders = [];
  // Именованные плейсхолдеры (основной формат)
  if (/\{CATEGORY_NAME\}/i.test(prompt)) foundPlaceholders.push('{CATEGORY_NAME}');
  if (/\{ID\}/i.test(prompt)) foundPlaceholders.push('{ID}');
  if (/\{URL\}/i.test(prompt)) foundPlaceholders.push('{URL}');
  if (/\{PAGE_CONTENT\}/i.test(prompt)) foundPlaceholders.push('{PAGE_CONTENT}');
  if (/\{KEYWORDS\}/i.test(prompt)) foundPlaceholders.push('{KEYWORDS}');
  if (/\{LSI\}/i.test(prompt)) foundPlaceholders.push('{LSI}');
  if (/\{LSI_TOP30\}/i.test(prompt)) foundPlaceholders.push('{LSI_TOP30}');
  if (/\{QA\}/i.test(prompt)) foundPlaceholders.push('{QA}');
  if (/\{COMPETITORS_TITLE\}/i.test(prompt)) foundPlaceholders.push('{COMPETITORS_TITLE}');
  if (/\{COMPETITORS_DESC\}/i.test(prompt)) foundPlaceholders.push('{COMPETITORS_DESC}');
  if (/\{USP\}/i.test(prompt)) foundPlaceholders.push('{USP}');
  if (/\{ANALYSIS_REPORT\}/i.test(prompt)) foundPlaceholders.push('{ANALYSIS_REPORT}');
  if (/\{DRAFT_TAGS\}/i.test(prompt)) foundPlaceholders.push('{DRAFT_TAGS}');
  // Устаревшие [X2] плейсхолдеры (обратная совместимость)
  if (/\[C\d*\]/i.test(prompt)) foundPlaceholders.push('[C] (legacy)');
  if (/\[G\d*\]/i.test(prompt)) foundPlaceholders.push('[G] (legacy)');
  if (/\[F\d*\]/i.test(prompt)) foundPlaceholders.push('[F] (legacy)');
  if (/\[H\d*\]/i.test(prompt)) foundPlaceholders.push('[H] (legacy)');
  if (/\[I\d*\]/i.test(prompt)) foundPlaceholders.push('[I] (legacy)');
  if (/\[J\d*\]/i.test(prompt)) foundPlaceholders.push('[J] (legacy)');
  if (/\[K\d*\]/i.test(prompt)) foundPlaceholders.push('[K] (legacy)');
  if (/\[L\d*\]/i.test(prompt)) foundPlaceholders.push('[L] (legacy)');

  Logger.log('📋 Найденные плейсхолдеры: ' + (foundPlaceholders.length > 0 ? foundPlaceholders.join(', ') : 'НЕТ'));

  // Маппинг столбцов на данные (для формата [X] и [X2])
  const columnMap = {
    'B': rowData.id || '',
    'C': rowData.pageName || '',
    'D': rowData.url || '',
    'F': rowData.pageContent || '',
    'G': rowData.semanticCore || '',
    'H': rowData.lsi || '',
    'I': rowData.qa || '',
    'J': rowData.competitorsTitle || '',
    'K': rowData.competitorsDesc || '',
    'L': rowData.usp || ''
  };

  // Заменяем [X], [X2], [X3], ... на значения
  for (const [col, value] of Object.entries(columnMap)) {
    const regex = new RegExp(`\\[${col}\\d*\\]`, 'gi');
    result = result.replace(regex, value);
  }

  // Заменяем {PLACEHOLDER} формат
  result = result.replace(/\{CATEGORY_NAME\}/gi, rowData.pageName || '');
  result = result.replace(/\{ID\}/gi, rowData.id || '');
  result = result.replace(/\{URL\}/gi, rowData.url || '');
  result = result.replace(/\{KEYWORDS\}/gi, rowData.semanticCore || '');
  result = result.replace(/\{PAGE_CONTENT\}/gi, rowData.pageContent || '');
  result = result.replace(/\{DESCRIPTION\}/gi, rowData.pageContent || ''); // алиас для обратной совместимости
  result = result.replace(/\{LSI\}/gi, rowData.lsi || '');
  // {LSI_TOP30} — топ-30 самых частотных LSI-слов (для компактных промптов)
  if (/\{LSI_TOP30\}/i.test(result)) {
    const lsiTop30 = extractLsiTop_(rowData.lsi, 30);
    result = result.replace(/\{LSI_TOP30\}/gi, lsiTop30);
  }
  result = result.replace(/\{QA\}/gi, rowData.qa || '');
  result = result.replace(/\{COMPETITORS_TITLE\}/gi, rowData.competitorsTitle || '');
  result = result.replace(/\{COMPETITORS_DESC\}/gi, rowData.competitorsDesc || '');
  result = result.replace(/\{USP\}/gi, rowData.usp || '');

  // Feedback
  if (rowData.feedback && rowData.feedback.trim() !== '') {
    result = result.replace(/\{FEEDBACK\}/gi, `\nЗАМЕЧАНИЯ ПОЛЬЗОВАТЕЛЯ:\n${rowData.feedback}`);
  } else {
    result = result.replace(/\{FEEDBACK\}/gi, '');
  }

  // Результаты этапов с отладкой (очищаем ролевые артефакты перед подстановкой)
  if (/\{ANALYSIS_REPORT\}/i.test(prompt)) {
    if (stageResults.stage1) {
      const cleanedReport = stripStageComments_(cleanStageOutput_(stageResults.stage1));
      Logger.log('✅ {ANALYSIS_REPORT} → очищен и заменён (' + stageResults.stage1.length + ' → ' + cleanedReport.length + ' симв.)');
      result = result.replace(/\{ANALYSIS_REPORT\}/gi, cleanedReport);
    } else {
      Logger.log('⚠️ {ANALYSIS_REPORT} найден, но stageResults.stage1 ПУСТОЙ!');
    }
  }

  if (/\{DRAFT_TAGS\}/i.test(prompt)) {
    if (stageResults.stage2) {
      const cleanedDraft = stripStageComments_(stageResults.stage2);
      Logger.log('✅ {DRAFT_TAGS} → очищен от комментариев и заменён (' + stageResults.stage2.length + ' → ' + cleanedDraft.length + ' симв.)');
      result = result.replace(/\{DRAFT_TAGS\}/gi, cleanedDraft);
    } else {
      Logger.log('⚠️ {DRAFT_TAGS} найден, но stageResults.stage2 ПУСТОЙ!');
    }
  }

  // Проверяем, остались ли незаменённые плейсхолдеры
  const unreplaced = result.match(/\{[A-Z_]+\}|\[[A-Z]\d*\]/gi);
  if (unreplaced && unreplaced.length > 0) {
    Logger.log('⚠️ НЕЗАМЕНЁННЫЕ плейсхолдеры: ' + unreplaced.join(', '));
  }

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  return result;
}

/**
 * Извлекает топ-N LSI-слов из ячейки H.
 * Ячейка содержит слова по строкам в формате: "слово (частота)" или просто "слово".
 * Данные уже отсортированы по убыванию частотности — берём первые N строк.
 *
 * @param {string} lsiText - Содержимое ячейки LSI
 * @param {number} limit - Сколько слов взять (по умолчанию 30)
 * @returns {string} - Топ-N слов через запятую (без частотности, только слова)
 */
function extractLsiTop_(lsiText, limit) {
  if (!lsiText) return '';
  limit = limit || 30;

  const lines = lsiText.toString().split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const words = lines
    .slice(0, limit)
    .map(line => line.replace(/\s*\(\d+\)\s*$/, '').trim()) // убираем "(135)" в конце
    .filter(w => w.length > 0);

  return words.join(', ');
}

/**
 * Парсит ответ Gemini в Title и Description
 * Поддерживает форматы:
 * - Текстовый: Title: ... Description: ...
 * - JSON: {"title": "...", "description": "..."}
 *
 * @param {string} response - Ответ от Gemini
 * @returns {Object|null} - {title, description} или null
 */
function parseSeoTagsResult_(response) {
  const context = 'parseSeoTagsResult';

  if (!response || typeof response !== 'string') {
    logError('❌ Пустой или некорректный ответ', null, context);
    return null;
  }

  let cleanResponse = response.trim();

  // Удаляем комментарии этапов (<!-- КОММЕНТАРИЙ: ... -->)
  cleanResponse = stripStageComments_(cleanResponse);

  // Очищаем от markdown обёртки
  cleanResponse = cleanResponse.replace(/^```json\s*/i, '');
  cleanResponse = cleanResponse.replace(/\s*```$/i, '');
  cleanResponse = cleanResponse.trim();

  // Попытка 1: JSON формат
  try {
    const parsed = JSON.parse(cleanResponse);
    if (parsed.title && parsed.description) {
      logInfo('✅ Распарсен JSON формат', null, context);
      return {
        title: parsed.title.trim(),
        description: parsed.description.trim()
      };
    }
  } catch (e) {
    // Не JSON, продолжаем
  }

  // Попытка 2: Текстовый формат с пробелом "Title: ... Description: ..."
  const spaceMatch = cleanResponse.match(/Title:\s*(.+?)\s+Description:\s*(.+)/is);
  if (spaceMatch) {
    logInfo('✅ Распарсен текстовый формат (пробел)', null, context);
    return {
      title: spaceMatch[1].trim(),
      description: spaceMatch[2].trim()
    };
  }

  // Попытка 3: Текстовый формат с | "Title: ... | Description: ..."
  const pipeMatch = cleanResponse.match(/Title:\s*(.+?)\s*\|\s*Description:\s*(.+)/is);
  if (pipeMatch) {
    logInfo('✅ Распарсен текстовый формат (pipe)', null, context);
    return {
      title: pipeMatch[1].trim(),
      description: pipeMatch[2].trim()
    };
  }

  // Попытка 4: Многострочный формат
  const titleMatch = cleanResponse.match(/Title:\s*(.+?)(?:\n|$)/i);
  const descMatch = cleanResponse.match(/Description:\s*(.+?)(?:\n|$)/is);

  if (titleMatch && descMatch) {
    logInfo('✅ Распарсен многострочный формат', null, context);
    return {
      title: titleMatch[1].trim(),
      description: descMatch[1].trim()
    };
  }

  // Попытка 5: Извлечь JSON из текста
  const jsonMatch = cleanResponse.match(/\{[\s\S]*"title"[\s\S]*"description"[\s\S]*\}/);
  if (jsonMatch) {
    try {
      const extracted = JSON.parse(jsonMatch[0]);
      if (extracted.title && extracted.description) {
        logInfo('✅ Извлечён JSON из текста', null, context);
        return {
          title: extracted.title.trim(),
          description: extracted.description.trim()
        };
      }
    } catch (e) {
      // Продолжаем
    }
  }

  logError('❌ Не удалось распарсить ответ', { response: cleanResponse.substring(0, 200) }, context);
  return null;
}

// ========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ========================================

/**
 * Ищет категорию на листе "SEO-теги" по ID
 * @param {string|number} categoryId - ID категории
 * @param {string} categoryName - Название категории
 * @param {boolean} [useRow2Prompts=false] - Если true, промпты и УТП берутся из строки 2
 */
function findCategoryInSeoTagsSheet_(categoryId, categoryName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);

  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();
  const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

  const row2Prompts = getPromptsFromRow2_(sheet);

  for (let i = SEO_TAGS_CONFIG.MASS_DATA_START_ROW - 1; i < data.length; i++) {
    const rowId = data[i][cols.ID - 1];

    if (rowId && (rowId == categoryId || String(rowId) === String(categoryId))) {
      return {
        rowNumber: i + 1,
        rowIndex: i,
        id: rowId,
        pageName: data[i][cols.PAGE_NAME - 1] || '',
        url: data[i][cols.URL - 1] || '',
        pageContent: data[i][cols.PAGE_CONTENT - 1] || '',
        semanticCore: data[i][cols.SEMANTIC_CORE - 1] || '',
        lsi: data[i][cols.LSI - 1] || '',
        qa: data[i][cols.QA - 1] || '',
        competitorsTitle: data[i][cols.COMPETITORS_TITLE - 1] || '',
        competitorsDesc: data[i][cols.COMPETITORS_DESC - 1] || '',
        usp: (data[i][cols.USP - 1] || '') || row2Prompts.usp,
        promptSingle: (data[i][cols.PROMPT_SINGLE - 1] || '') || row2Prompts.promptSingle,
        promptStage1: (data[i][cols.PROMPT_STAGE_1 - 1] || '') || row2Prompts.promptStage1,
        promptStage2: (data[i][cols.PROMPT_STAGE_2 - 1] || '') || row2Prompts.promptStage2,
        promptStage3: (data[i][cols.PROMPT_STAGE_3 - 1] || '') || row2Prompts.promptStage3
      };
    }
  }

  return null;
}

/**
 * Читает промпты и УТП из фиксированной строки 2 (шаблон для всех строк)
 * @param {Sheet} sheet - Лист "SEO-теги"
 * @returns {Object} {usp, promptSingle, promptStage1, promptStage2, promptStage3}
 */
function getPromptsFromRow2_(sheet) {
  const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
  const row2 = sheet.getRange(2, 1, 1, cols.PROMPT_STAGE_3).getValues()[0];
  return {
    usp: row2[cols.USP - 1] || '',
    promptSingle: row2[cols.PROMPT_SINGLE - 1] || '',
    promptStage1: row2[cols.PROMPT_STAGE_1 - 1] || '',
    promptStage2: row2[cols.PROMPT_STAGE_2 - 1] || '',
    promptStage3: row2[cols.PROMPT_STAGE_3 - 1] || ''
  };
}

/**
 * Получает отмеченные строки с листа "SEO-теги"
 * @param {Sheet} sheet - Лист "SEO-теги"
 * @param {boolean} [useRow2Prompts=false] - Если true, промпты и УТП берутся из строки 2
 */
function getCheckedRowsFromSeoTagsSheet_(sheet) {
  const data = sheet.getDataRange().getValues();
  const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
  const checkedRows = [];

  const row2Prompts = getPromptsFromRow2_(sheet);

  for (let i = SEO_TAGS_CONFIG.MASS_DATA_START_ROW - 1; i < data.length; i++) {
    if (data[i][cols.CHECKBOX - 1] === true) {
      checkedRows.push({
        rowNumber: i + 1,
        rowIndex: i,
        id: data[i][cols.ID - 1],
        pageName: data[i][cols.PAGE_NAME - 1] || '',
        url: data[i][cols.URL - 1] || '',
        pageContent: data[i][cols.PAGE_CONTENT - 1] || '',
        semanticCore: data[i][cols.SEMANTIC_CORE - 1] || '',
        lsi: data[i][cols.LSI - 1] || '',
        qa: data[i][cols.QA - 1] || '',
        competitorsTitle: data[i][cols.COMPETITORS_TITLE - 1] || '',
        competitorsDesc: data[i][cols.COMPETITORS_DESC - 1] || '',
        usp: (data[i][cols.USP - 1] || '') || row2Prompts.usp,
        promptSingle: (data[i][cols.PROMPT_SINGLE - 1] || '') || row2Prompts.promptSingle,
        promptStage1: (data[i][cols.PROMPT_STAGE_1 - 1] || '') || row2Prompts.promptStage1,
        promptStage2: (data[i][cols.PROMPT_STAGE_2 - 1] || '') || row2Prompts.promptStage2,
        promptStage3: (data[i][cols.PROMPT_STAGE_3 - 1] || '') || row2Prompts.promptStage3
      });
    }
  }

  return checkedRows;
}

/**
 * Записывает результат на лист "SEO-теги"
 */
function writeSeoTagsResult_(rowNumber, parsed) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);

  if (!sheet) return;

  const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

  sheet.getRange(rowNumber, cols.RESULT_TITLE).setValue(parsed.title);
  sheet.getRange(rowNumber, cols.RESULT_DESC).setValue(parsed.description);
  sheet.getRange(rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#e8f5e9');
}

// Дефолтные промпты удалены — используются только промпты из ячеек листа "SEO-теги"

// ========================================
// МИГРАЦИЯ ПЛЕЙСХОЛДЕРОВ
// ========================================

/**
 * Заменяет старые плейсхолдеры [C2], [G2] и т.д. на именованные {CATEGORY_NAME}, {KEYWORDS}
 * в ячейках промптов (K, L, M, N) листа "SEO-теги"
 */
function migrateToNamedPlaceholders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Лист "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '" не найден');
    return;
  }

  const replacements = [
    { pattern: /\[B\d*\]/gi, replacement: '{ID}' },
    { pattern: /\[C\d*\]/gi, replacement: '{CATEGORY_NAME}' },
    { pattern: /\[D\d*\]/gi, replacement: '{URL}' },
    { pattern: /\[F\d*\]/gi, replacement: '{PAGE_CONTENT}' },
    { pattern: /\[G\d*\]/gi, replacement: '{KEYWORDS}' },
    { pattern: /\[H\d*\]/gi, replacement: '{LSI}' },
    { pattern: /\[I\d*\]/gi, replacement: '{QA}' },
    { pattern: /\[J\d*\]/gi, replacement: '{COMPETITORS_TITLE}' },
    { pattern: /\[K\d*\]/gi, replacement: '{COMPETITORS_DESC}' },
    { pattern: /\[L\d*\]/gi, replacement: '{USP}' }
  ];

  const promptColumns = [
    SEO_TAGS_CONFIG.MASS_COLUMNS.PROMPT_SINGLE,    // K
    SEO_TAGS_CONFIG.MASS_COLUMNS.PROMPT_STAGE_1,   // L
    SEO_TAGS_CONFIG.MASS_COLUMNS.PROMPT_STAGE_2,   // M
    SEO_TAGS_CONFIG.MASS_COLUMNS.PROMPT_STAGE_3    // N
  ];

  const lastRow = sheet.getLastRow();
  if (lastRow < SEO_TAGS_CONFIG.MASS_DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('Нет данных для миграции');
    return;
  }

  let totalReplaced = 0;

  for (const col of promptColumns) {
    const range = sheet.getRange(1, col, lastRow, 1);
    const values = range.getValues();
    let changed = false;

    for (let i = 0; i < values.length; i++) {
      if (!values[i][0] || typeof values[i][0] !== 'string') continue;

      let text = values[i][0];
      let rowChanged = false;

      for (const r of replacements) {
        if (r.pattern.test(text)) {
          text = text.replace(r.pattern, r.replacement);
          rowChanged = true;
        }
      }

      if (rowChanged) {
        values[i][0] = text;
        changed = true;
        totalReplaced++;
      }
    }

    if (changed) {
      range.setValues(values);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Миграция завершена',
    'Обновлено ячеек: ' + totalReplaced + '\n\nСтарые [C2], [G2] и т.д. заменены на {CATEGORY_NAME}, {KEYWORDS} и т.д.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ========================================
// УТИЛИТЫ: СОЗДАНИЕ И ЗАГРУЗКА ЛИСТА
// ========================================

/**
 * Создаёт лист "SEO-теги" с правильной структурой
 */
function createSeoTagsSheet() {
  const context = 'createSeoTagsSheet';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    let sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);

    if (sheet) {
      const confirm = ui.alert('Лист уже существует',
        `Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" уже существует. Пересоздать?`,
        ui.ButtonSet.YES_NO);

      if (confirm !== ui.Button.YES) return;
      ss.deleteSheet(sheet);
    }

    // Создаём новый лист
    sheet = ss.insertSheet(SEO_TAGS_CONFIG.MASS_SHEET_NAME);

    // Заголовки
    const headers = [
      '☑️',                      // A
      'ID',                      // B
      'H1 / Маркерный запрос',   // C
      'URL',                     // D
      'Админка',                 // E
      'Содержание страницы',     // F
      'Запрос: подбор товаров',  // G [NEW]
      'Семантическое ядро',      // H
      'Конкуренты (Ссылки)',     // I [RENAMED]
      'Description конкурентов', // J
      'Наше УТП',                // K
      'Общий запрос',            // L
      'Запрос: Этап 1',          // M
      'Запрос: Этап 2',          // N
      'Запрос: Этап 3',          // O
      'Title',                   // P
      'Description',             // Q
      'Товары (результат AI)',   // R [NEW]
      'Дата обновления'          // S [NEW]
    ];

    // Записываем заголовки
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Стилизация заголовков
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange
      .setBackground('#1976d2')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setWrap(true);

    // Результатные столбцы (P, Q, R) - зелёный заголовок
    sheet.getRange(1, 16, 1, 3).setBackground('#4caf50');

    // Ширина столбцов
    sheet.setColumnWidth(1, 40);   // Чекбокс
    sheet.setColumnWidth(2, 80);   // ID
    sheet.setColumnWidth(3, 200);  // Название
    sheet.setColumnWidth(4, 200);  // URL
    sheet.setColumnWidth(5, 100);  // Админка
    sheet.setColumnWidth(6, 300);  // Содержание
    sheet.setColumnWidth(7, 250);  // Запрос подбора [NEW]
    sheet.setColumnWidth(8, 200);  // Семантика
    sheet.setColumnWidth(9, 200);  // Конкуренты
    sheet.setColumnWidth(10, 250); // Desc конкурентов
    sheet.setColumnWidth(11, 200); // УТП
    sheet.setColumnWidth(12, 300); // Общий запрос
    sheet.setColumnWidth(13, 300); // Этап 1
    sheet.setColumnWidth(14, 300); // Этап 2
    sheet.setColumnWidth(15, 300); // Этап 3
    sheet.setColumnWidth(16, 300); // Title результат
    sheet.setColumnWidth(17, 350); // Description результат
    sheet.setColumnWidth(18, 300); // Товары результат [NEW]
    sheet.setColumnWidth(19, 150); // Дата обновления [NEW]

    // Фиксируем заголовок
    sheet.setFrozenRows(1);

    // Добавляем чекбоксы (100 строк)
    sheet.getRange(2, 1, 100, 1).insertCheckboxes();

    logInfo(`✅ Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" создан`, null, context);
    ui.alert('Готово', `Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" создан`, ui.ButtonSet.OK);

  } catch (error) {
    logError('❌ Ошибка создания листа', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

// Ключ для хранения прогресса загрузки H1
const SEO_H1_PROGRESS_KEY = 'seo_h1_load_progress';
const SEO_H1_BATCH_SIZE = 100; // Категорий за один цикл (~2 мин с учётом задержек)
const SEO_H1_TRIGGER_NAME = 'continueLoadingH1_';

// Ключ для хранения прогресса полной загрузки (с SEO)
const SEO_FULL_PROGRESS_KEY = 'seo_full_load_progress';
const SEO_FULL_BATCH_SIZE = 10; // Меньший батч, т.к. делаем API-запросы для каждой категории

// ========================================
// CHUNKED STORAGE (обход лимита 9KB на свойство Script Properties)
// ========================================

/**
 * Сохраняет данные в Script Properties, разбивая на чанки по 8KB
 */
function saveProgressChunked_(key, data) {
  const props = PropertiesService.getScriptProperties();
  const json = JSON.stringify(data);
  const chunkSize = 8000;

  if (json.length <= chunkSize) {
    // Помещается в одно свойство
    props.setProperty(key, json);
    return;
  }

  // Разбиваем на чанки
  const numChunks = Math.ceil(json.length / chunkSize);
  for (let i = 0; i < numChunks; i++) {
    props.setProperty(`${key}_chunk_${i}`, json.substring(i * chunkSize, (i + 1) * chunkSize));
  }
  props.setProperty(key, JSON.stringify({ _chunked: true, chunks: numChunks }));
}

/**
 * Загружает данные из Script Properties, собирая из чанков
 */
function loadProgressChunked_(key) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(key);
  if (!raw) return null;

  const parsed = JSON.parse(raw);
  if (!parsed._chunked) return parsed; // Данные влезли в одно свойство

  let json = '';
  for (let i = 0; i < parsed.chunks; i++) {
    json += props.getProperty(`${key}_chunk_${i}`) || '';
  }
  return JSON.parse(json);
}

/**
 * Удаляет прогресс из Script Properties (включая чанки)
 */
function clearProgressChunked_(key) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(key);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (parsed._chunked) {
        for (let i = 0; i < parsed.chunks; i++) {
          props.deleteProperty(`${key}_chunk_${i}`);
        }
      }
    } catch (e) { /* ignore parse errors */ }
  }
  props.deleteProperty(key);
}

// ========================================
// ФУНКЦИИ СИМУЛЯЦИИ SEO-ШАБЛОНА INSALES
// ========================================

/**
 * Генерирует Title по шаблону InSales (если в API пусто)
 * Шаблон: {{ collection.title }} купить в Москве и СПБ в интернет-магазине
 */
function generateTemplateTitle_(categoryTitle) {
  return `${categoryTitle} купить в Москве и СПБ в интернет-магазине`;
}

/**
 * Склонение слова "модель" по числу
 */
function getModelWord_(count) {
  const n = count % 100;
  const n1 = count % 10;
  if (n > 10 && n < 20) return 'моделей';
  if (n1 === 1) return 'модель';
  if (n1 >= 2 && n1 <= 4) return 'модели';
  return 'моделей';
}

/**
 * Генерирует Description по шаблону InSales (если в API пусто)
 * Формат: В категории {H1} — {count} моделей от {brands}, ценой от {min} до {max} ₽ 🚚 Быстрая доставка ☎ 8 (800) 101-38-05
 */
function generateTemplateDescription_(h1, count, brands, minPrice, maxPrice) {
  const word = getModelWord_(count);
  let desc = `В категории ${h1}`;
  if (count > 0) {
    desc += ` — ${count} ${word}`;
  }
  if (brands && brands.length > 0) {
    desc += ` от ${brands.slice(0, 3).join(', ')}`;
  }
  if (minPrice && maxPrice && minPrice > 0) {
    const formatPrice = (p) => Math.round(p).toLocaleString('ru-RU');
    desc += `, ценой от ${formatPrice(minPrice)} до ${formatPrice(maxPrice)} ₽`;
  }
  desc += ' 🚚 Быстрая доставка ☎ 8 (800) 101-38-05';
  return desc;
}

/**
 * Загружает категории из главного листа в лист "SEO-теги"
 * Для каждой категории подтягивает H1 из InSales API (field_values)
 * Поддерживает автоматический перезапуск при превышении лимита времени
 */
function loadCategoriesToSeoTags() {
  const context = 'loadCategoriesToSeoTags';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  try {
    // Проверяем, есть ли незавершённая загрузка SEO
    const savedSeoProgress = props.getProperty(SEO_FULL_PROGRESS_KEY);
    if (savedSeoProgress) {
      const confirm = ui.alert('Продолжить загрузку?',
        'Обнаружена незавершённая загрузка SEO-данных.\n\n' +
        'Да - продолжить с места остановки\n' +
        'Нет - начать заново',
        ui.ButtonSet.YES_NO);

      if (confirm === ui.Button.YES) {
        continueLoadingSeoFullData_();
        return;
      } else {
        // Очищаем прогресс и начинаем заново
        clearSeoFullLoadProgress_();
      }
    }

    // Также проверяем старую загрузку H1 (для обратной совместимости)
    const savedProgress = props.getProperty(SEO_H1_PROGRESS_KEY);
    if (savedProgress) {
      const confirm = ui.alert('Продолжить загрузку H1?',
        'Обнаружена незавершённая загрузка H1 (старый формат).\n\n' +
        'Да - продолжить с места остановки\n' +
        'Нет - начать заново с новым форматом',
        ui.ButtonSet.YES_NO);

      if (confirm === ui.Button.YES) {
        continueLoadingH1_();
        return;
      } else {
        // Очищаем прогресс и начинаем заново
        clearH1LoadProgress_();
      }
    }

    // Получаем главный лист
    const mainSheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    if (!mainSheet) {
      ui.alert('Ошибка', 'Главный лист категорий не найден', ui.ButtonSet.OK);
      return;
    }

    // Получаем или создаём SEO-лист
    let seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      createSeoTagsSheet();
      seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    }

    // Читаем категории
    const mainData = mainSheet.getDataRange().getValues();
    const mainCols = MAIN_LIST_COLUMNS;

    // Собираем список категорий для загрузки (БЕЗ данных о товарах - они загрузятся позже)
    const categoriesList = [];
    for (let i = 2; i < mainData.length; i++) {
      const categoryId = mainData[i][mainCols.CATEGORY_ID - 1];
      const title = mainData[i][mainCols.TITLE - 1];
      const url = mainData[i][mainCols.URL - 1];
      const adminLink = mainData[i][mainCols.ADMIN_LINK - 1];
      const productsCount = mainData[i][mainCols.PRODUCTS_COUNT - 1] || 0;
      const inStockCount = mainData[i][mainCols.IN_STOCK_COUNT - 1] || 0;

      if (categoryId && title) {
        categoriesList.push({
          id: categoryId,
          title: title.replace(/^[└├─│\s]+/, ''),
          url: url || '',
          adminLink: adminLink || '',
          pageContent: '', // Заполнится после загрузки товаров
          productsCount: productsCount,
          inStockCount: inStockCount
        });
      }
    }

    if (categoriesList.length === 0) {
      ui.alert('Предупреждение', 'Нет категорий для загрузки', ui.ButtonSet.OK);
      return;
    }

    // Подтверждение ПЕРЕД тяжёлой загрузкой (чтобы не терять время на ожидание клика)
    const batches = Math.ceil(categoriesList.length / SEO_FULL_BATCH_SIZE);
    const confirm = ui.alert('Загрузка категорий с SEO-тегами',
      `Будет загружено ${categoriesList.length} категорий.\n\n` +
      `Загрузка батчами по ${SEO_FULL_BATCH_SIZE} категорий (${batches} батч(ей)).\n` +
      `Каждый батч загружает полные данные (SEO + H1 + товары).\n\n` +
      `Можно начать работу сразу после первого батча.\n\n` +
      `Продолжить?`,
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    // Загружаем данные товаров (цены и бренды) ПОСЛЕ подтверждения пользователем
    ss.toast('Загрузка данных о товарах...', '⏳', -1);
    logInfo('📊 Загружаем цены и бренды товаров для всех категорий', null, context);
    const productsDataByCategory = loadProductsDataByCategory_();

    // Дополняем categoriesList данными о товарах (pageContent)
    for (const cat of categoriesList) {
      const productData = productsDataByCategory[cat.id];
      let priceRange = '';

      if (productData) {
        const formatPrice = (price) => Math.round(price).toLocaleString('ru-RU');
        if (productData.min === productData.max && productData.min > 0) {
          priceRange = `Цена ${formatPrice(productData.min)} ₽`;
        } else if (productData.min < productData.max) {
          priceRange = `Цены от ${formatPrice(productData.min)} до ${formatPrice(productData.max)} ₽`;
        }
      }

      if (cat.productsCount > 0) {
        let pageContent = `В категории ${cat.productsCount} моделей (из них ${cat.inStockCount} в наличии)`;
        if (priceRange) {
          pageContent += `\n${priceRange}`;
        }
        if (productData && productData.brands && productData.brands.size > 0) {
          const brandsList = Array.from(productData.brands).slice(0, 15).join(', ');
          pageContent += `\nБренды: ${brandsList}`;
        }
        cat.pageContent = pageContent;
      }
    }

    // ШАГ 1: Записываем СТРУКТУРУ с плейсхолдером "⏳ Загрузка..."
    ss.toast('Записываем структуру категорий...', '⏳', -1);
    logInfo(`📝 Записываем структуру ${categoriesList.length} категорий`, null, context);

    const initialData = categoriesList.map(cat => [
      false, cat.id, cat.title, cat.url, cat.adminLink,
      '⏳ Загрузка...', '', '', '', '', '', '', '', '', '', '', '', '', ''
    ]);

    const startRow = SEO_TAGS_CONFIG.MASS_DATA_START_ROW;
    seoSheet.getRange(startRow, 1, initialData.length, 19).setValues(initialData);
    seoSheet.getRange(startRow, 1, initialData.length, 1).insertCheckboxes();

    ss.toast(`Записано ${categoriesList.length} категорий. Загружаем SEO-данные батчами...`, '⏳', 5);
    logInfo(`✅ Записана структура ${categoriesList.length} категорий`, null, context);

    // ШАГ 2: Сохраняем прогресс для батч-загрузки SEO + H1
    // productsSection уже предвычислена в каждой категории, productsDataByCategory не нужна
    const progressData = {
      categories: categoriesList.map(cat => ({
        id: cat.id,
        title: cat.title,
        url: cat.url,
        adminLink: cat.adminLink,
        productsSection: cat.pageContent || '', // предвычисленная секция товаров
        productsCount: cat.productsCount,
        inStockCount: cat.inStockCount
      })),
      needsProductData: false, // данные товаров уже вычислены в productsSection
      currentIndex: 0,
      successCount: 0,
      errorCount: 0,
      startTime: new Date().toISOString()
    };
    saveProgressChunked_(SEO_FULL_PROGRESS_KEY, progressData);

    logInfo(`📥 Начинаем загрузку SEO для ${categoriesList.length} категорий`, null, context);

    // Запускаем батч-загрузку SEO + H1
    continueLoadingSeoFullData_();

  } catch (error) {
    logError('❌ Ошибка инициализации загрузки', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Продолжает загрузку H1 с сохранённой позиции
 * Вызывается автоматически триггером или вручную
 */
function continueLoadingH1_() {
  const context = 'continueLoadingH1';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // Удаляем триггер, если он был
  deleteH1Triggers_();

  try {
    // Загружаем прогресс
    const savedProgress = props.getProperty(SEO_H1_PROGRESS_KEY);
    if (!savedProgress) {
      logWarning('⚠️ Нет сохранённого прогресса', null, context);
      return;
    }

    const progress = JSON.parse(savedProgress);
    const categoriesList = progress.categories;
    let currentIndex = progress.currentIndex;
    let successCount = progress.successCount;
    let errorCount = progress.errorCount;

    // Получаем SEO-лист
    let seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      createSeoTagsSheet();
      seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    }

    // Получаем credentials
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учётные данные InSales');
    }

    const startIndex = currentIndex;
    const endIndex = Math.min(currentIndex + SEO_H1_BATCH_SIZE, categoriesList.length);
    const categoriesToAdd = [];

    ss.toast(`Загрузка H1: ${startIndex + 1}-${endIndex} из ${categoriesList.length}...`, '⏳', -1);
    logInfo(`📥 Загрузка батча: ${startIndex + 1}-${endIndex} из ${categoriesList.length}`, null, context);

    // Обрабатываем порцию категорий
    for (let i = startIndex; i < endIndex; i++) {
      const cat = categoriesList[i];

      // Показываем прогресс каждые 20 категорий
      if ((i - startIndex) % 20 === 0) {
        ss.toast(`Загрузка H1: ${i + 1}/${categoriesList.length}...`, '⏳', -1);
      }

      try {
        const h1Value = fetchCategoryH1FromInsales_(cat.id, credentials);
        const displayName = h1Value && h1Value.trim() !== '' ? h1Value : cat.title;

        categoriesToAdd.push([
          false, cat.id, displayName, cat.url, cat.adminLink,
          cat.pageContent || '', '', '', '', '', '', '', '', '', '', '', ''
        ]);
        successCount++;
      } catch (error) {
        logWarning(`⚠️ Ошибка H1 для ${cat.id}: ${error.message}`, null, context);
        errorCount++;
        categoriesToAdd.push([
          false, cat.id, cat.title, cat.url, cat.adminLink,
          cat.pageContent || '', '', '', '', '', '', '', '', '', '', '', '', '', ''
        ]);
      }

      // Задержка для rate limiting
      if (i < endIndex - 1) {
        Utilities.sleep(500);
      }
    }

    // Записываем данные порции в таблицу
    const startRow = SEO_TAGS_CONFIG.MASS_DATA_START_ROW + startIndex;
    seoSheet.getRange(startRow, 1, categoriesToAdd.length, 19).setValues(categoriesToAdd);
    seoSheet.getRange(startRow, 1, categoriesToAdd.length, 1).insertCheckboxes();

    currentIndex = endIndex;

    // Проверяем, завершена ли загрузка
    if (currentIndex >= categoriesList.length) {
      // Загрузка завершена!
      clearH1LoadProgress_();

      const resultMessage = `✅ Загрузка завершена!\n\n` +
        `Всего категорий: ${categoriesList.length}\n` +
        `H1 получен: ${successCount}\n` +
        `Ошибок (использован title): ${errorCount}`;

      ss.toast('Загрузка H1 завершена!', '✅', 10);
      logInfo(`📥 Загрузка завершена. Успешно: ${successCount}, Ошибок: ${errorCount}`, null, context);

      // Показываем результат
      try {
        SpreadsheetApp.getUi().alert('Готово', resultMessage, SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) {
        // UI недоступен при запуске из триггера - просто логируем
        logInfo(resultMessage, null, context);
      }

    } else {
      // Сохраняем прогресс и планируем продолжение
      progress.currentIndex = currentIndex;
      progress.successCount = successCount;
      progress.errorCount = errorCount;
      props.setProperty(SEO_H1_PROGRESS_KEY, JSON.stringify(progress));

      ss.toast(`Загружено ${currentIndex}/${categoriesList.length}. Продолжение через 1 мин...`, '⏳', 60);
      logInfo(`📥 Батч завершён. Прогресс: ${currentIndex}/${categoriesList.length}. Запланирован перезапуск.`, null, context);

      // Создаём триггер для продолжения через 1 минуту
      ScriptApp.newTrigger('continueLoadingH1_')
        .timeBased()
        .after(60 * 1000) // 1 минута
        .create();
    }

  } catch (error) {
    logError('❌ Ошибка при загрузке батча', error, context);

    // Сохраняем прогресс даже при ошибке, чтобы можно было продолжить
    ss.toast(`Ошибка: ${error.message}. Попробуйте запустить снова.`, '❌', 10);
  }
}

/**
 * Удаляет все триггеры для загрузки H1
 */
function deleteH1Triggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueLoadingH1_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Очищает прогресс загрузки H1
 */
function clearH1LoadProgress_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(SEO_H1_PROGRESS_KEY);
  deleteH1Triggers_();
}

// ========================================
// НОВАЯ БАТЧ-ЗАГРУЗКА С SEO-ТЕГАМИ
// ========================================

/**
 * Продолжает загрузку SEO-данных батчами
 * Каждый батч загружает полные данные (SEO + H1) и записывает в лист
 * В первом батче загружает данные о товарах (цены, бренды) если нужно
 * Формат pageContent: Title, Description, В категории, Цены, Бренды - каждый с новой строки
 */
function continueLoadingSeoFullData_() {
  const context = 'continueLoadingSeoFullData';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Удаляем старый триггер
  deleteSeoFullTriggers_();

  try {
    const progress = loadProgressChunked_(SEO_FULL_PROGRESS_KEY);
    if (!progress) {
      logWarning('⚠️ Нет сохранённого прогресса SEO', null, context);
      return;
    }

    const { categories } = progress;
    let { currentIndex, successCount, errorCount } = progress;

    const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      logError('Лист SEO-теги не найден', null, context);
      return;
    }

    const startRow = SEO_TAGS_CONFIG.MASS_DATA_START_ROW;
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      logError('Не удалось получить credentials InSales', null, context);
      return;
    }

    const authHeader = 'Basic ' + Utilities.base64Encode(
      `${credentials.apiKey}:${credentials.password}`
    );

    // ШАГ 0: Если нужны данные товаров — загружаем в первом батче
    if (progress.needsProductData && currentIndex === 0) {
      ss.toast('Загрузка данных о товарах (цены, бренды)...', '⏳', -1);
      logInfo('📊 Загружаем цены и бренды товаров', null, context);

      const productsDataByCategory = loadProductsDataByCategory_();

      // Вычисляем productsSection для каждой категории
      for (const cat of categories) {
        const productData = productsDataByCategory[cat.id];
        let priceRange = '';

        if (productData) {
          const formatPrice = (price) => Math.round(price).toLocaleString('ru-RU');
          if (productData.min === productData.max && productData.min > 0) {
            priceRange = `Цена ${formatPrice(productData.min)} ₽`;
          } else if (productData.min < productData.max) {
            priceRange = `Цены от ${formatPrice(productData.min)} до ${formatPrice(productData.max)} ₽`;
          }
        }

        if (cat.productsCount > 0) {
          let section = `В категории ${cat.productsCount} моделей (из них ${cat.inStockCount} в наличии)`;
          if (priceRange) {
            section += `\n${priceRange}`;
          }
          if (productData && productData.brands && productData.brands.size > 0) {
            const brandsList = Array.from(productData.brands).slice(0, 15).join(', ');
            section += `\nБренды: ${brandsList}`;
          }
          cat.productsSection = section;
        }
      }

      progress.needsProductData = false;
      logInfo('✅ Данные о товарах загружены и распределены', null, context);

      // Сохраняем обновлённый прогресс с productsSection
      progress.currentIndex = currentIndex;
      progress.successCount = successCount;
      progress.errorCount = errorCount;
      saveProgressChunked_(SEO_FULL_PROGRESS_KEY, progress);

      ss.toast('Данные о товарах готовы. Загружаем SEO...', '⏳', 3);
    }

    // Обрабатываем батч
    const batch = categories.slice(currentIndex, currentIndex + SEO_FULL_BATCH_SIZE);
    logInfo(`📦 Батч: категории ${currentIndex + 1}-${currentIndex + batch.length} из ${categories.length}`, null, context);

    for (let i = 0; i < batch.length; i++) {
      const cat = batch[i];
      // rowNumber: из прогресса (update-режим) или вычисляемый (load-режим)
      const rowNumber = cat.rowNumber || (startRow + currentIndex + i);

      try {
        // API-запрос для получения SEO-данных и H1
        const response = UrlFetchApp.fetch(
          `https://${credentials.shop}/admin/collections/${cat.id}.json`,
          {
            headers: { 'Authorization': authHeader },
            muteHttpExceptions: true
          }
        );

        let seoTitle = '', metaDesc = '', h1 = '';

        if (response.getResponseCode() === 200) {
          const data = JSON.parse(response.getContentText());
          seoTitle = data.html_title || '';
          metaDesc = data.meta_description || '';
          h1 = getFieldValueByName(data, 'H1') || data.title || cat.title;
        } else {
          logWarning(`API ошибка для категории ${cat.id}: ${response.getResponseCode()}`, null, context);
        }

        // Симуляция шаблона если SEO пусто
        const cleanTitle = cat.title.replace(/^[└├─│\s]+/, '');
        if (!seoTitle) {
          seoTitle = generateTemplateTitle_(cleanTitle);
        }
        if (!metaDesc) {
          // Парсим бренды из productsSection для шаблона Description
          const brands = [];
          if (cat.productsSection) {
            const brandsMatch = cat.productsSection.match(/Бренды: (.+)/);
            if (brandsMatch) {
              brands.push(...brandsMatch[1].split(', ').slice(0, 3));
            }
          }
          // Парсим цены из productsSection
          let minPrice = 0, maxPrice = 0;
          if (cat.productsSection) {
            const priceMatch = cat.productsSection.match(/от ([\d\s]+) до ([\d\s]+) ₽/);
            if (priceMatch) {
              minPrice = parseInt(priceMatch[1].replace(/\s/g, '')) || 0;
              maxPrice = parseInt(priceMatch[2].replace(/\s/g, '')) || 0;
            }
          }
          metaDesc = generateTemplateDescription_(
            h1 || cleanTitle,
            cat.productsCount || 0,
            brands,
            minPrice,
            maxPrice
          );
        }

        // Формируем полный pageContent (каждый блок с новой строки)
        let pageContent = `Title: ${seoTitle}\nDescription: ${metaDesc}`;

        // Добавляем предвычисленную секцию товаров
        if (cat.productsSection) {
          pageContent += `\n${cat.productsSection}`;
        }

        // Записываем в лист: C = H1, F = полное содержание
        seoSheet.getRange(rowNumber, 3).setValue(h1 || cleanTitle);  // C: H1
        seoSheet.getRange(rowNumber, 6).setValue(pageContent);       // F: Содержание

        successCount++;
        Utilities.sleep(200); // Rate limiting

      } catch (error) {
        logError(`Ошибка загрузки категории ${cat.id}`, error, context);
        seoSheet.getRange(rowNumber, 6).setValue('❌ Ошибка загрузки');
        errorCount++;
      }
    }

    // Обновляем прогресс
    currentIndex += SEO_FULL_BATCH_SIZE;
    const remaining = Math.max(0, categories.length - currentIndex);

    ss.toast(
      `Загружено ${Math.min(currentIndex, categories.length)} из ${categories.length}. Осталось: ${remaining}`,
      '⏳ Загрузка SEO',
      5
    );

    if (currentIndex < categories.length) {
      // Сохраняем прогресс и создаём триггер для продолжения
      progress.currentIndex = currentIndex;
      progress.successCount = successCount;
      progress.errorCount = errorCount;
      saveProgressChunked_(SEO_FULL_PROGRESS_KEY, progress);

      // Создаём триггер для продолжения через 1 секунду
      ScriptApp.newTrigger('continueLoadingSeoFullData_')
        .timeBased()
        .after(1000)
        .create();

      logInfo(`⏳ Следующий батч через 1 сек. Прогресс: ${currentIndex}/${categories.length}`, null, context);
    } else {
      // Загрузка завершена — очищаем прогресс и триггеры
      clearProgressChunked_(SEO_FULL_PROGRESS_KEY);
      deleteSeoFullTriggers_();
      ss.toast(
        `✅ Загружено ${successCount} категорий` +
        (errorCount > 0 ? `, ошибок: ${errorCount}` : ''),
        '✅ Готово',
        10
      );
      logInfo(`✅ Загрузка SEO завершена: ${successCount} успешно, ${errorCount} ошибок`, null, context);
    }

  } catch (error) {
    logError('❌ Ошибка в continueLoadingSeoFullData_', error, context);
    clearProgressChunked_(SEO_FULL_PROGRESS_KEY);
    ss.toast('Ошибка загрузки: ' + error.message, '❌', 10);
  }
}

/**
 * Удаляет триггеры для загрузки SEO
 */
function deleteSeoFullTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueLoadingSeoFullData_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Очищает прогресс загрузки SEO
 */
function clearSeoFullLoadProgress_() {
  clearProgressChunked_(SEO_FULL_PROGRESS_KEY);
  deleteSeoFullTriggers_();
}

/**
 * Получает ID свойства "Бренд" по permalink (brend или brand)
 * @param {Object} credentials - Учетные данные
 * @returns {number|null} - ID свойства или null
 */
// Функция для получения ID свойства "Бренд" (proizvoditel)
function getBrandPropertyId_(credentials) {
  try {
    const url = `${credentials.baseUrl}/admin/properties.json?permalink=proizvoditel`;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const properties = JSON.parse(response.getContentText());
      if (properties && properties.length > 0) {
        return properties[0].id;
      }
    }
  } catch (e) {
    console.warn('Ошибка при получении ID свойства Бренд (proizvoditel):', e);
  }
  return null;
}

/**
 * Загружает данные о товарах (цены и бренды) для всех категорий из InSales
 * @returns {Object} - Объект {categoryId: {min: number, max: number, brands: Set<string>}}
 */
function loadProductsDataByCategory_() {
  const context = 'loadProductsDataByCategory';
  const dataByCategory = {};

  try {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      logWarning('⚠️ Не удалось получить учётные данные InSales', null, context);
      return dataByCategory;
    }

    // Загружаем товары постранично и сразу обрабатываем
    // const allProducts = []; // Убираем массив для экономии памяти
    let page = 1;
    const perPage = 250;

    // Сначала получим ID свойства "Бренд" (proizvoditel)
    // ИСПРАВЛЕНИЕ: Жестко задаем ID свойства "Бренд" (19403809), так как по пермалинку иногда подтягивался "Тип товара"
    const brandPropId = 19403809;
    // const brandPropId = getBrandPropertyId_(credentials); // Отключаем динамический поиск

    if (!brandPropId) {
      logWarning('⚠️ Не удалось найти свойство "proizvoditel", бренды не будут загружены', null, context);
    } else {
      logInfo(`ℹ️ Используем ID свойства "proizvoditel": ${brandPropId}`, null, context);
    }

    // Ограничим загрузку первыми 50 страницами (12500 товаров), чтобы не упираться в лимиты,
    // так как нам нужны бренды с первой страницы, а они скорее всего будут в начале списка
    // (хотя сортировка может отличаться, но для SEO тегов выборка будет репрезентативной)
    // Ограничим загрузку первыми 50 страницами (12500 товаров)
    while (page <= 50) {
      const url = `${credentials.baseUrl}/admin/products.json?per_page=${perPage}&page=${page}&fields=id,title,collections_ids,variants,characteristics,vendor`;

      const options = {
        method: 'GET',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() !== 200) {
        logWarning(`⚠️ Ошибка загрузки страницы ${page}: ${response.getResponseCode()}`, null, context);
        break;
      }

      const products = JSON.parse(response.getContentText());

      if (!products || products.length === 0) {
        break;
      }

      // Обрабатываем полученную пачку товаров сразу, чтобы не занимать память
      // Считаем все товары, но цены и бренды — ТОЛЬКО по товарам в наличии
      products.forEach(product => {
        const collectionIds = product.collections_ids || [];
        const hasStock = product.variants && product.variants.some(v => v.quantity > 0);

        // Считаем totalCount/inStockCount для ВСЕХ товаров
        collectionIds.forEach(collectionId => {
          if (!dataByCategory[collectionId]) {
            dataByCategory[collectionId] = {
              min: Infinity, max: -Infinity, brands: new Set(),
              totalCount: 0, inStockCount: 0
            };
          }
          dataByCategory[collectionId].totalCount++;
          if (hasStock) dataByCategory[collectionId].inStockCount++;
        });

        // Цены и бренды — только для товаров в наличии
        if (!hasStock) return;

        // 1. Пробуем взять из стандартного поля vendor
        let brandName = product.vendor;
        if (brandName === 'undefined') brandName = null; // Fix for potential string 'undefined'

        // 2. Если пусто, ищем в характеристиках (fallback)
        if (!brandName && brandPropId && product.characteristics) {
          const brandChar = product.characteristics.find(c => c.property_id === brandPropId);
          if (brandChar && brandChar.title) {
            brandName = brandChar.title.replace(/\(/g, '–').split('–')[0].trim();
          }
        }



        // Получаем минимальную цену только из вариантов В НАЛИЧИИ
        let productPrice = null;
        if (product.variants && product.variants.length > 0) {
          const prices = product.variants
            .filter(v => v.quantity > 0) // Только варианты в наличии
            .map(v => parseFloat(v.price))
            .filter(p => !isNaN(p) && p > 0);
          if (prices.length > 0) {
            productPrice = Math.min(...prices);
          }
        }

        collectionIds.forEach(collectionId => {
          // dataByCategory[collectionId] уже инициализирован выше (для всех товаров)

          // Обновляем цены
          if (productPrice !== null) {
            dataByCategory[collectionId].min = Math.min(dataByCategory[collectionId].min, productPrice);
            dataByCategory[collectionId].max = Math.max(dataByCategory[collectionId].max, productPrice);
          }

          // Добавляем бренд
          if (brandName) {
            dataByCategory[collectionId].brands.add(brandName);
          }

          // DEBUG: Logging for specific category (29773031 - Коллиматорные прицелы быстросъемные)
          if (collectionId == 29773031) {
            Logger.log(`[DEBUG] Product ${product.id} (${product.title}) -> Category 29773031. Brand: "${brandName}"`);
          }
        });
      });

      if (products.length < perPage) {
        break;
      }

      page++;
      // Увеличим паузу, так как payload большой
      Utilities.sleep(300);
    }

    // Итоговая статистика
    // logInfo(`📦 Загружено ${allProducts.length} товаров для анализа данных`, null, context); // allProducts теперь пуст
    logInfo(`✅ Данные собраны для ${Object.keys(dataByCategory).length} категорий`, null, context);

    // Очищаем некорректные цены (Infinity)
    Object.keys(dataByCategory).forEach(catId => {
      if (dataByCategory[catId].min === Infinity) {
        dataByCategory[catId].min = 0;
        dataByCategory[catId].max = 0;
      }
    });

    logInfo(`✅ Данные собраны для ${Object.keys(dataByCategory).length} категорий`, null, context);

  } catch (error) {
    logError('❌ Ошибка загрузки данных о товарах', error, context);
  }

  return dataByCategory;
}

/**
 * Получает H1 категории из InSales API
 * @param {number|string} categoryId - ID категории
 * @param {Object} credentials - Учётные данные InSales
 * @returns {string} - Значение H1 или пустая строка
 */
function fetchCategoryH1FromInsales_(categoryId, credentials) {
  const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  // Обработка rate limiting
  if (responseCode === 429) {
    Utilities.sleep(2000);
    return fetchCategoryH1FromInsales_(categoryId, credentials);
  }

  if (responseCode !== 200) {
    throw new Error(`API error ${responseCode}`);
  }

  const responseData = JSON.parse(response.getContentText());

  // InSales возвращает данные в { collection: { ... } }
  const categoryData = responseData.collection || responseData;

  // Извлекаем H1 из field_values
  return getH1FromFieldValues_(categoryData);
}

/**
 * Извлекает значение H1 из field_values категории
 * Использует getFieldValueByName из 13_categories_search.js (работает через справочник полей)
 * @param {Object} categoryData - Данные категории из API
 * @returns {string} - Значение H1 или пустая строка
 */
function getH1FromFieldValues_(categoryData) {
  // Используем существующую функцию getFieldValueByName, которая работает через справочник полей
  return getFieldValueByName(categoryData, 'H1');
}

/**
 * Очищает результаты (столбцы O и P) для отмеченных строк
 */
function clearSeoTagsResults() {
  const context = 'clearSeoTagsResults';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) {
      ui.alert('Ошибка', 'Лист не найден', ui.ButtonSet.OK);
      return;
    }

    const data = sheet.getDataRange().getValues();
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    let clearedCount = 0;

    for (let i = SEO_TAGS_CONFIG.MASS_DATA_START_ROW - 1; i < data.length; i++) {
      if (data[i][cols.CHECKBOX - 1] === true) {
        const rowNumber = i + 1;
        // Очищаем Title + Description (P, Q)
        sheet.getRange(rowNumber, cols.RESULT_TITLE, 1, 2).clearContent();
        sheet.getRange(rowNumber, cols.RESULT_TITLE, 1, 2).setBackground('#ffffff');
        // Очищаем критерии подбора (G) и результаты подбора товаров (R)
        sheet.getRange(rowNumber, cols.PRODUCT_CRITERIA).clearContent();
        sheet.getRange(rowNumber, cols.PRODUCT_CRITERIA).setBackground('#ffffff');
        sheet.getRange(rowNumber, cols.PRODUCTS_RESULT).clearContent();
        sheet.getRange(rowNumber, cols.PRODUCTS_RESULT).setBackground('#ffffff');
        clearedCount++;
      }
    }

    ui.alert('Готово', `Очищено строк: ${clearedCount}`, ui.ButtonSet.OK);
    logInfo(`🗑️ Очищено результатов: ${clearedCount}`, null, context);

  } catch (error) {
    logError('❌ Ошибка очистки', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

// ========================================
// СБОР КЛЮЧЕВИКОВ ИЗ ЛИСТА "СЕМАНТИКА"
// ========================================

/**
 * Собирает ключевики из листа "Семантика" в колонку G листа "SEO-теги"
 * Для категорий без данных автоматически запускает импорт из Метрики/GSC
 */
function pullKeywordsFromSemanticsToSeoTags() {
  const context = 'pullKeywordsFromSemanticsToSeoTags';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    logInfo('📥 Запуск сбора ключевиков для SEO-тегов', null, context);

    // Получаем листы
    const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      ui.alert('Ошибка', `Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" не найден`, ui.ButtonSet.OK);
      return;
    }

    let semanticsSheet = ss.getSheetByName(CATEGORY_SHEETS.SEMANTICS);
    if (!semanticsSheet) {
      // Создаём лист "Семантика" если не существует
      semanticsSheet = ss.insertSheet(CATEGORY_SHEETS.SEMANTICS);
      semanticsSheet.getRange('A1:K1').setValues([['✅', 'Ключевая фраза', 'Частотность (Шир.)', 'Частотность ("")', 'Частотность (!)', 'Частотность ([!])', 'Визиты', 'Отказы (%)', 'Заказы', 'Google', 'Период']]);
      semanticsSheet.getRange('A1:K1').setFontWeight('bold').setBackground('#e0e0e0');
    }

    // Читаем данные с листа SEO-теги
    const seoData = seoSheet.getDataRange().getValues();
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

    // Находим отмеченные строки
    const checkedRows = [];
    for (let i = SEO_TAGS_CONFIG.MASS_DATA_START_ROW - 1; i < seoData.length; i++) {
      if (seoData[i][cols.CHECKBOX - 1] === true) {
        checkedRows.push({
          rowNumber: i + 1,
          categoryId: String(seoData[i][cols.ID - 1]),
          pageName: seoData[i][cols.PAGE_NAME - 1] || '',
          url: seoData[i][cols.URL - 1] || ''
        });
      }
    }

    if (checkedRows.length === 0) {
      ui.alert('Ошибка', 'Не выбрано ни одной строки (чекбокс в столбце A)', ui.ButtonSet.OK);
      return;
    }

    // Спрашиваем период для импорта (если понадобится)
    const userProperties = PropertiesService.getUserProperties();
    const lastPeriod = userProperties.getProperty('LAST_SEMANTICS_PERIOD') || '2024-01-01:2025-12-31';

    const dateResponse = ui.prompt(
      'Сбор ключевиков',
      `Выбрано категорий: ${checkedRows.length}\n\n` +
      `Период для автоимпорта (если нет данных):\n` +
      `Формат: YYYY-MM-DD:YYYY-MM-DD\n` +
      `Оставьте пустым для: ${lastPeriod}`,
      ui.ButtonSet.OK_CANCEL
    );

    if (dateResponse.getSelectedButton() !== ui.Button.OK) return;

    let inputDate = dateResponse.getResponseText().trim();
    if (!inputDate) inputDate = lastPeriod;
    userProperties.setProperty('LAST_SEMANTICS_PERIOD', inputDate);

    const parts = inputDate.split(':');
    const date1 = parts[0].trim();
    const date2 = parts[1].trim();
    const periodStr = `${date1}:${date2}`;

    // Статистика
    let successCount = 0;
    let importedCount = 0;
    let errorCount = 0;

    ss.toast(`Обработка ${checkedRows.length} категорий...`, '⏳', -1);

    // Обрабатываем каждую отмеченную строку
    for (let i = 0; i < checkedRows.length; i++) {
      const row = checkedRows[i];
      ss.toast(`${i + 1}/${checkedRows.length}: ${row.pageName}`, '⏳', -1);

      try {
        // Ищем блок категории на листе "Семантика"
        let blockData = findSemanticsBlockForCategory_(semanticsSheet, row.categoryId);

        // Если блок не найден или пустой - запускаем автоимпорт
        if (!blockData || blockData.keywords.length === 0) {
          logInfo(`📥 Автоимпорт для категории ${row.categoryId}`, null, context);

          // Получаем данные категории для импорта
          const catData = {
            id: row.categoryId,
            path: row.url,
            name: row.pageName,
            breadcrumbs: row.pageName
          };

          // Вызываем функцию импорта из analytics_wordstat.js
          processCategoryImport(catData, semanticsSheet, date1, date2, periodStr);
          importedCount++;

          // Перечитываем блок после импорта
          SpreadsheetApp.flush();
          blockData = findSemanticsBlockForCategory_(semanticsSheet, row.categoryId);
        }

        // Если после импорта всё ещё нет данных - пропускаем
        if (!blockData || blockData.keywords.length === 0) {
          logWarning(`⚠️ Нет данных для категории ${row.categoryId}`, null, context);
          continue;
        }

        // Сортируем по (визиты + клики) DESC
        blockData.keywords.sort((a, b) => (b.visits + b.clicks) - (a.visits + a.clicks));

        // Формируем строку ключевиков (один под другим)
        const keywordsText = blockData.keywords
          .map(kw => kw.phrase)
          .join('\n');

        // Записываем в колонку G
        seoSheet.getRange(row.rowNumber, cols.SEMANTIC_CORE).setValue(keywordsText);
        successCount++;

        logInfo(`✅ Категория ${row.categoryId}: ${blockData.keywords.length} ключевиков`, null, context);

      } catch (error) {
        logError(`❌ Ошибка для категории ${row.categoryId}`, error, context);
        errorCount++;
      }

      // Небольшая пауза между категориями
      if (i < checkedRows.length - 1) {
        Utilities.sleep(200);
      }
    }

    // Итоговая статистика
    const resultMessage = `Готово!\n\n` +
      `Обработано: ${successCount}\n` +
      `Импортировано: ${importedCount}\n` +
      `Ошибок: ${errorCount}`;

    ss.toast('Сбор ключевиков завершён!', '✅', 5);
    ui.alert('Результат', resultMessage, ui.ButtonSet.OK);

    logInfo(`📊 Итого: успешно ${successCount}, импортировано ${importedCount}, ошибок ${errorCount}`, null, context);

  } catch (error) {
    logError('❌ Критическая ошибка сбора ключевиков', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Находит блок категории на листе "Семантика" и извлекает ключевые фразы
 * @param {Sheet} semanticsSheet - Лист "Семантика"
 * @param {string} categoryId - ID категории
 * @returns {Object|null} - {startRow, endRow, keywords: [{phrase, visits, clicks}]}
 */
function findSemanticsBlockForCategory_(semanticsSheet, categoryId) {
  const lastRow = semanticsSheet.getLastRow();
  if (lastRow < 2) return null;

  // Читаем колонку A для поиска заголовков
  const colAValues = semanticsSheet.getRange(1, 1, lastRow, 1).getValues();

  let blockStartRow = -1;
  let blockEndRow = -1;

  // Ищем заголовок блока по паттерну `📂 ... (categoryId)`
  for (let i = 0; i < colAValues.length; i++) {
    const val = String(colAValues[i][0]);
    if (val.includes(`(${categoryId})`) && val.startsWith('📂')) {
      blockStartRow = i + 1; // 1-based

      // Ищем конец блока (следующий заголовок или конец данных)
      for (let j = i + 1; j < colAValues.length; j++) {
        const nextVal = String(colAValues[j][0]);
        if (nextVal.startsWith('📂')) {
          blockEndRow = j; // Не включаем следующий заголовок
          break;
        }
      }
      if (blockEndRow === -1) blockEndRow = lastRow;
      break;
    }
  }

  if (blockStartRow === -1) return null;

  // Извлекаем ключевые фразы из блока
  const keywords = [];
  const dataRowsCount = blockEndRow - blockStartRow;

  if (dataRowsCount > 0) {
    // Колонки: B=фраза(2), G=визиты(7), J=google клики(10)
    const blockData = semanticsSheet.getRange(blockStartRow + 1, 1, dataRowsCount, 11).getValues();

    for (const row of blockData) {
      const phrase = String(row[1]).trim(); // Колонка B (индекс 1)
      if (phrase && phrase !== '' && !phrase.startsWith('📂')) {
        const visits = parseInt(row[6]) || 0;  // Колонка G (индекс 6)
        const clicks = parseInt(row[9]) || 0;  // Колонка J (индекс 9)

        keywords.push({
          phrase: phrase,
          visits: visits,
          clicks: clicks
        });
      }
    }
  }

  return {
    startRow: blockStartRow,
    endRow: blockEndRow,
    keywords: keywords
  };
}

// ========================================
// УМНОЕ ОБНОВЛЕНИЕ ЛИСТА SEO-ТЕГИ
// ========================================

/**
 * Умное обновление листа "SEO-теги":
 * - Быстро обновляет структуру (D, E колонки, удалённые, новые)
 * - Запускает батчевую загрузку SEO-данных (Title, Description, H1) для всех категорий
 * - Загрузка товаров (цены, бренды) происходит в первом батче, не блокируя меню
 */
function updateCategoriesToSeoTags() {
  const context = 'updateCategoriesToSeoTags';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    logInfo('🔄 Запуск умного обновления SEO-теги', null, context);

    // Получаем главный лист
    const mainSheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    if (!mainSheet) {
      ui.alert('Ошибка', 'Главный лист категорий не найден', ui.ButtonSet.OK);
      return;
    }

    // Получаем SEO-лист
    let seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      ui.alert('Информация',
        'Лист "SEO-теги" не найден.\n\nБудет выполнена полная загрузка.',
        ui.ButtonSet.OK);
      loadCategoriesToSeoTags();
      return;
    }

    // Читаем существующие данные с SEO-листа
    const seoData = seoSheet.getDataRange().getValues();
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    const startRow = SEO_TAGS_CONFIG.MASS_DATA_START_ROW;

    // Создаём карту существующих категорий: ID -> номер строки
    const existingCategories = new Map();
    for (let i = startRow - 1; i < seoData.length; i++) {
      const id = seoData[i][cols.ID - 1];
      if (id) {
        existingCategories.set(String(id), i + 1);
      }
    }

    logInfo(`📋 Найдено ${existingCategories.size} категорий на листе SEO-теги`, null, context);

    // Читаем категории из главного листа
    const mainData = mainSheet.getDataRange().getValues();
    const mainCols = MAIN_LIST_COLUMNS;

    const currentCategoryIds = new Set();
    const allCategories = [];

    for (let i = 2; i < mainData.length; i++) {
      const categoryId = mainData[i][mainCols.CATEGORY_ID - 1];
      const title = mainData[i][mainCols.TITLE - 1];
      const url = mainData[i][mainCols.URL - 1];
      const adminLink = mainData[i][mainCols.ADMIN_LINK - 1];

      if (!categoryId || !title) continue;

      currentCategoryIds.add(String(categoryId));

      const productsCount = mainData[i][mainCols.PRODUCTS_COUNT - 1] || 0;
      const inStockCount = mainData[i][mainCols.IN_STOCK_COUNT - 1] || 0;
      const cleanTitle = title.replace(/^[└├─│\s]+/, '');
      const isExisting = existingCategories.has(String(categoryId));

      allCategories.push({
        id: categoryId,
        title: cleanTitle,
        url: url || '',
        adminLink: adminLink || '',
        productsSection: '', // Заполнится в первом батче загрузки
        productsCount: productsCount,
        inStockCount: inStockCount,
        rowNumber: isExisting ? existingCategories.get(String(categoryId)) : null,
        isNew: !isExisting
      });
    }

    // Находим удалённые категории
    const deletedCategoryRows = [];
    for (const [id, rowNumber] of existingCategories) {
      if (!currentCategoryIds.has(id)) {
        deletedCategoryRows.push(rowNumber);
      }
    }

    const newCount = allCategories.filter(c => c.isNew).length;
    const updateCount = allCategories.filter(c => !c.isNew).length;

    // Подтверждение (мгновенное — до тяжёлых операций)
    const confirm = ui.alert('Умное обновление SEO-теги',
      `Категорий на листе: ${existingCategories.size}\n` +
      `Категорий в каталоге: ${currentCategoryIds.size}\n\n` +
      `• Обновить: ${updateCount} категорий\n` +
      `• Добавить новых: ${newCount} категорий\n` +
      `• Пометить удалёнными: ${deletedCategoryRows.length}\n\n` +
      `Будут обновлены: C (H1), D (URL), E (Админка), F (Title + Description + товары)\n` +
      `Сохранены: G-P (семантика, промпты, результаты)\n\n` +
      `Загрузка SEO-данных пойдёт батчами по ${SEO_FULL_BATCH_SIZE} категорий.\n\n` +
      `Продолжить?`,
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    ss.toast('Подготовка структуры...', '⏳', -1);

    // ШАГ 1: Помечаем удалённые категории серым
    if (deletedCategoryRows.length > 0) {
      for (const rowNumber of deletedCategoryRows) {
        seoSheet.getRange(rowNumber, 1, 1, 16)
          .setFontColor('#9e9e9e')
          .setBackground('#f5f5f5');
      }
      logInfo(`⚠️ Помечено ${deletedCategoryRows.length} удалённых категорий`, null, context);
    }

    // ШАГ 2: Добавляем новые категории в конец списка
    const newCategories = allCategories.filter(c => c.isNew);
    if (newCategories.length > 0) {
      const lastRow = seoSheet.getLastRow();
      const newRowStart = lastRow + 1;

      const newData = newCategories.map(cat => [
        false, cat.id, cat.title, cat.url, cat.adminLink,
        '⏳ Загрузка SEO...', '', '', '', '', '', '', '', '', '', ''
      ]);

      seoSheet.getRange(newRowStart, 1, newData.length, 16).setValues(newData);
      seoSheet.getRange(newRowStart, 1, newData.length, 1).insertCheckboxes();
      seoSheet.getRange(newRowStart, 1, newData.length, 16).setBackground('#e8f5e9');

      // Назначаем rowNumber для новых категорий
      newCategories.forEach((cat, i) => { cat.rowNumber = newRowStart + i; });

      logInfo(`➕ Добавлено ${newCategories.length} новых категорий`, null, context);
    }

    // ШАГ 3: Обновляем D, E для существующих (быстро, без API)
    for (const cat of allCategories) {
      if (!cat.isNew && cat.rowNumber) {
        seoSheet.getRange(cat.rowNumber, cols.URL).setValue(cat.url);
        seoSheet.getRange(cat.rowNumber, cols.ADMIN_LINK).setValue(cat.adminLink);
        // Убираем серый формат если был
        seoSheet.getRange(cat.rowNumber, 1, 1, 16)
          .setFontColor('#000000')
          .setBackground(null);
      }
    }

    // ШАГ 4: Сохраняем прогресс и запускаем батч-загрузку SEO
    const progressData = {
      categories: allCategories.map(cat => ({
        id: cat.id,
        title: cat.title,
        url: cat.url,
        adminLink: cat.adminLink,
        productsSection: '', // Заполнится в первом батче
        productsCount: cat.productsCount,
        inStockCount: cat.inStockCount,
        rowNumber: cat.rowNumber
      })),
      needsProductData: true, // Товары загрузятся в первом батче
      currentIndex: 0,
      successCount: 0,
      errorCount: 0,
      startTime: new Date().toISOString()
    };

    saveProgressChunked_(SEO_FULL_PROGRESS_KEY, progressData);

    logInfo(`📥 Запуск батч-загрузки SEO для ${allCategories.length} категорий`, null, context);
    ss.toast(`Структура обновлена. Загрузка SEO-данных батчами (${allCategories.length} категорий)...`, '⏳', 5);

    // Запускаем батч-загрузку
    continueLoadingSeoFullData_();

  } catch (error) {
    logError('❌ Ошибка умного обновления', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Загрузка SEO-данных из InSales ТОЛЬКО для отмеченных чекбоксом строк
 * Заполняет столбцы C (H1), D (URL), E (Админка), F (Содержание страницы)
 */
function loadSeoDataForCheckedRows(silent) {
  const context = 'loadSeoDataForCheckedRows';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = silent ? null : SpreadsheetApp.getUi();

  try {
    const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      if (!silent) ui.alert('Ошибка', 'Лист "SEO-теги" не найден', ui.ButtonSet.OK);
      return { complete: false, pendingCount: 0 };
    }

    const progressKey = SEO_TAGS_CONFIG.LOAD_CHECKED_PROGRESS_KEY;

    // Проверяем незавершённый прогресс
    let progress = loadProgressChunked_(progressKey);

    if (progress && progress.rows) {
      const pendingCount = progress.rows.filter(r => r.status === 'pending').length;
      if (pendingCount > 0) {
        if (silent) {
          // В пайплайне автоматически продолжаем
        } else {
          const resume = ui.alert('Незавершённая загрузка',
            `Найдена незавершённая загрузка:\n` +
            `Всего: ${progress.rows.length}, Ожидают: ${pendingCount}\n\n` +
            `Продолжить с места остановки?`,
            ui.ButtonSet.YES_NO);

          if (resume !== ui.Button.YES) {
            clearProgressChunked_(progressKey);
            progress = null;
          }
        }
      } else {
        clearProgressChunked_(progressKey);
        progress = null;
      }
    }

    if (!progress) {
      const selected = getSelectedSeoTagRows_();
      if (selected.length === 0) {
        if (!silent) ui.alert('Нет данных', 'Отметьте чекбоксами строки для загрузки на листе "SEO-теги"', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      if (!silent) {
        const confirm = ui.alert('Загрузка данных из InSales',
          `Будет загружено ${selected.length} категорий.\n` +
          `Обновятся столбцы: C (H1), D (URL), E (Админка), F (Содержание)\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);

        if (confirm !== ui.Button.YES) return { complete: false, pendingCount: 0 };
      }

      progress = {
        rows: selected.map(r => ({ row: r.row, id: r.id, status: 'pending' })),
        needsProductData: true,
        productsDataLoaded: false,
        successCount: 0,
        errorCount: 0
      };
    }

    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      if (!silent) ui.alert('Ошибка', 'Не удалось получить credentials InSales', ui.ButtonSet.OK);
      return { complete: false, pendingCount: 0 };
    }

    const authHeader = 'Basic ' + Utilities.base64Encode(
      `${credentials.apiKey}:${credentials.password}`
    );

    // Загрузка данных о товарах (цены, бренды) — один раз
    let productsDataByCategory = {};
    if (progress.needsProductData && !progress.productsDataLoaded) {
      ss.toast('Загрузка данных о товарах (цены, бренды)...', '⏳', -1);
      productsDataByCategory = loadProductsDataByCategory_();
      progress.productsDataLoaded = true;
      // Сохраняем productsSection для каждой строки
      for (const rowInfo of progress.rows) {
        const productData = productsDataByCategory[rowInfo.id];
        rowInfo.productsSection = buildProductsSection_(productData);
      }
      saveProgressChunked_(progressKey, progress);
    }

    const startTime = Date.now();
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    let { successCount, errorCount } = progress;

    for (let i = 0; i < progress.rows.length; i++) {
      const rowInfo = progress.rows[i];
      if (rowInfo.status !== 'pending') continue;

      // Safety timeout (5.5 min)
      if (Date.now() - startTime > ARSENKIN_CONFIG.EXECUTION_SAFETY_MS) {
        progress.successCount = successCount;
        progress.errorCount = errorCount;
        saveProgressChunked_(progressKey, progress);
        const remaining = progress.rows.filter(r => r.status === 'pending').length;
        ss.toast(`Таймаут. Загружено ${successCount}, осталось ${remaining}. Запустите повторно.`, '⏳', 10);
        logInfo(`⏳ Таймаут загрузки. Прогресс сохранён.`, null, context);
        return { complete: false, pendingCount: remaining };
      }

      ss.toast(`Загрузка ${i + 1}/${progress.rows.length}: ID ${rowInfo.id}`, '⏳', 3);

      try {
        const response = UrlFetchApp.fetch(
          `https://${credentials.shop}/admin/collections/${rowInfo.id}.json`,
          {
            headers: { 'Authorization': authHeader },
            muteHttpExceptions: true
          }
        );

        if (response.getResponseCode() !== 200) {
          throw new Error(`API HTTP ${response.getResponseCode()}`);
        }

        const data = JSON.parse(response.getContentText());
        const seoTitle = data.html_title || '';
        const metaDesc = data.meta_description || '';
        const h1 = getFieldValueByName(data, 'H1') || data.title || '';
        const url = data.url ? `/collection/${data.url}` : '';
        const adminLink = `https://binokl.shop/admin2/collections/${rowInfo.id}`;

        // Формируем содержание страницы
        let pageContent = `Title: ${seoTitle}\nDescription: ${metaDesc}`;
        if (rowInfo.productsSection) {
          pageContent += `\n${rowInfo.productsSection}`;
        }

        // Записываем в лист
        seoSheet.getRange(rowInfo.row, cols.PAGE_NAME).setValue(h1);
        seoSheet.getRange(rowInfo.row, cols.URL).setValue(url);
        seoSheet.getRange(rowInfo.row, cols.ADMIN_LINK).setValue(adminLink);
        seoSheet.getRange(rowInfo.row, cols.PAGE_CONTENT).setValue(pageContent);

        rowInfo.status = 'done';
        successCount++;
        Utilities.sleep(300);

      } catch (error) {
        logError(`❌ Ошибка загрузки категории ${rowInfo.id}`, error, context);
        seoSheet.getRange(rowInfo.row, cols.PAGE_CONTENT).setValue('❌ ' + error.message);
        seoSheet.getRange(rowInfo.row, cols.PAGE_CONTENT).setBackground('#ffebee');
        rowInfo.status = 'error';
        rowInfo.error = error.message;
        errorCount++;
      }
    }

    // Завершено
    clearProgressChunked_(progressKey);
    ss.toast(`✅ Загружено: ${successCount}, ошибок: ${errorCount}`, '✅ Готово', 10);
    logInfo(`✅ Загрузка отмеченных завершена: ${successCount} успешно, ${errorCount} ошибок`, null, context);

    if (errorCount > 0 && !silent) {
      ui.alert('Загрузка завершена',
        `Успешно: ${successCount}\nОшибок: ${errorCount}\n\nСтроки с ошибками помечены красным.`,
        ui.ButtonSet.OK);
    }

    return { complete: true, pendingCount: 0 };

  } catch (error) {
    logError('❌ Ошибка загрузки отмеченных', error, context);
    if (!silent) ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
    return { complete: false, pendingCount: 0 };
  }
}

/**
 * Формирует строку с данными о товарах для столбца F
 * productData содержит { min, max, brands: Set, totalCount, inStockCount } из loadProductsDataByCategory_()
 * @private
 */
function buildProductsSection_(productData) {
  if (!productData) return '';

  const formatPrice = (price) => Math.round(price).toLocaleString('ru-RU');
  let parts = [];

  if (productData.totalCount > 0) {
    parts.push(`В категории ${productData.totalCount} моделей (из них ${productData.inStockCount || 0} в наличии)`);
  }

  if (productData.min > 0 && productData.max > 0) {
    if (productData.min === productData.max) {
      parts.push(`Цена ${formatPrice(productData.min)} ₽`);
    } else {
      parts.push(`Цены от ${formatPrice(productData.min)} до ${formatPrice(productData.max)} ₽`);
    }
  }
  if (productData.brands && productData.brands.size > 0) {
    parts.push(`Бренды: ${Array.from(productData.brands).slice(0, 15).join(', ')}`);
  }

  return parts.join('\n');
}

/**
 * Снять пометку "удалённая" с выбранных категорий
 * (восстановить нормальное форматирование)
 */
function restoreDeletedCategories() {
  const context = 'restoreDeletedCategories';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      ui.alert('Ошибка', 'Лист "SEO-теги" не найден', ui.ButtonSet.OK);
      return;
    }

    const data = seoSheet.getDataRange().getValues();
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    let restoredCount = 0;

    for (let i = SEO_TAGS_CONFIG.MASS_DATA_START_ROW - 1; i < data.length; i++) {
      if (data[i][cols.CHECKBOX - 1] === true) {
        const rowNumber = i + 1;
        seoSheet.getRange(rowNumber, 1, 1, 16)
          .setFontColor('#000000')
          .setBackground(null);
        restoredCount++;
      }
    }

    if (restoredCount === 0) {
      ui.alert('Информация', 'Не выбрано ни одной строки (чекбокс в столбце A)', ui.ButtonSet.OK);
      return;
    }

    ui.alert('Готово', `Восстановлено форматирование: ${restoredCount} строк`, ui.ButtonSet.OK);
    logInfo(`✅ Восстановлено форматирование ${restoredCount} строк`, null, context);

  } catch (error) {
    logError('❌ Ошибка восстановления', error, context);
    ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
  }
}
// ========================================
// ТРЁХФАЗНЫЙ ПОДБОР ТОВАРОВ ДЛЯ КАТЕГОРИЙ
// ========================================

// ----------------------------------------
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ: ЛИСТ "ВЫГРУЗКА"
// ----------------------------------------

/**
 * Читает заголовки листа "Выгрузка" и возвращает маппинг колонок
 * @returns {Object|null} { headers: string[], colMap: {name: index}, charColumns: [{name, index}] }
 */
function readExportSheetHeaders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.EXPORT_SHEET_NAME);
  if (!sheet) return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = {};
  headers.forEach(function (h, i) { colMap[String(h).trim()] = i; });

  // Колонки, которые НЕ являются характеристиками (тяжёлые / служебные)
  var skipPatterns = [
    'описание', 'description', 'seo', 'meta', 'url', 'permalink',
    'изображен', 'image', 'тег', 'tag', 'вариант', 'variant',
    'html', 'script', 'ссылк'
  ];

  // Основные поля (не характеристики)
  var corePatterns = [
    'id товара', 'артикул', 'название', 'цена', 'остаток',
    'в наличии', 'категори', 'бренд', 'вес', 'габарит',
    'sku', 'barcode', 'штрихкод'
  ];

  var charColumns = [];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (!h) continue;
    var hLower = h.toLowerCase();

    var isSkip = skipPatterns.some(function (p) { return hLower.indexOf(p) !== -1; });
    if (isSkip) continue;

    var isCore = corePatterns.some(function (p) { return hLower.indexOf(p) !== -1; });
    if (isCore) continue;

    charColumns.push({ name: h, index: i });
  }

  return { headers: headers, colMap: colMap, charColumns: charColumns, sheet: sheet };
}

/**
 * Читает все товары из листа "Выгрузка" с характеристиками
 * @param {Object} exportData - результат readExportSheetHeaders_()
 * @returns {Array} массив товаров с характеристиками
 */
function readExportSheetProducts_(exportData) {
  var sheet = exportData.sheet;
  var colMap = exportData.colMap;
  var charColumns = exportData.charColumns;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var products = [];

  // Гибкий поиск колонки по подстрокам (case-insensitive)
  function findCol_(patterns) {
    var allHeaders = Object.keys(colMap);
    for (var p = 0; p < patterns.length; p++) {
      var pattern = patterns[p].toLowerCase();
      // Сначала точное совпадение (case-insensitive)
      for (var h = 0; h < allHeaders.length; h++) {
        if (allHeaders[h].toLowerCase() === pattern) return colMap[allHeaders[h]];
      }
      // Потом поиск по подстроке
      for (var h = 0; h < allHeaders.length; h++) {
        if (allHeaders[h].toLowerCase().indexOf(pattern) !== -1) return colMap[allHeaders[h]];
      }
    }
    return -1;
  }

  var idCol = findCol_(['id товара', 'id']);
  var nameCol = findCol_(['название товара', 'название', 'наименование']);
  var priceCol = findCol_(['цена продажи', 'цена']);
  var brandCol = findCol_(['бренд', 'производитель']);
  var stockCol = findCol_(['в наличии', 'остаток', 'количество']);

  if (idCol === -1 || nameCol === -1) {
    var foundHeaders = Object.keys(colMap).slice(0, 10).join(', ');
    logWarning('⚠️ Не найдены обязательные колонки в "Выгрузка". Первые заголовки: ' + foundHeaders, null, 'readExportSheetProducts_');
    return [];
  }

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var id = row[idCol];
    var name = row[nameCol];
    if (!id || !name) continue;

    var chars = {};
    for (var j = 0; j < charColumns.length; j++) {
      var val = row[charColumns[j].index];
      if (val !== '' && val !== null && val !== undefined) {
        chars[charColumns[j].name] = String(val);
      }
    }

    products.push({
      id: id,
      name: String(name),
      price: priceCol !== -1 ? row[priceCol] : '',
      brand: brandCol !== -1 ? String(row[brandCol]) : '',
      inStock: stockCol !== -1 ? row[stockCol] : '',
      characteristics: chars
    });
  }

  return products;
}

/**
 * Собирает образцы уникальных значений ключевых характеристик
 * Для передачи AI в промпте Фазы 1 — чтобы AI видел реальные данные
 *
 * @param {Array} products — все товары из readExportSheetProducts_
 * @param {Array} charColumns — колонки характеристик [{name, index}]
 * @returns {string} форматированный текст с примерами значений
 */
function getCharacteristicSamples_(products, charColumns) {
  // Считаем заполненность и уникальные значения для каждой характеристики
  var stats = {};
  for (var i = 0; i < charColumns.length; i++) {
    var name = charColumns[i].name;
    stats[name] = { count: 0, values: {} };
  }

  for (var i = 0; i < products.length; i++) {
    var chars = products[i].characteristics;
    var keys = Object.keys(chars);
    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];
      if (stats[key]) {
        stats[key].count++;
        var val = String(chars[key]).trim();
        if (val && val.length < 100) {
          stats[key].values[val] = (stats[key].values[val] || 0) + 1;
        }
      }
    }
  }

  // Сортируем по заполненности (больше → важнее)
  var sorted = Object.keys(stats).sort(function (a, b) {
    return stats[b].count - stats[a].count;
  });

  // Берём топ-10 самых заполненных
  var lines = [];
  var top = Math.min(sorted.length, 10);
  for (var i = 0; i < top; i++) {
    var name = sorted[i];
    var s = stats[name];
    if (s.count < 10) continue; // пропускаем редкие

    // Сортируем значения по частоте
    var vals = Object.keys(s.values).sort(function (a, b) {
      return s.values[b] - s.values[a];
    });
    var sample = vals.slice(0, 15).join(', ');
    lines.push('• ' + name + ' (' + s.count + ' товаров): ' + sample);
  }

  return lines.join('\n');
}

/**
 * Диагностика: показывает заголовки листа "Выгрузка"
 */
function debugExportSheetHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SEO_TAGS_CONFIG.EXPORT_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Лист "' + SEO_TAGS_CONFIG.EXPORT_SHEET_NAME + '" не найден');
    return;
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var list = headers.map(function (h, i) { return (i + 1) + ': ' + h; }).join('\n');
  SpreadsheetApp.getUi().alert('Заголовки "Выгрузка" (' + headers.length + ' колонок):\n\n' + list);
}

/**
 * Форматирует пакет товаров для отправки в AI
 * @param {Array} products - массив товаров
 * @param {Array} charColumns - колонки характеристик [{name, index}]
 * @returns {string} текст для промпта
 */
function formatProductBatchForAI_(products, charColumns) {
  var charNames = charColumns.map(function (c) { return c.name; });
  var MAX_CHARS_PER_PRODUCT = 15; // Ограничиваем кол-во характеристик для экономии токенов

  return products.map(function (p) {
    var charParts = [];
    for (var i = 0; i < charNames.length && charParts.length < MAX_CHARS_PER_PRODUCT; i++) {
      var val = p.characteristics[charNames[i]];
      if (val) {
        // Обрезаем длинные значения
        var valStr = String(val).substring(0, 80);
        charParts.push(charNames[i] + '=' + valStr);
      }
    }
    var charStr = charParts.join('; ');
    var parts = [p.id, p.name, p.price, p.brand, p.inStock];
    if (charStr) parts.push(charStr);
    return parts.join('|');
  }).join('\n');
}

/**
 * Получает ID текущих товаров категории из InSales
 * @param {number} categoryId - ID категории
 * @returns {Array<number>} массив ID товаров
 */
function getExistingProductIds_(categoryId) {
  var context = 'getExistingProductIds_';
  try {
    var credentials = getInsalesCredentialsSync();
    if (!credentials) return [];

    var allIds = [];
    var page = 1;
    var hasMore = true;

    while (hasMore) {
      var url = credentials.baseUrl + '/admin/products.json?collection_id=' + categoryId + '&per_page=250&page=' + page;
      var response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(credentials.apiKey + ':' + credentials.password)
        },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) break;

      var products = JSON.parse(response.getContentText());
      if (!products || products.length === 0) break;

      for (var i = 0; i < products.length; i++) {
        allIds.push(products[i].id);
      }

      hasMore = products.length === 250;
      page++;
      Utilities.sleep(300);
    }

    return allIds;
  } catch (e) {
    logError('Ошибка получения товаров категории', e, context);
    return [];
  }
}

/**
 * Парсинг ответа подбора товаров (JSON с selected_ids или product_ids)
 */
function processProductSelectionResult_(response) {
  try {
    var clean = response.trim();
    clean = clean.replace(/^```json\s*/i, '').replace(/\s*```$/i, '');
    var data = JSON.parse(clean);
    // Поддержка обоих форматов: selected_ids и product_ids
    if (data.selected_ids) return { product_ids: data.selected_ids, comment: data.comment || '' };
    if (data.product_ids) return data;
    return null;
  } catch (e) {
    // Попытка найти JSON массив
    var match = response.match(/\[[\d,\s]+\]/);
    if (match) {
      try {
        var ids = JSON.parse(match[0]);
        return { product_ids: ids };
      } catch (e2) { }
    }
    return null;
  }
}

// ----------------------------------------
// ФАЗА 1: ГЕНЕРАЦИЯ КРИТЕРИЕВ ПОДБОРА
// ----------------------------------------

/**
 * ФАЗА 1: Генерация критериев подбора товаров через AI
 *
 * Для каждой отмеченной строки на листе "SEO-теги":
 * 1. Читает название категории (C) и семантическое ядро (G)
 * 2. Читает заголовки листа "Выгрузка" — список доступных характеристик
 * 3. Отправляет в Gemini → получает критерии подбора
 * 4. Записывает критерии в колонку Q (PRODUCT_CRITERIA)
 *
 * Пользователь затем проверяет/редактирует критерии перед Фазой 2.
 */
function generateProductCriteriaMass() {
  var context = 'generateProductCriteriaMass';
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    logInfo('🧠 Запуск генерации критериев подбора (Фаза 1)', null, context);

    var sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) {
      ui.alert('Ошибка', 'Лист "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '" не найден', ui.ButtonSet.OK);
      return;
    }

    var checkedRows = getCheckedRowsFromSeoTagsSheet_(sheet);
    if (checkedRows.length === 0) {
      ui.alert('Ошибка', 'Не выбрано ни одной строки (чекбокс в колонке A)', ui.ButtonSet.OK);
      return;
    }

    // Читаем товары "Выгрузка" — для заголовков и образцов значений
    ss.toast('Чтение каталога "Выгрузка"...', '⏳', -1);
    var exportData = readExportSheetHeaders_();
    var charList = '';
    var charSamples = '';
    if (exportData && exportData.charColumns.length > 0) {
      charList = exportData.charColumns.map(function (c) { return c.name; }).join(', ');
      // Загружаем товары для сбора образцов значений
      var allProducts = readExportSheetProducts_(exportData);
      if (allProducts.length > 0) {
        charSamples = getCharacteristicSamples_(allProducts, exportData.charColumns);
      }
    } else {
      charList = '(лист "Выгрузка" не найден или пуст — используйте общие критерии)';
    }

    var confirm = ui.alert('Фаза 1: Генерация критериев подбора',
      'Выбрано категорий: ' + checkedRows.length + '\n' +
      'Доступных характеристик: ' + (exportData ? exportData.charColumns.length : 0) + '\n\n' +
      'AI сформулирует критерии подбора для каждой категории.\n' +
      'Результат → колонка Q (текстовый формат ВКЛЮЧИТЬ/ИСКЛЮЧИТЬ).\n' +
      'После генерации вы сможете отредактировать критерии.\n\n' +
      'Продолжить?',
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    var cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    var successCount = 0;

    var systemPrompt = 'Ты подбираешь товары для категории интернет-магазина оптики binokl.shop.\n' +
      'Сформулируй КРИТЕРИИ ФИЛЬТРАЦИИ на основе ХАРАКТЕРИСТИК товаров.\n\n' +
      'ПРИОРИТЕТ ФИЛЬТРАЦИИ (от важного к менее важному):\n' +
      '1. Характеристика "Параметр: Тип товара" — ГЛАВНЫЙ фильтр\n' +
      '2. "Параметр: Назначение" — уточняющий фильтр (если релевантен)\n' +
      '3. Числовые параметры (кратность, диаметр) — если указаны в ключевиках\n' +
      '4. "Название" — ТОЛЬКО для исключений (содержит слова-маркеры нерелевантных товаров)\n\n' +
      'ПРАВИЛА:\n' +
      '- Используй ТОЧНЫЕ названия характеристик из предоставленного списка\n' +
      '- Используй ТОЧНЫЕ значения из предоставленных образцов\n' +
      '- НЕ выдумывай значения — бери только из образцов\n' +
      '- Минимум правил для точного попадания\n\n' +
      'ОПЕРАТОРЫ:\n' +
      '= точное совпадение (одно из значений через запятую)\n' +
      '~ содержит подстроку (одно из значений через запятую)\n' +
      '> числовое больше\n' +
      '< числовое меньше\n\n' +
      'ФОРМАТ ОТВЕТА (строго):\n' +
      'ВКЛЮЧИТЬ:\n' +
      'Параметр: Тип товара = Бинокль\n' +
      'Параметр: Назначение ~ военный, тактический\n\n' +
      'ИСКЛЮЧИТЬ:\n' +
      'Название ~ детский, игрушечный\n\n' +
      'Логика: товар подходит если ВСЕ правила ВКЛЮЧИТЬ выполнены И НИ ОДНО правило ИСКЛЮЧИТЬ не сработало.\n' +
      'Верни ТОЛЬКО правила в указанном формате, без пояснений.';

    for (var i = 0; i < checkedRows.length; i++) {
      var row = checkedRows[i];
      ss.toast('Генерация критериев: ' + row.pageName + ' (' + (i + 1) + '/' + checkedRows.length + ')', '🧠 AI', -1);

      try {
        var userPrompt = 'КАТЕГОРИЯ: ' + row.pageName + '\n' +
          'КЛЮЧЕВЫЕ СЛОВА: ' + (row.semanticCore || 'не указаны') + '\n\n' +
          'ДОСТУПНЫЕ ХАРАКТЕРИСТИКИ В КАТАЛОГЕ:\n' + charList + '\n\n';

        if (charSamples) {
          userPrompt += 'ПРИМЕРЫ ЗНАЧЕНИЙ КЛЮЧЕВЫХ ХАРАКТЕРИСТИК (реальные данные каталога):\n' + charSamples + '\n\n';
        }

        userPrompt += 'Сформулируй критерии подбора товаров для этой категории.';

        var response = callGeminiWithTemperature_(userPrompt, systemPrompt, 0.1);

        if (response && response.trim()) {
          // Валидируем текстовые критерии
          var parsed = parseTextCriteria_(response);
          if (parsed && (parsed.include.length > 0 || parsed.exclude.length > 0)) {
            sheet.getRange(row.rowNumber, cols.PRODUCT_CRITERIA).setValue(response.trim());
            sheet.getRange(row.rowNumber, cols.PRODUCT_CRITERIA).setBackground('#e3f2fd');
            successCount++;
          } else {
            sheet.getRange(row.rowNumber, cols.PRODUCT_CRITERIA).setValue('⚠️ AI вернул непарсируемый ответ:\n' + response.trim());
            sheet.getRange(row.rowNumber, cols.PRODUCT_CRITERIA).setBackground('#fff3e0');
          }
        } else {
          sheet.getRange(row.rowNumber, cols.PRODUCT_CRITERIA).setValue('⚠️ Пустой ответ AI');
          sheet.getRange(row.rowNumber, cols.PRODUCT_CRITERIA).setBackground('#fff3e0');
        }

        Utilities.sleep(SEO_TAGS_CONFIG.DELAY_BETWEEN_ROWS);

      } catch (e) {
        logError('❌ Ошибка генерации критериев для строки ' + row.rowNumber, e, context);
        sheet.getRange(row.rowNumber, cols.PRODUCT_CRITERIA).setValue('Ошибка: ' + e.message);
        sheet.getRange(row.rowNumber, cols.PRODUCT_CRITERIA).setBackground('#ffebee');
      }
    }

    ss.toast('');
    ui.alert('Фаза 1 завершена',
      'Критерии сгенерированы: ' + successCount + '/' + checkedRows.length + '\n\n' +
      'Проверьте критерии в колонке Q (формат ВКЛЮЧИТЬ/ИСКЛЮЧИТЬ).\n' +
      'Затем запустите "Подобрать товары (фильтр)" — подбор мгновенный, без AI.',
      ui.ButtonSet.OK);

    logInfo('✅ Фаза 1 завершена: ' + successCount + '/' + checkedRows.length, null, context);

  } catch (error) {
    logError('❌ Ошибка generateProductCriteriaMass', error, context);
    ui.alert('Ошибка: ' + error.message);
  }
}

// ----------------------------------------
// ФАЗА 2: ПОДБОР ТОВАРОВ — ПРОГРАММНЫЙ ФИЛЬТР
// ----------------------------------------

/**
 * Парсит текстовые критерии из колонки Q
 * Формат:
 *   ВКЛЮЧИТЬ:
 *   Параметр: Тип товара = Бинокль, Монокуляр
 *   Параметр: Назначение ~ военный, тактический
 *   Параметр: Кратность > 6
 *
 *   ИСКЛЮЧИТЬ:
 *   Название ~ детский, игрушечный
 *
 * Операторы: = (equals_any), ~ (contains_any), > (gt), < (lt), != (not_empty)
 *
 * @param {string} text — содержимое ячейки Q
 * @returns {Object|null} { include: [...], exclude: [...] }
 */
function parseTextCriteria_(text) {
  var clean = String(text).trim();
  if (!clean) return null;

  // Обратная совместимость: если JSON — парсим как JSON
  if (clean.charAt(0) === '{') {
    try {
      clean = clean.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '');
      var parsed = JSON.parse(clean);
      if (parsed.include || parsed.exclude) {
        if (!parsed.include) parsed.include = [];
        if (!parsed.exclude) parsed.exclude = [];
        return parsed;
      }
    } catch (e) { /* не JSON — продолжаем текстовый парсинг */ }
  }

  var lines = clean.split('\n');
  var include = [];
  var exclude = [];
  var currentSection = null; // 'include' или 'exclude'

  // Маппинг операторов
  var opMap = {
    '=': 'equals_any',
    '~': 'contains_any',
    '>': 'gt',
    '<': 'lt',
    '!=': 'not_empty'
  };

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line || line.charAt(0) === '/' || line === '---') continue;

    // Определяем секцию
    var lineLower = line.toLowerCase();
    if (lineLower.indexOf('включить') !== -1 && line.indexOf(':') !== -1) {
      currentSection = 'include';
      continue;
    }
    if (lineLower.indexOf('исключить') !== -1 && line.indexOf(':') !== -1) {
      currentSection = 'exclude';
      continue;
    }

    if (!currentSection) continue;

    // Парсим правило: "Field op value1, value2"
    // Пробуем операторы от длинных к коротким: !=, ~, =, >, <
    var rule = null;
    var ops = ['!=', '~', '=', '>', '<'];
    for (var oi = 0; oi < ops.length; oi++) {
      var op = ops[oi];
      var searchStr = ' ' + op + ' ';
      var idx = line.indexOf(searchStr);
      var matchLen = searchStr.length; // длина " op " для правильного substring

      if (idx === -1) {
        // Пробуем без пробела перед оператором
        searchStr = op + ' ';
        idx = line.indexOf(searchStr);
        matchLen = searchStr.length;
      }
      if (idx === -1 && op.length === 1) {
        idx = line.lastIndexOf(op);
        matchLen = op.length;
      }

      if (idx !== -1) {
        var field = line.substring(0, idx).trim();
        var valPart = line.substring(idx + matchLen).trim();

        var values = [];
        if (op === '!=') {
          values = []; // not_empty не требует значений
        } else {
          values = valPart.split(',').map(function (v) { return v.trim(); }).filter(function (v) { return v; });
        }

        rule = {
          field: field,
          op: opMap[op] || op,
          values: values
        };
        break;
      }
    }

    if (rule) {
      if (currentSection === 'include') include.push(rule);
      else exclude.push(rule);
    }
  }

  if (include.length === 0 && exclude.length === 0) return null;
  return { include: include, exclude: exclude };
}

/**
 * Проверяет одно правило для товара
 * @param {Object} product — товар из readExportSheetProducts_
 * @param {Object} rule — {field, op, values}
 * @returns {boolean}
 */
function checkProductRule_(product, rule) {
  var field = rule.field;
  var op = rule.op;
  var values = rule.values || [];

  // Получаем значение поля товара
  var fieldValue = '';
  var fieldLower = field.toLowerCase();
  if (fieldLower === 'название' || fieldLower === 'name') {
    fieldValue = product.name || '';
  } else if (fieldLower === 'бренд' || fieldLower === 'brand' || fieldLower === 'производитель') {
    fieldValue = product.brand || '';
  } else if (fieldLower === 'цена' || fieldLower === 'price' || fieldLower === 'цена продажи') {
    fieldValue = String(product.price || '');
  } else if (fieldLower === 'в наличии' || fieldLower === 'остаток') {
    fieldValue = String(product.inStock || '');
  } else {
    // Ищем в характеристиках (точное и нечёткое совпадение имени)
    fieldValue = product.characteristics[field] || '';
    if (!fieldValue) {
      var charKeys = Object.keys(product.characteristics);
      for (var k = 0; k < charKeys.length; k++) {
        if (charKeys[k].toLowerCase() === fieldLower) {
          fieldValue = product.characteristics[charKeys[k]];
          break;
        }
      }
    }
    if (!fieldValue) fieldValue = '';
  }

  var fvLower = String(fieldValue).toLowerCase().trim();

  switch (op) {
    case 'contains_any':
      for (var i = 0; i < values.length; i++) {
        if (fvLower.indexOf(values[i].toLowerCase()) !== -1) return true;
      }
      return false;

    case 'contains_all':
      for (var i = 0; i < values.length; i++) {
        if (fvLower.indexOf(values[i].toLowerCase()) === -1) return false;
      }
      return true;

    case 'equals_any':
      for (var i = 0; i < values.length; i++) {
        if (fvLower === values[i].toLowerCase()) return true;
      }
      return false;

    case 'not_empty':
      return fvLower !== '' && fvLower !== '0' && fvLower !== 'null' && fvLower !== 'false';

    case 'gt':
      var numVal = parseFloat(fieldValue);
      var threshold = parseFloat(values[0]);
      return !isNaN(numVal) && !isNaN(threshold) && numVal > threshold;

    case 'lt':
      var numVal = parseFloat(fieldValue);
      var threshold = parseFloat(values[0]);
      return !isNaN(numVal) && !isNaN(threshold) && numVal < threshold;

    default:
      return false;
  }
}

/**
 * Фильтрует товары по JSON-критериям (программно, без AI)
 * @param {Array} products — все товары из Выгрузка
 * @param {Object} criteria — { include: [...], exclude: [...] }
 * @returns {Array} отфильтрованные товары
 */
function filterProductsByCriteria_(products, criteria) {
  var result = [];
  for (var i = 0; i < products.length; i++) {
    var p = products[i];

    // Все include-правила должны выполниться (AND)
    var includeOk = true;
    for (var j = 0; j < criteria.include.length; j++) {
      if (!checkProductRule_(p, criteria.include[j])) {
        includeOk = false;
        break;
      }
    }
    if (!includeOk) continue;

    // Ни одно exclude-правило не должно сработать
    var excluded = false;
    for (var j = 0; j < criteria.exclude.length; j++) {
      if (checkProductRule_(p, criteria.exclude[j])) {
        excluded = true;
        break;
      }
    }
    if (excluded) continue;

    result.push(p);
  }
  return result;
}

/**
 * ФАЗА 2: Программный подбор товаров из "Выгрузка" по JSON-критериям из колонки Q
 *
 * Работает МГНОВЕННО — без вызовов AI.
 * 1. Парсит JSON-критерии из колонки Q
 * 2. Читает все товары из "Выгрузка"
 * 3. Применяет фильтры программно
 * 4. Записывает ID в колонку R
 */
function selectProductsByCriteriaMass() {
  var context = 'selectProductsByCriteriaMass';
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    logInfo('🛍️ Запуск программного подбора товаров (Фаза 2)', null, context);

    var sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) {
      ui.alert('Ошибка', 'Лист "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '" не найден', ui.ButtonSet.OK);
      return;
    }

    var checkedRows = getCheckedRowsFromSeoTagsSheet_(sheet);
    if (checkedRows.length === 0) {
      ui.alert('Ошибка', 'Не выбрано ни одной строки', ui.ButtonSet.OK);
      return;
    }

    var cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    var rowsWithCriteria = [];
    for (var i = 0; i < checkedRows.length; i++) {
      var rawCriteria = sheet.getRange(checkedRows[i].rowNumber, cols.PRODUCT_CRITERIA).getValue();
      if (!rawCriteria || String(rawCriteria).startsWith('⚠️') || String(rawCriteria).startsWith('Ошибка')) continue;

      var parsed = parseTextCriteria_(rawCriteria);
      if (parsed) {
        checkedRows[i].criteria = parsed;
        checkedRows[i].criteriaRaw = String(rawCriteria).trim();
        rowsWithCriteria.push(checkedRows[i]);
      }
    }

    if (rowsWithCriteria.length === 0) {
      ui.alert('Ошибка',
        'Ни в одной строке нет валидных критериев (колонка Q).\n' +
        'Сначала запустите "Сформировать критерии подбора (AI)".\n\n' +
        'Критерии должны содержать секции ВКЛЮЧИТЬ: и/или ИСКЛЮЧИТЬ:',
        ui.ButtonSet.OK);
      return;
    }

    // Читаем товары из "Выгрузка"
    ss.toast('Чтение каталога "Выгрузка"...', '⏳', -1);
    var exportData = readExportSheetHeaders_();
    if (!exportData) {
      ui.alert('Ошибка', 'Лист "' + SEO_TAGS_CONFIG.EXPORT_SHEET_NAME + '" не найден', ui.ButtonSet.OK);
      return;
    }

    var allProducts = readExportSheetProducts_(exportData);
    if (allProducts.length === 0) {
      ui.alert('Ошибка', 'Лист "Выгрузка" пуст или не содержит товаров', ui.ButtonSet.OK);
      return;
    }

    var confirm = ui.alert('Фаза 2: Программный подбор',
      'Категорий с критериями: ' + rowsWithCriteria.length + '\n' +
      'Товаров в "Выгрузка": ' + allProducts.length + '\n\n' +
      'Подбор мгновенный (программный фильтр, без AI).\n' +
      'Результат → колонка R.\n\n' +
      'Продолжить?',
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    var successCount = 0;

    for (var ri = 0; ri < rowsWithCriteria.length; ri++) {
      var row = rowsWithCriteria[ri];
      ss.toast('Фильтрация: ' + row.pageName + ' (' + (ri + 1) + '/' + rowsWithCriteria.length + ')', '🔍', -1);

      try {
        var matched = filterProductsByCriteria_(allProducts, row.criteria);

        if (matched.length > 0) {
          // Формат: ID | Название (каждый товар на отдельной строке)
          var lines = matched.map(function (p) {
            return p.id + ' | ' + p.name;
          });
          lines.push('---');
          lines.push('Найдено: ' + matched.length + ' из ' + allProducts.length);
          var resultText = lines.join('\n');
          sheet.getRange(row.rowNumber, cols.PRODUCTS_RESULT).setValue(resultText);
          sheet.getRange(row.rowNumber, cols.PRODUCTS_RESULT).setBackground('#e8f5e9');
          successCount++;
          logInfo('✅ Подобрано ' + matched.length + ' товаров для "' + row.pageName + '"', null, context);
        } else {
          // Диагностика: показываем что парсер увидел + образцы реальных значений
          var diagLines = ['Нет подходящих товаров (0 из ' + allProducts.length + ')', ''];
          diagLines.push('--- ДИАГНОСТИКА ---');
          diagLines.push('Распарсенные критерии:');
          var allRules = (row.criteria.include || []).concat(row.criteria.exclude || []);
          for (var di = 0; di < allRules.length; di++) {
            var r = allRules[di];
            diagLines.push('  ' + (di < (row.criteria.include || []).length ? 'ВКЛЮЧИТЬ' : 'ИСКЛЮЧИТЬ') +
              ': "' + r.field + '" ' + r.op + ' [' + (r.values || []).join(', ') + ']');
            // Показать реальные значения этого поля у первых 5 товаров
            var samples = [];
            for (var si = 0; si < Math.min(allProducts.length, 20); si++) {
              var pv = '';
              var fLower = r.field.toLowerCase();
              if (fLower === 'название' || fLower === 'name') pv = allProducts[si].name;
              else {
                pv = allProducts[si].characteristics[r.field] || '';
                if (!pv) {
                  var ck = Object.keys(allProducts[si].characteristics);
                  for (var ci = 0; ci < ck.length; ci++) {
                    if (ck[ci].toLowerCase() === fLower) { pv = allProducts[si].characteristics[ck[ci]]; break; }
                  }
                }
              }
              if (pv && samples.indexOf(String(pv)) === -1) samples.push(String(pv));
              if (samples.length >= 5) break;
            }
            diagLines.push('  Реальные значения: ' + (samples.length > 0 ? samples.join(' | ') : '(поле не найдено у товаров!)'));
          }
          // Показать доступные ключи характеристик из первого товара
          if (allProducts.length > 0) {
            var charKeys = Object.keys(allProducts[0].characteristics).slice(0, 15);
            diagLines.push('');
            diagLines.push('Доступные характеристики (первый товар):');
            diagLines.push('  ' + charKeys.join(', '));
          }
          sheet.getRange(row.rowNumber, cols.PRODUCTS_RESULT).setValue(diagLines.join('\n'));
          sheet.getRange(row.rowNumber, cols.PRODUCTS_RESULT).setBackground('#fff3e0');
          logWarning('⚠️ 0 товаров для "' + row.pageName + '". Проверьте критерии.', null, context);
        }

      } catch (e) {
        logError('❌ Ошибка фильтрации для строки ' + row.rowNumber, e, context);
        sheet.getRange(row.rowNumber, cols.PRODUCTS_RESULT).setValue('Ошибка: ' + e.message);
        sheet.getRange(row.rowNumber, cols.PRODUCTS_RESULT).setBackground('#ffebee');
      }
    }

    ss.toast('');
    ui.alert('Фаза 2 завершена',
      'Подобраны товары для ' + successCount + '/' + rowsWithCriteria.length + ' категорий.\n\n' +
      'Проверьте результаты в колонке R.\n' +
      'Затем запустите "Отправить подобранные товары в InSales".',
      ui.ButtonSet.OK);

    logInfo('✅ Фаза 2 завершена: ' + successCount + '/' + rowsWithCriteria.length, null, context);

  } catch (error) {
    logError('❌ Ошибка selectProductsByCriteriaMass', error, context);
    ui.alert('Ошибка: ' + error.message);
  }
}

// ----------------------------------------
// ФАЗА 3: ЭКСПОРТ ТОВАРОВ В INSALES
// ----------------------------------------

/**
 * ФАЗА 3: Отправка подобранных товаров в InSales
 *
 * 1. Читает отмеченные строки + ID товаров из колонки R
 * 2. Добавляет товары в категории через POST /admin/collects.json
 * 3. Обновляет статус в колонке R
 */
function sendSelectedProductsToInSales() {
  var context = 'sendSelectedProductsToInSales';
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    logInfo('📤 Запуск экспорта товаров в InSales (Фаза 3)', null, context);

    var sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) {
      ui.alert('Ошибка', 'Лист "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '" не найден', ui.ButtonSet.OK);
      return;
    }

    var checkedRows = getCheckedRowsFromSeoTagsSheet_(sheet);
    if (checkedRows.length === 0) {
      ui.alert('Ошибка', 'Не выбрано ни одной строки', ui.ButtonSet.OK);
      return;
    }

    var cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

    // Собираем строки, у которых есть товары в колонке R
    var rowsToSend = [];
    for (var i = 0; i < checkedRows.length; i++) {
      var row = checkedRows[i];
      var categoryId = parseInt(row.id);
      if (!categoryId || categoryId <= 0) continue;

      var rawValue = String(sheet.getRange(row.rowNumber, cols.PRODUCTS_RESULT).getValue());
      // Парсим ID из формата "ID | Название" (по строке на товар)
      // Также поддерживает старый формат "ID, ID, ID" (через запятую)
      var lines = rawValue.split('\n');
      var ids = [];
      for (var li = 0; li < lines.length; li++) {
        var line = lines[li].trim();
        if (!line || line === '---' || line.indexOf('Найдено:') !== -1 || line.indexOf('//') === 0 || line.indexOf('✅') !== -1) continue;
        // Формат "ID | Название" — берём часть до |
        if (line.indexOf('|') !== -1) {
          var idPart = parseInt(line.split('|')[0].trim());
          if (!isNaN(idPart) && idPart > 0) ids.push(idPart);
        } else {
          // Старый формат: ID через запятую
          var parts = line.split(/[,\s]+/);
          for (var pi = 0; pi < parts.length; pi++) {
            var n = parseInt(parts[pi].trim());
            if (!isNaN(n) && n > 0) ids.push(n);
          }
        }
      }

      if (ids.length > 0) {
        rowsToSend.push({
          rowNumber: row.rowNumber,
          categoryId: categoryId,
          categoryName: row.pageName,
          productIds: ids
        });
      }
    }

    if (rowsToSend.length === 0) {
      ui.alert('Ошибка',
        'Нет строк с подобранными товарами в колонке R.\n' +
        'Сначала запустите "Подобрать товары (AI)".',
        ui.ButtonSet.OK);
      return;
    }

    var totalProducts = 0;
    for (var j = 0; j < rowsToSend.length; j++) {
      totalProducts += rowsToSend[j].productIds.length;
    }

    var confirm = ui.alert('Фаза 3: Синхронизация товаров в InSales',
      'Категорий: ' + rowsToSend.length + '\n' +
      'Всего товаров: ' + totalProducts + '\n\n' +
      'ПОЛНАЯ СИНХРОНИЗАЦИЯ:\n' +
      '• Новые товары будут ДОБАВЛЕНЫ\n' +
      '• Товары, которых нет в списке, будут УДАЛЕНЫ из категории\n\n' +
      'Продолжить?',
      ui.ButtonSet.YES_NO);

    if (confirm !== ui.Button.YES) return;

    var startTime = Date.now();
    var MAX_RUNTIME_MS = 300000; // 5 минут
    var totalAdded = 0;
    var totalRemoved = 0;
    var totalKept = 0;
    var totalErrors = 0;
    var aborted = false;

    for (var ri = 0; ri < rowsToSend.length; ri++) {
      if (Date.now() - startTime > MAX_RUNTIME_MS) {
        aborted = true;
        break;
      }

      var item = rowsToSend[ri];
      ss.toast('Синхронизация ' + item.productIds.length + ' товаров в "' + item.categoryName + '" (' + (ri + 1) + '/' + rowsToSend.length + ')', '🔄 InSales', -1);

      try {
        var result = syncProductsInCategory(item.categoryId, item.productIds);

        totalAdded += result.added;
        totalRemoved += result.removed;
        totalKept += result.kept;
        totalErrors += result.errors;

        // Обновляем статус в R
        var currentVal = sheet.getRange(item.rowNumber, cols.PRODUCTS_RESULT).getValue();
        // Сохраняем только строки с ID товаров (без старых статусов)
        var cleanLines = [];
        var rawLines = String(currentVal).split('\n');
        for (var cl = 0; cl < rawLines.length; cl++) {
          var cLine = rawLines[cl].trim();
          if (cLine && cLine.indexOf('✅') === -1 && cLine.indexOf('❌') === -1 && cLine.indexOf('🔄') === -1) {
            cleanLines.push(cLine);
          }
        }
        var statusText = cleanLines.join('\n');
        statusText += '\n🔄 Синхронизировано: +' + result.added + ' / -' + result.removed + ' / =' + result.kept;
        if (result.errors > 0) {
          statusText += ' (ошибок: ' + result.errors + ')';
        }
        sheet.getRange(item.rowNumber, cols.PRODUCTS_RESULT).setValue(statusText);
        sheet.getRange(item.rowNumber, cols.PRODUCTS_RESULT).setBackground(
          result.errors === 0 ? '#c8e6c9' : '#fff9c4'
        );

        logInfo('🔄 Синхронизировано "' + item.categoryName + '": +' + result.added + ' / -' + result.removed + ' / =' + result.kept, null, context);

      } catch (e) {
        logError('❌ Ошибка синхронизации для "' + item.categoryName + '"', e, context);
        sheet.getRange(item.rowNumber, cols.PRODUCTS_RESULT).setBackground('#ffebee');
        totalErrors += item.productIds.length;
      }
    }

    ss.toast('');

    if (aborted) {
      ui.alert('Таймаут',
        'Обработано ' + (ri) + '/' + rowsToSend.length + ' категорий.\n' +
        'Добавлено: ' + totalAdded + ', удалено: ' + totalRemoved + ', без изменений: ' + totalKept + '\n' +
        'Ошибок: ' + totalErrors + '\n\n' +
        'Снимите галочки с обработанных строк и запустите повторно.',
        ui.ButtonSet.OK);
    } else {
      ui.alert('Фаза 3 завершена',
        'Добавлено: ' + totalAdded + '\n' +
        'Удалено: ' + totalRemoved + '\n' +
        'Без изменений: ' + totalKept + '\n' +
        'Ошибок: ' + totalErrors,
        ui.ButtonSet.OK);
    }

    logInfo('✅ Фаза 3 завершена. +' + totalAdded + ' / -' + totalRemoved + ' / =' + totalKept + ' / ошибок: ' + totalErrors, null, context);

  } catch (error) {
    logError('❌ Ошибка sendSelectedProductsToInSales', error, context);
    ui.alert('Ошибка: ' + error.message);
  }
}

// ----------------------------------------
// ВАРИАНТ 3: ДАМП ТОВАРОВ ДЛЯ АГЕНТА
// ----------------------------------------

/**
 * Создаёт компактный дамп товаров из "Выгрузка" на новый лист "Дамп для агента"
 * Пользователь может скопировать текст и передать агенту (Claude Code) для ручного анализа
 *
 * Формат: ID|Название|Бренд|Цена|Наличие|Хар1=Знач1; Хар2=Знач2
 * Ограничение: 5 ключевых характеристик на товар для компактности
 */
function dumpProductsForAgent() {
  var context = 'dumpProductsForAgent';
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    logInfo('📋 Дамп товаров для агента', null, context);

    var exportData = readExportSheetHeaders_();
    if (!exportData) {
      ui.alert('Ошибка', 'Лист "' + SEO_TAGS_CONFIG.EXPORT_SHEET_NAME + '" не найден', ui.ButtonSet.OK);
      return;
    }

    var allProducts = readExportSheetProducts_(exportData);
    if (allProducts.length === 0) {
      ui.alert('Ошибка', 'Лист "Выгрузка" пуст', ui.ButtonSet.OK);
      return;
    }

    // Собираем статистику по характеристикам (какие встречаются чаще всего)
    var charFreq = {};
    for (var i = 0; i < allProducts.length; i++) {
      var keys = Object.keys(allProducts[i].characteristics);
      for (var j = 0; j < keys.length; j++) {
        charFreq[keys[j]] = (charFreq[keys[j]] || 0) + 1;
      }
    }

    // Топ характеристик по частоте
    var charNames = Object.keys(charFreq).sort(function (a, b) { return charFreq[b] - charFreq[a]; });
    var topChars = charNames.slice(0, 20); // Топ-20 самых частых

    // Формируем компактный текст
    var lines = [];
    lines.push('КАТАЛОГ ТОВАРОВ (' + allProducts.length + ' шт.)');
    lines.push('Характеристики: ' + topChars.join(', '));
    lines.push('---');
    lines.push('ID|Название|Бренд|Цена|Наличие|Характеристики');

    for (var i = 0; i < allProducts.length; i++) {
      var p = allProducts[i];
      var charParts = [];
      for (var j = 0; j < topChars.length && charParts.length < 8; j++) {
        var val = p.characteristics[topChars[j]];
        if (val) charParts.push(topChars[j] + '=' + String(val).substring(0, 50));
      }
      lines.push(p.id + '|' + p.name + '|' + p.brand + '|' + p.price + '|' + p.inStock + '|' + charParts.join('; '));
    }

    // Записываем на лист
    var dumpSheetName = 'Дамп для агента';
    var dumpSheet = ss.getSheetByName(dumpSheetName);
    if (!dumpSheet) {
      dumpSheet = ss.insertSheet(dumpSheetName);
    }
    dumpSheet.clear();

    var text = lines.join('\n');
    dumpSheet.getRange(1, 1).setValue(text);
    dumpSheet.setColumnWidth(1, 800);

    // Также добавляем отмеченные категории если есть
    var seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (seoSheet) {
      var checkedRows = getCheckedRowsFromSeoTagsSheet_(seoSheet);
      if (checkedRows.length > 0) {
        var catLines = ['\n\nКАТЕГОРИИ ДЛЯ ПОДБОРА:'];
        for (var i = 0; i < checkedRows.length; i++) {
          var r = checkedRows[i];
          catLines.push('ID=' + r.id + ' | ' + r.pageName + ' | Ключевики: ' + (r.semanticCore || '—').substring(0, 200));
        }
        dumpSheet.getRange(1, 1).setValue(text + '\n' + catLines.join('\n'));
      }
    }

    ss.setActiveSheet(dumpSheet);
    ui.alert('Дамп готов',
      'Лист "' + dumpSheetName + '" создан.\n' +
      'Товаров: ' + allProducts.length + '\n\n' +
      'Скопируйте содержимое ячейки A1 и передайте агенту (Claude Code).\n' +
      'Агент проанализирует товары и подберёт ID для каждой категории.',
      ui.ButtonSet.OK);

    logInfo('✅ Дамп создан: ' + allProducts.length + ' товаров', null, context);

  } catch (error) {
    logError('❌ Ошибка dumpProductsForAgent', error, context);
    ui.alert('Ошибка: ' + error.message);
  }
}

/**
 * Объединение выделенных кластеров
 * 1. Берет выделенные строки
 * 2. Объединяет ключевики (H) и конкурентов (I) в первую строку
 * 3. Удаляет остальные строки
 */
function combineSelectedClusters() {
  const context = 'combineSelectedClusters';
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!sheet) return;

    const checkedRows = getCheckedRowsFromSeoTagsSheet_(sheet);
    if (checkedRows.length < 2) {
      ui.alert('Нужно выбрать минимум 2 строки для объединения');
      return;
    }

    // Сортируем по номеру строки, чтобы первая была основной
    checkedRows.sort((a, b) => a.rowNumber - b.rowNumber);

    const mainRow = checkedRows[0];
    const otherRows = checkedRows.slice(1);
    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

    // Сбор данных
    let combinedKeywords = (sheet.getRange(mainRow.rowNumber, cols.SEMANTIC_CORE).getValue() || '').toString();
    let combinedCompetitors = (sheet.getRange(mainRow.rowNumber, cols.COMPETITORS_TITLE).getValue() || '').toString();

    for (const row of otherRows) {
      const kw = sheet.getRange(row.rowNumber, cols.SEMANTIC_CORE).getValue();
      const comp = sheet.getRange(row.rowNumber, cols.COMPETITORS_TITLE).getValue();

      if (kw) combinedKeywords += '\n' + kw;
      if (comp) combinedCompetitors += '\n' + comp;
    }

    // Очистка дублей (простая, по строкам)
    const uniqueKw = [...new Set(combinedKeywords.split('\n').map(s => s.trim()).filter(s => s))].join('\n');
    const uniqueComp = [...new Set(combinedCompetitors.split('\n').map(s => s.trim()).filter(s => s))].join('\n');

    // Запись в основную строку
    sheet.getRange(mainRow.rowNumber, cols.SEMANTIC_CORE).setValue(uniqueKw);
    sheet.getRange(mainRow.rowNumber, cols.COMPETITORS_TITLE).setValue(uniqueComp);

    // Удаление остальных строк (с конца, чтобы не сбить индексы)
    // Но так как мы удаляем строки, нужно быть осторожным. API deleteRow удаляет 1 строку.
    // Лучше удалять пачкой или с конца.

    const rowsToDelete = otherRows.map(r => r.rowNumber).sort((a, b) => b - a);

    for (const rowNum of rowsToDelete) {
      sheet.deleteRow(rowNum);
    }

    ss.toast('Кластеры объединены', '✅ Готово', 3);

  } catch (error) {
    logError('❌ Ошибка объединения кластеров', error, context);
    ui.alert('Ошибка: ' + error.message);
  }
}


// ================================================
// ДИАЛОГ ПОДБОРА ТОВАРОВ
// ================================================

/**
 * Открывает диалог подбора товаров для текущей строки в SEO-теги
 */
function showProductSelectionDialog() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
  if (!sheet) {
    ui.alert('Ошибка', 'Лист "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '" не найден', ui.ButtonSet.OK);
    return;
  }

  var activeCell = sheet.getActiveCell();
  if (!activeCell || activeCell.getRow() < SEO_TAGS_CONFIG.MASS_DATA_START_ROW) {
    ui.alert('Ошибка', 'Выберите строку категории (от строки ' + SEO_TAGS_CONFIG.MASS_DATA_START_ROW + ')', ui.ButtonSet.OK);
    return;
  }

  var rowNumber = activeCell.getRow();
  var cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
  var pageName = sheet.getRange(rowNumber, cols.PAGE_NAME).getValue();
  if (!pageName) {
    ui.alert('Ошибка', 'Строка ' + rowNumber + ' не содержит названия категории', ui.ButtonSet.OK);
    return;
  }

  // Передаём rowNumber в HTML через scriptlet-подход (шаблон)
  var html = HtmlService.createHtmlOutputFromFile('30_product_selection_dialog')
    .setWidth(900)
    .setHeight(750);

  // Передаём rowNumber через PropertiesService (временно)
  PropertiesService.getScriptProperties().setProperty('DIALOG_ROW_NUMBER', String(rowNumber));

  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Подбор товаров: ' + pageName);
}

/**
 * Возвращает начальные данные для диалога подбора товаров
 * @returns {Object} {categoryName, semanticCore, rowNumber, characteristics, existingCriteria}
 */
function getProductSelectionInitData() {
  var rowNumber = parseInt(PropertiesService.getScriptProperties().getProperty('DIALOG_ROW_NUMBER') || '0');
  if (rowNumber < 2) throw new Error('Не указана строка категории');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
  var cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

  var categoryName = sheet.getRange(rowNumber, cols.PAGE_NAME).getValue() || '';
  var semanticCore = sheet.getRange(rowNumber, cols.SEMANTIC_CORE).getValue() || '';
  var existingCriteria = sheet.getRange(rowNumber, cols.PRODUCT_CRITERIA).getValue() || '';

  // Читаем характеристики из "Выгрузка"
  var exportData = readExportSheetHeaders_();
  var characteristics = [];

  if (exportData && exportData.charColumns.length > 0) {
    var allProducts = readExportSheetProducts_(exportData);

    // Собираем уникальные значения для каждой характеристики
    var stats = {};
    for (var i = 0; i < exportData.charColumns.length; i++) {
      var name = exportData.charColumns[i].name;
      stats[name] = { count: 0, values: {} };
    }

    for (var i = 0; i < allProducts.length; i++) {
      var chars = allProducts[i].characteristics;
      var keys = Object.keys(chars);
      for (var j = 0; j < keys.length; j++) {
        var key = keys[j];
        if (stats[key]) {
          stats[key].count++;
          var val = String(chars[key]).trim();
          if (val && val.length < 100) {
            stats[key].values[val] = (stats[key].values[val] || 0) + 1;
          }
        }
      }
    }

    // Сортируем по заполненности
    var sorted = Object.keys(stats).sort(function (a, b) {
      return stats[b].count - stats[a].count;
    });

    for (var i = 0; i < sorted.length; i++) {
      var name = sorted[i];
      var s = stats[name];
      if (s.count < 5) continue; // Пропускаем совсем редкие

      // Сортируем значения по частоте, берём топ-50
      var vals = Object.keys(s.values)
        .sort(function (a, b) { return s.values[b] - s.values[a]; })
        .slice(0, 50)
        .map(function (v) { return { value: v, count: s.values[v] }; });

      characteristics.push({
        name: name,
        count: s.count,
        values: vals
      });
    }
  }

  // Загружаем текущие товары категории из InSales API
  var categoryId = parseInt(sheet.getRange(rowNumber, cols.ID).getValue());
  var existingProducts = []; // [{id, name, inStock}]
  if (categoryId > 0) {
    try {
      var credentials = getInsalesCredentialsSync();
      if (credentials) {
        var page = 1;
        var hasMore = true;
        while (hasMore) {
          var url = credentials.baseUrl + '/admin/products.json?collection_id=' + categoryId + '&per_page=250&page=' + page;
          var resp = UrlFetchApp.fetch(url, {
            method: 'GET',
            headers: {
              'Authorization': 'Basic ' + Utilities.base64Encode(credentials.apiKey + ':' + credentials.password)
            },
            muteHttpExceptions: true
          });
          if (resp.getResponseCode() !== 200) break;
          var prods = JSON.parse(resp.getContentText());
          if (!prods || prods.length === 0) break;
          for (var pi = 0; pi < prods.length; pi++) {
            var p = prods[pi];
            // Определяем наличие по available или variants
            var inStock = false;
            if (p.available !== undefined) {
              inStock = !!p.available;
            } else if (p.variants && p.variants.length > 0) {
              for (var vi = 0; vi < p.variants.length; vi++) {
                if (p.variants[vi].quantity > 0) { inStock = true; break; }
              }
            }
            existingProducts.push({
              id: p.id,
              name: p.title || p.name || ('ID ' + p.id),
              inStock: inStock
            });
          }
          hasMore = prods.length === 250;
          page++;
          Utilities.sleep(300);
        }
      }
    } catch (e) {
      logError('Ошибка загрузки товаров категории из InSales', e, 'getProductSelectionInitData');
    }
  }

  return {
    rowNumber: rowNumber,
    categoryId: categoryId,
    categoryName: categoryName,
    semanticCore: semanticCore,
    existingCriteria: existingCriteria,
    characteristics: characteristics,
    existingProducts: existingProducts // [{id, name, inStock}] — реальные товары из InSales
  };
}

/**
 * Генерирует AI-критерии для диалога
 * @param {string} categoryName - название категории
 * @param {string} semanticCore - ключевые слова
 * @returns {string} текстовые критерии
 */
function generateAICriteriaForDialog(categoryName, semanticCore) {
  var exportData = readExportSheetHeaders_();
  var charList = '';
  var charSamples = '';

  if (exportData && exportData.charColumns.length > 0) {
    charList = exportData.charColumns.map(function (c) { return c.name; }).join(', ');
    var allProducts = readExportSheetProducts_(exportData);
    if (allProducts.length > 0) {
      charSamples = getCharacteristicSamples_(allProducts, exportData.charColumns);
    }
  }

  var systemPrompt = 'Ты подбираешь товары для категории интернет-магазина оптики binokl.shop.\n' +
    'Сформулируй КРИТЕРИИ ФИЛЬТРАЦИИ на основе ХАРАКТЕРИСТИК товаров.\n\n' +
    'ПРИОРИТЕТ ФИЛЬТРАЦИИ (от важного к менее важному):\n' +
    '1. Характеристика "Параметр: Тип товара" — ГЛАВНЫЙ фильтр\n' +
    '2. "Параметр: Назначение" — уточняющий фильтр (если релевантен)\n' +
    '3. Числовые параметры (кратность, диаметр) — если указаны в ключевиках\n' +
    '4. "Название" — ТОЛЬКО для исключений\n\n' +
    'ПРАВИЛА:\n' +
    '- Используй ТОЧНЫЕ названия характеристик из предоставленного списка\n' +
    '- Используй ТОЧНЫЕ значения из предоставленных образцов\n' +
    '- НЕ выдумывай значения — бери только из образцов\n' +
    '- Минимум правил для точного попадания\n\n' +
    'ОПЕРАТОРЫ:\n' +
    '= точное совпадение (одно из значений через запятую)\n' +
    '~ содержит подстроку (одно из значений через запятую)\n' +
    '> числовое больше\n' +
    '< числовое меньше\n\n' +
    'ФОРМАТ ОТВЕТА (строго):\n' +
    'ВКЛЮЧИТЬ:\n' +
    'Параметр: Тип товара = Бинокль\n' +
    'Параметр: Назначение ~ военный, тактический\n\n' +
    'ИСКЛЮЧИТЬ:\n' +
    'Название ~ детский, игрушечный\n\n' +
    'Логика: товар подходит если ВСЕ правила ВКЛЮЧИТЬ выполнены И НИ ОДНО правило ИСКЛЮЧИТЬ не сработало.\n' +
    'Верни ТОЛЬКО правила в указанном формате, без пояснений.';

  var userPrompt = 'КАТЕГОРИЯ: ' + categoryName + '\n' +
    'КЛЮЧЕВЫЕ СЛОВА: ' + (semanticCore || 'не указаны') + '\n\n' +
    'ДОСТУПНЫЕ ХАРАКТЕРИСТИКИ В КАТАЛОГЕ:\n' + charList + '\n\n';

  if (charSamples) {
    userPrompt += 'ПРИМЕРЫ ЗНАЧЕНИЙ КЛЮЧЕВЫХ ХАРАКТЕРИСТИК (реальные данные каталога):\n' + charSamples + '\n\n';
  }

  userPrompt += 'Сформулируй критерии подбора товаров для этой категории.';

  return callGeminiWithTemperature_(userPrompt, systemPrompt, 0.1);
}

/**
 * Фильтрует товары по правилам из диалога
 * @param {string} rulesJSON - JSON-строка с {include: [...], exclude: [...]}
 * @param {string} existingIdsJSON - JSON-строка с массивом ID существующих товаров
 * @returns {Object} {matched, existingNotMatched, matchedTotal, total, truncated}
 */
function filterProductsForDialog(rulesJSON, existingProductsJSON) {
  var criteria = JSON.parse(rulesJSON);
  // existingProducts — массив {id, name, inStock} из InSales API
  var existingArr = existingProductsJSON ? JSON.parse(existingProductsJSON) : [];
  var existingSet = {};
  var existingById = {};
  for (var ei = 0; ei < existingArr.length; ei++) {
    var ep = existingArr[ei];
    existingSet[String(ep.id)] = true;
    existingById[String(ep.id)] = ep;
  }

  var exportData = readExportSheetHeaders_();
  if (!exportData) throw new Error('Лист "Выгрузка" не найден');

  var allProducts = readExportSheetProducts_(exportData);
  if (allProducts.length === 0) throw new Error('Лист "Выгрузка" пуст');

  var matched = filterProductsByCriteria_(allProducts, criteria);

  // Множество ID подобранных товаров
  var matchedSet = {};
  for (var mi = 0; mi < matched.length; mi++) {
    matchedSet[String(matched[mi].id)] = true;
  }

  // Существующие товары, которые НЕ попали в новый подбор
  // Берём данные из InSales (existingArr), т.к. в "Выгрузке" может не быть всех товаров
  var existingNotMatched = [];
  for (var ei2 = 0; ei2 < existingArr.length; ei2++) {
    var eid = String(existingArr[ei2].id);
    if (!matchedSet[eid]) {
      existingNotMatched.push({
        id: existingArr[ei2].id,
        name: existingArr[ei2].name,
        inStock: existingArr[ei2].inStock
      });
    }
  }

  // Ограничиваем до 500 для диалога
  var limited = matched.slice(0, 500);
  var result = limited.map(function (p) {
    return {
      id: p.id,
      name: p.name,
      inStock: p.inStock,
      existing: existingSet[String(p.id)] || false
    };
  });

  return {
    matched: result,
    existingNotMatched: existingNotMatched,
    matchedTotal: matched.length,
    total: allProducts.length,
    truncated: matched.length > 500
  };
}

/**
 * Сохраняет результат подбора из диалога в SEO-теги
 * @param {number} rowNumber - номер строки
 * @param {string} criteriaText - текст критериев для колонки Q
 * @param {Array} selectedProducts - [{id, name}, ...] для колонки R
 */
function saveProductSelectionFromDialog(rowNumber, criteriaText, selectedProducts) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
  if (!sheet) throw new Error('Лист "' + SEO_TAGS_CONFIG.MASS_SHEET_NAME + '" не найден');

  var cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

  // Записываем критерии в Q
  sheet.getRange(rowNumber, cols.PRODUCT_CRITERIA).setValue(criteriaText);
  sheet.getRange(rowNumber, cols.PRODUCT_CRITERIA).setBackground('#e3f2fd');

  // Записываем товары в R: ID | Название (по строке)
  if (selectedProducts && selectedProducts.length > 0) {
    var exportData = readExportSheetHeaders_();
    var allProducts = exportData ? readExportSheetProducts_(exportData) : [];
    var productMap = {};
    for (var i = 0; i < allProducts.length; i++) {
      productMap[allProducts[i].id] = allProducts[i].name;
    }

    var lines = selectedProducts.map(function (p) {
      var name = p.name || productMap[p.id] || '';
      return p.id + ' | ' + name;
    });
    lines.push('---');
    lines.push('Найдено: ' + selectedProducts.length);
    sheet.getRange(rowNumber, cols.PRODUCTS_RESULT).setValue(lines.join('\n'));
    sheet.getRange(rowNumber, cols.PRODUCTS_RESULT).setBackground('#e8f5e9');
  } else {
    sheet.getRange(rowNumber, cols.PRODUCTS_RESULT).setValue('');
    sheet.getRange(rowNumber, cols.PRODUCTS_RESULT).setBackground(null);
  }

  return { success: true, count: selectedProducts ? selectedProducts.length : 0 };
}

// ========================================
// ПАЙПЛАЙН SEO-ТЕГОВ (Phase 1 + Phase 2)
// ========================================

/**
 * Удаляет все триггеры для пайплайна SEO
 */
function deletePipelineTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'runSeoPipelinePhase1' || handler === 'runSeoPipelinePhase2') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Создаёт триггер для авто-перезапуска пайплайна через 1 минуту
 * @param {string} functionName - имя функции для перезапуска
 */
function schedulePipelineRestart_(functionName) {
  try {
    deletePipelineTriggers_();
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .after(60 * 1000) // 1 минута
      .create();
    logInfo(`⏰ Триггер перезапуска создан: ${functionName} через 1 мин`, null, 'schedulePipelineRestart_');
  } catch (e) {
    logError('❌ Ошибка создания триггера перезапуска', e, 'schedulePipelineRestart_');
    SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка авто-перезапуска. Запустите вручную.', '⚠️', 10);
  }
}

/**
 * Фаза 1: Загрузка данных → Парсинг ключевиков → Минусация → Кластеризация
 *
 * State machine с автоматическим прогоном всех шагов.
 * При приближении к 6-мин таймауту — авто-перезапуск через триггер.
 * По завершении останавливается для ручной проверки кластеров.
 */
function runSeoPipelinePhase1() {
  const context = 'runSeoPipelinePhase1';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const progressKey = SEO_TAGS_CONFIG.PIPELINE_PROGRESS_KEY;
  const MAX_EXECUTION_TIME = 5 * 60 * 1000; // 5 минут — буфер до 6-мин лимита
  const startTime = Date.now();

  // Удаляем триггер, который нас запустил (если есть)
  deletePipelineTriggers_();

  try {
    const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      try { SpreadsheetApp.getUi().alert('Ошибка', 'Лист "SEO-теги" не найден', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { }
      return;
    }

    // Читаем или инициализируем прогресс пайплайна
    let pipeline = JSON.parse(props.getProperty(progressKey) || 'null');

    if (!pipeline) {
      const selected = getSelectedSeoTagRows_();
      if (selected.length === 0) {
        try { SpreadsheetApp.getUi().alert('Нет данных', 'Отметьте чекбоксами строки для обработки', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { }
        return;
      }

      try {
        const ui = SpreadsheetApp.getUi();
        const confirm = ui.alert('Фаза 1: SEO-пайплайн',
          `Будет обработано строк: ${selected.length}\n\n` +
          `Шаги:\n` +
          `1. Загрузка данных из InSales\n` +
          `2. Парсинг ключевиков (Arsenkin)\n` +
          `3. Импорт LSI — подсветки (Arsenkin)\n` +
          `4. Подсветка минус-слов\n` +
          `5. Кластеризация (Arsenkin)\n\n` +
          `Все шаги выполняются автоматически.\n` +
          `При таймауте — авто-перезапуск.\n` +
          `После завершения — ручная проверка.\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;
      } catch (e) {
        // Запуск из триггера — UI недоступен, продолжаем
      }

      pipeline = { phase: 1, step: 'loading', startedAt: new Date().toISOString() };
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('🚀 Запуск Фазы 1 пайплайна', { rows: selected.length }, context);
    }

    logInfo(`📍 Фаза 1, шаг: ${pipeline.step}`, null, context);

    if (pipeline.step === 'phase1_done') {
      ss.toast('Фаза 1 уже завершена. Проверьте кластеры и запустите Фазу 2.', '✅', 10);
      return;
    }

    // === State machine: автоматический прогон всех шагов ===
    // После каждого шага проверяем время — если мало, авто-перезапуск через триггер

    // --- ШАГ 1: Загрузка данных ---
    if (pipeline.step === 'loading') {
      ss.toast('Фаза 1, шаг 1/5: Загрузка данных из InSales...', '🔄', -1);
      const result = loadSeoDataForCheckedRows(true);

      if (!result || !result.complete) {
        // Не завершено — нужен повторный запуск для этого же шага
        ss.toast('Загрузка не завершена. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase1');
        return;
      }

      // Проверяем обязательные столбцы, помечаем пустые красным
      const mandatoryCols = SEO_TAGS_CONFIG.MANDATORY_COLUMNS;
      const selected = getSelectedSeoTagRows_();
      let skippedCount = 0;
      for (const row of selected) {
        const pageName = seoSheet.getRange(row.row, mandatoryCols[0]).getValue();
        const pageContent = seoSheet.getRange(row.row, mandatoryCols[1]).getValue();
        if (!pageName || !pageContent) {
          seoSheet.getRange(row.row, 1, 1, seoSheet.getLastColumn()).setBackground('#ffcdd2');
          seoSheet.getRange(row.row, 1).setValue(false);
          skippedCount++;
        }
      }
      if (skippedCount > 0) {
        logInfo(`⚠️ Пропущено ${skippedCount} строк без обязательных данных`, null, context);
      }

      pipeline.step = 'keywords';
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('✅ Шаг 1/5 завершён (загрузка данных)', null, context);
      ss.toast('✅ Шаг 1/5: загрузка данных завершена. Переход к ключевикам...', '🔄', 5);

      // Проверяем время перед следующим шагом
      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast('Таймаут. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase1');
        return;
      }
    }

    if (pipeline.step === 'keywords') {
      ss.toast('Фаза 1, шаг 2/5: Парсинг ключевиков (Arsenkin)...', '🔄', -1);
      const result = importKeywordsFromArsenkin(true);

      if (!result || !result.complete) {
        ss.toast('Парсинг не завершён. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase1');
        return;
      }

      pipeline.step = 'lsi';
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('✅ Шаг 2/5 завершён (ключевики)', null, context);
      ss.toast('✅ Шаг 2/5: ключевики собраны. Переход к LSI...', '🔄', 5);

      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast('Таймаут. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase1');
        return;
      }
    }

    // --- ШАГ 3: Импорт LSI (подсветки) ---
    if (pipeline.step === 'lsi') {
      ss.toast('Фаза 1, шаг 3/5: Импорт LSI — подсветки (Arsenkin)...', '🔄', -1);
      const result = importLsiFromArsenkin(true);

      if (!result || !result.complete) {
        ss.toast('Импорт LSI не завершён. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase1');
        return;
      }

      pipeline.step = 'minus_words';
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('✅ Шаг 3/5 завершён (LSI)', null, context);
      ss.toast('✅ Шаг 3/5: LSI собраны. Переход к минус-словам...', '🔄', 5);

      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast('Таймаут. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase1');
        return;
      }
    }

    // --- ШАГ 4: Подсветка минус-слов ---
    if (pipeline.step === 'minus_words') {
      ss.toast('Фаза 1, шаг 4/5: Подсветка минус-слов...', '🔄', -1);
      const result = highlightMinusWordsInCheckedRows(true);

      if (result && result.success) {
        logInfo(`🔍 Подсветка: найдено ${result.totalHighlighted} фраз`, null, context);
      }

      pipeline.step = 'clustering';
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('✅ Шаг 4/5 завершён (минус-слова подсвечены)', null, context);
      ss.toast('✅ Шаг 4/5: минус-слова подсвечены. Переход к кластеризации...', '🔄', 5);

      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast('Таймаут. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase1');
        return;
      }
    }

    // --- ШАГ 5: Кластеризация ---
    if (pipeline.step === 'clustering') {
      ss.toast('Фаза 1, шаг 5/5: Кластеризация (Arsenkin)...', '🔄', -1);
      const result = clusterCategoriesWithArsenkin(true);

      if (!result || !result.complete) {
        ss.toast('Кластеризация не завершена. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase1');
        return;
      }

      pipeline.step = 'phase1_done';
      props.setProperty(progressKey, JSON.stringify(pipeline));

      logInfo('✅ Фаза 1 пайплайна завершена', null, context);
      ss.toast('Фаза 1 завершена! Проверьте кластеры и минус-слова.', '✅', -1);
      try {
        SpreadsheetApp.getUi().alert('Фаза 1 завершена',
          'Все шаги Фазы 1 выполнены:\n' +
          '✅ Загрузка данных\n' +
          '✅ Парсинг ключевиков\n' +
          '✅ Импорт LSI (подсветки)\n' +
          '✅ Подсветка минус-слов\n' +
          '✅ Кластеризация\n\n' +
          'Проверьте кластеры, минус-слова, и запустите Фазу 2.',
          SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) {
        // Запуск из триггера — UI недоступен
      }
    }

  } catch (error) {
    logError('❌ Ошибка Фазы 1 пайплайна', error, context);
    try { SpreadsheetApp.getUi().alert('Ошибка Фазы 1', error.message, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { }
  }
}

/**
 * Фаза 2: Удаление минус-слов → Частотность → Конкуренты → SEO-теги
 *
 * Запускается после ручной проверки кластеров.
 * Автоматический прогон всех шагов с авто-перезапуском по таймауту.
 * Использует промпты из строки 2.
 */
function runSeoPipelinePhase2() {
  const context = 'runSeoPipelinePhase2';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const progressKey = SEO_TAGS_CONFIG.PIPELINE_PROGRESS_KEY;
  const MAX_EXECUTION_TIME = 5 * 60 * 1000; // 5 минут
  const startTime = Date.now();

  // Удаляем триггер, который нас запустил (если есть)
  deletePipelineTriggers_();

  try {
    const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      try { SpreadsheetApp.getUi().alert('Ошибка', 'Лист "SEO-теги" не найден', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { }
      return;
    }

    let pipeline = JSON.parse(props.getProperty(progressKey) || 'null');

    if (!pipeline || (pipeline.step !== 'phase1_done' && pipeline.phase !== 2)) {
      const selected = getSelectedSeoTagRows_();
      if (selected.length === 0) {
        try { SpreadsheetApp.getUi().alert('Нет данных', 'Отметьте чекбоксами строки для обработки', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { }
        return;
      }

      try {
        const ui = SpreadsheetApp.getUi();
        const confirm = ui.alert('Фаза 2: SEO-пайплайн',
          `Будет обработано строк: ${selected.length}\n\n` +
          `Шаги:\n` +
          `1. Удаление минус-слов (подсвеченных)\n` +
          `2. Сбор частотности (Arsenkin)\n` +
          `3. Импорт тегов конкурентов (Arsenkin)\n` +
          `4. Генерация SEO-тегов (3 этапа, промпты из строки 2)\n\n` +
          `Все шаги выполняются автоматически.\n` +
          `При таймауте — авто-перезапуск.\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;
      } catch (e) { }

      pipeline = { phase: 2, step: 'delete_minus', startedAt: new Date().toISOString() };
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('🚀 Запуск Фазы 2 пайплайна', { rows: selected.length }, context);

    } else if (pipeline.step === 'phase1_done') {
      const selected = getSelectedSeoTagRows_();
      try {
        const ui = SpreadsheetApp.getUi();
        const confirm = ui.alert('Фаза 2: SEO-пайплайн',
          `Фаза 1 завершена. Отмечено строк: ${selected.length}\n\n` +
          `Шаги Фазы 2:\n` +
          `1. Удаление минус-слов (подсвеченных)\n` +
          `2. Сбор частотности\n` +
          `3. Импорт тегов конкурентов\n` +
          `4. Генерация SEO-тегов (3 этапа)\n\n` +
          `Все шаги выполняются автоматически.\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;
      } catch (e) { }

      pipeline = { phase: 2, step: 'delete_minus', startedAt: new Date().toISOString() };
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('🚀 Переход к Фазе 2 пайплайна', { rows: selected.length }, context);
    }

    logInfo(`📍 Фаза 2, шаг: ${pipeline.step}`, null, context);

    if (pipeline.step === 'phase2_done') {
      ss.toast('Фаза 2 уже завершена.', '✅', 10);
      return;
    }

    // === State machine: автоматический прогон всех шагов ===

    // --- ШАГ 1: Удаление минус-слов ---
    if (pipeline.step === 'delete_minus') {
      ss.toast('Фаза 2, шаг 1/4: Удаление минус-слов...', '🔄', -1);
      const result = deleteMinusWordsFromCheckedRows(true);

      if (result && result.success) {
        logInfo(`🚫 Удалено ${result.totalRemoved} фраз`, null, context);
      }

      pipeline.step = 'frequency';
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('✅ Шаг 1/4 завершён (минус-слова удалены)', null, context);
      ss.toast('✅ Шаг 1/4: минус-слова удалены. Переход к частотности...', '🔄', 5);

      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast('Таймаут. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase2');
        return;
      }
    }

    // --- ШАГ 2: Частотность ---
    if (pipeline.step === 'frequency') {
      ss.toast('Фаза 2, шаг 2/4: Сбор частотности (Arsenkin)...', '🔄', -1);
      const result = collectFrequencyFromArsenkin(true);

      if (!result || !result.complete) {
        ss.toast('Сбор частотности не завершён. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase2');
        return;
      }

      pipeline.step = 'competitors';
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('✅ Шаг 2/4 завершён (частотность)', null, context);
      ss.toast('✅ Шаг 2/4: частотность собрана. Переход к конкурентам...', '🔄', 5);

      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast('Таймаут. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase2');
        return;
      }
    }

    // --- ШАГ 3: Конкуренты ---
    if (pipeline.step === 'competitors') {
      ss.toast('Фаза 2, шаг 3/4: Импорт тегов конкурентов (Arsenkin)...', '🔄', -1);
      const result = importCompetitorTagsFromArsenkin(true);

      if (!result || !result.complete) {
        ss.toast('Импорт конкурентов не завершён. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase2');
        return;
      }

      pipeline.step = 'generation';
      props.setProperty(progressKey, JSON.stringify(pipeline));
      logInfo('✅ Шаг 3/4 завершён (конкуренты)', null, context);
      ss.toast('✅ Шаг 3/4: конкуренты импортированы. Переход к генерации...', '🔄', 5);

      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        ss.toast('Таймаут. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase2');
        return;
      }
    }

    // --- ШАГ 4: Генерация SEO-тегов ---
    if (pipeline.step === 'generation') {
      ss.toast('Фаза 2, шаг 4/4: Генерация SEO-тегов (3 этапа)...', '🔄', -1);
      const result = generateSeoTagsStagedMass(true); // silent=true

      if (!result || !result.complete) {
        ss.toast('Генерация не завершена. Авто-перезапуск через 1 мин...', '⏳', -1);
        schedulePipelineRestart_('runSeoPipelinePhase2');
        return;
      }

      pipeline.step = 'phase2_done';
      props.setProperty(progressKey, JSON.stringify(pipeline));

      logInfo('✅ Фаза 2 пайплайна завершена', null, context);
      ss.toast('Фаза 2 завершена! Результаты в столбцах O и P.', '✅', -1);
      try {
        SpreadsheetApp.getUi().alert('Фаза 2 завершена',
          'Все шаги Фазы 2 выполнены:\n' +
          '✅ Удаление минус-слов\n' +
          '✅ Сбор частотности\n' +
          '✅ Импорт тегов конкурентов\n' +
          '✅ Генерация SEO-тегов\n\n' +
          'Результаты в столбцах O и P.',
          SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) { }
    }

  } catch (error) {
    logError('❌ Ошибка Фазы 2 пайплайна', error, context);
    try { SpreadsheetApp.getUi().alert('Ошибка Фазы 2', error.message, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { }
  }
}

/**
 * Сбрасывает прогресс пайплайна и все подпрогрессы Арсенкина.
 * Также удаляет триггеры авто-перезапуска.
 */
function clearPipelineProgress() {
  const context = 'clearPipelineProgress';
  const props = PropertiesService.getScriptProperties();

  // Удаляем триггеры авто-перезапуска
  deletePipelineTriggers_();

  // Очищаем основной прогресс пайплайна
  props.deleteProperty(SEO_TAGS_CONFIG.PIPELINE_PROGRESS_KEY);

  // Очищаем прогресс загрузки отмеченных
  clearProgressChunked_(SEO_TAGS_CONFIG.LOAD_CHECKED_PROGRESS_KEY);

  // Очищаем прогрессы Арсенкин-задач
  const arsenkinKeys = [
    ARSENKIN_CONFIG.PROGRESS_KEYS.WORDSTAT_PHRASES,
    ARSENKIN_CONFIG.PROGRESS_KEYS.WORDSTAT_FREQ,
    ARSENKIN_CONFIG.PROGRESS_KEYS.CHECK_H,
    ARSENKIN_CONFIG.PROGRESS_KEYS.CLUSTERING
  ];

  for (const key of arsenkinKeys) {
    clearProgressChunked_(key);
  }

  logInfo('🧹 Прогресс пайплайна и все подпрогрессы сброшены', null, context);
  SpreadsheetApp.getActiveSpreadsheet().toast('Прогресс пайплайна сброшен', '🧹', 5);
}

// ========================================
// МИНУСАЦИЯ КЛЮЧЕВИКОВ
// ========================================

/**
 * Читает список минус-слов с листа "Минус-слова"
 * Формат: одна запись на строку в столбце A
 * Поддерживает оператор ! для точной формы слова
 * @returns {Array<{pattern: RegExp, original: string}>} Массив скомпилированных минус-правил
 */
function getMinusWords_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SEO_TAGS_CONFIG.MINUS_WORDS_SHEET_NAME;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    logWarning(`⚠️ Лист "${sheetName}" не найден`, null, 'getMinusWords_');
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  const rules = [];

  for (let i = 0; i < values.length; i++) {
    const raw = values[i][0];
    if (!raw || String(raw).trim() === '') continue;

    const entry = String(raw).trim();
    rules.push(buildMinusRule_(entry));
  }

  return rules;
}

/**
 * Строит регулярное выражение для минус-правила
 * - Без !: word boundary matching, регистронезависимо
 * - С ! перед словом: точная форма (без морфологии), регистронезависимо
 * - Фраза (несколько слов): все слова в указанном порядке
 * @param {string} entry - Запись минус-слова (может содержать !)
 * @returns {{pattern: RegExp, original: string}}
 */
function buildMinusRule_(entry) {
  // Разбиваем на токены для поддержки фраз с !
  const tokens = entry.split(/\s+/);

  // ВАЖНО: \b в JavaScript не работает с кириллицей (только ASCII).
  // Используем Unicode-совместимые границы слов:
  // Начало слова: (?:^|\s) — начало строки или пробел
  // Конец слова: (?:\s|$) — пробел или конец строки
  const WB_START = '(?:^|\\s)';
  const WB_END = '(?:\\s|$)';

  // Для одиночного токена
  if (tokens.length === 1) {
    const token = tokens[0];
    const word = token.startsWith('!') ? token.substring(1) : token;
    const escaped = escapeRegex_(word);
    const fullPattern = WB_START + escaped + WB_END;
    return {
      pattern: new RegExp(fullPattern, 'i'),
      original: entry
    };
  }

  // Для фраз: все токены должны идти подряд (через \s+)
  const phraseTokens = tokens.map(t => {
    const word = t.startsWith('!') ? t.substring(1) : t;
    return escapeRegex_(word);
  });
  const fullPattern = WB_START + phraseTokens.join('\\s+') + WB_END;

  return {
    pattern: new RegExp(fullPattern, 'i'),
    original: entry
  };
}

/**
 * Экранирует спецсимволы для использования в RegExp
 */
function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Фильтрует ключевики в ячейке по минус-словам
 * @param {string} cellValue - Значение ячейки (формат: "фраза (частота)\nфраза2 (частота2)")
 * @param {Array} minusRules - Массив правил из getMinusWords_()
 * @returns {{filtered: string, removedCount: number}} Отфильтрованное значение и кол-во удалённых
 */
function filterKeywordsByMinusWords_(cellValue, minusRules) {
  if (!cellValue || !minusRules || minusRules.length === 0) {
    return { filtered: cellValue, removedCount: 0 };
  }

  const lines = String(cellValue).split('\n').map(s => s.trim()).filter(s => s);
  let removedCount = 0;
  const kept = [];

  for (const line of lines) {
    // Извлекаем фразу (без частотности в скобках)
    const phrase = line.replace(/\s*\(\d+\)\s*$/, '').trim().toLowerCase();

    if (!phrase) {
      kept.push(line);
      continue;
    }

    let isMinused = false;
    for (const rule of minusRules) {
      if (rule.pattern.test(phrase)) {
        isMinused = true;
        break;
      }
    }

    if (isMinused) {
      removedCount++;
    } else {
      kept.push(line);
    }
  }

  return {
    filtered: kept.join('\n'),
    removedCount: removedCount
  };
}

/**
 * Подсвечивает минус-слова в столбце G (семантическое ядро) через RichText.
 * Фразы, попадающие под минус-слова, получают зачёркивание и серый цвет.
 * Фаза 1 пайплайна использует эту функцию для проверки перед удалением.
 * @param {boolean} [silent=false] - Если true, не показывает UI-диалоги
 * @returns {{success: boolean, totalHighlighted: number, processedRows: number}}
 */
function highlightMinusWordsInCheckedRows(silent) {
  const context = 'highlightMinusWordsInCheckedRows';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = silent ? null : SpreadsheetApp.getUi();

  try {
    const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      if (!silent) ui.alert('Ошибка', 'Лист "SEO-теги" не найден', ui.ButtonSet.OK);
      return { success: false, totalHighlighted: 0, processedRows: 0 };
    }

    const minusRules = getMinusWords_();
    if (minusRules.length === 0) {
      if (!silent) ui.alert('Минус-слова', 'Лист "Минус-слова" пуст или не найден.\nДобавьте минус-слова в столбец A.', ui.ButtonSet.OK);
      return { success: false, totalHighlighted: 0, processedRows: 0 };
    }

    const selected = getSelectedSeoTagRows_();
    if (selected.length === 0) {
      if (!silent) ui.alert('Нет данных', 'Отметьте чекбоксами строки для подсветки', ui.ButtonSet.OK);
      return { success: false, totalHighlighted: 0, processedRows: 0 };
    }

    if (!silent) {
      const confirm = ui.alert('Подсветка минус-слов',
        `Будет обработано строк: ${selected.length}\n` +
        `Минус-правил: ${minusRules.length}\n\n` +
        `Фразы с минус-словами будут зачёркнуты серым цветом.\n` +
        `Для удаления используйте "Удалить минус-слова".\n\n` +
        `Продолжить?`,
        ui.ButtonSet.YES_NO);
      if (confirm !== ui.Button.YES) return { success: false, totalHighlighted: 0, processedRows: 0 };
    }

    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    const normalStyle = SpreadsheetApp.newTextStyle().setForegroundColor('#000000').setStrikethrough(false).build();
    const minusStyle = SpreadsheetApp.newTextStyle().setForegroundColor('#999999').setStrikethrough(true).build();

    let totalHighlighted = 0;
    let processedRows = 0;

    for (const row of selected) {
      const cell = seoSheet.getRange(row.row, cols.SEMANTIC_CORE);
      const cellValue = cell.getValue();
      if (!cellValue || String(cellValue).trim() === '') continue;

      const lines = String(cellValue).split('\n').map(s => s.trim()).filter(s => s);
      let hasHighlights = false;

      // Определяем какие строки минусуются
      const lineStatuses = lines.map(line => {
        const phrase = line.replace(/\s*\(\d+\)\s*$/, '').trim().toLowerCase();
        if (!phrase) return false;
        for (const rule of minusRules) {
          if (rule.pattern.test(phrase)) return true;
        }
        return false;
      });

      // Если нет совпадений — пропускаем
      if (!lineStatuses.some(s => s)) {
        processedRows++;
        continue;
      }

      // Собираем полный текст и строим RichText
      const fullText = lines.join('\n');
      const builder = SpreadsheetApp.newRichTextValue().setText(fullText);

      let offset = 0;
      for (let i = 0; i < lines.length; i++) {
        const lineLen = lines[i].length;
        if (lineStatuses[i]) {
          builder.setTextStyle(offset, offset + lineLen, minusStyle);
          hasHighlights = true;
          totalHighlighted++;
        } else {
          builder.setTextStyle(offset, offset + lineLen, normalStyle);
        }
        offset += lineLen + 1; // +1 за \n
      }

      if (hasHighlights) {
        cell.setRichTextValue(builder.build());
      }
      processedRows++;
    }

    logInfo(`🔍 Подсветка: обработано ${processedRows} строк, подсвечено ${totalHighlighted} фраз`, null, context);

    if (!silent) {
      ss.toast(`Подсветка: ${totalHighlighted} фраз в ${processedRows} строках`, '✅', 10);
    }

    return { success: true, totalHighlighted, processedRows };

  } catch (error) {
    logError('❌ Ошибка подсветки минус-слов', error, context);
    if (!silent) ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
    return { success: false, totalHighlighted: 0, processedRows: 0 };
  }
}

/**
 * Удаляет минус-слова из столбца G (семантическое ядро) для отмеченных строк.
 * Удаляет зачёркнутые (подсвеченные) строки, а также любые фразы по минус-правилам.
 * Фаза 2 пайплайна вызывает после ручной проверки подсветки.
 * @param {boolean} [silent=false] - Если true, не показывает UI-диалоги
 * @returns {{success: boolean, totalRemoved: number, processedRows: number}}
 */
function deleteMinusWordsFromCheckedRows(silent) {
  const context = 'deleteMinusWordsFromCheckedRows';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = silent ? null : SpreadsheetApp.getUi();

  try {
    const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
    if (!seoSheet) {
      if (!silent) ui.alert('Ошибка', 'Лист "SEO-теги" не найден', ui.ButtonSet.OK);
      return { success: false, totalRemoved: 0, processedRows: 0 };
    }

    const minusRules = getMinusWords_();
    if (minusRules.length === 0) {
      if (!silent) ui.alert('Минус-слова', 'Лист "Минус-слова" пуст или не найден.\nДобавьте минус-слова в столбец A.', ui.ButtonSet.OK);
      return { success: false, totalRemoved: 0, processedRows: 0 };
    }

    const selected = getSelectedSeoTagRows_();
    if (selected.length === 0) {
      if (!silent) ui.alert('Нет данных', 'Отметьте чекбоксами строки для минусации', ui.ButtonSet.OK);
      return { success: false, totalRemoved: 0, processedRows: 0 };
    }

    if (!silent) {
      const confirm = ui.alert('Удаление минус-слов',
        `Будет обработано строк: ${selected.length}\n` +
        `Минус-правил: ${minusRules.length}\n\n` +
        `Фразы с минус-словами будут УДАЛЕНЫ из столбца G.\n\n` +
        `Продолжить?`,
        ui.ButtonSet.YES_NO);
      if (confirm !== ui.Button.YES) return { success: false, totalRemoved: 0, processedRows: 0 };
    }

    const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
    let totalRemoved = 0;
    let processedRows = 0;

    for (const row of selected) {
      const cell = seoSheet.getRange(row.row, cols.SEMANTIC_CORE);
      const cellValue = cell.getValue();
      if (!cellValue || String(cellValue).trim() === '') continue;

      const result = filterKeywordsByMinusWords_(cellValue, minusRules);

      if (result.removedCount > 0) {
        // Записываем как обычный текст, сбрасывая RichText-форматирование
        cell.clearContent();
        cell.setValue(result.filtered);
        totalRemoved += result.removedCount;
      }
      processedRows++;
    }

    logInfo(`🚫 Удаление: обработано ${processedRows} строк, удалено ${totalRemoved} фраз`, null, context);

    if (!silent) {
      ss.toast(`Удалено ${totalRemoved} фраз из ${processedRows} строк`, '✅', 10);
    }

    return { success: true, totalRemoved, processedRows };

  } catch (error) {
    logError('❌ Ошибка удаления минус-слов', error, context);
    if (!silent) ui.alert('Ошибка', error.message, ui.ButtonSet.OK);
    return { success: false, totalRemoved: 0, processedRows: 0 };
  }
}

/**
 * @deprecated Используйте deleteMinusWordsFromCheckedRows() или highlightMinusWordsInCheckedRows()
 */
function applyMinusWordsToCheckedRows(silent) {
  return deleteMinusWordsFromCheckedRows(silent);
}
