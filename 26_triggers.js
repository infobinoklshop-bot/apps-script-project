/**
 * ============================================
 * МОДУЛЬ: ТРИГГЕРЫ GOOGLE SHEETS
 * ============================================
 *
 * Обработчики событий Apps Script:
 * - onSelectionChange: реагирует на выбор ячеек
 * - onEdit: реагирует на редактирование (если нужен)
 */

// ============================================
// ОБРАБОТЧИК ВЫБОРА ЯЧЕЕК
// ============================================

/**
 * Триггер onSelectionChange - срабатывает при выборе ячейки пользователем
 *
 * ВАЖНО: Это "простой триггер" (simple trigger) с ограничениями:
 * - Не может напрямую вызывать UrlFetchApp (API запросы)
 * - Не может использовать PropertiesService
 * - Работает только с данными из текущей таблицы
 *
 * Поэтому мы используем его только для открытия диалога,
 * а все API запросы выполняются внутри диалога с полными правами
 *
 * @param {Object} e - Объект события с информацией о выборе
 */
function onSelectionChange(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const row = range.getRow();
    const col = range.getColumn();

    // Проверяем, что это детальный лист категории
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      return;
    }

    // Проверяем, что выбрана колонка H (PICKER_LINK = 8)
    // И что это строка в диапазоне таблицы ключевых слов (20-69)
    const pickerColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.PICKER_LINK;
    const dataStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;
    const dataEndRow = dataStartRow + 49; // 50 строк для ввода

    if (col === pickerColumn && row >= dataStartRow && row <= dataEndRow) {
      // Пользователь кликнул на ссылку "🔍 Выбрать" в колонке H
      // Открываем диалог подбора категорий
      showCategoryPickerForTagTiles(row, sheetName);
    }

  } catch (error) {
    // Простые триггеры не могут показывать UI alerts при ошибках
    // Логируем в консоль для отладки
    console.error('[ERROR] onSelectionChange:', error.message);
  }
}

// ============================================
// ОБРАБОТЧИК РЕДАКТИРОВАНИЯ (ОПЦИОНАЛЬНО)
// ============================================

/**
 * Триггер onEdit - срабатывает при редактировании ячейки
 *
 * Можно использовать для дополнительной логики, например:
 * - Автоматическая валидация при вводе ID категории
 * - Форматирование текста
 *
 * Пока не используется, но оставлен для будущих улучшений
 *
 * @param {Object} e - Объект события с информацией о редактировании
 */
function onEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    const row = range.getRow();
    const col = range.getColumn();

    // Проверяем, что это детальный лист категории
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      return;
    }

    // Пример: автоматическая валидация при изменении колонки E (CATEGORY_LINK)
    const categoryLinkColumn = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS.CATEGORY_LINK;
    const dataStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;
    const dataEndRow = dataStartRow + 49;

    if (col === categoryLinkColumn && row >= dataStartRow && row <= dataEndRow) {
      // Пользователь ввел данные в колонку "URL/ID категории"
      // Можно добавить автоматическую валидацию здесь
      // Но пока оставляем ручную валидацию через меню
      //
      // Раскомментируйте следующую строку для автоматической валидации:
      // validateTagKeywords();
    }

  } catch (error) {
    console.error('[ERROR] onEdit:', error.message);
  }
}

// ============================================
// ПРИМЕЧАНИЯ ПО ИСПОЛЬЗОВАНИЮ
// ============================================

/**
 * УСТАНОВКА ТРИГГЕРОВ:
 *
 * Простые триггеры (onSelectionChange, onEdit) устанавливаются автоматически
 * при наличии функций с этими именами в проекте Apps Script.
 *
 * Никакой ручной настройки не требуется!
 *
 * Если триггер не срабатывает:
 * 1. Проверьте имя функции (должно быть точно onSelectionChange)
 * 2. Убедитесь, что файл 26_triggers.js загружен в Apps Script (clasp push)
 * 3. Проверьте права доступа к таблице
 *
 * ОТЛАДКА:
 *
 * Простые триггеры не показывают UI alerts при ошибках.
 * Для отладки используйте:
 * - console.log() - пишет в логи Apps Script
 * - Logger.log() - пишет в старый лог
 * - Инструменты разработчика: Apps Script Editor → Executions
 */
