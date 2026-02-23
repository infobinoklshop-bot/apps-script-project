// ========================================
// СОЗДАНИЕ СТРУКТУРЫ ЛИСТОВ
// ========================================

/**
 * Создает структуру листов для работы с категориями
 * Вызывается один раз при инициализации модуля
 */
function createCategoryManagementStructure() {
  const context = "Создание структуры категорий";

  try {
    logInfo('🏗️ Создаем структуру листов для управления категориями', null, context);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Создаем главный лист списка категорий
    let mainListSheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!mainListSheet) {
      mainListSheet = ss.insertSheet(CATEGORY_SHEETS.MAIN_LIST);
      logInfo('✅ Создан лист: ' + CATEGORY_SHEETS.MAIN_LIST);
    }

    // Настраиваем главный лист
    setupMainListSheet(mainListSheet);

    // 2. Создаем лист для ключевых слов
    let keywordsSheet = ss.getSheetByName(CATEGORY_SHEETS.KEYWORDS);

    if (!keywordsSheet) {
      keywordsSheet = ss.insertSheet(CATEGORY_SHEETS.KEYWORDS);
      logInfo('✅ Создан лист: ' + CATEGORY_SHEETS.KEYWORDS);
    }

    setupKeywordsSheet(keywordsSheet);

    // 3. Создаем лист каталога для подбора
    let catalogSheet = ss.getSheetByName(CATEGORY_SHEETS.PRODUCTS_CATALOG);

    if (!catalogSheet) {
      catalogSheet = ss.insertSheet(CATEGORY_SHEETS.PRODUCTS_CATALOG);
      logInfo('✅ Создан лист: ' + CATEGORY_SHEETS.PRODUCTS_CATALOG);
    }

    setupProductsCatalogSheet(catalogSheet);

    logInfo('✅ Структура листов категорий создана успешно', null, context);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Структура для управления категориями создана!',
      '✅ Готово',
      5
    );

    return true;

  } catch (error) {
    logError('❌ Ошибка создания структуры', error, context);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка: ' + error.message,
      '❌ Ошибка',
      10
    );
    return false;
  }
}

/**
 * Настройка главного листа со списком категорий
 */
function setupMainListSheet(sheet) {
  // Очищаем лист
  sheet.clear();

  // Заголовки
  const headers = [
    '☑️',                    // A - Checkbox
    'ID',                   // B
    'Parent ID',            // C
    'Уровень',              // D
    'Путь в иерархии',      // E
    'Название категории',   // F
    'URL',                  // G
    'Товаров',              // H
    'В наличии',            // I
    'SEO статус',           // J
    'AI статус',            // K
    'Обновлено',            // L
    'Админка',              // M
    'Трафик (Яндекс.Метрика)', // N
    'Отказы',               // O
    'Показы GSC',           // P
    'Показы Я.Веб',         // Q
    'Период',               // R
    'Правила наполнения'    // S (Index 18, 1-based 19)
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Форматирование заголовков
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Ширина колонок
  sheet.setColumnWidth(1, 40);   // Checkbox
  sheet.setColumnWidth(2, 80);   // ID
  sheet.setColumnWidth(3, 80);   // Parent ID
  sheet.setColumnWidth(4, 80);   // Уровень
  sheet.setColumnWidth(5, 300);  // Путь
  sheet.setColumnWidth(6, 250);  // Название
  sheet.setColumnWidth(7, 200);  // URL
  sheet.setColumnWidth(8, 80);   // Товаров
  sheet.setColumnWidth(9, 80);   // В наличии
  sheet.setColumnWidth(10, 120); // SEO статус
  sheet.setColumnWidth(11, 120); // AI статус
  sheet.setColumnWidth(12, 150); // Обновлено
  sheet.setColumnWidth(13, 150); // Админка
  sheet.setColumnWidth(14, 100); // Визиты
  sheet.setColumnWidth(15, 80);  // Отказы
  sheet.setColumnWidth(16, 100); // Показы GSC
  sheet.setColumnWidth(17, 100); // Показы Я.Веб
  sheet.setColumnWidth(18, 200); // Период
  sheet.setColumnWidth(19, 300); // Правила наполнения

  // Закрепляем заголовок
  sheet.setFrozenRows(1);

  // Добавляем инструкцию
  sheet.getRange('A2').setValue('👇 Используйте "Загрузить категории" в меню');
  sheet.getRange('A2:M2').merge().setBackground('#fff3cd').setFontStyle('italic');

  // Добавляем настройку "Товаров на странице"
  sheet.getRange('Q1').setValue('Товаров на 1 стр:').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange('R1').setValue(36).setHorizontalAlignment('center').setBackground('#fff3cd');
  sheet.getRange('R1').setNote('Количество товаров на первой странице листинга. Используется для выделения в детальном листе.');

  logInfo('✅ Главный лист категорий настроен');
}

/**
 * Настройка листа ключевых слов
 */
function setupKeywordsSheet(sheet) {
  sheet.clear();

  const headers = [
    'Категория ID',
    'Категория',
    'Ключевое слово',
    'Частотность',
    'Конкуренция',
    'Тип',
    'Назначение',
    'Статус использования'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 150);

  sheet.setFrozenRows(1);

  logInfo('✅ Лист ключевых слов настроен');
}

/**
 * Настройка листа каталога товаров для подбора
 */
function setupProductsCatalogSheet(sheet) {
  sheet.clear();

  const headers = [
    '☑️',
    'ID товара',
    'Артикул',
    'Название',
    'Категории',
    'В наличии',
    'Цена',
    'Бренд',
    'Основные характеристики',
    'Дата добавления'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#fbbc04')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 80);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 300);
  sheet.setColumnWidth(10, 120);

  sheet.setFrozenRows(1);

  logInfo('✅ Лист каталога товаров настроен');
}