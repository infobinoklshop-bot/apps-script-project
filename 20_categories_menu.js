/**
 * ========================================
 * МОДУЛЬ: МЕНЮ КАТЕГОРИЙ
 * ========================================
 */

/**
 * Добавляет полное меню категорий
 */
function addFullCategoryMenu(mainMenu) {
  const categoryMenu = SpreadsheetApp.getUi().createMenu('📁 Категории');
  
  // Основные операции
  categoryMenu
    .addItem('🏗️ Создать структуру листов', 'createCategoryManagementStructure')
    .addSeparator()
    .addItem('📥 Загрузить все категории', 'loadCategoriesWithHierarchy')
    .addItem('🔍 Найти и открыть категорию', 'showCategorySearchDialog')
    .addItem('➕ Создать категорию', 'showCreateCategoryDialog')  // НОВОЕ!
    .addItem('🔄 Обновить данные категорий', 'updateCategoriesData')
    .addSeparator();
  
  // Семантическое ядро + LSI
  const semanticsMenu = SpreadsheetApp.getUi().createMenu('🔑 Семантика')
    .addItem('⚙️ Настроить API', 'configureSemanticsAPI')
    .addSeparator()
    .addItem('📊 Собрать ключевики', 'collectKeywordsForActiveCategory')
    .addItem('🎨 Собрать LSI и тематику', 'collectLSIForActiveCategory')  // НОВОЕ
    .addItem('🤖 Распределить (AI)', 'distributeKeywordsForActiveCategory')
    .addSeparator()
    .addItem('📋 Показать ключевики', 'showKeywordsSheet')
    .addItem('📋 Показать LSI', 'showLSISheet');  // НОВОЕ
  
  categoryMenu.addSubMenu(semanticsMenu);
  
  // AI генерация
  const aiMenu = SpreadsheetApp.getUi().createMenu('🤖 AI Генерация')
    .addItem('🔍 Анализ категории', 'analyzeActiveCategoryFromMenu')
    .addSeparator()
    .addItem('🎯 Генерировать SEO', 'generateSEOForActiveCategory')
    .addItem('📄 Генерировать описание', 'generateDescriptionForActiveCategory')
    .addSeparator()
    .addItem('🎨 Генерировать изображения', 'generateCategoryImagesWithAI');

  categoryMenu.addSubMenu(aiMenu);

  // Плитки тегов (ручное управление)
  const tilesMenu = SpreadsheetApp.getUi().createMenu('🏷️ Плитки тегов')
    .addItem('📝 Инициализировать таблицу', 'initializeTagKeywordsTable')
    .addSeparator()
    .addItem('✅ Проверить категории', 'validateTagKeywords')
    .addItem('➕ Создать новые категории', 'createCategoriesForTags')
    .addSeparator()
    .addItem('🎨 Сгенерировать плитки', 'generateTilesFromManualData')
    .addItem('👁️ Предпросмотр плитки', 'showTilesPreviewManual');

  categoryMenu.addSubMenu(tilesMenu);
  
  // НОВОЕ: Отслеживание позиций
  const positionsMenu = SpreadsheetApp.getUi().createMenu('📊 Позиции')
    .addItem('⚙️ Настроить маркерный запрос', 'setupMarkerQuery')
    .addItem('🔍 Проверить позиции', 'checkPositionsForActiveCategory')
    .addSeparator()
    .addItem('📈 Показать динамику', 'showPositionsDynamicsReport')
    .addItem('📋 История позиций', 'showPositionsHistorySheet');
  
  categoryMenu.addSubMenu(positionsMenu);

    // Управление товарами
  categoryMenu
    .addSeparator()
    .addItem('➕ Добавить товары', 'showAddProductsDialog')
    .addItem('🗑️ Удалить выбранные товары', 'removeSelectedProductsFromCategory')
    .addSeparator();
  
  // Отправка и утилиты
  categoryMenu
    .addItem('📤 Сохранить SEO в InSales', 'sendCategoryChangesToInSales')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠️ Утилиты')
      .addItem('📊 Статистика', 'showCategoriesStatistics')
      .addItem('💾 Экспорт в JSON', 'exportCategoryToJSON')
      .addItem('📋 Клонировать', 'cloneCategorySheet')
      .addItem('🗑️ Удалить лист', 'deleteCategorySheet'));
  
  mainMenu.addSubMenu(categoryMenu);
  
  logInfo('✅ Полное меню категорий добавлено');
}

/**
 * НОВЫЕ функции для меню
 */

function showLSISheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.LSI_WORDS);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert(
        'Лист не найден',
        'Сначала соберите LSI и тематические слова',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    ss.setActiveSheet(sheet);
    
  } catch (error) {
    logError('❌ Ошибка показа LSI', error);
  }
}

function showPositionsHistorySheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.POSITIONS_HISTORY);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert(
        'Лист не найден',
        'Сначала настройте отслеживание позиций',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    ss.setActiveSheet(sheet);
    
  } catch (error) {
    logError('❌ Ошибка показа истории позиций', error);
  }
}

// ========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ МЕНЮ
// ========================================

/**
 * Собирает ключевые слова для активной категории
 */
function collectKeywordsForActiveCategory() {
  try {
    const categoryData = getActiveCategoryData();
    
    if (!categoryData) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        'Откройте детальный лист категории перед выполнением этой операции',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    collectKeywordsForCategory(categoryData);
    
  } catch (error) {
    logError('❌ Ошибка сбора ключевых слов', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Распределяет ключевики для активной категории
 */
function distributeKeywordsForActiveCategory() {
  try {
    const categoryData = getActiveCategoryData();
    
    if (!categoryData) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        'Откройте детальный лист категории',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Получаем ключевые слова из листа
    const keywords = getKeywordsForCategory(categoryData.id);
    
    if (keywords.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Нет ключевых слов',
        'Сначала соберите семантическое ядро для категории',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    distributeKeywordsWithAI(categoryData.id, categoryData.title, keywords);
    
  } catch (error) {
    logError('❌ Ошибка распределения ключевиков', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Анализирует активную категорию
 */
function analyzeActiveCategory() {
  try {
    const categoryData = getActiveCategoryData();
    
    if (!categoryData) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        'Откройте детальный лист категории',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Получаем товары категории из листа
    const products = getProductsFromDetailSheet();
    
    // Выполняем анализ
    const analysis = analyzeCurrentCategoryState(categoryData, products);
    
    // Показываем результаты анализа
    showAnalysisResults(analysis);
    
  } catch (error) {
    logError('❌ Ошибка анализа категории', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Генерирует SEO для активной категории
 */
function generateSEOForActiveCategory() {
  try {
    const categoryData = getActiveCategoryData();
    
    if (!categoryData) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        'Откройте детальный лист категории',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Генерируем SEO теги через AI...',
      '⏳ Генерация',
      -1
    );
    
    // Получаем ключевые слова с назначением "seo_title" и "meta_description"
    const keywords = getKeywordsForCategory(categoryData.id);
    const seoKeywords = keywords.filter(kw => 
      kw.assignment === 'seo_title' || kw.assignment === 'meta_description'
    );
    
    // Генерируем SEO через AI
    const seoData = generateSEOWithAI(categoryData, seoKeywords);
    
    // Записываем в детальный лист
    updateSEOInDetailSheet(seoData);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'SEO теги сгенерированы!',
      '✅ Готово',
      5
    );
    
  } catch (error) {
    logError('❌ Ошибка генерации SEO', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Создает плитку тегов для активной категории
 */
function generateTagTilesForActiveCategory() {
  try {
    const categoryData = getActiveCategoryData();
    
    if (!categoryData) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        'Откройте детальный лист категории',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Создаем плитку тегов через AI...',
      '⏳ Создание',
      -1
    );
    
    // Получаем все категории для подбора релевантных ссылок
    const allCategories = getAllCategoriesFromMainList();
    
    // Получаем ключевые слова с назначением "tag_tiles"
    const keywords = getKeywordsForCategory(categoryData.id);
    const tagKeywords = keywords.filter(kw => kw.assignment === 'tag_tiles');
    
    // Генерируем плитку тегов через AI
    const tagTiles = generateTagTilesWithAI(categoryData, tagKeywords, allCategories);
    
    // Записываем в детальный лист
    writeTagTilesToDetailSheet(tagTiles);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Создано ${tagTiles.length} тегов!`,
      '✅ Готово',
      5
    );
    
  } catch (error) {
    logError('❌ Ошибка создания плитки тегов', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Показывает лист ключевых слов
 */
function showKeywordsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.KEYWORDS);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert(
        'Лист не найден',
        'Сначала создайте структуру листов',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    ss.setActiveSheet(sheet);
    
  } catch (error) {
    logError('❌ Ошибка показа листа ключевых слов', error);
  }
}

/**
 * Обёртка для вызова AI-анализа из меню
 */
function analyzeActiveCategoryFromMenu() {
  try {
    // Проверяем, что открыт правильный лист
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        'Откройте детальный лист категории для анализа.\n\n' +
        'Используйте: InSales Manager → Категории → 🔍 Найти и открыть категорию',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Вызываем функцию анализа
    analyzeActiveCategory();
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Ошибка',
      `Не удалось выполнить анализ:\n\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}