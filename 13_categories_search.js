// ========================================
// ПОИСК КАТЕГОРИЙ
// ========================================

/**
 * МАППИНГ полей: Google Sheets → InSales API
 * Возвращает правильное название поля для InSales
 */
function getInSalesFieldName(sheetFieldName) {
  const mapping = {
    'Раздел для страниц через запятую': 'Раздел для страниц через запятую',
    'Название для пунктов Д-меню': 'Название для пунктов Д-меню',
    'Показывать на главной?': 'Показывать на главной?',
    'Товары и noindex': 'Товары и noindex',
    'Название для Баннер 2': 'Название для Баннер 2',
    'ID товаров через запятую': 'ID товаров через запятую',
    'Показывать иконку для меню': 'Показывать иконку для меню',
    'Баннер 1': 'Баннер 1 изображение',
    'Баннер 2': 'Баннер 2 изображение',
    'Баннер 2 - ссылка': 'Баннер 2 - ссылка',
    'Баннер 2 - мобильное изображение': 'Баннер 2 - мобильное изображение',
    'Баннер 1 - ссылка': 'Баннер 1 - ссылка',
    'Официальное изображение': 'Официальное изображение',
    'Показывать Баннер – Изображение': 'Показывать Баннер – Изображение',
    'Вертикальное изображение': 'Вертикальное изображение',
    'Баннер – Мобильное изображение': 'Баннер – Мобильное изображение',
    'Товары и категории в ссылке': 'Товары и категории в ссылке',
    'Ссылка категории в слайдере': 'Ссылка категории в слайдере',
    'Название категории в меню': 'Название категории в меню',
    'Название категории в списке': 'Название категории в списке'
  };
  
  return mapping[sheetFieldName] || sheetFieldName;
}

/**
 * Диалог поиска категорий по названию
 * Вызывается из меню
 */
function showCategorySearchDialog() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { 
            font-family: Arial, sans-serif; 
            padding: 20px;
            margin: 0;
          }
          h3 { 
            margin-top: 0;
            color: #1976d2;
          }
          .search-container {
            margin: 20px 0;
          }
          .search-field { 
            width: 100%; 
            padding: 12px; 
            margin: 10px 0; 
            border: 2px solid #ddd; 
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
          }
          .search-field:focus {
            outline: none;
            border-color: #1976d2;
          }
          .hint {
            color: #666;
            font-size: 12px;
            margin: 5px 0;
          }
          .results {
            margin: 20px 0;
            max-height: 400px;
            overflow-y: auto;
            border: 1px solid #ddd;
            border-radius: 4px;
            display: none;
          }
          .results.visible {
            display: block;
          }
          .result-item {
            padding: 12px;
            border-bottom: 1px solid #eee;
            cursor: pointer;
            transition: background 0.2s;
          }
          .result-item:hover {
            background: #f5f5f5;
          }
          .result-item:last-child {
            border-bottom: none;
          }
          .result-title {
            font-weight: bold;
            color: #333;
            margin-bottom: 4px;
          }
          .result-path {
            font-size: 12px;
            color: #666;
          }
          .result-stats {
            font-size: 11px;
            color: #999;
            margin-top: 4px;
          }
          .buttons { 
            margin-top: 20px; 
            text-align: right;
            border-top: 1px solid #eee;
            padding-top: 15px;
          }
          button { 
            padding: 10px 20px; 
            margin-left: 10px;
            cursor: pointer;
            border: none;
            border-radius: 4px;
            font-size: 14px;
          }
          .btn-cancel {
            background: #f1f3f4;
            color: #333;
          }
          .btn-cancel:hover {
            background: #e0e0e0;
          }
          .primary { 
            background: #1976d2; 
            color: white;
          }
          .primary:hover {
            background: #1565c0;
          }
          .primary:disabled {
            background: #ccc;
            cursor: not-allowed;
          }
          .loading {
            text-align: center;
            padding: 20px;
            color: #666;
          }
          .no-results {
            text-align: center;
            padding: 40px;
            color: #999;
          }
          .selected {
            background: #e3f2fd !important;
            border-left: 3px solid #1976d2;
          }
        </style>
      </head>
      <body>
        <h3>🔍 Поиск категорий</h3>
        
        <div class="search-container">
          <input 
            type="text" 
            id="searchInput" 
            class="search-field" 
            placeholder="Введите название категории..."
            autocomplete="off"
          >
          <div class="hint">
            💡 Можно искать по части названия: "бинокли", "10x25", "полевые"
          </div>
        </div>
        
        <div id="results" class="results">
          <div class="loading">Введите текст для поиска...</div>
        </div>
        
        <div class="buttons">
          <button class="btn-cancel" onclick="google.script.host.close()">
            Отмена
          </button>
          <button 
            id="openBtn" 
            class="primary" 
            onclick="openSelectedCategory()" 
            disabled
          >
            Открыть детальный лист
          </button>
        </div>
        
        <script>
          let allCategories = [];
          let selectedCategory = null;
          
          // Загружаем категории при открытии
          window.onload = function() {
            loadCategories();
            
            // Поиск при вводе
            document.getElementById('searchInput').addEventListener('input', function(e) {
              performSearch(e.target.value);
            });
            
            // Enter для поиска
            document.getElementById('searchInput').addEventListener('keypress', function(e) {
              if (e.key === 'Enter') {
                performSearch(e.target.value);
              }
            });
          };
          
          function loadCategories() {
            google.script.run
              .withSuccessHandler(function(categories) {
                allCategories = categories;
                console.log('Загружено категорий:', categories.length);
              })
              .withFailureHandler(function(error) {
                showError('Ошибка загрузки категорий: ' + error.message);
              })
              .getCategoriesForSearch();
          }
          
          function performSearch(query) {
            const resultsDiv = document.getElementById('results');
            const openBtn = document.getElementById('openBtn');
            
            if (!query || query.length < 2) {
              resultsDiv.innerHTML = '<div class="loading">Введите минимум 2 символа...</div>';
              resultsDiv.classList.remove('visible');
              openBtn.disabled = true;
              return;
            }
            
            const searchQuery = query.toLowerCase();
            const filtered = allCategories.filter(cat => 
              cat.title.toLowerCase().includes(searchQuery) ||
              cat.path.toLowerCase().includes(searchQuery)
            );
            
            if (filtered.length === 0) {
              resultsDiv.innerHTML = '<div class="no-results">😔 Категории не найдены</div>';
              resultsDiv.classList.add('visible');
              openBtn.disabled = true;
              return;
            }
            
            // Отображаем результаты
            let html = '';
            filtered.forEach((cat, index) => {
              const indent = '&nbsp;&nbsp;'.repeat(cat.level);
              const prefix = cat.level > 0 ? '└─ ' : '';
              
              html += \`
                <div class="result-item" onclick="selectCategory(\${index})" id="result_\${index}">
                  <div class="result-title">
                    \${indent}\${prefix}\${cat.title}
                  </div>
                  <div class="result-path">📂 \${cat.path}</div>
                  <div class="result-stats">
                    📦 Товаров: \${cat.productsCount} | ✅ В наличии: \${cat.inStockCount}
                  </div>
                </div>
              \`;
            });
            
            resultsDiv.innerHTML = html;
            resultsDiv.classList.add('visible');
            
            // Сохраняем отфильтрованные результаты
            window.searchResults = filtered;
          }
          
          function selectCategory(index) {
            selectedCategory = window.searchResults[index];
            
            // Снимаем выделение со всех
            document.querySelectorAll('.result-item').forEach(item => {
              item.classList.remove('selected');
            });
            
            // Выделяем выбранный
            document.getElementById('result_' + index).classList.add('selected');
            
            // Активируем кнопку
            document.getElementById('openBtn').disabled = false;
          }
          
          function openSelectedCategory() {
            if (!selectedCategory) {
              alert('Выберите категорию');
              return;
            }
            
            console.log('Открываем категорию:', selectedCategory);
            
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('Ошибка: ' + error.message);
              })
              .createDetailedCategorySheet(selectedCategory);
          }
          
          function showError(message) {
            document.getElementById('results').innerHTML = 
              '<div class="no-results">❌ ' + message + '</div>';
          }
        </script>
      </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Поиск категорий InSales');
}

/**
 * Получает категории для поиска (вызывается из диалога)
 */
function getCategoriesForSearch() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    
    if (!sheet) {
      throw new Error('Лист категорий не найден. Сначала загрузите категории.');
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      throw new Error('Нет загруженных категорий. Используйте "Загрузить категории".');
    }
    
    const categories = [];
    
    // Пропускаем заголовок (строка 1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      categories.push({
        id: row[MAIN_LIST_COLUMNS.CATEGORY_ID - 1],
        parent_id: row[MAIN_LIST_COLUMNS.PARENT_ID - 1],
        level: row[MAIN_LIST_COLUMNS.LEVEL - 1],
        path: row[MAIN_LIST_COLUMNS.HIERARCHY_PATH - 1],
        title: row[MAIN_LIST_COLUMNS.TITLE - 1]
        .replace(/[\s└─]/g, '')  // Убираем визуальные символы
        .replace(/([а-яё])([А-ЯЁ])/g, '$1 $2')  // Разделяем слитные слова
        .trim(),
        url: row[MAIN_LIST_COLUMNS.URL - 1],
        productsCount: row[MAIN_LIST_COLUMNS.PRODUCTS_COUNT - 1] || 0,
        inStockCount: row[MAIN_LIST_COLUMNS.IN_STOCK_COUNT - 1] || 0
      });
    }
    
    logInfo(`📋 Подготовлено ${categories.length} категорий для поиска`);
    
    return categories;
    
  } catch (error) {
    logError('❌ Ошибка получения категорий для поиска', error);
    throw error;
  }
}

// ========================================
// СОЗДАНИЕ ДЕТАЛЬНОГО ЛИСТА КАТЕГОРИИ
// ========================================

/**
 * Создает детальный лист для редактирования категории
 */
function createDetailedCategorySheet(categoryData) {
  const context = "Создание детального листа";
  
  try {
    logInfo(`🎨 Создаем детальный лист для категории: ${categoryData.title}`, null, context);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Создаем лист для "${categoryData.title}"...`,
      '⏳ Создание',
      -1
    );
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `${CATEGORY_SHEETS.DETAIL_PREFIX}${categoryData.title}`;
    
    // Проверяем, существует ли уже такой лист
    let sheet = ss.getSheetByName(sheetName);
    
    if (sheet) {
      // Лист существует - просто активируем его
      ss.setActiveSheet(sheet);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Лист уже существует, переключились на него',
        'ℹ️ Информация',
        3
      );
      
      return;
    }
    
    // Создаем новый лист
    sheet = ss.insertSheet(sheetName);
    
    // 1. Загружаем полные данные категории из InSales
    const fullCategoryData = loadFullCategoryData(categoryData.id);
    
    // 2. Загружаем товары категории
    const categoryProducts = loadCategoryProducts(categoryData.id);
    
    // 3. Настраиваем структуру листа
    setupDetailedCategorySheet(sheet, fullCategoryData, categoryProducts);
    
    // Активируем созданный лист
    ss.setActiveSheet(sheet);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Детальный лист создан! Товаров: ${categoryProducts.length}`,
      '✅ Готово',
      5
    );
    
    logInfo('✅ Детальный лист успешно создан', null, context);
    
  } catch (error) {
    logError('❌ Ошибка создания детального листа', error, context);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка: ' + error.message,
      '❌ Ошибка',
      10
    );
    throw error;
  }
}

/**
 * Загружает полные данные категории из InSales
 */
function loadFullCategoryData(categoryId) {
  const context = `Загрузка данных категории ${categoryId}`;
  
  try {
    logInfo('📥 Загружаем полные данные категории', null, context);
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }
    
    const url = `${credentials.baseUrl}${CATEGORY_API_ENDPOINTS.COLLECTION_BY_ID.replace('{id}', categoryId)}`;
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Ошибка API: ${response.getResponseCode()}`);
    }
    
    const categoryData = JSON.parse(response.getContentText());
    
    logInfo('✅ Данные категории загружены', null, context);
    
    return categoryData;
    
  } catch (error) {
    logError('❌ Ошибка загрузки данных категории', error, context);
    throw error;
  }
}

/**
 * Загружает товары категории
 */
function loadCategoryProducts(categoryId) {
  const context = `Загрузка товаров категории ${categoryId}`;
  
  try {
    logInfo('📦 Загружаем товары категории', null, context);
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }
    
    const allProducts = [];
    let page = 1;
    const perPage = 100;
    
    while (true) {
      const url = `${credentials.baseUrl}${CATEGORY_API_ENDPOINTS.COLLECTION_PRODUCTS.replace('{id}', categoryId)}&per_page=${perPage}&page=${page}`;
      
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
        break;
      }
      
      const products = JSON.parse(response.getContentText());
      
      if (!products || products.length === 0) {
        break;
      }
      
      allProducts.push(...products);
      
      if (products.length < perPage || page >= 10) {
        break;
      }
      
      page++;
      Utilities.sleep(300);
    }
    
    logInfo(`✅ Загружено товаров: ${allProducts.length}`, null, context);
    
    return allProducts;
    
  } catch (error) {
    logError('❌ Ошибка загрузки товаров категории', error, context);
    return [];
  }
}

/**
 * Настраивает детальный лист категории (ИСПРАВЛЕНО - ПРАВИЛЬНЫЕ ПОЛЯ!)
 */
function setupDetailedCategorySheet(sheet, categoryData, products) {
  const context = "Настройка детального листа";
  
  try {
    console.log('🎨 Настраиваем детальный лист категории');
    
    // Очищаем лист
    sheet.clear();
    
    // ========================================
    // СЕКЦИЯ 1: ИНФОРМАЦИЯ О КАТЕГОРИИ
    // ========================================
    
    sheet.getRange('A1').setValue('📁 ИНФОРМАЦИЯ О КАТЕГОРИИ')
         .setFontWeight('bold')
         .setFontSize(14)
         .setBackground('#1976d2')
         .setFontColor('#ffffff');
    sheet.getRange('A1:F1').merge();
    
    const infoData = [
      ['ID категории:', categoryData.id],
      ['Название:', categoryData.title],
      ['URL:', categoryData.url || categoryData.permalink],
      ['Путь в иерархии:', categoryData.path || 'Корневая категория'],
      ['', ''],
      ['Маркерный запрос:', '']
    ];
    
    sheet.getRange(2, 1, infoData.length, 2).setValues(infoData);
    sheet.getRange('A2:A7').setFontWeight('bold').setBackground('#f1f3f4');
    sheet.getRange('B2:B7').setWrap(true);
    
    // Ссылка на админку InSales
    const adminLink = `https://myshop-on665.myinsales.ru/admin2/collections/${categoryData.id}`;
    sheet.getRange('A8').setValue('Админка InSales:').setFontWeight('bold').setBackground('#f1f3f4');
    const adminCell = sheet.getRange('B8');
    adminCell.setValue(adminLink);
    adminCell.setFontColor('#1155cc');
    adminCell.setFontLine('underline');
    
    // Ширина колонок
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 500);
    
    // ========================================
    // СЕКЦИЯ 2: SEO ДАННЫЕ (ИСПРАВЛЕНО!)
    // ========================================
    
    sheet.getRange('A12').setValue('🎯 SEO ДАННЫЕ')
         .setFontWeight('bold')
         .setFontSize(14)
         .setBackground('#4caf50')
         .setFontColor('#ffffff');
    sheet.getRange('A12:F12').merge();
    
    // ИСПРАВЛЕНО: Извлекаем H1 через справочник полей
    let h1Value = getFieldValueByName(categoryData, 'H1');

    // Если H1 нет - используем html_title как запасной вариант
    if (!h1Value || h1Value.trim() === '') {
      h1Value = categoryData.html_title || categoryData.title || '';
    }

    const seoData = [
      ['SEO Title:', categoryData.html_title || categoryData.title || ''],
      ['Meta Description:', categoryData.meta_description || ''],
      ['H1 заголовок:', h1Value],
      ['Ключевые слова:', categoryData.meta_keywords || '']
    ];
    
    sheet.getRange(13, 1, seoData.length, 2).setValues(seoData);
    sheet.getRange('A13:A16').setFontWeight('bold').setBackground('#f1f3f4');
    sheet.getRange('B13:B16').setWrap(true);
    
    // Описание категории
    sheet.getRange('A17').setValue('Описание категории:').setFontWeight('bold').setBackground('#f1f3f4');
    sheet.getRange('B17').setValue(categoryData.description || '').setWrap(true);
    sheet.setRowHeight(17, 150);

    // ========================================
    // СЕКЦИЯ 3: ТАБЛИЦА КЛЮЧЕВЫХ СЛОВ ДЛЯ ПЛИТОК (НОВОЕ!)
    // ========================================

    const keywordsStartRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_START;
    const keywordsHeaderRow = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_HEADER_ROW;
    const keywordsDataStart = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_DATA_START;

    // Заголовок секции
    sheet.getRange(keywordsStartRow, 1, 1, 7)
      .merge()
      .setValue('📝 КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПЛИТОК ТЕГОВ (ручной ввод)')
      .setBackground('#E8F0FE')
      .setFontWeight('bold')
      .setFontSize(12);

    // Заголовки столбцов
    const keywordHeaders = [
      '☑️',                    // A: Чекбокс
      'Ключевое слово',        // B
      'Тип плитки',            // C
      'Текст анкора',          // D
      'URL/ID категории',      // E
      'Статус категории',      // F
      'ID родителя (новая)'    // G
    ];

    sheet.getRange(keywordsHeaderRow, 1, 1, keywordHeaders.length).setValues([keywordHeaders]);
    sheet.getRange(keywordsHeaderRow, 1, 1, keywordHeaders.length)
      .setBackground('#D0E1F9')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    // Форматирование столбцов
    sheet.setColumnWidth(1, 40);   // Чекбокс
    sheet.setColumnWidth(2, 200);  // Ключевое слово
    sheet.setColumnWidth(3, 120);  // Тип плитки
    sheet.setColumnWidth(4, 250);  // Текст анкора
    sheet.setColumnWidth(5, 150);  // URL/ID категории
    sheet.setColumnWidth(6, 150);  // Статус категории
    sheet.setColumnWidth(7, 120);  // ID родителя

    // ДИНАМИЧЕСКИЙ РАЗМЕР: минимум 5 строк для будущего заполнения
    const cols = DETAIL_SHEET_SECTIONS.TAG_KEYWORDS_COLUMNS;
    const keywordsRowCount = Math.max(5, 5); // Пока минимум 5, позже можно будет передавать количество ключевых слов

    console.log(`[DEBUG] Создаём таблицу ключевых слов с ${keywordsRowCount} строками`);

    // Добавляем валидацию для столбца "Тип плитки"
    const tileTypeColumn = cols.TILE_TYPE;
    const tileTypeRange = sheet.getRange(keywordsDataStart, tileTypeColumn, keywordsRowCount, 1);
    const tileTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Верхняя', 'Нижняя'], true)
      .setAllowInvalid(false)
      .build();
    tileTypeRange.setDataValidation(tileTypeRule);

    // Добавляем чекбоксы в первый столбец
    const checkboxColumn = cols.CHECKBOX;
    const checkboxRange = sheet.getRange(keywordsDataStart, checkboxColumn, keywordsRowCount, 1);
    checkboxRange.insertCheckboxes();

    // Добавляем формулы для автоматической проверки статуса категории (столбец F)
    const statusColumn = cols.CATEGORY_STATUS;
    const categoryLinkColumn = cols.CATEGORY_LINK;

    for (let i = 0; i < keywordsRowCount; i++) {
      const row = keywordsDataStart + i;
      const linkCell = columnToLetterInternal(categoryLinkColumn) + row;
      const formula = `=IF(ISBLANK(${linkCell}), "Не указана", "Проверить")`;
      sheet.getRange(row, statusColumn).setFormula(formula);
    }

    console.log('✅ Таблица ключевых слов инициализирована');

    // ========================================
    // СЕКЦИЯ 4: СТАТИСТИКА ТОВАРОВ (обновлена позиция!)
    // ========================================

    // ДИНАМИЧЕСКОЕ ПОЗИЦИОНИРОВАНИЕ: статистика начинается после фактического количества строк ключевых слов
    const statsStartRow = keywordsDataStart + keywordsRowCount + 3;

    sheet.getRange(statsStartRow, 1).setValue('📊 СТАТИСТИКА ТОВАРОВ')
         .setFontWeight('bold')
         .setFontSize(14)
         .setBackground('#ff9800')
         .setFontColor('#ffffff');
    sheet.getRange(statsStartRow, 1, 1, 6).merge();
    
    const inStockCount = products.filter(p => 
      p.variants && p.variants.some(v => v.quantity && v.quantity > 0)
    ).length;
    
    const statsData = [
      ['Всего товаров:', products.length],
      ['В наличии:', inStockCount],
      ['Нет в наличии:', products.length - inStockCount],
      ['Процент наличия:', products.length > 0 ? Math.round(inStockCount / products.length * 100) + '%' : '0%']
    ];

    sheet.getRange(statsStartRow + 1, 1, statsData.length, 2).setValues(statsData);
    sheet.getRange(statsStartRow + 1, 1, statsData.length, 1).setFontWeight('bold').setBackground('#f1f3f4');

    // ========================================
    // СЕКЦИЯ 5: ТЕКУЩИЕ ТОВАРЫ (обновлена позиция!)
    // ========================================

    const productsStartRow = statsStartRow + 1 + statsData.length + 2;

    sheet.getRange(productsStartRow, 1).setValue('🛒 ТЕКУЩИЕ ТОВАРЫ В КАТЕГОРИИ')
         .setFontWeight('bold')
         .setFontSize(14)
         .setBackground('#9c27b0')
         .setFontColor('#ffffff');
    sheet.getRange(productsStartRow, 1, 1, 6).merge();

    const productHeaders = [
      'Название', 'Артикул', 'Цена', 'В наличии', 'ID', '☑️'
    ];

    sheet.getRange(productsStartRow + 1, 1, 1, productHeaders.length).setValues([productHeaders]);
    sheet.getRange(productsStartRow + 1, 1, 1, productHeaders.length)
         .setFontWeight('bold')
         .setBackground('#e1bee7')
         .setHorizontalAlignment('center');

    const productsDataRow = productsStartRow + 2;

    if (products.length > 0) {
      const productRows = products.map(product => {
        const variant = product.variants && product.variants[0];
        const inStock = variant && variant.quantity > 0 ? 'Да' : 'Нет';
        const price = variant ? variant.price : (product.price || '');

        return [
          product.title,                // A - Название
          variant ? variant.sku : '',   // B - Артикул
          price || '',                   // C - Цена
          inStock,                      // D - В наличии
          product.id,                   // E - ID
          false                         // F - Чекбокс
        ];
      });

      sheet.getRange(productsDataRow, 1, productRows.length, productRows[0].length).setValues(productRows);
      sheet.getRange(productsDataRow, 6, productRows.length, 1).insertCheckboxes();  // Чекбоксы в колонке F
      sheet.getRange(productsDataRow, 3, productRows.length, 1).setNumberFormat('#,##0.00 ₽');  // Цена в колонке C

      // Ширина колонок
      sheet.setColumnWidth(1, 450);  // Название
      sheet.setColumnWidth(2, 120);  // Артикул
      sheet.setColumnWidth(3, 100);  // Цена
      sheet.setColumnWidth(4, 100);  // В наличии
      sheet.setColumnWidth(5, 100);  // ID
      sheet.setColumnWidth(6, 40);   // Чекбокс
    }

    // ========================================
    // СЕКЦИЯ 6: ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ КАТЕГОРИИ (обновлена позиция!)
    // ========================================

    const extraFieldsStartRow = productsDataRow + products.length + 3;

    sheet.getRange(extraFieldsStartRow, 1).setValue('⚙️ ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ КАТЕГОРИИ')
        .setFontWeight('bold')
        .setFontSize(14)
        .setBackground('#673ab7')
        .setFontColor('#ffffff');
    sheet.getRange(extraFieldsStartRow, 1, 1, 2).merge();

    // Загружаем справочник полей
    const collectionFields = loadCollectionFieldsDictionary();
    const extraFieldsData = [];

    if (categoryData.field_values && Array.isArray(categoryData.field_values)) {
      for (let i = 0; i < categoryData.field_values.length; i++) {
        const fieldValue = categoryData.field_values[i];
        
        if (!fieldValue || !fieldValue.collection_field_id) continue;
        
        // Находим название поля из справочника
        const fieldInfo = collectionFields.find(f => f.id === fieldValue.collection_field_id);
        const fieldName = fieldInfo ? fieldInfo.title : `Неизвестное поле (ID: ${fieldValue.collection_field_id})`;
        
        // Получаем значение
        let value = fieldValue.value || '';
        
        // Для длинных значений обрезаем
        let displayValue = value.toString();
        if (displayValue.length > 100) {
          displayValue = displayValue.substring(0, 100) + '...';
        }
        
        extraFieldsData.push([
          fieldName + ':',
          displayValue || '(пусто)'
        ]);
      }
    }

    if (extraFieldsData.length > 0) {
      sheet.getRange(extraFieldsStartRow + 2, 1, extraFieldsData.length, 2).setValues(extraFieldsData);
      sheet.getRange(extraFieldsStartRow + 2, 1, extraFieldsData.length, 1)
          .setFontWeight('bold')
          .setBackground('#f1f3f4');
      sheet.getRange(extraFieldsStartRow + 2, 2, extraFieldsData.length, 1).setWrap(true);
    }

    console.log('✅ Показано дополнительных полей:', extraFieldsData.length);

    // ========================================
    // СЕКЦИЯ 7: ПЛИТКА ТЕГОВ - ВЕРХНЯЯ (существующие + новые)
    // ========================================

    const topTagsStartRow = extraFieldsStartRow + 2 + extraFieldsData.length + 3;

    sheet.getRange(topTagsStartRow, 1).setValue('🏷️ ПЛИТКА ТЕГОВ - ВЕРХНЯЯ (над описанием категории)')
        .setFontWeight('bold')
        .setFontSize(14)
        .setBackground('#00bcd4')
        .setFontColor('#ffffff');
    sheet.getRange(topTagsStartRow, 1, 1, 8).merge();

    sheet.getRange(topTagsStartRow + 1, 1).setValue('Инструкция:')
        .setFontWeight('bold')
        .setBackground('#e0f7fa');
    sheet.getRange(topTagsStartRow + 1, 2, 1, 7).merge()
        .setValue('Слева - существующие теги из InSales. Справа - новые сгенерированные теги. Сравните и выберите лучший вариант.')
        .setWrap(true)
        .setBackground('#e0f7fa');

    // Заголовки: Существующие теги | Новые теги
    const topTagsHeaders = [
      'БЫЛО: Текст ссылки', 'БЫЛО: URL', 'БЫЛО: ☑️ Включить', '',
      'СТАЛО: Текст анкора', 'СТАЛО: URL', 'СТАЛО: ID категории', 'СТАЛО: Примечание'
    ];
    sheet.getRange(topTagsStartRow + 3, 1, 1, topTagsHeaders.length).setValues([topTagsHeaders]);
    sheet.getRange(topTagsStartRow + 3, 1, 1, 3)
        .setFontWeight('bold')
        .setFontFamily('Roboto')
        .setFontSize(8)
        .setBackground('#ffccbc')  // Красноватый для старых
        .setHorizontalAlignment('center')
        .setVerticalAlignment('top');
    sheet.getRange(topTagsStartRow + 3, 4, 1, 1)
        .setBackground('#ffffff');  // Разделитель
    sheet.getRange(topTagsStartRow + 3, 5, 1, 4)
        .setFontWeight('bold')
        .setFontFamily('Roboto')
        .setFontSize(8)
        .setBackground('#c8e6c9')  // Зеленоватый для новых
        .setHorizontalAlignment('center')
        .setVerticalAlignment('top');

    // Загружаем существующие теги из InSales
    const topTagsHTML = getFieldValueByName(categoryData, 'Блок ссылок сверху');
    const topTags = parseTagsFromHTML(topTagsHTML);

    console.log(`📋 Парсинг верхних тегов: найдено ${topTags.length} существующих ссылок`);

    // ДИНАМИЧЕСКИЙ РАЗМЕР: минимум 3 строки для будущего заполнения
    const topTagsRowCount = Math.max(3, topTags.length);
    console.log(`[DEBUG] Создаём верхнюю плитку с ${topTagsRowCount} строками`);

    const topTagsData = [];
    for (let i = 0; i < topTagsRowCount; i++) {
      if (i < topTags.length) {
        topTagsData.push([
          topTags[i].text,   // A: БЫЛО - Текст
          topTags[i].url,    // B: БЫЛО - URL
          true,              // C: БЫЛО - Чекбокс (по умолчанию включен)
          '',                // D: Разделитель
          '',                // E: СТАЛО - Текст анкора (будет заполнено при генерации)
          '',                // F: СТАЛО - URL
          '',                // G: СТАЛО - ID категории
          ''                 // H: СТАЛО - Примечание
        ]);
      } else {
        topTagsData.push(['', '', false, '', '', '', '', '']);
      }
    }

    sheet.getRange(topTagsStartRow + 4, 1, topTagsData.length, topTagsData[0].length).setValues(topTagsData);

    // Добавляем чекбоксы в колонку C
    sheet.getRange(topTagsStartRow + 4, 3, topTagsData.length, 1).insertCheckboxes();

    // Форматирование: старые теги - красноватый фон, новые - зеленоватый
    sheet.getRange(topTagsStartRow + 4, 1, topTagsData.length, 3)
        .setBackground('#ffebee')
        .setFontFamily('Roboto')
        .setFontSize(8)
        .setVerticalAlignment('top')
        .setHorizontalAlignment('left');
    sheet.getRange(topTagsStartRow + 4, 5, topTagsData.length, 4)
        .setBackground('#e8f5e9')
        .setFontFamily('Roboto')
        .setFontSize(8)
        .setVerticalAlignment('top')
        .setHorizontalAlignment('left');

    // Ширина колонок
    sheet.setColumnWidth(1, 250);  // БЫЛО: Текст
    sheet.setColumnWidth(2, 300);  // БЫЛО: URL
    sheet.setColumnWidth(3, 80);   // БЫЛО: Чекбокс
    sheet.setColumnWidth(4, 30);   // Разделитель
    sheet.setColumnWidth(5, 250);  // СТАЛО: Текст
    sheet.setColumnWidth(6, 300);  // СТАЛО: URL
    sheet.setColumnWidth(7, 120);  // СТАЛО: ID
    sheet.setColumnWidth(8, 150);  // СТАЛО: Примечание

    // Группировка строк (временно отключена из-за ошибок API)
    // TODO: Реализовать группировку строк после исправления API

    // Добавляем поле для финального HTML кода верхней плитки
    const topHTMLRow = topTagsStartRow + 4 + topTagsRowCount + 1;
    sheet.getRange(topHTMLRow, 1, 1, 1).setValue('HTML код (финальный):').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.getRange(topHTMLRow, 2, 1, 7)
        .merge()
        .setValue('HTML код будет сгенерирован после создания плиток')
        .setWrap(true)
        .setBackground('#e8f5e9')
        .setFontFamily('Roboto')
        .setFontSize(9)
        .setFontColor('#666666');

    // ========================================
    // СЕКЦИЯ 8: ПЛИТКА ТЕГОВ - НИЖНЯЯ (существующие + новые)
    // ========================================

    // ДИНАМИЧЕСКИЙ РАСЧЕТ ПОЗИЦИИ для нижней плитки:
    // Структура верхней плитки: заголовок(1) + инструкция(1) + пустая(1) + заголовки(1) + данные(topTagsRowCount) + пустая(1) + "HTML код (финальный)"(1) + отступ(3)
    const bottomTagsStartRow = topTagsStartRow + 1 + 1 + 1 + 1 + topTagsRowCount + 1 + 1 + 3;
    console.log(`[DEBUG] Нижняя плитка начинается со строки ${bottomTagsStartRow} (после ${topTagsRowCount} строк верхней плитки + 1 строка для HTML)`);

    sheet.getRange(bottomTagsStartRow, 1).setValue('🏷️ ПЛИТКА ТЕГОВ - НИЖНЯЯ (под описанием категории)')
        .setFontWeight('bold')
        .setFontSize(14)
        .setBackground('#ff5722')
        .setFontColor('#ffffff');
    sheet.getRange(bottomTagsStartRow, 1, 1, 8).merge();

    sheet.getRange(bottomTagsStartRow + 1, 1).setValue('Инструкция:')
        .setFontWeight('bold')
        .setBackground('#ffe0db');
    sheet.getRange(bottomTagsStartRow + 1, 2, 1, 7).merge()
        .setValue('Слева - существующие теги из InSales. Справа - новые сгенерированные теги. Сравните и выберите лучший вариант.')
        .setWrap(true)
        .setBackground('#ffe0db');

    // Заголовки: Существующие теги | Новые теги
    const bottomTagsHeaders = [
      'БЫЛО: Текст ссылки', 'БЫЛО: URL', 'БЫЛО: ☑️ Включить', '',
      'СТАЛО: Текст анкора', 'СТАЛО: URL', 'СТАЛО: ID категории', 'СТАЛО: Примечание'
    ];
    sheet.getRange(bottomTagsStartRow + 3, 1, 1, bottomTagsHeaders.length).setValues([bottomTagsHeaders]);
    sheet.getRange(bottomTagsStartRow + 3, 1, 1, 3)
        .setFontWeight('bold')
        .setFontFamily('Roboto')
        .setFontSize(8)
        .setBackground('#ffccbc')  // Красноватый для старых
        .setHorizontalAlignment('center')
        .setVerticalAlignment('top');
    sheet.getRange(bottomTagsStartRow + 3, 4, 1, 1)
        .setBackground('#ffffff');  // Разделитель
    sheet.getRange(bottomTagsStartRow + 3, 5, 1, 4)
        .setFontWeight('bold')
        .setFontFamily('Roboto')
        .setFontSize(8)
        .setBackground('#c8e6c9')  // Зеленоватый для новых
        .setHorizontalAlignment('center')
        .setVerticalAlignment('top');

    // Загружаем существующие теги из InSales
    const bottomTagsHTML = getFieldValueByName(categoryData, 'Блок ссылок');
    const bottomTags = parseTagsFromHTML(bottomTagsHTML);

    console.log(`📋 Парсинг нижних тегов: найдено ${bottomTags.length} существующих ссылок`);

    // ДИНАМИЧЕСКИЙ РАЗМЕР: минимум 5 строк для будущего заполнения (нижняя плитка обычно больше)
    const bottomTagsRowCount = Math.max(5, bottomTags.length);
    console.log(`[DEBUG] Создаём нижнюю плитку с ${bottomTagsRowCount} строками`);

    const bottomTagsData = [];
    for (let i = 0; i < bottomTagsRowCount; i++) {
      if (i < bottomTags.length) {
        bottomTagsData.push([
          bottomTags[i].text,   // A: БЫЛО - Текст
          bottomTags[i].url,    // B: БЫЛО - URL
          true,                 // C: БЫЛО - Чекбокс (по умолчанию включен)
          '',                   // D: Разделитель
          '',                   // E: СТАЛО - Текст анкора (будет заполнено при генерации)
          '',                   // F: СТАЛО - URL
          '',                   // G: СТАЛО - ID категории
          ''                    // H: СТАЛО - Примечание
        ]);
      } else {
        bottomTagsData.push(['', '', false, '', '', '', '', '']);
      }
    }

    sheet.getRange(bottomTagsStartRow + 4, 1, bottomTagsData.length, bottomTagsData[0].length).setValues(bottomTagsData);

    // Добавляем чекбоксы в колонку C
    sheet.getRange(bottomTagsStartRow + 4, 3, bottomTagsData.length, 1).insertCheckboxes();

    // Форматирование: старые теги - красноватый фон, новые - зеленоватый
    sheet.getRange(bottomTagsStartRow + 4, 1, bottomTagsData.length, 3)
        .setBackground('#ffebee')
        .setFontFamily('Roboto')
        .setFontSize(8)
        .setVerticalAlignment('top')
        .setHorizontalAlignment('left');
    sheet.getRange(bottomTagsStartRow + 4, 5, bottomTagsData.length, 4)
        .setBackground('#e8f5e9')
        .setFontFamily('Roboto')
        .setFontSize(8)
        .setVerticalAlignment('top')
        .setHorizontalAlignment('left');

    // Ширина колонок (уже установлены выше, не дублируем)

    // Группировка строк (временно отключена из-за ошибок API)
    // TODO: Реализовать группировку строк после исправления API

    // Добавляем поле для финального HTML кода нижней плитки
    const bottomHTMLRow = bottomTagsStartRow + 4 + bottomTagsRowCount + 1;
    sheet.getRange(bottomHTMLRow, 1, 1, 1).setValue('HTML код (финальный):').setFontWeight('bold').setBackground('#f0f0f0');
    sheet.getRange(bottomHTMLRow, 2, 1, 7)
        .merge()
        .setValue('HTML код будет сгенерирован после создания плиток')
        .setWrap(true)
        .setBackground('#e8f5e9')
        .setFontFamily('Roboto')
        .setFontSize(9)
        .setFontColor('#666666');

    console.log('✅ Детальный лист настроен с блоками сравнения плиток (было/стало) и чекбоксами');
    
  } catch (error) {
    console.error('❌ Ошибка настройки детального листа:', error);
    throw error;
  }
}

/**
 * Получает учетные данные InSales синхронно
 */
function getInsalesCredentialsSync() {
  try {
    const config = getInsalesConfig();
    
    if (!config || !config.apiKey || !config.password || !config.shop) {
      throw new Error('Учетные данные InSales не настроены в 01_config.gs');
    }
    
    return {
      apiKey: config.apiKey,
      password: config.password,
      shop: config.shop,
      baseUrl: config.baseUrl
    };
    
  } catch (error) {
    console.error('❌ Ошибка получения учетных данных InSales:', error);
    return null;
  }
}

/**
 * ТЕСТ: ПОЛНАЯ диагностика field_values из InSales API
 */
function testCategoryDataFromAPI() {
  const categoryId = 9071624; // ID вашей категории "Театральные"
  
  Logger.clear();
  Logger.log('=== ТЕСТ ЗАГРУЗКИ ДАННЫХ КАТЕГОРИИ ===');
  Logger.log(`ID категории: ${categoryId}`);
  
  try {
    const credentials = getInsalesCredentialsSync();
    const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
    
    Logger.log(`URL: ${url}`);
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    Logger.log(`HTTP статус: ${response.getResponseCode()}`);
    
    if (response.getResponseCode() === 200) {
      const categoryData = JSON.parse(response.getContentText());
      
      Logger.log('\n=== ОСНОВНЫЕ ПОЛЯ ===');
      Logger.log(`ID: ${categoryData.id}`);
      Logger.log(`title: ${categoryData.title}`);
      Logger.log(`html_title: ${categoryData.html_title}`);
      Logger.log(`meta_description: ${categoryData.meta_description}`);
      
      Logger.log('\n=== FIELD_VALUES (ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ) ===');
      
      if (!categoryData.field_values) {
        Logger.log('❌ field_values отсутствует в ответе API!');
      } else if (!Array.isArray(categoryData.field_values)) {
        Logger.log('❌ field_values не является массивом!');
        Logger.log(`Тип: ${typeof categoryData.field_values}`);
        Logger.log(`Значение: ${JSON.stringify(categoryData.field_values)}`);
      } else if (categoryData.field_values.length === 0) {
        Logger.log('⚠️ field_values пустой массив (нет дополнительных полей)');
      } else {
        Logger.log(`✅ Найдено field_values: ${categoryData.field_values.length} полей`);
        Logger.log('\nПОЛНЫЙ СПИСОК FIELD_VALUES:');
        
        for (let i = 0; i < categoryData.field_values.length; i++) {
          const field = categoryData.field_values[i];
          Logger.log(`\n${i + 1}. ---------------------`);
          Logger.log(`   name: ${field.name || 'null'}`);
          Logger.log(`   title: ${field.title || 'null'}`);
          Logger.log(`   handle: ${field.handle || 'null'}`);
          Logger.log(`   value: ${field.value ? field.value.substring(0, 100) : 'null'}`);
          Logger.log(`   type: ${field.type || 'null'}`);
        }
      }
      
      Logger.log('\n=== ВСЕ КЛЮЧИ ВЕРХНЕГО УРОВНЯ ===');
      const allKeys = Object.keys(categoryData).sort();
      for (const key of allKeys) {
        const value = categoryData[key];
        if (typeof value === 'object' && value !== null) {
          Logger.log(`${key}: [${Array.isArray(value) ? 'array' : 'object'}]`);
        } else {
          Logger.log(`${key}: ${value}`);
        }
      }
      
      Logger.log('\n✅ ТЕСТ ЗАВЕРШЁН');
      
    } else {
      Logger.log('❌ ОШИБКА загрузки категории');
    }
    
  } catch (error) {
    Logger.log(`❌ ОШИБКА: ${error.message}`);
  }
  
  const logOutput = Logger.getLog();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<pre style="font-family: monospace; font-size: 11px; white-space: pre-wrap;">${logOutput}</pre>`)
      .setWidth(900)
      .setHeight(700),
    'Диагностика field_values'
  );
}

/**
 * Вспомогательная функция: конвертирует номер столбца в букву (1 → A, 2 → B, и т.д.)
 */
function columnToLetterInternal(column) {
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
 * Парсит блок ссылок из HTML в массив тегов
 */
function parseTagsFromHTML(html) {
  if (!html || html.trim() === '') {
    return [];
  }
  
  const tags = [];
  
  try {
    // Ищем все <a> теги
    const linkRegex = /<a[^>]*href=["']([^"']*)["'][^>]*>([^<]*)<\/a>/gi;
    let match;
    
    while ((match = linkRegex.exec(html)) !== null) {
      const url = match[1];
      const text = match[2];
      
      if (text && url) {
        tags.push({
          text: text.trim(),
          url: url.trim()
        });
      }
    }
    
    console.log(`📋 Найдено тегов: ${tags.length}`);
    
  } catch (error) {
    console.error('❌ Ошибка парсинга тегов:', error);
  }
  
  return tags;
}

/**
 * Извлекает значение поля из field_values
 */
function getFieldValueFromCategory(categoryData, fieldName) {
  if (!categoryData.field_values || !Array.isArray(categoryData.field_values)) {
    return null;
  }
  
  const field = categoryData.field_values.find(f => 
    f.name === fieldName || 
    f.handle === fieldName.toLowerCase().replace(/ /g, '_') ||
    f.title === fieldName
  );
  
  return field ? field.value : null;
}
/**
 * Загружает справочник полей категорий из InSales
 * Кэширует результат в Script Properties на 24 часа
 */
function loadCollectionFieldsDictionary() {
  try {
    // Проверяем кэш
    const cache = PropertiesService.getScriptProperties();
    const cachedData = cache.getProperty('collection_fields_cache');
    const cacheTime = cache.getProperty('collection_fields_cache_time');
    
    // Если кэш свежий (< 24 часов) - используем его
    if (cachedData && cacheTime) {
      const age = Date.now() - parseInt(cacheTime);
      if (age < 24 * 60 * 60 * 1000) {
        console.log('✅ Используем кэшированный справочник полей');
        return JSON.parse(cachedData);
      }
    }
    
    // Загружаем справочник из API
    console.log('📥 Загружаем справочник полей из InSales...');
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }
    
    const url = `${credentials.baseUrl}/admin/collection_fields.json`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Ошибка загрузки справочника: ${response.getResponseCode()}`);
    }
    
    const fields = JSON.parse(response.getContentText());
    
    console.log(`✅ Загружено полей: ${fields.length}`);
    
    // Сохраняем в кэш
    cache.setProperty('collection_fields_cache', JSON.stringify(fields));
    cache.setProperty('collection_fields_cache_time', Date.now().toString());
    
    return fields;
    
  } catch (error) {
    console.error('❌ Ошибка загрузки справочника полей:', error);
    return [];
  }
}

/**
 * Находит ID поля по его названию
 */
function getFieldIdByName(fieldName) {
  try {
    const fields = loadCollectionFieldsDictionary();
    
    const field = fields.find(f => 
      f.title === fieldName || 
      f.name === fieldName ||
      f.handle === fieldName
    );
    
    return field ? field.id : null;
    
  } catch (error) {
    console.error(`❌ Ошибка поиска ID поля "${fieldName}":`, error);
    return null;
  }
}

/**
 * Находит значение поля в categoryData по названию поля
 */
function getFieldValueByName(categoryData, fieldName) {
  try {
    // Получаем ID поля по названию
    const fieldId = getFieldIdByName(fieldName);
    
    if (!fieldId) {
      console.warn(`⚠️ Поле "${fieldName}" не найдено в справочнике`);
      return '';
    }
    
    // Ищем значение в field_values по collection_field_id
    if (!categoryData.field_values || !Array.isArray(categoryData.field_values)) {
      return '';
    }
    
    const fieldValue = categoryData.field_values.find(fv => 
      fv.collection_field_id === fieldId
    );
    
    return fieldValue && fieldValue.value ? fieldValue.value : '';
    
  } catch (error) {
    console.error(`❌ Ошибка получения значения поля "${fieldName}":`, error);
    return '';
  }
}

function testH1Reading() {
  const categoryId = 9171538;
  
  // Загружаем категорию
  const categoryData = loadFullCategoryData(categoryId);
  
  // Пробуем получить H1 новым способом
  const h1Value = getFieldValueByName(categoryData, 'H1');
  
  console.log('=== ТЕСТ ЧТЕНИЯ H1 ===');
  console.log('H1 из field_values:', h1Value);
  console.log('html_title:', categoryData.html_title);
  console.log('title:', categoryData.title);
  
  // Проверяем какое значение будет использовано
  const finalH1 = h1Value && h1Value.trim() !== '' ? h1Value : (categoryData.html_title || categoryData.title);
  console.log('\n✅ Итоговый H1:', finalH1);
}