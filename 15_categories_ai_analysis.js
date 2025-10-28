// ========================================
// AI АНАЛИЗ И РАСПРЕДЕЛЕНИЕ КЛЮЧЕВИКОВ
// ========================================

/**
 * AI-анализ текущего состояния категории
 */
function analyzeCurrentCategoryState(categoryData, products) {
  const context = `AI анализ категории ${categoryData.title}`;
  
  try {
    logInfo('🤖 Начинаем AI-анализ текущего состояния категории', null, context);
    
    // Собираем текущее состояние категории
    const currentState = {
      title: categoryData.title,
      description: categoryData.description || '',
      seoTitle: categoryData.seo_title || '',
      seoDescription: categoryData.seo_description || '',
      h1: categoryData.h1_title || categoryData.title,
      productsCount: products.length,
      inStockCount: products.filter(p => p.variants && p.variants.some(v => v.quantity > 0)).length,
      
      // Анализ листинга товаров
      productTitles: products.slice(0, 20).map(p => p.title),
      
      // Анализ характеристик товаров
      commonCharacteristics: extractCommonCharacteristics(products),
      
      // Анализ брендов
      brands: [...new Set(products.map(p => p.brand_title).filter(Boolean))],
      
      // Анализ ценового диапазона
      priceRange: calculatePriceRange(products)
    };
    
    logInfo('📊 Собрана информация о текущем состоянии');
    
    // Отправляем на анализ в OpenAI
    const analysis = performAIAnalysis(currentState);
    
    logInfo('✅ AI-анализ завершен', null, context);
    
    return analysis;
    
  } catch (error) {
    logError('❌ Ошибка AI-анализа', error, context);
    throw error;
  }
}

/**
 * Извлекает общие характеристики из товаров
 */
function extractCommonCharacteristics(products) {
  const charCounts = {};
  
  products.forEach(product => {
    if (product.characteristics && Array.isArray(product.characteristics)) {
      product.characteristics.forEach(char => {
        const propName = char.property_title || char.property_name || '';
        if (propName) {
          charCounts[propName] = (charCounts[propName] || 0) + 1;
        }
      });
    }
  });
  
  // Возвращаем характеристики, которые встречаются у >50% товаров
  const threshold = products.length * 0.5;
  const common = Object.entries(charCounts)
    .filter(([prop, count]) => count >= threshold)
    .map(([prop, count]) => ({ name: prop, frequency: count }))
    .sort((a, b) => b.frequency - a.frequency);
  
  return common;
}

/**
 * Вычисляет ценовой диапазон
 */
function calculatePriceRange(products) {
  const prices = products
    .map(p => {
      const variant = p.variants && p.variants[0];
      return variant ? variant.price : p.price;
    })
    .filter(price => price > 0);
  
  if (prices.length === 0) {
    return { min: 0, max: 0, avg: 0 };
  }
  
  const min = Math.min(...prices);
  const max = Math.max(...prices);
  const avg = prices.reduce((a, b) => a + b, 0) / prices.length;
  
  return { min, max, avg: Math.round(avg) };
}

/**
 * Отправляет данные на анализ в OpenAI
 */
function performAIAnalysis(currentState) {
  const context = "OpenAI анализ";
  
  try {
    const apiKey = getSetting('openaiApiKey');
    
    if (!apiKey) {
      throw new Error('OpenAI API ключ не настроен');
    }
    
    // Формируем промпт для анализа
    const prompt = `Проанализируй текущее состояние категории интернет-магазина и дай рекомендации по оптимизации.

ТЕКУЩЕЕ СОСТОЯНИЕ:
- Название категории: ${currentState.title}
- Товаров в категории: ${currentState.productsCount}
- В наличии: ${currentState.inStockCount}
- Ценовой диапазон: ${currentState.priceRange.min} - ${currentState.priceRange.max} руб.

SEO:
- SEO Title: ${currentState.seoTitle || 'отсутствует'}
- SEO Description: ${currentState.seoDescription || 'отсутствует'}
- H1: ${currentState.h1}

Описание категории: ${currentState.description || 'отсутствует'}

Примеры товаров:
${currentState.productTitles.slice(0, 10).join('\n')}

Общие характеристики товаров:
${currentState.commonCharacteristics.map(c => `- ${c.name} (${c.frequency} товаров)`).join('\n')}

Бренды: ${currentState.brands.join(', ')}

ЗАДАЧА:
Проанализируй категорию и выдай структурированный ответ в JSON формате:
{
  "analysis": {
    "strengths": ["список сильных сторон"],
    "weaknesses": ["список слабых мест"],
    "opportunities": ["возможности для улучшения"]
  },
  "recommendations": {
    "seo_title": "рекомендуемый SEO title",
    "seo_description": "рекомендуемое SEO description",
    "h1": "рекомендуемый H1",
    "description_structure": ["какие разделы должны быть в описании"],
    "keywords_focus": ["на какие ключевые слова стоит сфокусироваться"]
  }
}`;
    
    // Запрос к OpenAI API
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-4',
        messages: [
          {
            role: 'system',
            content: 'Ты - эксперт по SEO и оптимизации интернет-магазинов. Даешь конкретные, применимые рекомендации.'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.7,
        response_format: { type: 'json_object' }
      }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`OpenAI API ошибка: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    const analysisText = data.choices[0].message.content;
    const analysis = JSON.parse(analysisText);
    
    logInfo('✅ AI-анализ получен', null, context);
    
    return analysis;
    
  } catch (error) {
    logError('❌ Ошибка OpenAI анализа', error, context);
    throw error;
  }
}

/**
 * Распределяет ключевые слова по элементам страницы с помощью AI
 */
function distributeKeywordsWithAI(categoryId, categoryTitle, keywords) {
  const context = `AI распределение ключевиков для ${categoryTitle}`;
  
  try {
    logInfo('🤖 Начинаем AI-распределение ключевых слов', null, context);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'AI анализирует ключевые слова...',
      '⏳ Обработка',
      -1
    );
    
    const apiKey = getSetting('openaiApiKey');
    
    if (!apiKey) {
      throw new Error('OpenAI API ключ не настроен');
    }
    
    // Формируем промпт для распределения
    const keywordsText = keywords.slice(0, 50).map(kw => 
      `${kw.keyword} (частотность: ${kw.frequency})`
    ).join('\n');
    
    const prompt = `Распредели ключевые слова по элементам страницы категории интернет-магазина.

КАТЕГОРИЯ: ${categoryTitle}

КЛЮЧЕВЫЕ СЛОВА:
${keywordsText}

ЭЛЕМЕНТЫ СТРАНИЦЫ:
1. SEO Title (до 60 символов) - самые важные ключевики
2. Meta Description (до 160 символов) - расширенная семантика
3. H1 заголовок - главный запрос
4. Описание категории - LSI ключевики и естественные вхождения
5. Плитка тегов (фильтры) - дополнительные запросы
6. Фильтратор (характеристики) - технические параметры

ЗАДАЧА:
Распредели ключевые слова логично по этим элементам. Выдай ответ в JSON:
{
  "distribution": {
    "seo_title": ["ключевик 1", "ключевик 2"],
    "meta_description": ["ключевик 1", "ключевик 2", "ключевик 3"],
    "h1": "главный ключевик",
    "description": ["ключевик для текста 1", "ключевик для текста 2"],
    "tag_tiles": ["тег 1", "тег 2", "тег 3"],
    "filters": ["параметр 1", "параметр 2"]
  },
  "reasoning": "краткое объяснение логики распределения"
}`;
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        model: 'gpt-4',
        messages: [
          {
            role: 'system',
            content: 'Ты - эксперт по SEO и распределению семантического ядра. Понимаешь специфику e-commerce.'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.5,
        response_format: { type: 'json_object' }
      }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`OpenAI API ошибка: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    const distributionText = data.choices[0].message.content;
    const distribution = JSON.parse(distributionText);
    
    // Сохраняем распределение в листе ключевых слов
    updateKeywordsWithDistribution(categoryId, distribution.distribution);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ключевые слова распределены по элементам!',
      '✅ Готово',
      5
    );
    
    logInfo('✅ AI-распределение завершено', null, context);
    
    return distribution;
    
  } catch (error) {
    logError('❌ Ошибка AI-распределения', error, context);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка: ' + error.message,
      '❌ Ошибка',
      10
    );
    throw error;
  }
}

/**
 * Обновляет лист ключевых слов с информацией о назначении
 */
function updateKeywordsWithDistribution(categoryId, distribution) {
  const context = "Обновление назначения ключевиков";
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.KEYWORDS);
    
    if (!sheet) {
      throw new Error('Лист ключевых слов не найден');
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Создаем map ключевик -> назначение
    const assignmentMap = {};
    
    Object.entries(distribution).forEach(([element, keywords]) => {
      if (Array.isArray(keywords)) {
        keywords.forEach(kw => {
          assignmentMap[kw.toLowerCase()] = element;
        });
      } else if (typeof keywords === 'string') {
        assignmentMap[keywords.toLowerCase()] = element;
      }
    });
    
    // Обновляем назначение в листе
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCategoryId = row[0];
      const keyword = row[2];
      
      if (rowCategoryId == categoryId && keyword) {
        const assignment = assignmentMap[keyword.toLowerCase()];
        
        if (assignment) {
          // Обновляем колонку "Назначение" (индекс 6)
          sheet.getRange(i + 1, 7).setValue(assignment);
          // Обновляем статус (индекс 7)
          sheet.getRange(i + 1, 8).setValue('Назначено');
        }
      }
    }
    
    logInfo(`✅ Обновлены назначения ключевых слов`, null, context);
    
  } catch (error) {
    logError('❌ Ошибка обновления назначений', error, context);
    throw error;
  }
}

// ========================================
// ДОБАВЛЕНИЕ ФУНКЦИЙ В МЕНЮ
// ========================================

/**
 * Расширенное меню категорий с AI функциями
 */
function addExtendedCategoryMenu(mainMenu) {
  const categoryMenu = SpreadsheetApp.getUi().createMenu('📁 Категории');
  
  categoryMenu
    .addItem('🏗️ Создать структуру листов', 'createCategoryManagementStructure')
    .addSeparator()
    .addItem('📥 Загрузить категории с иерархией', 'loadCategoriesWithHierarchy')
    .addItem('🔍 Найти и открыть категорию', 'showCategorySearchDialog')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔑 Семантическое ядро')
      .addItem('⚙️ Настроить API семантики', 'configureSemanticsAPI')
      .addItem('📊 Собрать ключевые слова', 'collectKeywordsForActiveCategory')
      .addItem('🤖 Распределить ключевики (AI)', 'distributeKeywordsForActiveCategory')
      .addItem('📋 Показать все ключевики', 'showKeywordsSheet'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🤖 AI Генерация')
      .addItem('📝 Анализ текущего состояния', 'analyzeActiveCategory')
      .addItem('🎯 Генерировать SEO теги', 'generateSEOForActiveCategory')
      .addItem('📄 Генерировать описание', 'generateDescriptionForActiveCategory')
      .addItem('🏷️ Создать плитку тегов', 'generateTagTilesForActiveCategory'))
    .addSeparator()
    .addItem('🔄 Обновить данные категорий', 'updateCategoriesData')
    .addItem('📊 Показать статистику', 'showCategoriesStatistics');
  
  mainMenu.addSubMenu(categoryMenu);
  
  logInfo('✅ Расширенное меню категорий добавлено');
}

// ========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ РАБОТЫ С ДАННЫМИ
// ========================================

/**
 * Получает данные активной категории из детального листа
 */
function getActiveCategoryData() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Проверяем, что это детальный лист категории
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      return null;
    }
    
    // Извлекаем данные из листа
    const categoryId = sheet.getRange(DETAIL_SHEET_SECTIONS.CATEGORY_ID_CELL).getValue();
    const title = sheet.getRange(DETAIL_SHEET_SECTIONS.TITLE_CELL).getValue();
    const url = sheet.getRange(DETAIL_SHEET_SECTIONS.URL_CELL).getValue();
    const path = sheet.getRange(DETAIL_SHEET_SECTIONS.PARENT_PATH_CELL).getValue();
    
    return {
      id: categoryId,
      title: title,
      url: url,
      path: path
    };
    
  } catch (error) {
    logError('❌ Ошибка получения данных активной категории', error);
    return null;
  }
}

/**
 * Получает товары из детального листа
 */
function getProductsFromDetailSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    const products = [];
    const startRow = DETAIL_SHEET_SECTIONS.CURRENT_PRODUCTS_START;
    
    // Пропускаем заголовки и собираем товары
    for (let i = startRow; i < data.length; i++) {
      const row = data[i];
      
      // Прерываем, если дошли до пустой строки или следующего блока
      if (!row[1]) break; // Проверяем ID товара
      
      products.push({
        id: row[1],
        sku: row[2],
        title: row[3],
        in_stock: row[4] === 'Да',
        price: row[5],
        brand_title: row[6],
        characteristics_text: row[7]
      });
    }
    
    return products;
    
  } catch (error) {
    logError('❌ Ошибка получения товаров из листа', error);
    return [];
  }
}

/**
 * Получает ключевые слова для категории
 */
function getKeywordsForCategory(categoryId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.KEYWORDS);
    
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const keywords = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[0] == categoryId) {
        keywords.push({
          keyword: row[2],
          frequency: row[3],
          competition: row[4],
          type: row[5],
          assignment: row[6],
          status: row[7]
        });
      }
    }
    
    return keywords;
    
  } catch (error) {
    logError('❌ Ошибка получения ключевых слов', error);
    return [];
  }
}

/**
 * Получает все категории из главного листа
 */
function getAllCategoriesFromMainList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const categories = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      categories.push({
        id: row[MAIN_LIST_COLUMNS.CATEGORY_ID - 1],
        title: row[MAIN_LIST_COLUMNS.TITLE - 1].replace(/[\s└─]/g, '').trim(),
        url: row[MAIN_LIST_COLUMNS.URL - 1],
        path: row[MAIN_LIST_COLUMNS.HIERARCHY_PATH - 1],
        level: row[MAIN_LIST_COLUMNS.LEVEL - 1]
      });
    }
    
    return categories;
    
  } catch (error) {
    logError('❌ Ошибка получения списка категорий', error);
    return [];
  }
}  

/**
 * AI-анализ активной категории
 */
function analyzeActiveCategory() {
  const context = "AI анализ категории";
  
  try {
    logInfo('🔍 Начинаем AI-анализ активной категории', null, context);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Собираем данные категории...',
      '🔍 Анализ',
      3
    );
    
    // 1. Получаем активный лист
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    
    // 2. Проверяем, что это детальный лист категории
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      throw new Error('Откройте детальный лист категории для анализа');
    }
    
    // 3. Получаем данные категории (БЕЗОПАСНО)
    const categoryTitle = String(sheet.getRange('B3').getValue() || 'Бинокли').trim();
    
    console.log('📊 Категория:', categoryTitle);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Отправляем запрос к AI...',
      '🤖 OpenAI',
      -1
    );
    
    // 4. ПРОСТОЙ промпт без лишних данных
    const simplePrompt = `Проанализируй SEO категории ${categoryTitle} в интернет-магазине. Дай 3 совета.`;
    
    console.log('📝 Промпт:', simplePrompt);
    
    // 5. Получаем конфиг
    const config = getOpenAIConfig();
    const url = 'https://api.openai.com/v1/chat/completions';
    
    // 6. Формируем запрос
    const payload = {
      model: 'gpt-4o-mini',
      messages: [
        { role: 'user', content: simplePrompt }
      ],
      max_tokens: 200
    };
    
    console.log('📦 Отправляем запрос...');
    
    // 7. Отправляем
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('📥 HTTP код:', responseCode);
    
    if (responseCode !== 200) {
      console.error('❌ Ошибка:', responseText);
      throw new Error(`OpenAI API ошибка ${responseCode}: ${responseText}`);
    }
    
    const result = JSON.parse(responseText);
    const analysis = result.choices[0].message.content.trim();
    
    console.log('✅ Анализ получен:', analysis.substring(0, 100) + '...');
    
    logInfo('✅ Анализ получен от OpenAI', null, context);
    
    // 8. Показываем результат
    const ui = SpreadsheetApp.getUi();
    
    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body { 
              font-family: Arial, sans-serif; 
              padding: 20px;
              line-height: 1.6;
            }
            h3 { 
              color: #1976d2; 
              margin-top: 0;
            }
            .info {
              background: #f5f5f5;
              padding: 15px;
              border-radius: 4px;
              margin-bottom: 20px;
            }
            .analysis {
              background: white;
              padding: 15px;
              border-left: 4px solid #4caf50;
              white-space: pre-wrap;
              line-height: 1.8;
            }
            button {
              background: #1976d2;
              color: white;
              border: none;
              padding: 10px 20px;
              border-radius: 4px;
              cursor: pointer;
              margin-top: 20px;
            }
            button:hover {
              background: #1565c0;
            }
          </style>
        </head>
        <body>
          <h3>🤖 AI-анализ категории</h3>
          
          <div class="info">
            <strong>Категория:</strong> ${categoryTitle}
          </div>
          
          <div class="analysis">${analysis}</div>
          
          <button onclick="google.script.host.close()">Закрыть</button>
        </body>
      </html>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(600)
      .setHeight(500);
    
    ui.showModalDialog(htmlOutput, 'AI-анализ категории');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '✅ Анализ завершен',
      'Готово',
      5
    );
    
    return {
      success: true,
      analysis: analysis
    };
    
  } catch (error) {
    console.error('❌ ОШИБКА:', error.message);
    
    logError('❌ Ошибка AI-анализа', error, context);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Ошибка: ${error.message}`,
      '❌ Ошибка',
      10
    );
    
    throw error;
  }
}

function testAnalysisMinimal() {
  try {
    console.log('🧪 Минимальный тест анализа');
    
    const config = getOpenAIConfig();
    const url = 'https://api.openai.com/v1/chat/completions';
    
    // МАКСИМАЛЬНО ПРОСТОЙ промпт
    const simplePrompt = 'Проанализируй SEO категории Бинокли в интернет-магазине. Дай 3 совета.';
    
    const payload = {
      model: 'gpt-4o-mini',
      messages: [
        { role: 'user', content: simplePrompt }
      ],
      max_tokens: 200
    };
    
    console.log('Отправляем:', JSON.stringify(payload, null, 2));
    
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const text = response.getContentText();
    
    console.log('HTTP код:', code);
    console.log('Ответ:', text);
    
    if (code === 200) {
      const result = JSON.parse(text);
      console.log('✅ РАБОТАЕТ! Ответ:', result.choices[0].message.content);
      SpreadsheetApp.getActiveSpreadsheet().toast('✅ Работает!', 'Успех', 5);
    } else {
      console.error('❌ Ошибка:', text);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Ошибка ${code}`, 'Ошибка', 10);
    }
    
  } catch (error) {
    console.error('❌ Критическая ошибка:', error.message, error.stack);
  }
}

/**
 * УПРОЩЕННАЯ ТЕСТОВАЯ ВЕРСИЯ
 * Точная копия testAnalysisMinimal, но для реальной категории
 */
function testSimpleAnalysisForCategory() {
  try {
    console.log('🧪 Упрощенный тест анализа категории');
    
    // 1. Получаем активный лист
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();
    
    // 2. Проверяем, что это детальный лист
    if (!sheetName.startsWith('Категория — ')) {
      SpreadsheetApp.getUi().alert('Откройте детальный лист категории');
      return;
    }
    
    // 3. Читаем название категории
    let categoryTitle = 'Категория';
    try {
      const titleValue = sheet.getRange('B3').getValue();
      if (titleValue) {
        categoryTitle = String(titleValue).trim();
      }
    } catch (e) {
      console.warn('Не удалось прочитать название, используем дефолт');
    }
    
    console.log('📊 Категория:', categoryTitle);
    
    // 4. Получаем конфиг
    const config = getOpenAIConfig();
    const url = 'https://api.openai.com/v1/chat/completions';
    
    // 5. МАКСИМАЛЬНО ПРОСТОЙ промпт (как в тесте)
    const simplePrompt = `Проанализируй SEO категории "${categoryTitle}" в интернет-магазине. Дай 3 совета.`;
    
    // 6. Точно такой же payload как в тесте
    const payload = {
      model: 'gpt-4o-mini',
      messages: [
        { role: 'user', content: simplePrompt }
      ],
      max_tokens: 200
    };
    
    console.log('📤 Отправляем:', JSON.stringify(payload, null, 2));
    
    // 7. Отправка (точно как в тесте)
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${config.apiKey}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const text = response.getContentText();
    
    console.log('📥 HTTP код:', code);
    console.log('📥 Ответ:', text.substring(0, 200));
    
    if (code === 200) {
      const result = JSON.parse(text);
      const analysis = result.choices[0].message.content;
      
      console.log('✅ РАБОТАЕТ! Анализ:', analysis);
      
      // Показываем результат
      const ui = SpreadsheetApp.getUi();
      ui.alert('✅ Анализ получен', analysis, ui.ButtonSet.OK);
      
      SpreadsheetApp.getActiveSpreadsheet().toast('✅ Работает!', 'Успех', 5);
    } else {
      console.error('❌ Ошибка:', code, text);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Ошибка ${code}`, 'Ошибка', 10);
    }
    
  } catch (error) {
    console.error('❌ Критическая ошибка:', error.message, error.stack);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}