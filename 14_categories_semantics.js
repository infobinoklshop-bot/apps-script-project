/**
 * ========================================
 * МОДУЛЬ: СЕМАНТИКА + LSI + ТЕМАТИКА
 * ========================================
 */

// СКОПИРУЙТЕ СЮДА ИЗ ОСНОВНОГО КОДА:
// - configureSemanticsAPI()
// - showManualSemanticsAPISetup()
// - saveSemanticsAPISettings()
// - collectKeywordsForCategory()
// - generateBaseQueries()
// - fetchKeywordsFromAPI()
// - fetchKeywordsFromSerpstat()
// - deduplicateKeywords()
// - saveKeywordsToSheet()

// ========================================
// НОВЫЙ ФУНКЦИОНАЛ: LSI И ТЕМАТИКА
// ========================================

/**
 * Создает лист LSI и тематических слов
 */
function setupLSISheet(sheet) {
  sheet.clear();
  
  const headers = [
    'Категория ID',
    'Категория',
    'Слово/фраза',
    'Тип',  // LSI, Тематика, Синоним
    'Источник',  // OpenAI, API, Ручной
    'Релевантность',
    'Использовано',
    'Где использовано'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 200);
  
  sheet.setFrozenRows(1);
  
  logInfo('✅ Лист LSI и тематики настроен');
}

/**
 * Собирает LSI и тематические слова через AI
 */
function collectLSIAndThematicWords(categoryData, keywords) {
  const context = `Сбор LSI для ${categoryData.title}`;
  
  try {
    logInfo('🔍 Начинаем сбор LSI и тематических слов через AI', null, context);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'AI собирает LSI и тематику...',
      '⏳ Анализ',
      -1
    );
    
    const apiKey = getSetting('openaiApiKey');
    
    if (!apiKey) {
      throw new Error('OpenAI API ключ не настроен');
    }
    
    // Формируем список основных ключевиков
    const mainKeywords = keywords
      .slice(0, 20)
      .map(kw => kw.keyword)
      .join(', ');
    
    const prompt = `Ты - эксперт по SEO и семантическому анализу текстов.

ЗАДАЧА: Подобрать LSI-слова и тематические термины для категории интернет-магазина.

КАТЕГОРИЯ: ${categoryData.title}
ПУТЬ: ${categoryData.path || 'Корневая'}

ОСНОВНЫЕ КЛЮЧЕВЫЕ СЛОВА:
${mainKeywords}

ТРЕБУЕТСЯ:

1. **LSI-слова** (Latent Semantic Indexing):
   - Слова и фразы, семантически связанные с основными ключевиками
   - Слова, которые обычно встречаются в текстах на эту тему
   - Синонимы и вариации написания
   - Примеры: для "бинокль" → "оптика", "увеличение", "объектив", "призма"

2. **Тематические термины**:
   - Специфичные для ниши термины и жаргон
   - Названия характеристик и параметров
   - Технические термины
   - Названия типов и видов товаров
   - Примеры: для биноклей → "Porro-призмы", "Roof-призмы", "просветление", "выходной зрачок"

3. **Синонимы и вариации**:
   - Разговорные варианты терминов
   - Профессиональная терминология
   - Устаревшие названия
   - Иностранные термины в русской транскрипции

ФОРМАТ ОТВЕТА - JSON:
{
  "lsi_words": [
    {"word": "слово/фраза", "relevance": 0.95, "usage_example": "где можно использовать"},
    ...
  ],
  "thematic_terms": [
    {"term": "термин", "relevance": 0.90, "category": "характеристики/типы/бренды"},
    ...
  ],
  "synonyms": [
    {"synonym": "синоним", "for_word": "для какого слова", "relevance": 0.85},
    ...
  ]
}

Подбери минимум 30-50 релевантных слов и терминов.`;
    
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
            content: 'Ты - эксперт по SEO, семантическому анализу и LSI. Отлично знаешь специфику интернет-магазинов.'
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
    const resultText = data.choices[0].message.content;
    const result = JSON.parse(resultText);
    
    // Сохраняем в лист
    saveLSIAndThematicToSheet(categoryData.id, categoryData.title, result);
    
    const totalWords = 
      (result.lsi_words ? result.lsi_words.length : 0) +
      (result.thematic_terms ? result.thematic_terms.length : 0) +
      (result.synonyms ? result.synonyms.length : 0);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Собрано ${totalWords} LSI и тематических слов!`,
      '✅ Готово',
      5
    );
    
    logInfo(`✅ Собрано LSI и тематики: ${totalWords} слов`, null, context);
    
    return result;
    
  } catch (error) {
    logError('❌ Ошибка сбора LSI', error, context);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка: ' + error.message,
      '❌ Ошибка',
      10
    );
    throw error;
  }
}

/**
 * Сохраняет LSI и тематику в лист
 */
function saveLSIAndThematicToSheet(categoryId, categoryTitle, data) {
  const context = "Сохранение LSI и тематики";
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CATEGORY_SHEETS.LSI_WORDS);
    
    if (!sheet) {
      sheet = ss.insertSheet(CATEGORY_SHEETS.LSI_WORDS);
      setupLSISheet(sheet);
    }
    
    // Удаляем старые данные для категории
    const existingData = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    for (let i = existingData.length - 1; i >= 1; i--) {
      if (existingData[i][0] == categoryId) {
        rowsToDelete.push(i + 1);
      }
    }
    
    rowsToDelete.forEach(row => {
      sheet.deleteRow(row);
    });
    
    // Подготавливаем новые данные
    const rows = [];
    
    // LSI слова
    if (data.lsi_words && Array.isArray(data.lsi_words)) {
      data.lsi_words.forEach(item => {
        rows.push([
          categoryId,
          categoryTitle,
          item.word,
          'LSI',
          'OpenAI',
          item.relevance || 0,
          'Нет',
          item.usage_example || ''
        ]);
      });
    }
    
    // Тематические термины
    if (data.thematic_terms && Array.isArray(data.thematic_terms)) {
      data.thematic_terms.forEach(item => {
        rows.push([
          categoryId,
          categoryTitle,
          item.term,
          'Тематика (' + (item.category || 'общее') + ')',
          'OpenAI',
          item.relevance || 0,
          'Нет',
          ''
        ]);
      });
    }
    
    // Синонимы
    if (data.synonyms && Array.isArray(data.synonyms)) {
      data.synonyms.forEach(item => {
        rows.push([
          categoryId,
          categoryTitle,
          item.synonym,
          'Синоним: ' + (item.for_word || ''),
          'OpenAI',
          item.relevance || 0,
          'Нет',
          ''
        ]);
      });
    }
    
    // Записываем данные
    if (rows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rows.length, 8).setValues(rows);
      
      // Форматируем релевантность
      sheet.getRange(lastRow + 1, 6, rows.length, 1).setNumberFormat('0.00');
    }
    
    logInfo(`✅ Сохранено ${rows.length} LSI и тематических слов`, null, context);
    
  } catch (error) {
    logError('❌ Ошибка сохранения LSI', error, context);
    throw error;
  }
}

/**
 * Получает LSI и тематику для категории
 */
function getLSIForCategory(categoryId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.LSI_WORDS);
    
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const lsiWords = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == categoryId) {
        lsiWords.push({
          word: data[i][2],
          type: data[i][3],
          source: data[i][4],
          relevance: data[i][5],
          used: data[i][6],
          where_used: data[i][7]
        });
      }
    }
    
    return lsiWords;
    
  } catch (error) {
    logError('❌ Ошибка получения LSI', error);
    return [];
  }
}

/**
 * Функция для меню - собрать LSI для активной категории
 */
function collectLSIForActiveCategory() {
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
    
    const keywords = getKeywordsForCategory(categoryData.id);
    
    if (keywords.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Нет ключевых слов',
        'Сначала соберите семантическое ядро',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    collectLSIAndThematicWords(categoryData, keywords);
    
  } catch (error) {
    logError('❌ Ошибка сбора LSI', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}