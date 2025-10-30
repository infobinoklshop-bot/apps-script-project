/**
 * ========================================
 * МОДУЛЬ: ГЕНЕРАЦИЯ АНКОРОВ ДЛЯ ПЛИТОК ТЕГОВ
 * ========================================
 *
 * Генерирует анкоры для верхней и нижней плитки тегов
 * используя AI (Gemini)
 */

/**
 * Получает дочерние категории для текущей категории
 * @param {number} categoryId - ID родительской категории
 * @returns {Array<Object>} Массив дочерних категорий
 */
function getChildCategories(categoryId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!sheet) {
      throw new Error('Лист категорий не найден');
    }

    const data = sheet.getDataRange().getValues();
    const children = [];

    // Пропускаем заголовок
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const parentId = row[MAIN_LIST_COLUMNS.PARENT_ID - 1];

      if (parentId == categoryId) {
        children.push({
          id: row[MAIN_LIST_COLUMNS.CATEGORY_ID - 1],
          title: String(row[MAIN_LIST_COLUMNS.TITLE - 1]).replace(/[\s└─]/g, '').trim(),
          url: row[MAIN_LIST_COLUMNS.URL - 1],
          level: row[MAIN_LIST_COLUMNS.LEVEL - 1],
          path: row[MAIN_LIST_COLUMNS.HIERARCHY_PATH - 1]
        });
      }
    }

    console.log('[INFO] Найдено дочерних категорий:', children.length);

    return children;

  } catch (error) {
    console.error('[ERROR] Ошибка получения дочерних категорий:', error.message);
    return [];
  }
}

/**
 * Получает все категории магазина для маппинга
 * @returns {Array<Object>} Массив всех категорий
 */
function getAllCategoriesForMapping() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!sheet) {
      throw new Error('Лист категорий не найден');
    }

    const data = sheet.getDataRange().getValues();
    const categories = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      categories.push({
        id: row[MAIN_LIST_COLUMNS.CATEGORY_ID - 1],
        title: String(row[MAIN_LIST_COLUMNS.TITLE - 1]).replace(/[\s└─]/g, '').trim(),
        url: row[MAIN_LIST_COLUMNS.URL - 1],
        path: row[MAIN_LIST_COLUMNS.HIERARCHY_PATH - 1]
      });
    }

    return categories;

  } catch (error) {
    console.error('[ERROR] Ошибка получения списка категорий:', error.message);
    return [];
  }
}

/**
 * Генерирует анкоры для ВЕРХНЕЙ плитки (навигация по дочерним категориям)
 * @param {number} categoryId - ID категории
 * @param {Object} densityAnalysis - Результаты анализа плотности
 * @returns {Array<Object>} Массив анкоров
 */
function generateTopTileAnchors(categoryId, densityAnalysis) {
  const context = `Генерация анкоров верхней плитки для ${categoryId}`;

  try {
    console.log('[INFO] 🎨 Генерируем анкоры для верхней плитки (навигация)');

    // 1. Получаем дочерние категории
    const children = getChildCategories(categoryId);

    if (children.length === 0) {
      console.log('[INFO] ⚠️ У категории нет дочерних категорий, верхняя плитка не нужна');
      return [];
    }

    console.log('[INFO] Дочерних категорий:', children.length);

    // 2. Получаем ключевые слова с недостатком
    const deficitKeywords = densityAnalysis.results
      .filter(r => r.status === '❌ Недостаточно')
      .map(r => r.keyword)
      .slice(0, 10); // Берем топ-10

    // 3. Подготавливаем данные для AI
    const categoryData = {
      id: densityAnalysis.categoryId,
      name: densityAnalysis.categoryName,
      childCategories: children
    };

    // 4. Генерируем промпт для AI
    const prompt = buildTopTilePrompt(categoryData, deficitKeywords);

    // 5. Вызываем AI API
    console.log('[INFO] Отправляем запрос в AI...');
    const aiResponse = callAIForGeneration(prompt);

    // 6. Парсим ответ
    const anchors = parseAIResponse(aiResponse);

    // 7. Ограничиваем количество
    const limit = TAG_TILES_CONFIG.TOP_TILE_LIMIT.max;
    const limitedAnchors = anchors.slice(0, limit);

    console.log('[INFO] ✅ Сгенерировано анкоров для верхней плитки:', limitedAnchors.length);

    return limitedAnchors;

  } catch (error) {
    console.error('[ERROR] Ошибка генерации анкоров верхней плитки:', error.message);
    throw error;
  }
}

/**
 * Генерирует анкоры для НИЖНЕЙ плитки (SEO, ключевые слова)
 * @param {number} categoryId - ID категории
 * @param {Object} densityAnalysis - Результаты анализа плотности
 * @returns {Array<Object>} Массив анкоров
 */
function generateBottomTileAnchors(categoryId, densityAnalysis) {
  const context = `Генерация анкоров нижней плитки для ${categoryId}`;

  try {
    console.log('[INFO] 🎨 Генерируем анкоры для нижней плитки (SEO)');

    // 1. Получаем ключевые слова с недостатком
    const deficitKeywords = densityAnalysis.results
      .filter(r => r.status === '❌ Недостаточно')
      .sort((a, b) => a.tfPercent - b.tfPercent); // Сортируем по возрастанию плотности

    console.log('[INFO] Ключевых слов с недостатком:', deficitKeywords.length);

    if (deficitKeywords.length === 0) {
      console.log('[INFO] ✅ Все ключевые слова имеют достаточную плотность');
      return [];
    }

    // 2. Получаем все категории для маппинга
    const allCategories = getAllCategoriesForMapping();

    // 3. Получаем фрагмент текущего контента
    const contentFragment = densityAnalysis.contentData ?
      (densityAnalysis.contentData.description || '').substring(0, 500) : '';

    // 4. Подготавливаем данные для AI
    const categoryData = {
      id: densityAnalysis.categoryId,
      name: densityAnalysis.categoryName,
      path: densityAnalysis.contentData.path || 'Корневая',
      deficitKeywords: deficitKeywords.slice(0, 20), // Топ-20
      currentContent: contentFragment,
      availableCategories: allCategories.slice(0, 50) // Топ-50 для контекста
    };

    // 5. Генерируем промпт для AI
    const prompt = buildBottomTilePrompt(categoryData);

    // 6. Вызываем AI API
    console.log('[INFO] Отправляем запрос в AI...');
    const aiResponse = callAIForGeneration(prompt);

    // 7. Парсим ответ
    const anchors = parseAIResponse(aiResponse);

    // 8. Ограничиваем количество
    const limit = TAG_TILES_CONFIG.BOTTOM_TILE_LIMIT.max;
    const limitedAnchors = anchors.slice(0, limit);

    console.log('[INFO] ✅ Сгенерировано анкоров для нижней плитки:', limitedAnchors.length);

    return limitedAnchors;

  } catch (error) {
    console.error('[ERROR] Ошибка генерации анкоров нижней плитки:', error.message);
    throw error;
  }
}

/**
 * Строит промпт для генерации анкоров верхней плитки
 * @param {Object} categoryData - Данные категории
 * @param {Array<string>} deficitKeywords - Ключевые слова с недостатком
 * @returns {string} Промпт для AI
 */
function buildTopTilePrompt(categoryData, deficitKeywords) {
  const childList = categoryData.childCategories
    .map(c => `- ${c.title} (ID: ${c.id}, URL: ${c.url})`)
    .join('\n');

  const keywordsList = deficitKeywords.length > 0 ?
    deficitKeywords.join(', ') : 'нет данных';

  return `Ты - эксперт по навигации интернет-магазинов и SEO.

ЗАДАЧА: Создать 5-8 естественных анкоров для ВЕРХНЕЙ плитки навигации.

КОНТЕКСТ:
Категория: "${categoryData.name}"
ID: ${categoryData.id}

ДОЧЕРНИЕ ПОДКАТЕГОРИИ:
${childList}

КЛЮЧЕВЫЕ СЛОВА С НЕДОСТАТОЧНОЙ ПЛОТНОСТЬЮ:
${keywordsList}

ТРЕБОВАНИЯ:
1. Анкор должен быть понятным для пользователя (2-4 слова)
2. Использовать ключевые слова естественно, где это уместно
3. Каждый анкор - ссылка на ОДНУ из дочерних подкатегорий
4. Приоритет: удобство навигации > SEO
5. Не более 8 анкоров (лучше 5-6)
6. Не изобретай категории - используй ТОЛЬКО те, что в списке

ФОРМАТ ОТВЕТА (строго JSON):
\`\`\`json
[
  {
    "anchor": "Текст анкора",
    "category_id": 12345,
    "category_name": "Название категории",
    "keywords_used": ["ключевик1", "ключевик2"],
    "reasoning": "Краткое объяснение почему этот анкор хорош"
  }
]
\`\`\`

Верни ТОЛЬКО валидный JSON массив без дополнительного текста.`;
}

/**
 * Строит промпт для генерации анкоров нижней плитки
 * @param {Object} categoryData - Данные категории
 * @returns {string} Промпт для AI
 */
function buildBottomTilePrompt(categoryData) {
  const deficitList = categoryData.deficitKeywords
    .map(k => `- "${k.keyword}" (TF: ${k.tfPercent}%, цель: ${k.targetMin}-${k.targetMax}%)`)
    .join('\n');

  const categoriesList = categoryData.availableCategories
    .map(c => `- ${c.title} (ID: ${c.id}, URL: ${c.url})`)
    .join('\n');

  return `Ты - эксперт по SEO-оптимизации для интернет-магазинов.

ЗАДАЧА: Создать 15-30 естественных анкоров для НИЖНЕЙ плитки тегов (SEO-оптимизация).

КОНТЕКСТ:
Категория: "${categoryData.name}"
ID: ${categoryData.id}
Путь: ${categoryData.path}

ТЕКУЩИЙ КОНТЕНТ (фрагмент):
"${categoryData.currentContent}"

КЛЮЧЕВЫЕ СЛОВА С НЕДОСТАТОЧНОЙ ПЛОТНОСТЬЮ:
${deficitList}

СУЩЕСТВУЮЩИЕ КАТЕГОРИИ МАГАЗИНА (примеры):
${categoriesList}
... и другие

ТРЕБОВАНИЯ:
1. Анкор должен включать ключевые слова с недостатком
2. Звучать естественно (2-5 слов, не спам)
3. Быть полезным для пользователя
4. Привязать к существующей категории если она подходит
5. Если подходящей категории нет - пометить как "CREATE_NEW"
6. Создать 15-30 анкоров (приоритет самым "дефицитным" ключевикам)
7. Не повторять одинаковые анкоры

ФОРМАТ ОТВЕТА (строго JSON):
\`\`\`json
[
  {
    "anchor": "Текст анкора (2-5 слов)",
    "keywords_covered": ["ключевик1", "ключевик2"],
    "target_category": {
      "id": 12345,
      "name": "Название категории",
      "url": "/collection/url",
      "action": "LINK"
    },
    "reasoning": "Краткое объяснение"
  },
  {
    "anchor": "Другой анкор",
    "keywords_covered": ["ключевик3"],
    "target_category": {
      "id": null,
      "name": "Название новой категории",
      "url": "/collection/suggested-url",
      "action": "CREATE_NEW"
    },
    "reasoning": "Почему нужна новая категория"
  }
]
\`\`\`

Верни ТОЛЬКО валидный JSON массив без дополнительного текста.`;
}

/**
 * Вызывает AI API для генерации анкоров (поддерживает OpenAI, Claude и Gemini)
 * @param {string} prompt - Промпт для AI
 * @returns {string} Ответ от AI
 */
function callAIForGeneration(prompt) {
  const provider = TAG_TILES_CONFIG.AI_PROVIDER || 'openai';

  if (provider === 'openai') {
    return callOpenAIForGeneration(prompt);
  } else if (provider === 'claude') {
    return callClaudeAPI(prompt);
  } else if (provider === 'gemini') {
    return callGeminiAPI(prompt);
  } else {
    throw new Error(`Неизвестный AI провайдер: ${provider}`);
  }
}

/**
 * Вызывает OpenAI API для генерации анкоров
 * Использует настроенный в проекте getOpenAIConfig()
 * @param {string} prompt - Промпт для AI
 * @returns {string} Ответ от AI
 */
function callOpenAIForGeneration(prompt) {
  try {
    // Используем уже настроенную функцию из проекта
    const config = getOpenAIConfig();
    const apiKey = config.apiKey;

    if (!apiKey || apiKey.includes('YOUR_')) {
      throw new Error('OpenAI API ключ не настроен. Добавьте OPENAI_API_KEY в Script Properties.');
    }

    const url = 'https://api.openai.com/v1/chat/completions';

    const payload = {
      model: config.model || 'gpt-4o-mini',
      messages: [{
        role: 'system',
        content: 'Ты - эксперт по навигации интернет-магазинов и SEO. Отвечай ТОЛЬКО валидным JSON без дополнительного текста.'
      }, {
        role: 'user',
        content: prompt
      }],
      temperature: TAG_TILES_CONFIG.AI_TEMPERATURE,
      max_tokens: TAG_TILES_CONFIG.AI_MAX_TOKENS
    };

    const options = {
      method: 'POST',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${apiKey}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    console.log('[INFO] Отправляем запрос в OpenAI API...');

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      const errorBody = response.getContentText();
      console.error('[ERROR] OpenAI API ответ:', errorBody);
      throw new Error(`OpenAI API error: ${statusCode} - ${errorBody}`);
    }

    const result = JSON.parse(response.getContentText());

    if (!result.choices || !result.choices[0] || !result.choices[0].message) {
      throw new Error('Неожиданный формат ответа от OpenAI API');
    }

    const text = result.choices[0].message.content;

    console.log('[INFO] ✅ Получен ответ от OpenAI');

    return text;

  } catch (error) {
    console.error('[ERROR] Ошибка вызова OpenAI API:', error.message);
    throw error;
  }
}

/**
 * Вызывает Claude API для генерации анкоров
 * @param {string} prompt - Промпт для AI
 * @returns {string} Ответ от AI
 */
function callClaudeAPI(prompt) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

    if (!apiKey) {
      throw new Error('CLAUDE_API_KEY не настроен в Script Properties');
    }

    const url = 'https://api.anthropic.com/v1/messages';

    const payload = {
      model: TAG_TILES_CONFIG.AI_MODEL,
      max_tokens: TAG_TILES_CONFIG.AI_MAX_TOKENS,
      temperature: TAG_TILES_CONFIG.AI_TEMPERATURE,
      messages: [{
        role: 'user',
        content: prompt
      }]
    };

    const options = {
      method: 'POST',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    console.log('[INFO] Отправляем запрос в Claude API...');

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      throw new Error(`Claude API error: ${statusCode} - ${response.getContentText()}`);
    }

    const result = JSON.parse(response.getContentText());

    if (!result.content || !result.content[0] || !result.content[0].text) {
      throw new Error('Неожиданный формат ответа от Claude API');
    }

    const text = result.content[0].text;

    console.log('[INFO] ✅ Получен ответ от Claude');

    return text;

  } catch (error) {
    console.error('[ERROR] Ошибка вызова Claude API:', error.message);
    throw error;
  }
}

/**
 * Вызывает Gemini API для генерации анкоров
 * @param {string} prompt - Промпт для AI
 * @returns {string} Ответ от AI
 */
function callGeminiAPI(prompt) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

    if (!apiKey) {
      throw new Error('GEMINI_API_KEY не настроен в Script Properties');
    }

    const url = `https://generativelanguage.googleapis.com/v1/models/${TAG_TILES_CONFIG.AI_MODEL}:generateContent?key=${apiKey}`;

    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: TAG_TILES_CONFIG.AI_TEMPERATURE,
        maxOutputTokens: TAG_TILES_CONFIG.AI_MAX_TOKENS,
        topP: 0.95,
        topK: 40
      }
    };

    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode !== 200) {
      throw new Error(`Gemini API error: ${statusCode} - ${response.getContentText()}`);
    }

    const result = JSON.parse(response.getContentText());

    if (!result.candidates || !result.candidates[0] || !result.candidates[0].content) {
      throw new Error('Неожиданный формат ответа от Gemini API');
    }

    const text = result.candidates[0].content.parts[0].text;

    console.log('[INFO] ✅ Получен ответ от Gemini');

    return text;

  } catch (error) {
    console.error('[ERROR] Ошибка вызова Gemini API:', error.message);
    throw error;
  }
}

/**
 * Парсит ответ от AI и извлекает JSON
 * @param {string} aiResponse - Текстовый ответ от AI
 * @returns {Array<Object>} Массив анкоров
 */
function parseAIResponse(aiResponse) {
  try {
    // Убираем markdown форматирование если есть
    let jsonText = aiResponse.trim();

    // Извлекаем JSON из markdown блока
    const jsonMatch = jsonText.match(/```json\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      jsonText = jsonMatch[1];
    }

    // Убираем возможные префиксы
    jsonText = jsonText.replace(/^[^[\{]*/, '');
    jsonText = jsonText.replace(/[^\]\}]*$/, '');

    // Парсим JSON
    const anchors = JSON.parse(jsonText);

    if (!Array.isArray(anchors)) {
      throw new Error('Ответ AI не является массивом');
    }

    console.log('[INFO] ✅ Распарсено анкоров:', anchors.length);

    return anchors;

  } catch (error) {
    console.error('[ERROR] Ошибка парсинга ответа AI:', error.message);
    console.error('[ERROR] Ответ AI:', aiResponse);
    throw new Error('Не удалось распарсить ответ AI: ' + error.message);
  }
}

/**
 * ГЛАВНАЯ ФУНКЦИЯ: Генерирует анкоры для обеих плиток
 * @param {number} categoryId - ID категории
 * @returns {Object} Объект с анкорами для верхней и нижней плитки
 */
function generateTileAnchors(categoryId) {
  try {
    console.log('[INFO] 🎨 Начинаем генерацию анкоров для категории:', categoryId);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Генерируем анкоры для плиток тегов...',
      '🎨 AI генерация',
      -1
    );

    // 1. Сначала анализируем плотность
    const densityAnalysis = analyzeKeywordDensity(categoryId, getKeywordsForCategory(categoryId));

    // 2. Генерируем анкоры для верхней плитки
    console.log('[INFO] === ВЕРХНЯЯ ПЛИТКА (Навигация) ===');
    const topAnchors = generateTopTileAnchors(categoryId, densityAnalysis);

    // 3. Генерируем анкоры для нижней плитки
    console.log('[INFO] === НИЖНЯЯ ПЛИТКА (SEO) ===');
    const bottomAnchors = generateBottomTileAnchors(categoryId, densityAnalysis);

    // 4. Формируем результат
    const result = {
      categoryId: categoryId,
      categoryName: densityAnalysis.categoryName,
      topTile: {
        anchors: topAnchors,
        count: topAnchors.length
      },
      bottomTile: {
        anchors: bottomAnchors,
        count: bottomAnchors.length
      },
      densityAnalysis: densityAnalysis,
      generatedAt: new Date()
    };

    console.log('[INFO] ✅ Генерация завершена');
    console.log('[INFO] Верхняя плитка:', result.topTile.count, 'анкоров');
    console.log('[INFO] Нижняя плитка:', result.bottomTile.count, 'анкоров');

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Сгенерировано:
Верхняя плитка: ${result.topTile.count} анкоров
Нижняя плитка: ${result.bottomTile.count} анкоров`,
      '✅ Готово',
      10
    );

    return result;

  } catch (error) {
    console.error('[ERROR] Ошибка генерации анкоров:', error.message);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка: ' + error.message,
      '❌ Ошибка',
      10
    );

    throw error;
  }
}

/**
 * ТЕСТОВАЯ ФУНКЦИЯ: Генерация анкоров для категории Бинокли
 */
function testGenerateAnchors() {
  const categoryId = 9175197; // Бинокли

  const result = generateTileAnchors(categoryId);

  console.log('=== РЕЗУЛЬТАТЫ ГЕНЕРАЦИИ ===');
  console.log('Категория:', result.categoryName);
  console.log('\n=== ВЕРХНЯЯ ПЛИТКА ===');
  result.topTile.anchors.forEach((a, i) => {
    console.log(`${i+1}. "${a.anchor}" → ${a.category_name} (${a.category_id})`);
  });
  console.log('\n=== НИЖНЯЯ ПЛИТКА ===');
  result.bottomTile.anchors.forEach((a, i) => {
    console.log(`${i+1}. "${a.anchor}" → ${a.target_category.name} (${a.target_category.action})`);
  });

  return result;
}
