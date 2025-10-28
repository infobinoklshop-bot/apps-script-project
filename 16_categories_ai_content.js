// ========================================
// AI ГЕНЕРАЦИЯ КОНТЕНТА
// ========================================

/**
 * Генерирует SEO теги через AI
 */
function generateSEOWithAI(categoryData, keywords) {
  const context = "AI генерация SEO";
  
  try {
    logInfo('🤖 Генерируем SEO теги через AI', null, context);
    
    const apiKey = getSetting('openaiApiKey');
    
    if (!apiKey) {
      throw new Error('OpenAI API ключ не настроен');
    }
    
    const keywordsText = keywords.map(kw => kw.keyword).join(', ');
    
    const prompt = `Создай SEO-оптимизированные теги для категории интернет-магазина.

КАТЕГОРИЯ: ${categoryData.title}
URL: ${categoryData.url}
ИЕРАРХИЯ: ${categoryData.path || 'Корневая категория'}

КЛЮЧЕВЫЕ СЛОВА:
${keywordsText}

ТРЕБОВАНИЯ:
1. SEO Title: до 60 символов, включить главный ключевик и бренд "Binokl.shop"
2. Meta Description: до 160 символов, призыв к действию, USP
3. H1: естественный заголовок с главным ключевиком

Верни JSON:
{
  "seo_title": "текст",
  "meta_description": "текст",
  "h1": "текст",
  "keywords": "через запятую топ-10 ключевиков"
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
            content: 'Ты - эксперт по SEO для интернет-магазинов. Создаешь цепляющие и оптимизированные тексты.'
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
    const seoText = data.choices[0].message.content;
    const seoData = JSON.parse(seoText);
    
    logInfo('✅ SEO теги сгенерированы', null, context);
    
    return seoData;
    
  } catch (error) {
    logError('❌ Ошибка генерации SEO', error, context);
    throw error;
  }
}

/**
 * Генерирует описание категории через AI
 */
function generateCategoryDescriptionWithAI_OLD(categoryData, products, keywords) {
  const context = "AI генерация описания";
  
  try {
    logInfo('🤖 Генерируем описание категории через AI', null, context);
    
    const apiKey = getSetting('openaiApiKey');
    
    if (!apiKey) {
      throw new Error('OpenAI API ключ не настроен');
    }
    
    // Подготавливаем данные для промпта
    const productsInfo = products.slice(0, 10).map(p => p.title).join('\n');
    const keywordsForDescription = keywords
      .filter(kw => kw.assignment === 'description')
      .map(kw => kw.keyword)
      .join(', ');
    
    const prompt = `Напиши SEO-оптимизированное описание для категории интернет-магазина оптических приборов.

КАТЕГОРИЯ: ${categoryData.title}
ПУТЬ: ${categoryData.path || 'Корневая'}

ТОВАРЫ В КАТЕГОРИИ (примеры):
${productsInfo}

КЛЮЧЕВЫЕ СЛОВА ДЛЯ ТЕКСТА:
${keywordsForDescription}

СТРУКТУРА ОПИСАНИЯ:
1. Вводный абзац (2-3 предложения) - что это за категория, для кого
2. Преимущества покупки в Binokl.shop (3-4 пункта)
3. Как выбрать товар из категории (практические советы)
4. Призыв к действию

ТРЕБОВАНИЯ:
- Естественное включение ключевых слов
- Живой, продающий язык
- Полезная информация для покупателя
- Объем: 300-500 слов
- HTML разметка: <h2>, <p>, <ul>, <li>

Верни только HTML код описания без комментариев.`;
    
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
            content: 'Ты - копирайтер для e-commerce. Пишешь продающие, информативные и SEO-оптимизированные тексты.'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.8
      }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`OpenAI API ошибка: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    const description = data.choices[0].message.content;
    
    logInfo('✅ Описание категории сгенерировано', null, context);
    
    return description;
    
  } catch (error) {
    logError('❌ Ошибка генерации описания', error, context);
    throw error;
  }
}

/**
 * Генерирует плитку тегов с AI релевантностью
 */
function generateTagTilesWithAI(categoryData, keywords, allCategories) {
  const context = "AI генерация плитки тегов";
  
  try {
    logInfo('🤖 Генерируем плитку тегов через AI', null, context);
    
    const apiKey = getSetting('openaiApiKey');
    
    if (!apiKey) {
      throw new Error('OpenAI API ключ не настроен');
    }
    
    // Подготавливаем список категорий для AI
    const categoriesText = allCategories
      .filter(cat => cat.id !== categoryData.id) // Исключаем текущую
      .slice(0, 50) // Ограничиваем для токенов
      .map(cat => `ID: ${cat.id}, Название: ${cat.title}, URL: ${cat.url}`)
      .join('\n');
    
    const keywordsText = keywords.map(kw => kw.keyword).join(', ');
    
    const prompt = `Создай плитку тегов (теговую навигацию) для категории интернет-магазина.

ТЕКУЩАЯ КАТЕГОРИЯ: ${categoryData.title}
ПУТЬ: ${categoryData.path || 'Корневая'}

КЛЮЧЕВЫЕ СЛОВА ДЛЯ ТЕГОВ:
${keywordsText}

ДОСТУПНЫЕ КАТЕГОРИИ ДЛЯ ССЫЛОК:
${categoriesText}

ЗАДАЧА:
1. Создай 8-12 тегов на основе ключевых слов
2. Для каждого тега подбери наиболее релевантную категорию для ссылки
3. Оцени релевантность связи от 0 до 1

Верни JSON массив:
[
  {
    "anchor": "текст ссылки/тега",
    "target_category_id": ID категории или null,
    "target_url": "URL категории" или null для внутренних фильтров,
    "relevance": 0.95,
    "type": "category_link" или "filter"
  }
]

Приоритет: сначала ссылки на категории, потом теги-фильтры.`;
    
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
            content: 'Ты - эксперт по внутренней перелинковке сайтов. Понимаешь семантику и связи между категориями.'
          },
          {
            role: 'user',
            content: prompt
          }
        ],
        temperature: 0.6,
        response_format: { type: 'json_object' }
      }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`OpenAI API ошибка: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    const tagsText = data.choices[0].message.content;
    const tagsData = JSON.parse(tagsText);
    
    // AI возвращает объект, нам нужен массив
    const tags = tagsData.tags || tagsData.tag_tiles || tagsData;
    
    logInfo(`✅ Сгенерировано ${Array.isArray(tags) ? tags.length : 0} тегов`, null, context);
    
    return Array.isArray(tags) ? tags : [];
    
  } catch (error) {
    logError('❌ Ошибка генерации плитки тегов', error, context);
    throw error;
  }
}

// ========================================
// ФУНКЦИИ ЗАПИСИ ДАННЫХ В ДЕТАЛЬНЫЙ ЛИСТ
// ========================================

/**
 * Обновляет SEO данные в детальном листе
 */
function updateSEOInDetailSheet(seoData) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    sheet.getRange(DETAIL_SHEET_SECTIONS.SEO_TITLE_CELL).setValue(seoData.seo_title);
    sheet.getRange(DETAIL_SHEET_SECTIONS.SEO_DESCRIPTION_CELL).setValue(seoData.meta_description);
    sheet.getRange(DETAIL_SHEET_SECTIONS.SEO_H1_CELL).setValue(seoData.h1);
    sheet.getRange(DETAIL_SHEET_SECTIONS.SEO_KEYWORDS_CELL).setValue(seoData.keywords);
    
    // Подсвечиваем обновленные ячейки
    sheet.getRange('B13:B16').setBackground('#e8f5e9');
    
    logInfo('✅ SEO данные записаны в детальный лист');
    
  } catch (error) {
    logError('❌ Ошибка записи SEO данных', error);
    throw error;
  }
}

/**
 * Обновляет описание в детальном листе
 */
function updateDescriptionInDetailSheet(description) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Записываем в ячейку описания (B17)
    sheet.getRange('B17').setValue(description);
    sheet.getRange('B17').setWrap(true);
    sheet.getRange('B17').setBackground('#fff3e0');
    
    logInfo('✅ Описание записано в детальный лист');
    
  } catch (error) {
    logError('❌ Ошибка записи описания', error);
    throw error;
  }
}

/**
 * Записывает плитку тегов в детальный лист
 */
function writeTagTilesToDetailSheet(tagTiles) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const startRow = DETAIL_SHEET_SECTIONS.TAG_TILES_START;
    
    // Заголовок блока
    sheet.getRange(startRow, 1).setValue('🏷️ ПЛИТКА ТЕГОВ').setFontWeight('bold').setFontSize(14).setBackground('#e1bee7').setFontColor('#000000');
    sheet.getRange(startRow, 1, 1, 6).merge();
    
    // Заголовки таблицы
    const headers = ['Анкор', 'URL ссылки', 'Тип', 'Релевантность', 'ID категории', 'Статус'];
    sheet.getRange(startRow + 1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(startRow + 1, 1, 1, headers.length)
         .setFontWeight('bold')
         .setBackground('#f3e5f5')
         .setHorizontalAlignment('center');
    
    // Данные тегов
    if (tagTiles.length > 0) {
      const tagRows = tagTiles.map(tag => [
        tag.anchor,
        tag.target_url || 'Фильтр',
        tag.type === 'category_link' ? 'Ссылка на категорию' : 'Тег-фильтр',
        tag.relevance || 0,
        tag.target_category_id || '',
        'Не добавлено'
      ]);
      
      sheet.getRange(startRow + 2, 1, tagRows.length, tagRows[0].length).setValues(tagRows);
      
      // Форматирование релевантности
      sheet.getRange(startRow + 2, 4, tagRows.length, 1).setNumberFormat('0.00');
      
      // Условное форматирование по релевантности
      for (let i = 0; i < tagRows.length; i++) {
        const relevance = tagRows[i][3];
        const row = startRow + 2 + i;
        
        let bgColor = '#ffebee'; // Красноватый для низкой релевантности
        if (relevance >= 0.8) bgColor = '#e8f5e9'; // Зеленый для высокой
        else if (relevance >= 0.5) bgColor = '#fff9c4'; // Желтый для средней
        
        sheet.getRange(row, 1, 1, 6).setBackground(bgColor);
      }
    }
    
    logInfo(`✅ Плитка тегов записана: ${tagTiles.length} тегов`);
    
  } catch (error) {
    logError('❌ Ошибка записи плитки тегов', error);
    throw error;
  }
}

/**
 * Показывает результаты AI-анализа
 */
function showAnalysisResults(analysis) {
  try {
    const strengths = analysis.analysis.strengths.map((s, i) => `${i + 1}. ${s}`).join('\n');
    const weaknesses = analysis.analysis.weaknesses.map((w, i) => `${i + 1}. ${w}`).join('\n');
    const opportunities = analysis.analysis.opportunities.map((o, i) => `${i + 1}. ${o}`).join('\n');
    
    const message = `📊 РЕЗУЛЬТАТЫ AI-АНАЛИЗА КАТЕГОРИИ\n\n` +
                   `✅ СИЛЬНЫЕ СТОРОНЫ:\n${strengths}\n\n` +
                   `⚠️ СЛАБЫЕ МЕСТА:\n${weaknesses}\n\n` +
                   `💡 ВОЗМОЖНОСТИ ДЛЯ УЛУЧШЕНИЯ:\n${opportunities}\n\n` +
                   `\n📝 РЕКОМЕНДАЦИИ:\n` +
                   `• SEO Title: ${analysis.recommendations.seo_title}\n` +
                   `• H1: ${analysis.recommendations.h1}\n\n` +
                   `Подробные рекомендации записаны в лист категории.`;
    
    SpreadsheetApp.getUi().alert('AI-Анализ категории', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    // Записываем рекомендации в лист
    writeAnalysisToSheet(analysis);
    
  } catch (error) {
    logError('❌ Ошибка показа результатов анализа', error);
  }
}

/**
 * Записывает результаты анализа в лист
 */
function writeAnalysisToSheet(analysis) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const startRow = DETAIL_SHEET_SECTIONS.AI_GENERATION_START;
    
    // Заголовок блока
    sheet.getRange(startRow, 1).setValue('🤖 AI-АНАЛИЗ И РЕКОМЕНДАЦИИ')
         .setFontWeight('bold')
         .setFontSize(14)
         .setBackground('#4285f4')
         .setFontColor('#ffffff');
    sheet.getRange(startRow, 1, 1, 6).merge();
    
    // Сильные стороны
    sheet.getRange(startRow + 2, 1).setValue('✅ Сильные стороны:').setFontWeight('bold');
    sheet.getRange(startRow + 3, 1, 1, 6).merge()
         .setValue(analysis.analysis.strengths.join('\n• '))
         .setWrap(true)
         .setBackground('#e8f5e9');
    
    // Слабые места
    sheet.getRange(startRow + 5, 1).setValue('⚠️ Слабые места:').setFontWeight('bold');
    sheet.getRange(startRow + 6, 1, 1, 6).merge()
         .setValue(analysis.analysis.weaknesses.join('\n• '))
         .setWrap(true)
         .setBackground('#ffebee');
    
    // Возможности
    sheet.getRange(startRow + 8, 1).setValue('💡 Возможности:').setFontWeight('bold');
    sheet.getRange(startRow + 9, 1, 1, 6).merge()
         .setValue(analysis.analysis.opportunities.join('\n• '))
         .setWrap(true)
         .setBackground('#fff9c4');
    
    // Рекомендации
    sheet.getRange(startRow + 11, 1).setValue('📝 Рекомендации:').setFontWeight('bold');
    
    const recommendations = [
      ['SEO Title:', analysis.recommendations.seo_title],
      ['SEO Description:', analysis.recommendations.seo_description],
      ['H1:', analysis.recommendations.h1],
      ['Фокус ключевиков:', analysis.recommendations.keywords_focus.join(', ')]
    ];
    
    sheet.getRange(startRow + 12, 1, recommendations.length, 2).setValues(recommendations);
    sheet.getRange(startRow + 12, 1, recommendations.length, 1).setFontWeight('bold').setBackground('#f1f3f4');
    sheet.getRange(startRow + 12, 2, recommendations.length, 1).setWrap(true);
    
    logInfo('✅ Результаты анализа записаны в лист');
    
  } catch (error) {
    logError('❌ Ошибка записи анализа', error);
    throw error;
  }
}