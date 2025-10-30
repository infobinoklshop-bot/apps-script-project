/**
 * ========================================
 * МОДУЛЬ: АНАЛИЗ ПЛОТНОСТИ КЛЮЧЕВЫХ СЛОВ
 * ========================================
 *
 * Анализирует контент страницы категории и определяет
 * плотность ключевых слов (TF - Term Frequency)
 */

/**
 * Получает весь текстовый контент категории
 * @param {number} categoryId - ID категории
 * @returns {Object} Объект с контентом страницы
 */
function getCategoryPageContent(categoryId) {
  const context = `Получение контента категории ${categoryId}`;

  try {
    console.log('[INFO] Загружаем контент категории:', categoryId);

    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }

    // 1. Загружаем категорию
    const categoryUrl = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
    const categoryOptions = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const categoryResponse = UrlFetchApp.fetch(categoryUrl, categoryOptions);

    if (categoryResponse.getResponseCode() !== 200) {
      throw new Error(`Ошибка загрузки категории: ${categoryResponse.getResponseCode()}`);
    }

    const category = JSON.parse(categoryResponse.getContentText());

    // 2. Извлекаем текстовые поля
    const textContent = {
      title: category.title || '',
      html_title: category.html_title || category.title || '',
      meta_description: category.meta_description || '',
      description: stripHtmlTags(category.description || ''),

      // Для статистики
      rawDescription: category.description || '',

      // Метаданные
      categoryId: categoryId,
      categoryName: category.title,
      url: category.url
    };

    // 3. Загружаем товары категории (первые 20 для анализа)
    console.log('[INFO] Загружаем товары категории...');

    const productsUrl = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=20`;
    const productsResponse = UrlFetchApp.fetch(productsUrl, categoryOptions);

    if (productsResponse.getResponseCode() === 200) {
      const products = JSON.parse(productsResponse.getContentText());
      textContent.productTitles = products.map(p => p.title || '').filter(t => t);
      console.log('[INFO] Загружено товаров:', textContent.productTitles.length);
    } else {
      textContent.productTitles = [];
      console.log('[WARNING] Не удалось загрузить товары');
    }

    console.log('[INFO] ✅ Контент категории загружен');

    return textContent;

  } catch (error) {
    console.error('[ERROR] Ошибка загрузки контента:', error.message);
    throw error;
  }
}

/**
 * Удаляет HTML теги из текста
 * @param {string} html - HTML текст
 * @returns {string} Чистый текст
 */
function stripHtmlTags(html) {
  if (!html) return '';

  // Удаляем script и style блоки
  let text = html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, ' ');
  text = text.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, ' ');

  // Заменяем <br>, <p>, <div> на пробелы
  text = text.replace(/<br\s*\/?>/gi, ' ');
  text = text.replace(/<\/p>/gi, ' ');
  text = text.replace(/<\/div>/gi, ' ');

  // Удаляем все остальные теги
  text = text.replace(/<[^>]+>/g, ' ');

  // Декодируем HTML entities
  text = text.replace(/&nbsp;/gi, ' ');
  text = text.replace(/&quot;/gi, '"');
  text = text.replace(/&amp;/gi, '&');
  text = text.replace(/&lt;/gi, '<');
  text = text.replace(/&gt;/gi, '>');

  // Убираем множественные пробелы
  text = text.replace(/\s+/g, ' ');

  return text.trim();
}

/**
 * Объединяет весь контент в единую строку для анализа
 * @param {Object} contentData - Данные контента
 * @returns {string} Объединенный текст
 */
function combineContentText(contentData) {
  const parts = [
    contentData.title,
    contentData.html_title,
    contentData.meta_description,
    contentData.description,
    (contentData.productTitles || []).join(' ')
  ];

  return parts.filter(p => p).join(' ').toLowerCase();
}

/**
 * Токенизирует текст (разбивает на слова)
 * @param {string} text - Текст для токенизации
 * @returns {Array<string>} Массив слов
 */
function tokenizeText(text) {
  // Приводим к нижнему регистру и разбиваем по словам
  return text
    .toLowerCase()
    .replace(/[^\u0400-\u04FFa-z0-9\s-]/g, ' ') // Оставляем только буквы, цифры, пробелы, дефисы
    .split(/\s+/)
    .filter(word => word.length > 1); // Убираем слова из 1 символа
}

/**
 * Подсчитывает вхождения ключевого слова/фразы в тексте
 * @param {string} text - Текст для поиска
 * @param {string} keyword - Ключевое слово или фраза
 * @returns {number} Количество вхождений
 */
function countKeywordOccurrences(text, keyword) {
  const normalizedText = text.toLowerCase();
  const normalizedKeyword = keyword.toLowerCase().trim();

  // Если это фраза (несколько слов)
  if (normalizedKeyword.includes(' ')) {
    const regex = new RegExp(normalizedKeyword.replace(/\s+/g, '\\s+'), 'gi');
    const matches = normalizedText.match(regex);
    return matches ? matches.length : 0;
  }

  // Если это одно слово - ищем как целое слово (word boundary)
  const regex = new RegExp(`\\b${normalizedKeyword}\\b`, 'gi');
  const matches = normalizedText.match(regex);
  return matches ? matches.length : 0;
}

/**
 * Анализирует плотность ключевых слов
 * @param {number} categoryId - ID категории
 * @param {Array<Object>} keywords - Массив ключевых слов [{keyword, type}]
 * @returns {Object} Результаты анализа
 */
function analyzeKeywordDensity(categoryId, keywords) {
  const context = `Анализ плотности для категории ${categoryId}`;

  try {
    console.log('[INFO] 🔍 Начинаем анализ плотности ключевых слов');
    console.log('[INFO] Ключевых слов для анализа:', keywords.length);

    // 1. Получаем контент страницы
    const contentData = getCategoryPageContent(categoryId);
    const fullText = combineContentText(contentData);

    // 2. Токенизируем текст
    const tokens = tokenizeText(fullText);
    const totalWords = tokens.length;

    console.log('[INFO] Всего слов на странице:', totalWords);

    if (totalWords === 0) {
      throw new Error('На странице нет текстового контента для анализа');
    }

    // 3. Анализируем каждое ключевое слово
    const results = [];
    const config = TAG_TILES_CONFIG.DENSITY_THRESHOLDS;

    keywords.forEach(kwData => {
      const keyword = kwData.keyword || kwData;
      const type = kwData.type || 'ADDITIONAL';

      // Подсчитываем вхождения
      const occurrences = countKeywordOccurrences(fullText, keyword);

      // Вычисляем TF в процентах
      const tfPercent = (occurrences / totalWords) * 100;

      // Определяем целевые пороги в зависимости от типа
      let targetMin, targetMax;

      if (type === 'MAIN_KEYWORD' || type === 'Целевой') {
        targetMin = config.MAIN_KEYWORD.min;
        targetMax = config.MAIN_KEYWORD.max;
      } else if (type === 'LSI' || type === 'Тематика') {
        targetMin = config.LSI.min;
        targetMax = config.LSI.max;
      } else {
        targetMin = config.ADDITIONAL.min;
        targetMax = config.ADDITIONAL.max;
      }

      // Определяем статус
      let status;
      if (tfPercent > config.SPAM_THRESHOLD) {
        status = '⚠️ Переспам';
      } else if (tfPercent >= targetMin && tfPercent <= targetMax) {
        status = '✅ Достаточно';
      } else if (tfPercent < targetMin) {
        status = '❌ Недостаточно';
      } else {
        status = '⚠️ Много';
      }

      results.push({
        keyword: keyword,
        type: type,
        occurrences: occurrences,
        totalWords: totalWords,
        tfPercent: parseFloat(tfPercent.toFixed(2)),
        targetMin: targetMin,
        targetMax: targetMax,
        status: status
      });
    });

    console.log('[INFO] ✅ Анализ завершен. Проанализировано ключевых слов:', results.length);

    // 4. Формируем итоговый результат
    const summary = {
      categoryId: categoryId,
      categoryName: contentData.categoryName,
      totalWords: totalWords,
      analyzedKeywords: results.length,

      // Статистика
      sufficient: results.filter(r => r.status === '✅ Достаточно').length,
      insufficient: results.filter(r => r.status === '❌ Недостаточно').length,
      overspam: results.filter(r => r.status === '⚠️ Переспам').length,
      tooMany: results.filter(r => r.status === '⚠️ Много').length,

      // Детальные результаты
      results: results,

      // Метаданные
      analyzedAt: new Date(),
      contentData: contentData
    };

    return summary;

  } catch (error) {
    console.error('[ERROR] Ошибка анализа плотности:', error.message);
    throw error;
  }
}

/**
 * Сохраняет результаты анализа в лист "Анализ плотности"
 * @param {Object} analysisResult - Результат анализа
 */
function saveDensityAnalysisToSheet(analysisResult) {
  try {
    console.log('[INFO] 💾 Сохраняем результаты анализа плотности');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(TAG_TILES_CONFIG.DENSITY_ANALYSIS_SHEET);

    // Создаем лист если не существует
    if (!sheet) {
      sheet = ss.insertSheet(TAG_TILES_CONFIG.DENSITY_ANALYSIS_SHEET);
      setupDensityAnalysisSheet(sheet);
    }

    // Удаляем старые данные для этой категории
    const data = sheet.getDataRange().getValues();
    const categoryId = analysisResult.categoryId;

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][DENSITY_ANALYSIS_COLUMNS.CATEGORY_ID - 1] == categoryId) {
        sheet.deleteRow(i + 1);
      }
    }

    // Добавляем новые данные
    const rows = analysisResult.results.map(r => [
      analysisResult.categoryId,
      analysisResult.categoryName,
      r.keyword,
      r.type,
      r.occurrences,
      r.totalWords,
      r.tfPercent,
      r.targetMin,
      r.targetMax,
      r.status,
      'Нет',           // Использовано в плитке
      '',              // Анкор
      new Date()       // Дата анализа
    ]);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

      // Форматируем процентные значения
      const tfColumn = DENSITY_ANALYSIS_COLUMNS.TF_PERCENT;
      sheet.getRange(sheet.getLastRow() - rows.length + 1, tfColumn, rows.length, 1)
        .setNumberFormat('0.00%');
    }

    console.log('[INFO] ✅ Сохранено записей:', rows.length);

  } catch (error) {
    console.error('[ERROR] Ошибка сохранения анализа:', error.message);
    throw error;
  }
}

/**
 * Настраивает лист "Анализ плотности"
 * @param {Sheet} sheet - Лист для настройки
 */
function setupDensityAnalysisSheet(sheet) {
  sheet.clear();

  const headers = [
    'ID категории',
    'Категория',
    'Ключевое слово',
    'Тип',
    'Вхождений',
    'Всего слов',
    'TF%',
    'Мин%',
    'Макс%',
    'Статус',
    'В плитке',
    'Анкор',
    'Дата анализа'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#2196f3')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 100);   // ID
  sheet.setColumnWidth(2, 200);   // Категория
  sheet.setColumnWidth(3, 250);   // Ключевое слово
  sheet.setColumnWidth(4, 120);   // Тип
  sheet.setColumnWidth(5, 80);    // Вхождений
  sheet.setColumnWidth(6, 100);   // Всего слов
  sheet.setColumnWidth(7, 80);    // TF%
  sheet.setColumnWidth(8, 60);    // Мин%
  sheet.setColumnWidth(9, 60);    // Макс%
  sheet.setColumnWidth(10, 150);  // Статус
  sheet.setColumnWidth(11, 80);   // В плитке
  sheet.setColumnWidth(12, 200);  // Анкор
  sheet.setColumnWidth(13, 150);  // Дата

  sheet.setFrozenRows(1);

  console.log('[INFO] ✅ Лист "Анализ плотности" настроен');
}

/**
 * Получает список ключевых слов для категории из листа "Ключевые слова"
 * @param {number} categoryId - ID категории
 * @returns {Array<Object>} Массив ключевых слов
 */
function getKeywordsForCategory(categoryId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.KEYWORDS);

    if (!sheet) {
      console.log('[WARNING] Лист "Ключевые слова" не найден');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const keywords = [];

    // Пропускаем заголовок
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Проверяем ID категории (колонка A)
      if (row[0] == categoryId) {
        keywords.push({
          keyword: row[2],        // Ключевое слово
          frequency: row[3],      // Частотность
          type: row[5] || 'ADDITIONAL'  // Тип
        });
      }
    }

    console.log('[INFO] Найдено ключевых слов для категории:', keywords.length);

    return keywords;

  } catch (error) {
    console.error('[ERROR] Ошибка получения ключевых слов:', error.message);
    return [];
  }
}

/**
 * ГЛАВНАЯ ФУНКЦИЯ: Анализирует плотность и сохраняет результаты
 * @param {number} categoryId - ID категории
 * @returns {Object} Результаты анализа
 */
function runKeywordDensityAnalysis(categoryId) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Анализируем плотность ключевых слов...',
      '🔍 Анализ',
      -1
    );

    // 1. Получаем ключевые слова
    const keywords = getKeywordsForCategory(categoryId);

    if (keywords.length === 0) {
      throw new Error('Для этой категории нет ключевых слов. Добавьте их в лист "Ключевые слова".');
    }

    // 2. Анализируем плотность
    const analysisResult = analyzeKeywordDensity(categoryId, keywords);

    // 3. Сохраняем результаты
    saveDensityAnalysisToSheet(analysisResult);

    // 4. Показываем результат
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Проанализировано: ${analysisResult.analyzedKeywords} слов
Недостаточно: ${analysisResult.insufficient}
Достаточно: ${analysisResult.sufficient}
Переспам: ${analysisResult.overspam}`,
      '✅ Анализ завершен',
      10
    );

    return analysisResult;

  } catch (error) {
    console.error('[ERROR] Ошибка анализа:', error.message);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка: ' + error.message,
      '❌ Ошибка',
      10
    );

    throw error;
  }
}

/**
 * ТЕСТОВАЯ ФУНКЦИЯ: Анализ плотности для категории Бинокли
 */
function testDensityAnalysis() {
  const categoryId = 9175197; // Бинокли
  const result = runKeywordDensityAnalysis(categoryId);

  console.log('=== РЕЗУЛЬТАТЫ АНАЛИЗА ===');
  console.log('Категория:', result.categoryName);
  console.log('Всего слов:', result.totalWords);
  console.log('Недостаточно:', result.insufficient);
  console.log('Достаточно:', result.sufficient);
  console.log('Переспам:', result.overspam);

  return result;
}
