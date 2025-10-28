/**
 * ========================================
 * МОДУЛЬ: ГЕНЕРАЦИЯ ОПИСАНИЙ КАТЕГОРИЙ ЧЕРЕЗ ASSISTANTS API
 * ========================================
 */

/**
 * Генерирует описание категории через OpenAI Assistants API
 * ЛОГИКА:
 * - Если описание в B17 есть → отправляет на РЕРАЙТ в C17
 * - Если описания в B17 нет → генерирует НОВОЕ в C17
 * - После генерации проверяет ОЧИЩЕННЫЙ текст детектором ИИ
 * - В ячейку сохраняется HTML-версия для отправки в InSales
 */
function generateDescriptionForActiveCategory() {
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
    
    // Проверяем настройку ассистента
    if (typeof CATEGORY_DESCRIPTION_ASSISTANT_ID === 'undefined' || 
        CATEGORY_DESCRIPTION_ASSISTANT_ID === 'asst_XXXXX') {
      SpreadsheetApp.getUi().alert(
        'Ошибка настройки',
        'Не настроен ID ассистента.\n\nОткройте 01_config.gs и укажите ID ассистента',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Собираем данные
    const sheet = SpreadsheetApp.getActiveSheet();
    const h1 = sheet.getRange('B15').getValue() || categoryData.title;
    const existingDescription = sheet.getRange('B17').getValue() || '';
    const productsCount = sheet.getRange('B21').getValue() || 0;
    
    // Определяем режим
    const isRewrite = existingDescription && existingDescription.toString().trim().length > 100;
    const mode = isRewrite ? 'РЕРАЙТ' : 'ГЕНЕРАЦИЯ С НУЛЯ';
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${mode} описания через AI...`,
      '⏳ Работаем',
      -1
    );
    
    console.log(`📝 Режим: ${mode}`);
    
    // Получаем товары и формируем промпт
    const productsData = getProductsDataFromSheet(sheet);
    const userPrompt = isRewrite 
      ? buildRewritePrompt(categoryData.title, h1, existingDescription, productsCount, productsData)
      : buildNewDescriptionPrompt(categoryData.title, h1, productsCount, productsData);
    
    console.log('📤 Отправляем запрос ассистенту');
    
    // Вызываем ассистента - ПОЛУЧАЕМ HTML
    const newDescriptionHTML = callOpenAIAssistantForCategory(
      userPrompt, 
      CATEGORY_DESCRIPTION_ASSISTANT_ID
    );
    
    console.log(`✅ Получен HTML-текст: ${newDescriptionHTML.length} символов`);
    
    // ========================================
    // ПРОВЕРКА ДЕТЕКТОРОМ ИИ (ОЧИЩЕННЫЙ ТЕКСТ)
    // ========================================
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Проверяем текст детектором ИИ...',
      '🤖 Анализ',
      -1
    );
    
    // ВАЖНО: Очищаем HTML ТОЛЬКО для проверки
    const cleanTextForAI = stripHtmlTags(newDescriptionHTML);
    
    console.log('🔍 Запускаем детектор ИИ на ОЧИЩЕННОМ тексте');
    console.log(`📏 HTML-версия: ${newDescriptionHTML.length} символов`);
    console.log(`📏 Чистая версия для ИИ: ${cleanTextForAI.length} символов`);
    console.log(`📊 Удалено HTML: ${newDescriptionHTML.length - cleanTextForAI.length} символов`);
    
    let aiScore = null;
    let aiCheckStatus = '';
    
    try {
      // Проверяем ОЧИЩЕННЫЙ текст
      const apiKey = 'bc50023f63cbd635f6bb291c16395545';
      aiScore = checkAIDetectionScore(cleanTextForAI, apiKey);
      
      console.log(`✅ Детектор ИИ: ${aiScore}%`);
      
      // Оценка результата
      if (aiScore <= 30) {
        aiCheckStatus = `✅ Отлично (${aiScore}%)`;
      } else if (aiScore <= 50) {
        aiCheckStatus = `⚠️ Приемлемо (${aiScore}%)`;
      } else if (aiScore <= 70) {
        aiCheckStatus = `⚠️ Высокий (${aiScore}%)`;
      } else {
        aiCheckStatus = `❌ Очень высокий (${aiScore}%)`;
      }
      
    } catch (aiError) {
      console.error('❌ Ошибка детектора ИИ:', aiError.message);
      aiCheckStatus = `⚠️ Ошибка: ${aiError.message}`;
    }
    
    // Записываем HTML-версию в ячейку
    writeNewDescriptionToColumn(sheet, newDescriptionHTML, aiCheckStatus);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${mode} завершена!`,
      '✅ Готово',
      5
    );
    
    // Уведомление с результатом проверки
    const message = isRewrite 
      ? `✅ Описание отредактировано!\n\n📍 Старая версия: колонка B\n📍 Новая версия: колонка C (HTML)\n🤖 Детектор ИИ: ${aiCheckStatus}\n\nПроверьте и выберите лучший вариант.`
      : `✅ Описание создано с нуля!\n\n📍 Результат в колонке C (HTML)\n🤖 Детектор ИИ: ${aiCheckStatus}\n\nПроверьте перед отправкой в InSales.`;
    
    SpreadsheetApp.getUi().alert('Готово', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('❌ Ошибка генерации описания:', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Промпт для РЕРАЙТА существующего описания
 */
function buildRewritePrompt(title, h1, existingDescription, totalCount, productsData) {
  const priceRange = productsData.minPrice > 0 
    ? `от ${productsData.minPrice}₽ до ${productsData.maxPrice}₽`
    : 'не указана';
  
  return `ЗАДАЧА: Отредактируй существующее описание категории интернет-магазина оптических приборов.

ДАННЫЕ О КАТЕГОРИИ:
- Название: ${title}
- H1: ${h1}
- Количество товаров: ${totalCount} (в наличии: ${productsData.count})
- Диапазон цен: ${priceRange}

ТЕКУЩЕЕ ОПИСАНИЕ (ТРЕБУЕТ УЛУЧШЕНИЯ):
${existingDescription}

ЧТО НУЖНО СДЕЛАТЬ:
1. Сохрани структуру и основную информацию
2. Улучши читаемость и естественность текста
3. Убери канцеляризмы типа "является", "представляет собой", "позволяет осуществлять"
4. Добавь конкретики и пользы для покупателя
5. Сделай текст более живым и продающим
6. Сохрани ключевые слова, но вплети их естественнее
7. Убедись, что призыв к действию яркий и убедительный

ТРЕБОВАНИЯ:
- Объем: 300-500 слов (как в оригинале ±50 слов)
- HTML разметка: <h2>, <p>, <ul>, <li>
- Живой язык без штампов
- НЕ упоминай конкретные цены

Верни только улучшенный HTML код описания.`;
}

/**
 * Промпт для НОВОЙ генерации (когда описания нет)
 */
function buildNewDescriptionPrompt(title, h1, totalCount, productsData) {
  const productsList = productsData.products.length > 0 
    ? productsData.products.map(p => `• ${p.title} (${p.price}₽, ${p.brand || 'без бренда'})`).join('\n')
    : 'Нет товаров в наличии';
  
  const priceRange = productsData.minPrice > 0 
    ? `от ${productsData.minPrice}₽ до ${productsData.maxPrice}₽`
    : 'не указана';
  
  return `ЗАДАЧА: Создай SEO-оптимизированное описание для категории интернет-магазина оптических приборов с нуля.

ДАННЫЕ О КАТЕГОРИИ:
- Название: ${title}
- H1: ${h1}
- Количество товаров: ${totalCount} (в наличии: ${productsData.count})
- Диапазон цен: ${priceRange}

ПРИМЕРЫ ТОВАРОВ В КАТЕГОРИИ:
${productsList}

СТРУКТУРА ОПИСАНИЯ:
1. Вводный абзац (2-3 предложения) - что это за категория, для кого, зачем
2. Описание типов товаров и их применения
3. Преимущества покупки в Binokl.shop:
   - Широкий ассортимент профессиональных приборов
   - Консультации экспертов
   - Быстрая доставка
   - Гарантия качества
4. Советы по выбору товара
5. Призыв к действию

ТРЕБОВАНИЯ:
- Объем: 300-500 слов
- Живой, продающий язык
- HTML разметка: <h2>, <p>, <ul>, <li>
- Естественное включение ключевых слов из названия
- НЕ упоминай конкретные цены

Верни только готовый HTML код описания.`;
}

/**
 * Собирает данные товаров из листа категории
 */
function getProductsDataFromSheet(sheet) {
  try {
    const startRow = 28;
    const data = sheet.getRange(startRow, 1, 20, 8).getValues();
    
    const products = [];
    let minPrice = Infinity;
    let maxPrice = 0;
    
    for (let row of data) {
      if (!row[3] || row[3].toString().trim() === '') break;
      
      const title = row[3];
      const inStock = row[4];
      const price = parseFloat(row[5]) || 0;
      const brand = row[6];
      
      if (inStock && inStock.toString().toLowerCase() !== 'нет' && price > 0) {
        products.push({ title, price, brand });
        
        if (price < minPrice) minPrice = price;
        if (price > maxPrice) maxPrice = price;
      }
    }
    
    return {
      products: products.slice(0, 10),
      count: products.length,
      minPrice: minPrice === Infinity ? 0 : minPrice,
      maxPrice: maxPrice
    };
    
  } catch (error) {
    console.error('❌ Ошибка чтения товаров из листа:', error);
    return { products: [], count: 0, minPrice: 0, maxPrice: 0 };
  }
}

/**
 * Вызов OpenAI Assistant для категорий
 */
function callOpenAIAssistantForCategory(userContent, assistantId) {
  try {
    const apiKey = getOpenAIConfig().apiKey;
    
    if (!apiKey || apiKey.includes('ВАШ')) {
      throw new Error('OpenAI API ключ не настроен в файле 01_config.gs');
    }
    
    console.log('🤖 Вызываем ассистента:', assistantId);
    
    const result = executeOpenAIRequestForCategory(userContent, assistantId, apiKey);
    
    if (!result || result.trim().length === 0) {
      throw new Error('Ассистент вернул пустой результат');
    }
    
    const cleaned = cleanMarkdownFromResult(result);
    console.log('✅ Получен ответ длиной:', cleaned.length, 'символов');
    
    return cleaned;
    
  } catch (error) {
    console.error('❌ Ошибка вызова ассистента:', error);
    throw new Error(`Ошибка OpenAI Assistant: ${error.message}`);
  }
}

/**
 * Выполнение запроса к OpenAI Assistants API
 */
function executeOpenAIRequestForCategory(userContent, assistantId, apiKey) {
  let threadId;
  
  // 1. Создание треда
  try {
    console.log('📝 Создаем thread...');
    
    const threadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/threads', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      },
      payload: JSON.stringify({
        messages: [{ role: 'user', content: userContent }]
      }),
      muteHttpExceptions: true
    });

    if (threadResponse.getResponseCode() >= 400) {
      throw new Error(`Thread error: ${threadResponse.getContentText()}`);
    }

    threadId = JSON.parse(threadResponse.getContentText()).id;
    console.log('✅ Thread создан:', threadId);
    
  } catch (error) {
    throw new Error(`Ошибка создания thread: ${error.message}`);
  }

  // 2. Запуск
  let runId;
  try {
    const runResponse = UrlFetchApp.fetch(`https://api.openai.com/v1/threads/${threadId}/runs`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
        'OpenAI-Beta': 'assistants=v2'
      },
      payload: JSON.stringify({ assistant_id: assistantId }),
      muteHttpExceptions: true
    });

    runId = JSON.parse(runResponse.getContentText()).id;
    console.log('✅ Run запущен:', runId);
    
  } catch (error) {
    throw new Error(`Ошибка run: ${error.message}`);
  }

  // 3. Ожидание
  let status = 'queued';
  let attempts = 0;
  
  while (status !== 'completed' && attempts < 40) {
    Utilities.sleep(3000);
    attempts++;
    
    const statusResponse = UrlFetchApp.fetch(
      `https://api.openai.com/v1/threads/${threadId}/runs/${runId}`,
      {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${apiKey}`,
          'OpenAI-Beta': 'assistants=v2'
        }
      }
    );
    
    status = JSON.parse(statusResponse.getContentText()).status;
    console.log(`⏳ ${attempts}/40: ${status}`);
  }

  if (status !== 'completed') {
    throw new Error(`Не завершился: ${status}`);
  }

  // 4. Результат
  const messagesResponse = UrlFetchApp.fetch(
    `https://api.openai.com/v1/threads/${threadId}/messages`,
    {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'OpenAI-Beta': 'assistants=v2'
      }
    }
  );

  const messages = JSON.parse(messagesResponse.getContentText()).data;
  
  for (let msg of messages) {
    if (msg.role === 'assistant' && msg.content[0]?.text?.value) {
      return msg.content[0].text.value;
    }
  }

  throw new Error('Нет ответа от ассистента');
}

/**
 * Очистка markdown из ответа
 */
function cleanMarkdownFromResult(text) {
  if (!text || typeof text !== 'string') {
    return text;
  }

  let cleaned = text.trim();
  
  // Удаляем markdown блоки
  cleaned = cleaned.replace(/^```html\s*/i, '');
  cleaned = cleaned.replace(/^```\s*/i, '');
  cleaned = cleaned.replace(/\s*```\s*$/i, '');
  
  // Удаляем префиксы ассистентов
  cleaned = cleaned.replace(/^json\s*/i, '');
  cleaned = cleaned.replace(/^html\s*/i, '');
  cleaned = cleaned.replace(/^response:\s*/i, '');
  cleaned = cleaned.replace(/^result:\s*/i, '');
  
  return cleaned.trim();
}

/**
 * Проверка текста детектором ИИ через Text.ru API
 * ВАЖНО: принимает ОЧИЩЕННЫЙ от HTML текст
 */
function checkAIDetectionScore(cleanText, apiKey) {
  if (!cleanText || cleanText.length < 100) {
    throw new Error('Текст слишком короткий для проверки (минимум 100 символов)');
  }
  
  if (cleanText.length > 20000) {
    throw new Error('Текст слишком длинный (максимум 20000 символов)');
  }
  
  try {
    console.log('🔍 Создаём задачу детектора ИИ...');
    console.log(`📏 Длина чистого текста: ${cleanText.length} символов`);
    
    // 1. Создание задачи
    const createResponse = UrlFetchApp.fetch('https://api.text.ru/neurotools/api/v1/task/detector', {
      method: 'POST',
      headers: {
        'X-USERKEY': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ text: cleanText }),
      muteHttpExceptions: true
    });

    if (createResponse.getResponseCode() >= 400) {
      throw new Error(`Ошибка создания задачи: ${createResponse.getContentText()}`);
    }

    const createResult = JSON.parse(createResponse.getContentText());
    const taskId = createResult.taskId || createResult.task_id || createResult.id;
    
    if (!taskId) {
      throw new Error('Не получен ID задачи от API');
    }
    
    console.log(`✅ Задача создана: ${taskId}`);

    // 2. Ожидание результата
    let attempts = 0;
    const maxAttempts = 20;
    
    while (attempts < maxAttempts) {
      Utilities.sleep(4000);
      attempts++;
      
      console.log(`⏳ Проверка ${attempts}/${maxAttempts}...`);
      
      const statusResponse = UrlFetchApp.fetch(
        `https://api.text.ru/neurotools/api/v1/task/detector/${taskId}`,
        {
          method: 'GET',
          headers: {
            'X-USERKEY': apiKey,
            'Content-Type': 'application/json'
          },
          muteHttpExceptions: true
        }
      );

      if (statusResponse.getResponseCode() >= 400) {
        console.warn(`⚠️ Ошибка получения статуса: ${statusResponse.getContentText()}`);
        continue;
      }

      const statusResult = JSON.parse(statusResponse.getContentText());
      const status = statusResult.status || 'UNKNOWN';

      console.log(`📊 Статус: ${status}`);

      if (status === 'IN_PROCESS' || status === 'PROCESSING') {
        continue;
      }

      if (status === 'FAILED' || status === 'ERROR') {
        throw new Error(`Задача завершилась с ошибкой: ${statusResult.error || 'неизвестная'}`);
      }

      if (status === 'READY' || status === 'COMPLETED') {
        const resultData = statusResult.result || statusResult.data || {};
        let aiPercent = resultData.percent || resultData.ai_percent || resultData.percentage;
        
        if (aiPercent === undefined || isNaN(aiPercent)) {
          throw new Error('Не найден процент в результате');
        }
        
        aiPercent = Math.max(0, Math.min(100, parseFloat(aiPercent)));
        
        console.log(`✅ Детектор вернул результат: ${aiPercent}%`);
        
        return Math.round(aiPercent);
      }
    }

    throw new Error('Детектор не завершился за разумное время');

  } catch (error) {
    console.error('❌ Ошибка детектора ИИ:', error.message);
    throw new Error(`Детектор ИИ: ${error.message}`);
  }
}

/**
 * Удаляет HTML-теги из текста ТОЛЬКО ДЛЯ ПРОВЕРКИ
 * Оригинальный HTML сохраняется в ячейках
 */
function stripHtmlTags(html) {
  if (!html || typeof html !== 'string') {
    return '';
  }
  
  let text = html;
  
  console.log('🧹 Очистка HTML для детектора...');
  console.log(`📏 До очистки: ${html.length} символов`);
  
  // Заменяем структурные элементы на читаемые
  text = text.replace(/<\/p>/gi, '\n\n');
  text = text.replace(/<p[^>]*>/gi, '');
  text = text.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<li[^>]*>/gi, '• ');
  text = text.replace(/<\/li>/gi, '\n');
  text = text.replace(/<\/(ul|ol)>/gi, '\n');
  text = text.replace(/<h2[^>]*>/gi, '\n');
  text = text.replace(/<\/h2>/gi, '\n');
  
  // Удаляем все остальные теги
  text = text.replace(/<[^>]+>/g, '');
  
  // Декодируем HTML-сущности
  text = text.replace(/&nbsp;/g, ' ');
  text = text.replace(/&quot;/g, '"');
  text = text.replace(/&apos;/g, "'");
  text = text.replace(/&lt;/g, '<');
  text = text.replace(/&gt;/g, '>');
  text = text.replace(/&amp;/g, '&');
  text = text.replace(/&mdash;/g, '—');
  
  // Очистка лишних пробелов
  text = text.replace(/\n\s*\n\s*\n/g, '\n\n');
  text = text.replace(/[ \t]+/g, ' ');
  text = text.trim();
  
  console.log(`📏 После очистки: ${text.length} символов`);
  console.log(`📊 Удалено: ${html.length - text.length} символов (HTML-разметка)`);
  
  return text;
}

/**
 * Записывает новое описание в колонку C с результатом проверки ИИ
 * ВАЖНО: сохраняет HTML-версию для отправки в InSales
 */
function writeNewDescriptionToColumn(sheet, htmlDescription, aiCheckStatus) {
  try {
    console.log('💾 Записываем результаты...');
    console.log(`📏 HTML-описание: ${htmlDescription.length} символов`);
    
    // Записываем HTML-версию в C17
    sheet.getRange('C17').setValue(htmlDescription);
    sheet.getRange('C17').setWrap(true);
    sheet.getRange('C17').setBackground('#c8e6c9');
    
    console.log('✅ HTML-описание записано в C17');
    
    // Заголовок для C16
    const headerC = sheet.getRange('C16').getValue();
    if (!headerC || headerC.toString().trim() === '') {
      sheet.getRange('C16').setValue('Новая версия описания (HTML):');
      sheet.getRange('C16').setFontWeight('bold');
      sheet.getRange('C16').setBackground('#81c784');
    }
    
    // Записываем результат детектора ИИ в D17
    if (aiCheckStatus) {
      sheet.getRange('D16').setValue('Детектор ИИ:');
      sheet.getRange('D16').setFontWeight('bold');
      sheet.getRange('D16').setBackground('#90caf9');
      
      sheet.getRange('D17').setValue(aiCheckStatus);
      sheet.getRange('D17').setWrap(true);
      
      // Цвет в зависимости от результата
      if (aiCheckStatus.startsWith('✅')) {
        sheet.getRange('D17').setBackground('#c8e6c9'); // Зеленый
      } else if (aiCheckStatus.startsWith('⚠️')) {
        sheet.getRange('D17').setBackground('#fff9c4'); // Желтый
      } else if (aiCheckStatus.startsWith('❌')) {
        sheet.getRange('D17').setBackground('#ffcdd2'); // Красный
      } else {
        sheet.getRange('D17').setBackground('#f5f5f5'); // Серый
      }
      
      sheet.setColumnWidth(4, 200);
      console.log('✅ Результат детектора ИИ записан в D17');
    }
    
    // Подсвечиваем старую версию
    const oldDescription = sheet.getRange('B17').getValue();
    if (oldDescription && oldDescription.toString().trim().length > 0) {
      sheet.getRange('B17').setBackground('#fff9c4');
      sheet.getRange('B16').setValue('Описание категории (старая версия):');
      sheet.getRange('B16').setFontWeight('bold');
      sheet.getRange('B16').setBackground('#ffeb3b');
    }
    
    console.log('✅ Все данные записаны успешно');
    
  } catch (error) {
    console.error('❌ Ошибка записи:', error);
    throw error;
  }
}