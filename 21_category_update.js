/**
 * ========================================
 * ОТПРАВКА ВСЕХ ИЗМЕНЕНИЙ КАТЕГОРИИ В INSALES
 * ========================================
 */

/**
 * Отправляет ВСЕ изменения категории в InSales (ПОЛНАЯ ВЕРСИЯ)
 */
function sendCategoryChangesToInSales() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    if (!sheetName.startsWith('Категория — ')) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        'Эта функция работает только на детальном листе категории',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    console.log('[INFO] Отправка всех изменений категории в InSales');
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Начинаем отправку изменений категории...',
      '⏳ Отправка',
      -1
    );
    
    // Читаем данные из листа
    const categoryId = sheet.getRange('B2').getValue();
    
    if (!categoryId) {
      throw new Error('ID категории не найден в ячейке B2');
    }
    
    console.log('[INFO] Отправка изменений категории', categoryId);

    // ИСПРАВЛЕНО: Определяем номер строки начала доп.полей динамически
    const productsValue = sheet.getRange('B21').getValue();
    const products = parseInt(productsValue) || 0; // ВАЖНО: Принудительное преобразование в число!
    const productsStartRow = calculateSheetSections(sheet).productsStart;
    const extraFieldsStartRow = productsStartRow + products + 5; // Строка после товаров

    console.log(`[DEBUG] products: ${products} (type: ${typeof products}), productsStartRow: ${productsStartRow} (type: ${typeof productsStartRow})`);
    
    // Читаем ВЕРХНИЕ и НИЖНИЕ теги
    // ПРИОРИТЕТ: сначала пробуем из таблицы, потом из дополнительных полей
    let topTags = getTopTagsFromTable(sheet, products);
    if (topTags.length === 0) {
      console.log('⚠️ Верхняя таблица пуста, читаем из дополнительных полей...');
      topTags = getTopTagsFromSheet(sheet, products);
    }

    let bottomTags = getBottomTagsFromTable(sheet, products);
    if (bottomTags.length === 0) {
      console.log('⚠️ Нижняя таблица пуста, читаем из дополнительных полей...');
      bottomTags = getBottomTagsFromSheet(sheet, products);
    }
    
    // Получаем описание
    let description = getDescriptionForInSales(sheet);
    
    // Формируем финальное описание (описание идёт БЕЗ тегов, теги в отдельных полях)
    let finalDescription = description;
    
    // Собираем field_values для отправки
    const fieldValues = [];
    
    // H1 заголовок - отправляем с ID существующего field_value
    const h1Value = sheet.getRange('B15').getValue();
    if (h1Value && h1Value.toString().trim() !== '') {
      const h1FieldValueId = findFieldValueId(categoryId, 'H1');
      if (h1FieldValueId) {
        fieldValues.push({ 
          id: h1FieldValueId,
          value: h1Value.toString().trim() 
        });
      }
    }
    
    // Блок ссылок сверху (верхняя плитка)
    if (topTags.length > 0) {
      const topTagsHTML = generateTagsBlockHTML(topTags);
      const topTagsFieldValueId = findFieldValueId(categoryId, 'Блок ссылок сверху');
      if (topTagsFieldValueId) {
        fieldValues.push({ 
          id: topTagsFieldValueId,
          value: topTagsHTML 
        });
        console.log('[INFO] 🏷️ Верхних тегов:', topTags.length);
      }
    }

    // Блок ссылок (нижняя плитка)
    if (bottomTags.length > 0) {
      const bottomTagsHTML = generateTagsBlockHTML(bottomTags);
      const bottomTagsFieldValueId = findFieldValueId(categoryId, 'Блок ссылок');
      
      if (bottomTagsFieldValueId) {
        // ИСПРАВЛЕНО: Обновляем существующее поле через id
        fieldValues.push({ 
          id: bottomTagsFieldValueId,
          value: bottomTagsHTML 
        });
        console.log('[INFO] 🏷️ Нижних тегов:', bottomTags.length);
      } else {
        // Если поле не существует - создаём новое
        const bottomTagsFieldId = getFieldIdByName('Блок ссылок');
        if (bottomTagsFieldId) {
          fieldValues.push({ 
            collection_field_id: bottomTagsFieldId,
            value: bottomTagsHTML 
          });
          console.log('[INFO] 🏷️ Создаём новое поле "Блок ссылок" с тегами:', bottomTags.length);
        }
      }
    }
    
    // Дополнительные поля категории - ДИНАМИЧЕСКОЕ ЧТЕНИЕ
    console.log('[INFO] Читаем дополнительные поля из листа...');

    const collectionFields = loadCollectionFieldsDictionary();
    const startRow = extraFieldsStartRow + 2;

    // Читаем все строки из секции дополнительных полей
    for (let i = 0; i < 50; i++) {
      const fieldNameCell = sheet.getRange(startRow + i, 1).getValue();
      const fieldValueCell = sheet.getRange(startRow + i, 2).getValue();
      
      // Если строка с названием поля пустая - прекращаем чтение
      if (!fieldNameCell || fieldNameCell.toString().trim() === '') {
        break;
      }
      
      // Извлекаем название поля (убираем ":" в конце)
      const fieldName = fieldNameCell.toString().replace(':', '').trim();
      
      // Пропускаем пустые значения и "(пусто)"
      if (!fieldValueCell || 
          fieldValueCell.toString().trim() === '' || 
          fieldValueCell.toString().trim() === '(пусто)') {
        console.log(`⏭️ Пропускаем пустое поле: ${fieldName}`);
        continue;
      }
      
      // ИСПРАВЛЕНО: Пропускаем поля плиток тегов - они обрабатываются отдельно
      if (fieldName === 'Блок ссылок сверху' || fieldName === 'Блок ссылок') {
        console.log(`⏭️ Пропускаем поле плитки тегов: ${fieldName} (обрабатывается отдельно)`);
        continue;
      }

      // Находим ID поля в справочнике
      const fieldInfo = collectionFields.find(f => f.title === fieldName);
      
      if (fieldInfo) {
        // Находим существующий field_value_id для этого поля
        const existingFieldValueId = findFieldValueId(categoryId, fieldName);
        
        if (existingFieldValueId) {
          // Обновляем существующее
          fieldValues.push({
            id: existingFieldValueId,
            value: fieldValueCell.toString().trim()
          });
          console.log(`📝 Обновляем поле "${fieldName}"`);
        } else {
          // Создаем новое
          fieldValues.push({
            collection_field_id: fieldInfo.id,
            value: fieldValueCell.toString().trim()
          });
          console.log(`➕ Создаем новое поле "${fieldName}"`);
        }
      } else {
        console.warn(`⚠️ Поле "${fieldName}" не найдено в справочнике InSales`);
      }
    }

    console.log(`[INFO] Прочитано дополнительных полей для отправки`);
    
    // Собираем ВСЕ изменения для отправки
    const changes = {
      collection: {
      // SEO данные - ПРАВИЛЬНЫЕ ПОЛЯ!
      html_title: sheet.getRange('B13').getValue() || null,
      meta_description: sheet.getRange('B14').getValue() || null,
      meta_keywords: sheet.getRange('B16').getValue() !== 'Описание категории (старая версия):' ? (sheet.getRange('B16').getValue() || null) : null,
        
        // Описание
        description: finalDescription,
        
        // Дополнительные поля через field_values_attributes
        field_values_attributes: fieldValues
      }
    };
    
    console.log('[INFO] Отправляем обновление категории в InSales');
    console.log('[INFO] Полей для обновления:', Object.keys(changes.collection).filter(k => changes.collection[k]).length);
    console.log('[INFO] Field values:', fieldValues.length);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Отправляем все данные в InSales...',
      '📤 Отправка',
      -1
    );
    
    const result = sendCategoryUpdateToInSalesAPI(categoryId, changes);
    
    if (result.success) {
      const productStats = getProductChangeStats(sheet);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Категория успешно обновлена!',
        '✅ Готово',
        5
      );
      
      let message = `✅ Все данные категории успешно обновлены в InSales!\n\n`;
      message += `📁 ID категории: ${categoryId}\n`;
      message += `📝 Обновлено основных полей: ${Object.keys(changes.collection).filter(k => changes.collection[k] !== null).length}\n`;
      message += `⚙️ Обновлено дополнительных полей: ${fieldValues.length}\n`;
      message += `🏷️ Верхних тегов: ${topTags.length}\n`;
      message += `🏷️ Нижних тегов: ${bottomTags.length}\n\n`;
      
      message += `📊 СТАТУС ТОВАРОВ В КАТЕГОРИИ:\n`;
      message += `• Всего товаров: ${productStats.total}\n`;
      message += `• В наличии: ${productStats.inStock}\n`;
      message += `• Нет в наличии: ${productStats.outOfStock}`;
      
      SpreadsheetApp.getUi().alert(
        'Готово',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      console.log('[INFO] ✅ Категория успешно обновлена');
      
    } else {
      throw new Error(result.error || 'Неизвестная ошибка');
    }
    
  } catch (error) {
    console.error('[ERROR] ❌ Ошибка отправки:', error.message);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка отправки',
      '❌ Ошибка',
      5
    );
    
    SpreadsheetApp.getUi().alert(
      'Ошибка отправки',
      `Не удалось отправить изменения в InSales:\n\n${error.message}\n\n` +
      `Возможные причины:\n` +
      `• Сервер InSales временно недоступен (503)\n` +
      `• Превышен лимит запросов\n` +
      `• Проблемы с интернет-соединением\n\n` +
      `Попробуйте через несколько минут.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Получает статистику изменений товаров
 */
function getProductChangeStats(sheet) {
  try {
    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const lastRow = sheet.getLastRow();
    
    if (lastRow < startRow) {
      return { total: 0, inStock: 0, outOfStock: 0 };
    }
    
    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();
    
    let total = 0;
    let inStock = 0;
    let outOfStock = 0;
    
    for (let i = 0; i < data.length; i++) {
      const id = data[i][1]; // ID товара
      if (!id || id.toString().trim() === '') break;
      
      total++;
      
      const inStockStatus = data[i][4]; // Колонка "В наличии"
      if (inStockStatus === 'Да') {
        inStock++;
      } else {
        outOfStock++;
      }
    }
    
    return { total, inStock, outOfStock };
    
  } catch (error) {
    console.error('❌ Ошибка подсчёта статистики товаров:', error);
    return { total: 0, inStock: 0, outOfStock: 0 };
  }
}

/**
 * Получает описание для отправки в InSales
 * Приоритет: C17 (новая версия) → B17 (старая версия)
 */
function getDescriptionForInSales(sheet) {
  // Проверяем новую версию в C17
  const newDescription = sheet.getRange('C17').getValue();
  if (newDescription && newDescription.toString().trim().length > 0) {
    console.log('[INFO] Используем новое описание из C17');
    return newDescription.toString().trim();
  }
  
  // Если C17 пустая, берем из B17
  const oldDescription = sheet.getRange('B17').getValue();
  if (oldDescription && oldDescription.toString().trim().length > 0) {
    console.log('[INFO] Используем старое описание из B17');
    return oldDescription.toString().trim();
  }
  
  console.log('[INFO] Описание отсутствует');
  return '';
}

/**
 * Читает ВЕРХНИЕ теги из дополнительного поля (HTML)
 * FALLBACK когда таблица пустая
 */
function getTopTagsFromSheet(sheet, productsCount) {
  try {
    const extraFieldsCount = countExtraFields(sheet, productsCount);
    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const productsStartRow = calculateSheetSections(sheet).productsStart;
    const topTagsStartRow = productsStartRow + productsCount + 5 + 2 + extraFieldsCount + 3 + 1 + 2 + 1;
    
    const data = sheet.getRange(topTagsStartRow, 1, 10, 2).getValues(); // Колонки A и B
    
    const tags = [];
    
    for (let i = 0; i < data.length; i++) {
      const text = data[i][0];
      const url = data[i][1];
      
      if (text && text.toString().trim() !== '' && 
          url && url.toString().trim() !== '') {
        tags.push({
          text: text.toString().trim(),
          url: url.toString().trim()
        });
      }
    }
    
    console.log(`🏷️ Найдено верхних тегов: ${tags.length}`);
    return tags;
    
  } catch (error) {
    console.error('❌ Ошибка чтения верхних тегов:', error);
    return [];
  }
}

/**
 * Читает НИЖНИЕ теги из дополнительного поля (HTML)
 * FALLBACK когда таблица пустая
 */
function getBottomTagsFromSheet(sheet, productsCount) {
  try {
    const extraFieldsCount = countExtraFields(sheet, productsCount);
    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const productsStartRow = calculateSheetSections(sheet).productsStart;
    const topTagsStartRow = productsStartRow + productsCount + 5 + 2 + extraFieldsCount + 3 + 1 + 2 + 1;
    const bottomTagsStartRow = topTagsStartRow + 10 + 3 + 1 + 2 + 1;
    
    const data = sheet.getRange(bottomTagsStartRow, 1, 10, 2).getValues(); // Колонки A и B
    
    const tags = [];
    
    for (let i = 0; i < data.length; i++) {
      const text = data[i][0];
      const url = data[i][1];
      
      if (text && text.toString().trim() !== '' && 
          url && url.toString().trim() !== '') {
        tags.push({
          text: text.toString().trim(),
          url: url.toString().trim()
        });
      }
    }
    
    console.log(`🏷️ Найдено нижних тегов: ${tags.length}`);
    return tags;
    
  } catch (error) {
    console.error('❌ Ошибка чтения нижних тегов:', error);
    return [];
  }
}

/**
 * Генерирует HTML блока ссылок из массива тегов
 */
function generateTagsBlockHTML(tags) {
  if (!tags || tags.length === 0) {
    return '';
  }
  
  let html = '<ul>\n';
  
  for (const tag of tags) {
    html += `  <li><a href="${tag.url}">${tag.text}</a></li>\n`;
  }
  
  html += '</ul>';
  
  return html;
}

/**
 * Отправляет обновление категории в InSales API
 * ИСПРАВЛЕНО: обновляем только стандартные поля без field_values
 */
function sendCategoryUpdateToInSalesAPI(categoryId, changes) {
  const maxRetries = 3;
  const retryDelay = 2000;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`[INFO] 🔄 Попытка ${attempt}/${maxRetries} отправки в InSales...`);
      
      const credentials = getInsalesCredentialsSync();
      if (!credentials) {
        throw new Error('Не удалось получить учетные данные InSales');
      }
      
      const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
      
      // ИСПРАВЛЕНИЕ: Убираем field_values_attributes из основного запроса
      const cleanChanges = {
        collection: {
          html_title: changes.collection.html_title,
          meta_description: changes.collection.meta_description,
          meta_keywords: changes.collection.meta_keywords,
          description: changes.collection.description
        }
      };
      
      console.log('[INFO] 📡 URL запроса:', url);
      console.log('[INFO] 📦 Данные для отправки:', JSON.stringify(cleanChanges, null, 2).substring(0, 500));
      
      const options = {
        method: 'PUT',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'User-Agent': 'GoogleAppsScript-InSalesManager/2.0'
        },
        payload: JSON.stringify(cleanChanges),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      console.log('[INFO] 📥 HTTP статус:', statusCode);
      console.log('[INFO] 📥 Ответ сервера:', responseText.substring(0, 500));
      
      // Успешные коды
      if (statusCode === 200 || statusCode === 204) {
        console.log('[INFO] ✅ Категория успешно обновлена!');
        
        // НОВОЕ: Теперь обновляем H1 отдельным запросом
        if (changes.collection.field_values_attributes && 
            changes.collection.field_values_attributes.length > 0) {
          updateCategoryFieldValues(categoryId, changes.collection.field_values_attributes);
        }
        
        return { success: true };
      }
      
      // Ошибка 503 - повторяем попытку
      if (statusCode === 503) {
        console.warn(`[WARN] ⚠️ Попытка ${attempt}: Сервер InSales недоступен (503)`);
        
        if (attempt < maxRetries) {
          console.log(`[INFO] ⏳ Ожидание ${retryDelay / 1000} сек перед следующей попыткой...`);
          Utilities.sleep(retryDelay);
          continue;
        } else {
          throw new Error(`Сервер InSales недоступен после ${maxRetries} попыток (503 Service Unavailable)`);
        }
      }
      
      // Другие ошибки
      if (statusCode >= 400) {
        let errorMessage = `HTTP ${statusCode}`;
        
        try {
          const errorData = JSON.parse(responseText);
          if (errorData.errors) {
            errorMessage += ': ' + JSON.stringify(errorData.errors);
          }
        } catch (e) {
          errorMessage += ': ' + responseText.substring(0, 200);
        }
        
        throw new Error(errorMessage);
      }
      
      throw new Error(`Неожиданный код ответа: ${statusCode}`);
      
    } catch (error) {
      console.error(`[ERROR] ❌ Попытка ${attempt} неудачна:`, error.message);
      
      if (attempt === maxRetries) {
        return { success: false, error: error.message };
      }
      
      if (!error.message.includes('503')) {
        return { success: false, error: error.message };
      }
      
      Utilities.sleep(retryDelay);
    }
  }
  
  return { success: false, error: 'Превышено количество попыток' };
}

/**
 * ИСПРАВЛЕНО: Обновляет field_values через правильный endpoint
 */
function updateCategoryFieldValues(categoryId, fieldValuesArray) {
  try {
    console.log(`[INFO] 🔧 Обновляем дополнительные поля категории ${categoryId}`);
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      console.error('[ERROR] Не удалось получить учетные данные');
      return false;
    }
    
    // ИСПРАВЛЕНИЕ: Обновляем field_values через основной endpoint категории
    // но отдельным запросом для каждого поля
    
    let successCount = 0;
    
    for (const fieldValue of fieldValuesArray) {
      try {
        const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
        
        // Формируем payload только с одним field_value
        const payload = {
          collection: {
            field_values_attributes: [fieldValue]
          }
        };
        
        console.log(`[INFO] 🔧 Обновляем поле:`, JSON.stringify(fieldValue));
        
        const options = {
          method: 'PUT',
          headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(url, options);
        const statusCode = response.getResponseCode();
        
        if (statusCode === 200 || statusCode === 204) {
          console.log(`[INFO] ✅ Поле обновлено успешно`);
          successCount++;
        } else {
          const responseText = response.getContentText();
          console.error(`[ERROR] ❌ Ошибка обновления поля: ${statusCode}`);
          console.error(`[ERROR] Ответ:`, responseText.substring(0, 200));
        }
        
        Utilities.sleep(500); // Задержка между запросами
        
      } catch (error) {
        console.error('[ERROR] Ошибка обновления field_value:', error.message);
      }
    }
    
    console.log(`[INFO] 🎯 Обновлено дополнительных полей: ${successCount}/${fieldValuesArray.length}`);
    return successCount > 0;
    
  } catch (error) {
    console.error('[ERROR] Ошибка обновления field_values:', error.message);
    return false;
  }
}

/**
 * Вспомогательная функция
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
    console.error('❌ Ошибка получения учётных данных InSales:', error);
    return null;
  }
}

/**
 * Находит ID существующего field_value по названию поля
 */
function findFieldValueId(categoryId, fieldName) {
  try {
    // Загружаем категорию из InSales
    const categoryData = loadFullCategoryData(categoryId);
    
    if (!categoryData.field_values || !Array.isArray(categoryData.field_values)) {
      return null;
    }
    
    // Получаем collection_field_id для этого поля
    const fieldId = getFieldIdByName(fieldName);
    if (!fieldId) {
      return null;
    }
    
    // Ищем field_value с таким collection_field_id
    const fieldValue = categoryData.field_values.find(fv => 
      fv.collection_field_id === fieldId
    );
    
    return fieldValue ? fieldValue.id : null;
    
  } catch (error) {
    console.error(`❌ Ошибка поиска field_value ID для "${fieldName}":`, error);
    return null;
  }
}
/**
 * ========================================
 * РАБОТА С ПЛИТКАМИ ТЕГОВ ИЗ ТАБЛИЦ
 * ========================================
 */

/**
 * Читает верхнюю плитку тегов из таблицы (строки 76+)
 * Приоритет: таблица → существующее поле
 */
function getTopTagsFromTable(sheet, productsCount) {
  try {
    const extraFieldsCount = countExtraFields(sheet, productsCount);

    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const productsStartRow = calculateSheetSections(sheet).productsStart;
    const topTagsStartRow = productsStartRow + productsCount + 5 + 2 + extraFieldsCount + 3 + 1 + 2 + 1;
    
    console.log(`📋 Читаем верхнюю плитку тегов из таблицы (строка ${topTagsStartRow})...`);
    
    const data = sheet.getRange(topTagsStartRow, 1, 10, 2).getValues(); // ИСПРАВЛЕНО: Только A-B
    
    const tags = [];
    
    for (let i = 0; i < data.length; i++) {
      const text = data[i][0];  // Колонка A - Текст ссылки
      const url = data[i][1];   // Колонка B - URL
      
      if (text && text.toString().trim() !== '' && 
          url && url.toString().trim() !== '') {
        tags.push({
          text: text.toString().trim(),
          url: url.toString().trim()
        });
      }
    }
    
    console.log(`✅ Прочитано верхних тегов из таблицы: ${tags.length}`);
    return tags;
    
  } catch (error) {
    console.error('❌ Ошибка чтения верхней плитки из таблицы:', error);
    return [];
  }
}

/**
 * Читает нижнюю плитку тегов из таблицы
 * Приоритет: таблица → существующее поле
 */
function getBottomTagsFromTable(sheet, productsCount) {
  try {
    const extraFieldsCount = countExtraFields(sheet, productsCount);

    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const productsStartRow = calculateSheetSections(sheet).productsStart;
    const topTagsStartRow = productsStartRow + productsCount + 5 + 2 + extraFieldsCount + 3 + 1 + 2 + 1;
    const bottomTagsStartRow = topTagsStartRow + 10 + 3 + 1 + 2 + 1;
    
    console.log(`📋 Читаем нижнюю плитку тегов из таблицы (строка ${bottomTagsStartRow})...`);
    
    const data = sheet.getRange(bottomTagsStartRow, 1, 10, 2).getValues(); // ИСПРАВЛЕНО: Только A-B
    
    const tags = [];
    
    for (let i = 0; i < data.length; i++) {
      const text = data[i][0];
      const url = data[i][1];
      
      if (text && text.toString().trim() !== '' && 
          url && url.toString().trim() !== '') {
        tags.push({
          text: text.toString().trim(),
          url: url.toString().trim()
        });
      }
    }
    
    console.log(`✅ Прочитано нижних тегов из таблицы: ${tags.length}`);
    return tags;
    
  } catch (error) {
    console.error('❌ Ошибка чтения нижней плитки из таблицы:', error);
    return [];
  }
}

/**
 * Подсчитывает реальное количество дополнительных полей в листе
 */
function countExtraFields(sheet, productsCount) {
  try {
    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const productsStartRow = calculateSheetSections(sheet).productsStart;
    const extraFieldsStartRow = productsStartRow + productsCount + 5 + 2; // Строка начала данных доп.полей
    
    let count = 0;
    for (let i = 0; i < 50; i++) {
      const fieldName = sheet.getRange(extraFieldsStartRow + i, 1).getValue();
      if (!fieldName || fieldName.toString().trim() === '') {
        break;
      }
      count++;
    }
    
    return count;
    
  } catch (error) {
    console.error('❌ Ошибка подсчета дополнительных полей:', error);
    return 20; // Запасное значение
  }
}