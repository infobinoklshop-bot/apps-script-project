/**
 * ========================================
 * МОДУЛЬ: ОТПРАВКА В INSALES
 * ========================================
 * 
 * Функции для отправки изменений категорий в InSales
 * с автоматической фиксацией изменений для отслеживания позиций
 */

/**
 * Обновляет категорию в InSales через API
 */
function updateCategoryInInSales(categoryId, updateData) {
  const context = `Обновление категории ${categoryId} в InSales`;

  try {
    logInfo('📤 Отправляем обновление категории в InSales API', null, context);

    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }

    const url = `${credentials.baseUrl}${CATEGORY_API_ENDPOINTS.UPDATE_COLLECTION.replace('{id}', categoryId)}`;

    // Формируем payload для InSales API
    const payload = {
      collection: {
        title: updateData.title,
        seo_title: updateData.seo_title,
        seo_description: updateData.seo_description,
        h1_title: updateData.h1_title,
        meta_keywords: updateData.meta_keywords,
        description: updateData.description
      }
    };

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

    logInfo('📡 Отправляем PUT запрос к InSales API...');

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200 || responseCode === 204) {
      logInfo('✅ Категория успешно обновлена в InSales', null, context);
      return true;
    } else {
      const responseText = response.getContentText();
      logError(`❌ Ошибка API InSales: ${responseCode}`, responseText, context);

      // Показываем пользователю подробности ошибки
      SpreadsheetApp.getUi().alert(
        'Ошибка InSales API',
        `Код ошибки: ${responseCode}\n\nОтвет API:\n${responseText.substring(0, 500)}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );

      return false;
    }

  } catch (error) {
    logError('❌ Ошибка обновления категории в InSales', error, context);
    return false;
  }
}

/**
 * Обновляет статус категории в главном листе
 */
function updateCategoryStatusInMainList(categoryId, seoStatus) {
  const context = "Обновление статуса в главном листе";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!sheet) {
      logWarning('⚠️ Главный лист категорий не найден');
      return;
    }

    const data = sheet.getDataRange().getValues();

    // Ищем категорию по ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][MAIN_LIST_COLUMNS.CATEGORY_ID - 1] == categoryId) {
        // Обновляем SEO статус (колонка J)
        sheet.getRange(i + 1, MAIN_LIST_COLUMNS.SEO_STATUS).setValue(seoStatus);

        // Обновляем AI статус (колонка K)
        sheet.getRange(i + 1, MAIN_LIST_COLUMNS.AI_STATUS).setValue('✅ Обработано AI');

        // Обновляем дату (колонка L)
        sheet.getRange(i + 1, MAIN_LIST_COLUMNS.LAST_UPDATED).setValue(new Date());

        // Подсвечиваем строку зеленым
        sheet.getRange(i + 1, 1, 1, 13).setBackground('#e8f5e9');

        logInfo(`✅ Статус категории ${categoryId} обновлен в главном листе`, null, context);
        break;
      }
    }

  } catch (error) {
    logError('❌ Ошибка обновления статуса категории', error, context);
  }
}

/**
 * НОВАЯ ФУНКЦИЯ: Автоматически фиксирует изменения перед отправкой
 * Вызывается из sendCategoryChangesToInSales()
 */
function autoLogChangeBeforeSending(categoryData, updateData) {
  const context = "Автоматическая фиксация изменений";

  try {
    const changes = [];

    // Собираем список изменений
    if (updateData.seo_title) {
      changes.push(`SEO Title: "${updateData.seo_title}"`);
    }

    if (updateData.seo_description) {
      const shortDesc = updateData.seo_description.length > 50 ?
        updateData.seo_description.substring(0, 50) + '...' :
        updateData.seo_description;
      changes.push(`Meta Description: "${shortDesc}"`);
    }

    if (updateData.h1_title) {
      changes.push(`H1: "${updateData.h1_title}"`);
    }

    if (updateData.meta_keywords) {
      changes.push(`Keywords: "${updateData.meta_keywords}"`);
    }

    if (updateData.description) {
      changes.push('Обновлено полное описание категории');
    }

    if (changes.length > 0) {
      const changeLog = `📤 Отправка в InSales: ${changes.join('; ')}`;

      // Сохраняем в историю изменений
      logPageChange(categoryData.id, changeLog);

      logInfo(`✅ Автоматически зафиксировано изменений: ${changes.length}`, null, context);
    } else {
      logWarning('⚠️ Нет изменений для фиксации');
    }

  } catch (error) {
    logError('❌ Ошибка автоматической фиксации изменений', error, context);
    // Не прерываем процесс, просто логируем ошибку
  }
}

/**
 * Обновляет только SEO данные категории (быстрое обновление)
 */
function updateCategorySEOOnly() {
  const context = "Быстрое обновление SEO";

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

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Быстрое обновление SEO',
      'Отправить только SEO теги (title, description, H1) без описания категории?\n\nЭто быстрее, чем полное обновление.',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Обновляем SEO теги...',
      '⏳ Обновление',
      -1
    );

    const sheet = SpreadsheetApp.getActiveSheet();

    const updateData = {
      seo_title: sheet.getRange(DETAIL_SHEET_SECTIONS.SEO_TITLE_CELL).getValue(),
      seo_description: sheet.getRange(DETAIL_SHEET_SECTIONS.SEO_DESCRIPTION_CELL).getValue(),
      h1_title: sheet.getRange(DETAIL_SHEET_SECTIONS.SEO_H1_CELL).getValue(),
      meta_keywords: sheet.getRange(DETAIL_SHEET_SECTIONS.SEO_KEYWORDS_CELL).getValue()
    };

    // Фиксируем изменения
    autoLogChangeBeforeSending(categoryData, updateData);

    // Отправляем
    const success = updateCategoryInInSales(categoryData.id, updateData);

    if (success) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'SEO теги обновлены!',
        '✅ Готово',
        5
      );

      updateCategoryStatusInMainList(categoryData.id, '✅ SEO обновлено');
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Ошибка обновления',
        '❌ Ошибка',
        10
      );
    }

  } catch (error) {
    logError('❌ Ошибка быстрого обновления SEO', error, context);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Обновляет только описание категории
 */
function updateCategoryDescriptionOnly() {
  const context = "Обновление описания";

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

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Обновление описания',
      'Отправить только описание категории?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Обновляем описание...',
      '⏳ Обновление',
      -1
    );

    const sheet = SpreadsheetApp.getActiveSheet();

    const updateData = {
      description: sheet.getRange('B17').getValue()
    };

    // Фиксируем изменения
    autoLogChangeBeforeSending(categoryData, updateData);

    // Отправляем
    const success = updateCategoryInInSales(categoryData.id, updateData);

    if (success) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Описание обновлено!',
        '✅ Готово',
        5
      );

      logPageChange(categoryData.id, 'Обновлено описание категории через InSales API');
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Ошибка обновления',
        '❌ Ошибка',
        10
      );
    }

  } catch (error) {
    logError('❌ Ошибка обновления описания', error, context);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Массовое обновление категорий из главного листа
 */
function bulkUpdateSelectedCategories() {
  const context = "Массовое обновление";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!sheet) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        'Главный лист категорий не найден',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const data = sheet.getDataRange().getValues();
    const selectedCategories = [];

    // Собираем отмеченные категории
    for (let i = 2; i < data.length; i++) { // Начинаем с 3-й строки (данные)
      const checkbox = data[i][0];
      if (checkbox === true) {
        selectedCategories.push({
          row: i + 1,
          id: data[i][MAIN_LIST_COLUMNS.CATEGORY_ID - 1],
          title: data[i][MAIN_LIST_COLUMNS.TITLE - 1]
        });
      }
    }

    if (selectedCategories.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Нет выбранных категорий',
        'Отметьте чекбоксами категории, которые нужно обновить',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Массовое обновление',
      `Обновить ${selectedCategories.length} категорий в InSales?\n\n` +
      'Будут отправлены данные из детальных листов каждой категории.',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Обрабатываем категории
    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < selectedCategories.length; i++) {
      const cat = selectedCategories[i];

      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Обновление ${i + 1}/${selectedCategories.length}: ${cat.title}`,
        '⏳ Обработка',
        -1
      );

      try {
        // Открываем детальный лист категории
        const detailSheetName = `${CATEGORY_SHEETS.DETAIL_PREFIX}${cat.title}`;
        const detailSheet = ss.getSheetByName(detailSheetName);

        if (!detailSheet) {
          logWarning(`⚠️ Детальный лист не найден для ${cat.title}`);
          errorCount++;
          continue;
        }

        // Собираем данные
        const updateData = {
          seo_title: detailSheet.getRange(DETAIL_SHEET_SECTIONS.SEO_TITLE_CELL).getValue(),
          seo_description: detailSheet.getRange(DETAIL_SHEET_SECTIONS.SEO_DESCRIPTION_CELL).getValue(),
          h1_title: detailSheet.getRange(DETAIL_SHEET_SECTIONS.SEO_H1_CELL).getValue(),
          meta_keywords: detailSheet.getRange(DETAIL_SHEET_SECTIONS.SEO_KEYWORDS_CELL).getValue(),
          description: detailSheet.getRange('B17').getValue()
        };

        // Отправляем
        const success = updateCategoryInInSales(cat.id, updateData);

        if (success) {
          successCount++;
          sheet.getRange(cat.row, MAIN_LIST_COLUMNS.SEO_STATUS).setValue('✅ Готово');
        } else {
          errorCount++;
          sheet.getRange(cat.row, MAIN_LIST_COLUMNS.SEO_STATUS).setValue('❌ Ошибка');
        }

        Utilities.sleep(500); // Задержка между запросами

      } catch (catError) {
        logError(`❌ Ошибка обновления категории ${cat.title}`, catError);
        errorCount++;
      }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Готово! Успешно: ${successCount}, Ошибок: ${errorCount}`,
      '✅ Массовое обновление завершено',
      10
    );

    logInfo(`✅ Массовое обновление: успех ${successCount}, ошибки ${errorCount}`, null, context);

  } catch (error) {
    logError('❌ Ошибка массового обновления', error, context);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Получает позиции товаров в категории через API collects
 */
function getProductPositions(categoryId) {
  const context = `Получение позиций товаров для категории ${categoryId}`;

  try {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }

    const positions = {};
    let page = 1;
    const perPage = 250; // Максимум для collects

    while (true) {
      const url = `${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&per_page=${perPage}&page=${page}`;

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
        logWarning(`⚠️ Ошибка получения позиций (стр ${page}): ${response.getResponseCode()}`);
        break;
      }

      const collects = JSON.parse(response.getContentText());

      if (!collects || collects.length === 0) {
        break;
      }

      collects.forEach(item => {
        // item.product_id -> item.position
        positions[item.product_id] = item.position || 0;
      });

      if (collects.length < perPage) {
        break;
      }

      page++;
      Utilities.sleep(200);
    }

    logInfo(`✅ Загружены позиции для ${Object.keys(positions).length} товаров`, null, context);
    return positions;

  } catch (error) {
    logError('❌ Ошибка получения позиций товаров', error, context);
    return {}; // Возвращаем пустой объект, чтобы не ломать логику
  }
}

/**
 * Получает дополнительные поля категории (metafields)
 * @param {number|string} categoryId 
 * @returns {Array} Список полей [{id, key, value, ...}]
 */
function getCategoryAdditionalFields(categoryId) {
  const context = `Получение доп. полей категории ${categoryId}`;
  try {
    const credentials = getInsalesCredentialsSync();
    // В InSales доп. поля категорий обычно лежат в /admin/collections/{id}/fields_values.json
    // Но API может отличаться, проверим стандартный endpoint для полей
    const url = `${credentials.baseUrl}/admin/collections/${categoryId}/fields_values.json`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return data || []; // Обычно возвращает массив объектов
    }

    // Если endpoint fields_values недоступен, попробуем через саму категорию (иногда поля приходят в теле)
    // Но для записи точно нужен отдельный endpoint или вложенный объект
    logWarning(`⚠️ Не удалось получить поля через fields_values (Code ${response.getResponseCode()})`);
    return [];

  } catch (error) {
    logError('❌ Ошибка получения доп. полей', error, context);
    return [];
  }
}

/**
 * Обновляет (или создает) дополнительное поле категории
 * @param {number|string} categoryId 
 * @param {string} handle (ключ поля)
 * @param {string} value 
 */
function updateCategoryAdditionalField(categoryId, handle, value) {
  const context = `Обновление доп. поля ${handle} для ${categoryId}`;
  try {
    const credentials = getInsalesCredentialsSync();

    // 1. Сначала нужно найти ID поля по handle, если мы обновляем значение
    // Но в InSales API для fields_values обычно нужно знать field_id (ID самого свойства), а не только значение
    // Это усложняет задачу. Нам нужно знать ID свойства "smart_rules".

    // ПРЕДПОЛОЖЕНИЕ: Мы будем искать поле по handle (ключу) среди всех свойств
    // Если "smart_rules" еще нет, мы должны его знать ID.
    // Упрощение: Попробуем использовать API v2 или будем считать, что пользователь уже создал поле и мы ищем его ID.

    // ШАГ 1: Ищем ID свойства (Fields)
    // Это глобальный список полей для коллекций
    // Кэшируем или ищем каждый раз (пока ищем)
    const fieldsUrl = `${credentials.baseUrl}/admin/fields.json`;
    // Примечание: fields.json возвращает поля для всего. Нужно фильтровать по type?
    // Обычно для коллекций это может быть отдельный список.

    // Но InSales позволяет передавать fields_values при обновлении коллекции
    // Пробуем через updateCategory (PUT /collections/{id}.json)

    const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;

    const payload = {
      collection: {
        // В InSales API передача доп полей часто выглядит так:
        // fields_values_attributes: [{ field_id: 123, value: "..." }]
        // Нам нужно знать ID поля!
      }
    };

    // Чтобы не усложнять сейчас поиск ID поля (это лишний запрос), 
    // Добавим функцию поиска ID поля по handle.
    const fieldId = getFieldIdByHandle(handle);
    if (!fieldId) {
      throw new Error(`Не найдено доп. поле с handle "${handle}" в InSales`);
    }

    // Готовим payload для обновления значения
    payload.collection.fields_values_attributes = [
      {
        field_id: fieldId,
        value: value
      }
    ];

    const response = UrlFetchApp.fetch(url, {
      method: 'PUT',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      logInfo(`✅ Доп. поле ${handle} обновлено`, null, context);
      return true;
    } else {
      logError(`❌ Ошибка обновления поля: ${response.getResponseCode()}`, response.getContentText(), context);
      return false;
    }

  } catch (error) {
    logError('❌ Ошибка updateCategoryAdditionalField', error, context);
    return false;
  }
}

/**
 * Ищет ID дополнительного поля по его handle (системному имени)
 * Кэширует результат в ScriptProperties чтобы не дёргать API постоянно
 */
function getFieldIdByHandle(handle) {
  const cacheKey = `INSALES_FIELD_ID_${handle}`;
  const cachedId = PropertiesService.getScriptProperties().getProperty(cacheKey);

  if (cachedId) return parseInt(cachedId);

  // Ищем в API
  const credentials = getInsalesCredentialsSync();
  // Получаем список всех полей. 
  // В InSales поля для коллекций могут быть смешаны.
  const url = `${credentials.baseUrl}/admin/fields.json`;

  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
  });

  const fields = JSON.parse(response.getContentText());
  // Ищем поле, у которого desteny_type (или похожее) связано с Collections, и handle совпадает
  // В документации InSales поля имеют `handle`.

  const targetField = fields.find(f => f.handle === handle);

  if (targetField) {
    PropertiesService.getScriptProperties().setProperty(cacheKey, targetField.id.toString());
    return targetField.id;
  }

  return null;
}
