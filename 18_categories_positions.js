/**
 * ========================================
 * МОДУЛЬ: ОТСЛЕЖИВАНИЕ ПОЗИЦИЙ В ПОИСКЕ
 * ========================================
 * 
 * Функционал:
 * - Снятие позиций по маркерным запросам
 * - История изменений на странице
 * - Связь изменений с динамикой позиций
 */

/**
 * Создает лист истории позиций
 */
function setupPositionsHistorySheet(sheet) {
  sheet.clear();
  
  const headers = [
    'Дата проверки',
    'Категория ID',
    'Категория',
    'Маркерный запрос',
    'Позиция Yandex',
    'Позиция Google',
    'URL в выдаче',
    'Изменение',  // +5, -3, без изменений
    'Что изменено на странице',
    'Дата изменения страницы',
    'Комментарий',
    'Статус'  // Отслеживается, Приостановлено
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ff6f00')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 120);  // Дата
  sheet.setColumnWidth(2, 100);  // ID
  sheet.setColumnWidth(3, 200);  // Категория
  sheet.setColumnWidth(4, 250);  // Запрос
  sheet.setColumnWidth(5, 100);  // Yandex
  sheet.setColumnWidth(6, 100);  // Google
  sheet.setColumnWidth(7, 300);  // URL
  sheet.setColumnWidth(8, 100);  // Изменение
  sheet.setColumnWidth(9, 400);  // Что изменено
  sheet.setColumnWidth(10, 120); // Дата изменения
  sheet.setColumnWidth(11, 300); // Комментарий
  sheet.setColumnWidth(12, 120); // Статус
  
  sheet.setFrozenRows(1);
  
  logInfo('✅ Лист истории позиций настроен');
}

/**
 * Настраивает маркерный запрос для категории
 */
function setupMarkerQuery() {
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
    
    // Запрашиваем маркерный запрос
    const response = ui.prompt(
      'Настройка отслеживания позиций',
      `Введите маркерный запрос для категории "${categoryData.title}":\n\n` +
      'Это должен быть ОСНОВНОЙ запрос, по которому вы хотите ранжироваться.\n' +
      'Например: "купить бинокль 10x25"',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const markerQuery = response.getResponseText().trim();
    
    if (!markerQuery) {
      ui.alert('Запрос не может быть пустым');
      return;
    }
    
    // Сохраняем настройку
    saveMarkerQueryForCategory(categoryData.id, markerQuery);
    
    // Снимаем первоначальные позиции
    const confirmed = ui.alert(
      'Снять текущие позиции?',
      'Снять текущие позиции для фиксации начального состояния?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmed === ui.Button.YES) {
      checkPositionsForCategory(categoryData, markerQuery, 'Начальная проверка');
    }
    
  } catch (error) {
    logError('❌ Ошибка настройки маркерного запроса', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Сохраняет маркерный запрос для категории
 */
function saveMarkerQueryForCategory(categoryId, query) {
  const context = "Сохранение маркерного запроса";
  
  try {
    // Сохраняем в Script Properties
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`marker_query_${categoryId}`, query);
    
    // Обновляем в детальном листе (добавим поле)
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Проверяем, есть ли уже ячейка для маркерного запроса
    const markerCell = sheet.getRange('B6');
    if (markerCell.getValue() === '' || sheet.getRange('A6').getValue() !== 'Маркерный запрос:') {
      sheet.getRange('A6').setValue('Маркерный запрос:').setFontWeight('bold');
    }
    markerCell.setValue(query).setBackground('#fff3e0');
    
    logInfo(`✅ Маркерный запрос сохранен для категории ${categoryId}`, null, context);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Маркерный запрос "${query}" сохранен`,
      '✅ Готово',
      3
    );
    
  } catch (error) {
    logError('❌ Ошибка сохранения маркерного запроса', error, context);
    throw error;
  }
}

/**
 * Получает маркерный запрос для категории
 */
function getMarkerQueryForCategory(categoryId) {
  try {
    const props = PropertiesService.getScriptProperties();
    return props.getProperty(`marker_query_${categoryId}`) || null;
  } catch (error) {
    logError('❌ Ошибка получения маркерного запроса', error);
    return null;
  }
}

/**
 * Проверяет позиции категории в поиске
 */
function checkPositionsForCategory(categoryData, markerQuery, comment = '') {
  const context = `Проверка позиций для ${categoryData.title}`;
  
  try {
    logInfo('🔍 Начинаем проверку позиций', null, context);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Проверяем позиции по запросу "${markerQuery}"...`,
      '⏳ Проверка',
      -1
    );
    
    // Проверяем позиции в Яндекс и Google
    const yandexPosition = checkPositionInYandex(markerQuery, categoryData.url);
    const googlePosition = checkPositionInGoogle(markerQuery, categoryData.url);
    
    // Сохраняем результат
    savePositionCheck(categoryData, markerQuery, yandexPosition, googlePosition, comment);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Позиции: Яндекс ${yandexPosition || 'нет'}, Google ${googlePosition || 'нет'}`,
      '✅ Готово',
      8
    );
    
    logInfo(`✅ Позиции проверены: Y:${yandexPosition}, G:${googlePosition}`, null, context);
    
    return {
      yandex: yandexPosition,
      google: googlePosition
    };
    
  } catch (error) {
    logError('❌ Ошибка проверки позиций', error, context);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка проверки позиций: ' + error.message,
      '❌ Ошибка',
      10
    );
    throw error;
  }
}

/**
 * Проверяет позицию в Яндексе
 * ВАЖНО: Для реальной проверки нужен API Яндекс.XML или сервис проверки позиций
 */
function checkPositionInYandex(query, targetUrl) {
  const context = "Проверка Яндекс";
  
  try {
    // ВАРИАНТ 1: Использовать API Яндекс.XML (платный)
    // ВАРИАНТ 2: Использовать сервис проверки позиций (Serpstat, SE Ranking)
    // ВАРИАНТ 3: Ручная проверка через форму
    
    // Пример заглушки - замените на реальный API
    logWarning('⚠️ Проверка позиций Яндекс не реализована - добавьте API');
    
    // Показываем диалог для ручного ввода
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Ручной ввод позиции Яндекс',
      `Проверьте позицию вручную в Яндексе:\n\nЗапрос: "${query}"\nURL: ${targetUrl}\n\nВведите позицию (или 0 если нет в ТОП-100):`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const position = parseInt(response.getResponseText());
      return isNaN(position) ? null : (position === 0 ? null : position);
    }
    
    return null;
    
  } catch (error) {
    logError('❌ Ошибка проверки Яндекс', error, context);
    return null;
  }
}

/**
 * Проверяет позицию в Google
 * ВАЖНО: Google не любит автоматические запросы, используйте API
 */
function checkPositionInGoogle(query, targetUrl) {
  const context = "Проверка Google";
  
  try {
    // Аналогично Яндексу - нужен API или сервис
    logWarning('⚠️ Проверка позиций Google не реализована - добавьте API');
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Ручной ввод позиции Google',
      `Проверьте позицию вручную в Google:\n\nЗапрос: "${query}"\nURL: ${targetUrl}\n\nВведите позицию (или 0 если нет в ТОП-100):`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      const position = parseInt(response.getResponseText());
      return isNaN(position) ? null : (position === 0 ? null : position);
    }
    
    return null;
    
  } catch (error) {
    logError('❌ Ошибка проверки Google', error, context);
    return null;
  }
}

/**
 * Сохраняет результат проверки позиций
 */
function savePositionCheck(categoryData, query, yandexPos, googlePos, comment) {
  const context = "Сохранение проверки позиций";
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CATEGORY_SHEETS.POSITIONS_HISTORY);
    
    if (!sheet) {
      sheet = ss.insertSheet(CATEGORY_SHEETS.POSITIONS_HISTORY);
      setupPositionsHistorySheet(sheet);
    }
    
    // Получаем предыдущую проверку для расчета изменений
    const previousCheck = getLastPositionCheck(categoryData.id, query);
    
    // Вычисляем изменение
    let yandexChange = '';
    let googleChange = '';
    
    if (previousCheck) {
      if (previousCheck.yandex && yandexPos) {
        const diff = previousCheck.yandex - yandexPos;
        yandexChange = diff > 0 ? `+${diff}` : (diff < 0 ? `${diff}` : '=');
      }
      
      if (previousCheck.google && googlePos) {
        const diff = previousCheck.google - googlePos;
        googleChange = diff > 0 ? `+${diff}` : (diff < 0 ? `${diff}` : '=');
      }
    }
    
    const change = [yandexChange, googleChange].filter(Boolean).join(' | ') || 'Первая проверка';
    
    // Получаем информацию о последних изменениях на странице
    const pageChanges = getRecentPageChanges(categoryData.id);
    const lastChangeDate = pageChanges.length > 0 ? pageChanges[0].date : '';
    const changesText = pageChanges.length > 0 ? 
      pageChanges.slice(0, 3).map(ch => `${ch.date}: ${ch.description}`).join('\n') : 
      'Нет изменений';
    
    // Новая строка данных
    const newRow = [
      new Date(),                    // Дата проверки
      categoryData.id,               // ID категории
      categoryData.title,            // Название
      query,                         // Маркерный запрос
      yandexPos || 'Нет в ТОП-100', // Позиция Яндекс
      googlePos || 'Нет в ТОП-100', // Позиция Google
      categoryData.url,              // URL
      change,                        // Изменение
      changesText,                   // Что изменено
      lastChangeDate,                // Дата изменения
      comment || '',                 // Комментарий
      'Отслеживается'               // Статус
    ];
    
    // Добавляем строку
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    
    // Форматирование
    sheet.getRange(lastRow + 1, 1).setNumberFormat('dd.mm.yyyy hh:mm');
    sheet.getRange(lastRow + 1, 10).setNumberFormat('dd.mm.yyyy hh:mm');
    
    // Цветовое кодирование изменений
    const changeCell = sheet.getRange(lastRow + 1, 8);
    if (change.includes('+')) {
      changeCell.setBackground('#c8e6c9'); // Зеленый - рост
    } else if (change.includes('-')) {
      changeCell.setBackground('#ffcdd2'); // Красный - падение
    } else if (change === '=') {
      changeCell.setBackground('#fff9c4'); // Желтый - без изменений
    }
    
    logInfo('✅ Проверка позиций сохранена', null, context);
    
  } catch (error) {
    logError('❌ Ошибка сохранения проверки', error, context);
    throw error;
  }
}

/**
 * Получает последнюю проверку позиций
 */
function getLastPositionCheck(categoryId, query) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.POSITIONS_HISTORY);
    
    if (!sheet) {
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Ищем последнюю проверку для этой категории и запроса
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] == categoryId && data[i][3] === query) {
        return {
          date: data[i][0],
          yandex: typeof data[i][4] === 'number' ? data[i][4] : null,
          google: typeof data[i][5] === 'number' ? data[i][5] : null
        };
      }
    }
    
    return null;
    
  } catch (error) {
    logError('❌ Ошибка получения последней проверки', error);
    return null;
  }
}

/**
 * Фиксирует изменение на странице категории
 */
function logPageChange(categoryId, changeDescription) {
  const context = "Фиксация изменения страницы";
  
  try {
    // Сохраняем в специальное хранилище изменений
    const props = PropertiesService.getScriptProperties();
    const changesKey = `page_changes_${categoryId}`;
    
    let changes = [];
    const existingChanges = props.getProperty(changesKey);
    
    if (existingChanges) {
      try {
        changes = JSON.parse(existingChanges);
      } catch (e) {
        changes = [];
      }
    }
    
    // Добавляем новое изменение
    changes.unshift({
      date: new Date().toISOString(),
      description: changeDescription,
      timestamp: Date.now()
    });
    
    // Храним последние 50 изменений
    if (changes.length > 50) {
      changes = changes.slice(0, 50);
    }
    
    props.setProperty(changesKey, JSON.stringify(changes));
    
    logInfo(`✅ Изменение зафиксировано для категории ${categoryId}`, null, context);
    
    // Также добавляем запись в детальный лист
    addChangeLogToDetailSheet(changeDescription);
    
  } catch (error) {
    logError('❌ Ошибка фиксации изменения', error, context);
  }
}

/**
 * Получает недавние изменения страницы
 */
function getRecentPageChanges(categoryId, limit = 10) {
  try {
    const props = PropertiesService.getScriptProperties();
    const changesKey = `page_changes_${categoryId}`;
    const changesData = props.getProperty(changesKey);
    
    if (!changesData) {
      return [];
    }
    
    const changes = JSON.parse(changesData);
    
    return changes.slice(0, limit).map(ch => ({
      date: new Date(ch.timestamp).toLocaleDateString('ru-RU'),
      description: ch.description,
      timestamp: ch.timestamp
    }));
    
  } catch (error) {
    logError('❌ Ошибка получения истории изменений', error);
    return [];
  }
}

/**
 * Добавляет запись в лог изменений детального листа
 */
function addChangeLogToDetailSheet(changeDescription) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    if (!sheetName.startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
      return; // Не детальный лист
    }
    
    // Ищем секцию лога изменений (создаем если нет)
    const changeLogStart = 400; // Строка начала лога
    
    // Проверяем, есть ли заголовок
    const headerCell = sheet.getRange(changeLogStart, 1);
    if (headerCell.getValue() !== '📝 ИСТОРИЯ ИЗМЕНЕНИЙ') {
      // Создаем секцию
      headerCell.setValue('📝 ИСТОРИЯ ИЗМЕНЕНИЙ')
                .setFontWeight('bold')
                .setFontSize(14)
                .setBackground('#607d8b')
                .setFontColor('#ffffff');
      sheet.getRange(changeLogStart, 1, 1, 6).merge();
      
      // Заголовки таблицы
      const logHeaders = ['Дата', 'Время', 'Описание изменения', 'Пользователь'];
      sheet.getRange(changeLogStart + 1, 1, 1, logHeaders.length).setValues([logHeaders]);
      sheet.getRange(changeLogStart + 1, 1, 1, logHeaders.length)
           .setFontWeight('bold')
           .setBackground('#cfd8dc');
    }
    
    // Находим первую пустую строку в логе
    let logRow = changeLogStart + 2;
    while (sheet.getRange(logRow, 1).getValue() !== '') {
      logRow++;
    }
    
    // Добавляем запись
    const now = new Date();
    const user = Session.getActiveUser().getEmail();
    
    sheet.getRange(logRow, 1, 1, 4).setValues([[
      now.toLocaleDateString('ru-RU'),
      now.toLocaleTimeString('ru-RU'),
      changeDescription,
      user
    ]]);
    
    // Форматирование
    sheet.getRange(logRow, 3).setWrap(true);
    sheet.getRange(logRow, 1, 1, 4).setBackground('#f5f5f5');
    
  } catch (error) {
    logError('❌ Ошибка добавления в лог изменений', error);
  }
}

/**
 * Проверяет позиции для активной категории
 */
function checkPositionsForActiveCategory() {
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
    
    const markerQuery = getMarkerQueryForCategory(categoryData.id);
    
    if (!markerQuery) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Маркерный запрос не настроен',
        'Сначала настройте маркерный запрос для отслеживания позиций.',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response === ui.Button.OK) {
        setupMarkerQuery();
      }
      return;
    }
    
    // Запрашиваем комментарий
    const ui = SpreadsheetApp.getUi();
    const commentResponse = ui.prompt(
      'Комментарий к проверке',
      'Добавьте комментарий (необязательно):',
      ui.ButtonSet.OK_CANCEL
    );
    
    const comment = commentResponse.getSelectedButton() === ui.Button.OK ? 
      commentResponse.getResponseText() : '';
    
    checkPositionsForCategory(categoryData, markerQuery, comment);
    
  } catch (error) {
    logError('❌ Ошибка проверки позиций', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Показывает отчет по динамике позиций
 */
function showPositionsDynamicsReport() {
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
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.POSITIONS_HISTORY);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert(
        'Нет данных',
        'История проверок позиций отсутствует',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Собираем данные по категории
    const data = sheet.getDataRange().getValues();
    const categoryChecks = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == categoryData.id) {
        categoryChecks.push({
          date: data[i][0],
          query: data[i][3],
          yandex: data[i][4],
          google: data[i][5],
          change: data[i][7],
          pageChanges: data[i][8]
        });
      }
    }
    
    if (categoryChecks.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Нет данных',
        'Проверки позиций для этой категории не проводились',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Формируем отчет
    let report = `📊 ОТЧЕТ ПО ДИНАМИКЕ ПОЗИЦИЙ\n\nКатегория: ${categoryData.title}\n`;
    report += `Проверок: ${categoryChecks.length}\n\n`;
    
    report += `ПОСЛЕДНИЕ 5 ПРОВЕРОК:\n\n`;
    
    categoryChecks.slice(0, 5).forEach((check, index) => {
      const date = new Date(check.date).toLocaleDateString('ru-RU');
      report += `${index + 1}. ${date} - Запрос: "${check.query}"\n`;
      report += `   Яндекс: ${check.yandex} | Google: ${check.google}\n`;
      report += `   Изменение: ${check.change}\n`;
      if (check.pageChanges) {
        report += `   Изменения на странице: ${check.pageChanges.split('\n')[0]}\n`;
      }
      report += `\n`;
    });
    
    // Анализ тренда
    if (categoryChecks.length >= 3) {
      report += `\nАНАЛИЗ ТРЕНДА (последние 3 проверки):\n`;
      
      const last3 = categoryChecks.slice(0, 3);
      const yandexTrend = analyzeTrend(last3.map(c => 
        typeof c.yandex === 'number' ? c.yandex : null
      ));
      const googleTrend = analyzeTrend(last3.map(c => 
        typeof c.google === 'number' ? c.google : null
      ));
      
      report += `Яндекс: ${yandexTrend}\n`;
      report += `Google: ${googleTrend}\n`;
    }
    
    SpreadsheetApp.getUi().alert('Динамика позиций', report, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    logError('❌ Ошибка показа отчета', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Анализирует тренд позиций
 */
function analyzeTrend(positions) {
  const validPositions = positions.filter(p => p !== null);
  
  if (validPositions.length < 2) {
    return 'Недостаточно данных';
  }
  
  const first = validPositions[validPositions.length - 1];
  const last = validPositions[0];
  const diff = first - last;
  
  if (diff > 3) {
    return `📈 Рост (+${diff} позиций)`;
  } else if (diff < -3) {
    return `📉 Падение (${diff} позиций)`;
  } else if (diff !== 0) {
    return `➡️ Небольшое изменение (${diff > 0 ? '+' : ''}${diff})`;
  } else {
    return `⏸️ Стабильно`;
  }
}

/**
 * Автоматическая фиксация изменений при отправке в InSales
 * Вызывается из sendCategoryChangesToInSales() в модуле 19
 */
function autoLogChangeBeforeSending(categoryData, updateData) {
  try {
    const changes = [];
    
    if (updateData.seo_title) changes.push(`SEO Title: "${updateData.seo_title}"`);
    if (updateData.seo_description) changes.push(`Meta Description: "${updateData.seo_description.substring(0, 50)}..."`);
    if (updateData.h1_title) changes.push(`H1: "${updateData.h1_title}"`);
    if (updateData.description) changes.push('Обновлено описание категории');
    
    if (changes.length > 0) {
      const changeLog = `Обновление через AI: ${changes.join('; ')}`;
      logPageChange(categoryData.id, changeLog);
      
      logInfo(`✅ Изменения автоматически зафиксированы: ${changes.length} элементов`);
    }
    
  } catch (error) {
    logError('❌ Ошибка автоматической фиксации изменений', error);
  }
}