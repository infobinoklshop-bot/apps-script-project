// ========================================
// ЗАГРУЗКА КАТЕГОРИЙ С ИЕРАРХИЕЙ
// ========================================

/**
 * Главная функция загрузки категорий с иерархией
 * Вызывается из меню
 */
function loadCategoriesWithHierarchy() {
  const context = "Загрузка категорий";

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Загружаем категории из InSales...',
      '⏳ Загрузка',
      -1
    );

    logInfo('🚀 Начинаем загрузку категорий с иерархией', null, context);

    // 1. Загружаем все коллекции из InSales
    const allCollections = loadAllCollectionsFromInSales();

    if (!allCollections || allCollections.length === 0) {
      throw new Error('Не удалось загрузить категории из InSales');
    }

    logInfo(`📦 Загружено коллекций: ${allCollections.length}`);

    // 2. Строим иерархическую структуру
    const hierarchy = buildCategoryHierarchy(allCollections);

    logInfo(`🌳 Построена иерархия: ${hierarchy.length} корневых категорий`);

    // 3. Преобразуем в плоский список с информацией об иерархии
    const flatList = flattenHierarchyWithLevels(hierarchy);

    logInfo(`📋 Создан плоский список: ${flatList.length} категорий`);

    // 4. Загружаем количество товаров для каждой категории
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Подсчитываем товары в категориях...',
      '⏳ Анализ',
      -1
    );

    // ИСПРАВЛЕНО: убрал await
    const categoriesWithCounts = addProductCountsToCategories(flatList);

    // 5. Записываем в главный лист
    writeCategoriesToMainList(categoriesWithCounts);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Загружено ${categoriesWithCounts.length} категорий с иерархией!`,
      '✅ Готово',
      10
    );

    logInfo('✅ Категории успешно загружены и записаны', null, context);

    return {
      success: true,
      totalCategories: categoriesWithCounts.length,
      rootCategories: hierarchy.length
    };

  } catch (error) {
    logError('❌ Ошибка загрузки категорий', error, context);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Ошибка: ' + error.message,
      '❌ Ошибка',
      10
    );
    throw error;
  }
}

/**
 * Загружает все коллекции из InSales
 */
function loadAllCollectionsFromInSales() {
  const context = "Загрузка коллекций InSales";

  try {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }

    const allCollections = [];
    let page = 1;
    const perPage = 250;

    while (true) {
      const url = `${credentials.baseUrl}${CATEGORY_API_ENDPOINTS.COLLECTIONS}?per_page=${perPage}&page=${page}`;

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
      const responseCode = response.getResponseCode();

      if (responseCode === 429) {
        Utilities.sleep(2000);
        continue;
      }

      if (responseCode !== 200) {
        throw new Error(`Ошибка API: ${responseCode}`);
      }

      const collections = JSON.parse(response.getContentText());

      if (!collections || collections.length === 0) {
        break;
      }

      allCollections.push(...collections);

      logInfo(`📄 Страница ${page}: загружено ${collections.length} коллекций`);

      if (collections.length < perPage || page >= 10) {
        break;
      }

      page++;
      Utilities.sleep(300);
    }

    logInfo(`✅ Всего загружено коллекций: ${allCollections.length}`, null, context);

    return allCollections;

  } catch (error) {
    logError('❌ Ошибка загрузки коллекций', error, context);
    throw error;
  }
}

/**
 * Строит иерархическую структуру категорий
 */
function buildCategoryHierarchy(collections) {
  const context = "Построение иерархии";

  try {
    logInfo('🌳 Строим иерархическую структуру категорий', null, context);

    // Создаем map для быстрого доступа
    const categoryMap = {};
    collections.forEach(cat => {
      categoryMap[cat.id] = {
        ...cat,
        children: [],
        level: 0,
        path: []
      };
    });

    // Строим дерево
    const rootCategories = [];

    collections.forEach(category => {
      const cat = categoryMap[category.id];

      if (!category.parent_id) {
        // Корневая категория
        rootCategories.push(cat);
      } else {
        // Дочерняя категория
        const parent = categoryMap[category.parent_id];
        if (parent) {
          parent.children.push(cat);
          cat.level = parent.level + 1;
          cat.path = [...parent.path, parent.title];
        } else {
          // Если родитель не найден, добавляем как корневую
          rootCategories.push(cat);
        }
      }
    });

    // Сортируем по position
    const sortByPosition = (a, b) => (a.position || 0) - (b.position || 0);

    rootCategories.sort(sortByPosition);

    function sortChildren(categories) {
      categories.forEach(cat => {
        if (cat.children.length > 0) {
          cat.children.sort(sortByPosition);
          sortChildren(cat.children);
        }
      });
    }

    sortChildren(rootCategories);

    logInfo(`✅ Построена иерархия: ${rootCategories.length} корневых категорий`, null, context);

    return rootCategories;

  } catch (error) {
    logError('❌ Ошибка построения иерархии', error, context);
    throw error;
  }
}

/**
 * Преобразует иерархию в плоский список с информацией об уровнях
 */
function flattenHierarchyWithLevels(hierarchy, parentPath = []) {
  const flatList = [];

  hierarchy.forEach(category => {
    const currentPath = [...parentPath, category.title];
    const pathString = currentPath.join(' > ');

    // Добавляем текущую категорию
    flatList.push({
      id: category.id,
      parent_id: category.parent_id || null,
      level: category.level,
      path: pathString,
      title: category.title,
      url: category.url || category.permalink || '',
      position: category.position || 0,
      is_hidden: category.is_hidden || false,
      seo_title: category.seo_title || '',
      seo_description: category.seo_description || '',
      description: category.description || '',
      productsCount: 0,  // Будет заполнено позже
      inStockCount: 0    // Будет заполнено позже
    });

    // Рекурсивно добавляем дочерние категории
    if (category.children && category.children.length > 0) {
      const childrenFlat = flattenHierarchyWithLevels(category.children, currentPath);
      flatList.push(...childrenFlat);
    }
  });

  return flatList;
}

/**
 * Добавляет количество товаров к категориям
 * Оптимизированная версия - загружаем товары один раз
 * ИСПРАВЛЕНО: убрал async
 */
function addProductCountsToCategories(categories) {
  const context = "Подсчет товаров";

  try {
    logInfo('📊 Подсчитываем товары в категориях', null, context);

    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные');
    }

    // Загружаем все товары один раз
    // ИСПРАВЛЕНО: убрал await
    const allProducts = loadAllProductsQuick(credentials);

    logInfo(`📦 Загружено товаров для анализа: ${allProducts.length}`);

    if (allProducts.length > 0) {
      const firstProd = allProducts[0];
      console.log('🔍 Структура первого товара (DEBUG):', JSON.stringify(firstProd));
      if (!firstProd.variants) console.warn('⚠️ У товара нет поля variants!');
      if (firstProd.variants && firstProd.variants.length > 0 && firstProd.variants[0].quantity === undefined) {
        console.warn('⚠️ У варианта нет поля quantity!');
      }
    }

    // Подсчитываем товары по категориям
    const productCountsByCategory = {};
    const inStockCountsByCategory = {};

    allProducts.forEach(product => {
      // Товар может быть в нескольких коллекциях
      const collectionIds = product.collections_ids || [];

      collectionIds.forEach(collectionId => {
        // Общее количество
        productCountsByCategory[collectionId] = (productCountsByCategory[collectionId] || 0) + 1;

        // В наличии (проверяем через варианты)
        const hasStock = product.variants && product.variants.some(v => v.quantity > 0);
        if (hasStock) {
          inStockCountsByCategory[collectionId] = (inStockCountsByCategory[collectionId] || 0) + 1;
        }
      });
    });

    // Применяем счетчики к категориям
    categories.forEach(category => {
      category.productsCount = productCountsByCategory[category.id] || 0;
      category.inStockCount = inStockCountsByCategory[category.id] || 0;
    });

    logInfo('✅ Счетчики товаров добавлены к категориям', null, context);

    return categories;

  } catch (error) {
    logError('❌ Ошибка подсчета товаров', error, context);
    // Возвращаем категории с нулевыми счетчиками
    return categories;
  }
}

/**
 * Быстрая загрузка всех товаров (только ID и коллекции)
 * ИСПРАВЛЕНО: убрал async
 */
function loadAllProductsQuick(credentials) {
  const allProducts = [];
  let page = 1;
  const perPage = 250;

  while (page <= 100) { // Увеличиваем лимит до 25000 товаров
    const url = `${credentials.baseUrl}/admin/products.json?per_page=${perPage}&page=${page}&fields=id,collections_ids,variants`;

    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() !== 200) {
        break;
      }

      const products = JSON.parse(response.getContentText());

      if (!products || products.length === 0) {
        break;
      }

      allProducts.push(...products);

      if (products.length < perPage) {
        break;
      }

      page++;
      Utilities.sleep(300);

    } catch (error) {
      logWarning(`⚠️ Ошибка на странице ${page}: ${error.message}`);
      break;
    }
  }

  return allProducts;
}

/**
 * Записывает категории в главный лист
 */
function writeCategoriesToMainList(categories) {
  const context = 'Запись категорий в главный лист';

  try {
    console.log('📝 Записываем категории в главный лист');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!sheet) {
      throw new Error('Лист "Категории — Список" не найден');
    }

    // Очищаем старые данные (кроме заголовка и инструкции)
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.deleteRows(3, lastRow - 2);
    }

    // Подготавливаем данные для записи
    const dataRows = categories.map(cat => {
      // Вычисляем отступ через ступенчатую иерархию
      const indent = '  '.repeat(cat.level);
      const visualTitle = cat.level > 0 ? `${indent}└─ ${cat.title}` : cat.title;

      // ИСПРАВЛЕНО: правильная ссылка на админку InSales
      const adminLink = `https://${INSALES_CONFIG.adminDomain}${INSALES_CONFIG.adminPath}/collections/${cat.id}`;

      // Определяем статусы
      const seoStatus = (cat.seo_title && cat.seo_description) ? '✅ Готово' : '⚠️ Требуется';

      return [
        false,                    // A - Checkbox (пустой)
        cat.id,                   // B - ID
        cat.parent_id || '',      // C - Parent ID
        cat.level,                // D - Уровень
        cat.path,                 // E - Иерархия с отступами
        visualTitle,              // F - Название
        cat.url,                  // G - URL
        cat.productsCount,        // H - Товаров
        cat.inStockCount,         // I - В наличии
        seoStatus,                // J - SEO статус
        'Не обработано',          // K - AI статус
        new Date(),               // L - Дата обновления
        adminLink                 // M - Админка
      ];
    });

    // Записываем данные начиная со строки 3
    if (dataRows.length > 0) {
      sheet.getRange(3, 1, dataRows.length, dataRows[0].length).setValues(dataRows);

      // Добавляем чекбоксы в колонку A
      const checkboxRange = sheet.getRange(3, 1, dataRows.length, 1);
      checkboxRange.insertCheckboxes();

      // Форматируем числовые колонки
      sheet.getRange(3, 8, dataRows.length, 2).setNumberFormat('0'); // Товаров, В наличии
      sheet.getRange(3, 12, dataRows.length, 1).setNumberFormat('dd.mm.yyyy hh:mm'); // Дата
      // ПРОВЕРКА И ВОССТАНОВЛЕНИЕ НАСТРОЙКИ "ТОВАРОВ НА СТРАНИЦЕ"
      let itemsPerPage = 36;
      const settingLabel = sheet.getRange('O1').getValue();
      const settingValue = sheet.getRange('P1').getValue();

      // Убрали создание "Товаров на 1 стр" в O1, так как это вызывает дублирование
      // if (settingLabel !== 'Товаров на 1 стр:') {
      //   sheet.getRange('O1').setValue('Товаров на 1 стр:').setFontWeight('bold').setHorizontalAlignment('right');
      // }
      if (settingLabel !== 'Товаров на 1 стр:') {
        sheet.getRange('P1').setValue(36).setHorizontalAlignment('center').setBackground('#fff3cd');
        sheet.getRange('P1').setNote('Количество товаров на первой странице листинга. Используется для выделения в детальном листе.');
      } else if (settingValue && !isNaN(parseInt(settingValue))) {
        itemsPerPage = parseInt(settingValue);
      }

      // Форматируем уровни для визуализации и выделяем "проблемные" категории
      for (let i = 0; i < dataRows.length; i++) {
        const level = dataRows[i][3]; // Уровень (Column D)
        const inStock = dataRows[i][8]; // В наличии (Column I)
        const rowIndex = i + 3;

        // Получаем статус скрытости из исходного объекта
        const isHidden = categories[i].is_hidden;

        let bgColor = '#ffffff';
        let fontColor = '#000000';

        // Логика цветов:
        // 1. Если категория скрыта -> СЕРЫЙ (приоритет)
        if (isHidden) {
          bgColor = '#f5f5f5'; // Light Gray
          fontColor = '#808080'; // Gray Text
        }
        // 2. Если товаров в наличии меньше, чем на 1 странице -> КРАСНОВАТЫЙ (требует внимания)
        else if (inStock < itemsPerPage) {
          bgColor = '#ffccbc'; // Light Red
        }
        // 3. Иначе - стандартная иерархия
        else {
          if (level === 0) bgColor = '#ede7f6'; // Светло-синий для корневых
          else if (level === 1) bgColor = '#f3e5f5'; // Очень светло-серый для уровня 1
          else if (level === 2) bgColor = '#fce4ec'; // Очень светло-серый для уровня 2
        }

        const range = sheet.getRange(rowIndex, 1, 1, 13);
        range.setBackground(bgColor);
        range.setFontColor(fontColor);
      }
    }

    // Удаляем строку с инструкцией
    sheet.deleteRow(2);

    console.log(`✅ Записано ${dataRows.length} категорий в лист`);

  } catch (error) {
    console.error('❌ Ошибка записи категорий:', error);
    throw error;
  }
}
function testCategoryFields() {
  const categoryId = 9171538;

  const credentials = getInsalesCredentialsSync();
  const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;

  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
      'Content-Type': 'application/json'
    }
  });

  const categoryData = JSON.parse(response.getContentText());

  console.log('=== SEO ПОЛЯ ===');
  console.log('html_title:', categoryData.html_title);
  console.log('meta_description:', categoryData.meta_description);

  console.log('\n=== FIELD_VALUES ===');
  if (categoryData.field_values && categoryData.field_values.length > 0) {
    for (let i = 0; i < Math.min(5, categoryData.field_values.length); i++) {
      const field = categoryData.field_values[i];
      console.log(`[${i}] name: "${field.name}", value: "${field.value ? field.value.substring(0, 80) : 'null'}"`);
    }
  }
}

function testFullFieldStructure() {
  const categoryId = 9171538;

  const credentials = getInsalesCredentialsSync();
  const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;

  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
      'Content-Type': 'application/json'
    }
  });

  const categoryData = JSON.parse(response.getContentText());

  console.log('=== FIELD_VALUES ПОЛНАЯ СТРУКТУРА ===');
  console.log('Всего field_values:', categoryData.field_values ? categoryData.field_values.length : 0);

  if (categoryData.field_values && categoryData.field_values.length > 0) {
    for (let i = 0; i < categoryData.field_values.length; i++) {
      const field = categoryData.field_values[i];

      console.log(`\n[${i}] ------------------`);
      console.log('Все ключи:', Object.keys(field));
      console.log('Полный объект:', JSON.stringify(field, null, 2));
    }
  }

  console.log('\n=== ПРОВЕРЯЕМ HTML_TITLE ===');
  console.log('html_title:', categoryData.html_title);
  console.log('title:', categoryData.title);
}