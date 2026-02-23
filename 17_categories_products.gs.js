/**
 * ========================================
 * УПРАВЛЕНИЕ ТОВАРАМИ КАТЕГОРИИ (ФИНАЛЬНАЯ ВЕРСИЯ - COLLECTS API)
 * ========================================
 */

/**
 * Удаляет выбранные товары из категории
 */
function removeSelectedProductsFromCategory() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();

    if (!sheetName.startsWith('Категория — ')) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Только для детального листа категории', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const categoryId = sheet.getRange('B2').getValue();
    if (!categoryId) throw new Error('ID категории не найден');

    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      SpreadsheetApp.getUi().alert('Нет товаров в категории');
      return;
    }

    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues();
    const toDelete = [];

    for (let i = 0; i < data.length; i++) {
      const title = data[i][0];     // A - Название
      const sku = data[i][1];       // B - Артикул  
      const price = data[i][2];     // C - Цена
      const inStock = data[i][3];   // D - В наличии
      const id = data[i][4];        // E - ID
      const checked = data[i][5];   // F - Чекбокс ← ИСПРАВЛЕНО!

      if (!id || id.toString().trim() === '') break;

      if (checked === true) {
        toDelete.push({ id: parseInt(id), title: title, row: startRow + i });
      }
    }

    if (toDelete.length === 0) {
      SpreadsheetApp.getUi().alert('Нет отмеченных товаров');
      return;
    }

    const ui = SpreadsheetApp.getUi();
    const confirm = ui.alert(
      'Удалить товары?',
      `Будет удалено: ${toDelete.length} товаров\n\n${toDelete.slice(0, 3).map(t => t.title).join('\n')}\n\nПродолжить?`,
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    SpreadsheetApp.getActiveSpreadsheet().toast('Удаляем товары из категории...', '⏳ Обработка', -1);

    // Удаляем через Collects API
    const result = removeProductsFromCategoryAPI(categoryId, toDelete.map(t => t.id));

    // Удаляем строки в обратном порядке
    for (let i = toDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(toDelete[i].row);
    }

    // Обновляем статистику
    updateCategoryStatistics(sheet);

    SpreadsheetApp.getActiveSpreadsheet().toast(`Удалено: ${result.success}`, '✅ Готово', 5);
    ui.alert('Готово', `Удалено: ${result.success}\nОшибок: ${result.errors}`, ui.ButtonSet.OK);

  } catch (error) {
    console.error('❌ Ошибка:', error.message);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Удаляет товары через Collects API (РАБОЧИЙ МЕТОД!)
 */
function removeProductsFromCategoryAPI(categoryId, productIds) {
  try {
    console.log(`🗑️ Удаляем ${productIds.length} товаров из категории ${categoryId}`);

    const credentials = getInsalesCredentialsSync();
    if (!credentials) throw new Error('Учетные данные не настроены');

    // ШАГ 1: Получаем все связи категории
    const collectsUrl = `${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&per_page=250`;

    console.log(`📋 Загружаем связи категории...`);

    const collectsResponse = UrlFetchApp.fetch(collectsUrl, {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    if (collectsResponse.getResponseCode() !== 200) {
      throw new Error('Не удалось загрузить связи категории');
    }

    const collects = JSON.parse(collectsResponse.getContentText());
    console.log(`📋 Найдено связей: ${collects.length}`);

    // ШАГ 2: Находим ID связей для удаляемых товаров
    const collectsToDelete = collects.filter(collect =>
      productIds.includes(collect.product_id)
    );

    console.log(`🗑️ Связей для удаления: ${collectsToDelete.length}`);

    if (collectsToDelete.length === 0) {
      console.log('⚠️ Ни один товар не найден в категории');
      return { success: 0, errors: 0 };
    }

    // ШАГ 3: Удаляем каждую связь
    let successCount = 0;
    let errorCount = 0;

    for (const collect of collectsToDelete) {
      try {
        const deleteUrl = `${credentials.baseUrl}/admin/collects/${collect.id}.json`;

        console.log(`🗑️ Удаляем связь ${collect.id} (товар ${collect.product_id})...`);

        const deleteResponse = UrlFetchApp.fetch(deleteUrl, {
          method: 'DELETE',
          headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
          },
          muteHttpExceptions: true
        });

        const statusCode = deleteResponse.getResponseCode();

        if (statusCode === 200 || statusCode === 204) {
          console.log(`✅ Товар ${collect.product_id} удалён из категории`);
          successCount++;
        } else {
          console.error(`❌ Ошибка удаления товара ${collect.product_id}: ${statusCode}`);
          errorCount++;
        }

        Utilities.sleep(500);

      } catch (e) {
        console.error(`❌ Ошибка удаления связи ${collect.id}:`, e.message);
        errorCount++;
      }
    }

    console.log(`✅ Итого: удалено ${successCount}, ошибок ${errorCount}`);

    return { success: successCount, errors: errorCount };

  } catch (error) {
    console.error('❌ Ошибка удаления товаров:', error);
    throw error;
  }
}

/**
 * Добавляет товары в категорию через Collects API
 */
function addProductsToCategory(categoryId, productIds) {
  try {
    console.log(`➕ Добавляем ${productIds.length} товаров в категорию ${categoryId}`);

    const credentials = getInsalesCredentialsSync();
    if (!credentials) throw new Error('Учетные данные не настроены');

    let successCount = 0;
    let errorCount = 0;
    const addedProducts = [];

    for (const productId of productIds) {
      try {
        const createUrl = `${credentials.baseUrl}/admin/collects.json`;
        const payload = {
          collect: {
            collection_id: categoryId,
            product_id: productId
          }
        };

        console.log(`➕ Добавляем товар ${productId} в категорию ${categoryId}...`);

        const response = UrlFetchApp.fetch(createUrl, {
          method: 'POST',
          headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });

        if (response.getResponseCode() === 201) {
          console.log(`✅ Товар ${productId} добавлен`);
          successCount++;
          addedProducts.push(productId);
        } else {
          console.error(`❌ Ошибка добавления товара ${productId}: ${response.getResponseCode()} ${response.getContentText()}`);
          errorCount++;
        }

        Utilities.sleep(300); // Пауза для стабильности

      } catch (e) {
        console.error(`❌ Ошибка обработки товара ${productId}:`, e);
        errorCount++;
      }
    }

    console.log(`✅ Добавлено: ${successCount}, Ошибок: ${errorCount}`);
    return { success: successCount, errors: errorCount, added: addedProducts };

  } catch (error) {
    console.error('❌ Ошибка добавления товаров:', error);
    throw error;
  }
}

/**
 * Синхронизирует товары с листа в InSales (добавляет недостающие)
 * Используется, когда товары добавлены в таблицу вручную
 */
function syncProductsFromSheetToInSales() {
  const context = 'syncProductsFromSheetToInSales';
  const ui = SpreadsheetApp.getUi();

  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (!sheet.getName().startsWith('Категория — ')) {
      ui.alert('Ошибка', 'Запускайте только с детального листа категории', ui.ButtonSet.OK);
      return;
    }

    const categoryId = sheet.getRange('B2').getValue();
    if (!categoryId) {
      ui.alert('Ошибка', 'ID категории не найден в ячейке B2', ui.ButtonSet.OK);
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Анализируем товары...', '🔍 Поиск', -1);

    // 1. Получаем товары с листа (ID из колонки E)
    const sheetProductIds = getCurrentCategoryProductIds(sheet);
    if (sheetProductIds.length === 0) {
      ui.alert('В таблице нет товаров', 'Добавьте товары в список вручную или через поиск.', ui.ButtonSet.OK);
      return;
    }

    // 2. Получаем текущие товары в InSales
    const credentials = getInsalesCredentialsSync();
    if (!credentials) throw new Error('Нет доступов к InSales');

    const collects = getAllCollectsForCategory(categoryId, credentials);
    const insalesIdsSet = new Set(collects.map(c => c.product_id));

    // 3. Вычисляем разницу
    const toAdd = sheetProductIds.filter(id => !insalesIdsSet.has(id));

    if (toAdd.length === 0) {
      ui.alert('Синхронизация', 'Все товары из таблицы уже есть в категории InSales.', ui.ButtonSet.OK);
      return;
    }

    // 4. Подтверждение
    const confirm = ui.alert(
      'Отправка товаров',
      `Найдено товаров в таблице: ${sheetProductIds.length}\n` +
      `Отсутствуют в InSales: ${toAdd.length}\n\n` +
      `Добавить эти товары в категорию InSales?`,
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    SpreadsheetApp.getActiveSpreadsheet().toast(`Добавляем ${toAdd.length} товаров...`, '🚀 Отправка', -1);

    // 5. Добавляем
    const result = addProductsToCategory(categoryId, toAdd);

    // 6. Предлагаем обновить порядок
    const sortConfirm = ui.alert(
      'Готово',
      `Добавлено товаров: ${result.success}\nОшибок: ${result.errors}\n\n` +
      `Хотите также синхронизировать ПОРЯДОК сортировки как в таблице?`,
      ui.ButtonSet.YES_NO
    );

    if (sortConfirm === ui.Button.YES) {
      updateProductPositionsFromSheetV2(true);
    }

  } catch (error) {
    console.error(`❌ Ошибка синхронизации товаров:`, error);
    ui.alert('Ошибка: ' + error.message);
  }
}

/**
 * Показывает диалог добавления товаров
 */
function showAddProductsDialog() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();

    if (!sheetName.startsWith('Категория — ')) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Только для детального листа категории', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const categoryId = sheet.getRange('B2').getValue();
    const categoryTitle = sheet.getRange('B3').getValue();

    if (!categoryId) throw new Error('ID категории не найден');

    // Получаем текущие товары категории для исключения дублей
    const existingProductIds = getCurrentCategoryProductIds(sheet);

    const htmlContent = createAddProductsDialogHTML(categoryId, categoryTitle, existingProductIds);

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(900)
      .setHeight(650);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Добавить товары в категорию');

  } catch (error) {
    console.error('❌ Ошибка:', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Получает ID текущих товаров категории
 */
function getCurrentCategoryProductIds(sheet) {
  try {
    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) return [];

    // ИСПРАВЛЕНО: ID товара в колонке E (5), а не B (2)!
    const data = sheet.getRange(startRow, 5, lastRow - startRow + 1, 1).getValues();
    const ids = [];

    for (let i = 0; i < data.length; i++) {
      const id = data[i][0];
      if (id && id.toString().trim() !== '') {
        ids.push(parseInt(id));
      } else {
        break;
      }
    }

    console.log(`📋 Текущих товаров в категории: ${ids.length}`);
    return ids;

  } catch (error) {
    console.error('❌ Ошибка чтения товаров:', error);
    return [];
  }
}

/**
 * ========================================
 * СОРТИРОВКА ТОВАРОВ (НОВЫЙ ФУНКЦИОНАЛ)
 * ========================================
 */

/**
 * Обновляет порядок товаров в InSales на основе порядка строк в таблице
 */
/**
 * Обновляет порядок товаров в InSales на основе порядка строк в таблице
 * @param {boolean} isInteractive - показывать ли UI диалоги (по умолчанию true)
 * @param {number} limit - ограничение количества обновляемых товаров (0 = все)
 */
// Duplicate function removed


/**
 * Получает ВСЕ связи (collects) категории с пагинацией
 */
function getAllCollectsForCategory(categoryId, credentials) {
  let allCollects = [];
  let page = 1;
  const perPage = 250; // Максимум InSales

  while (true) {
    const url = `${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&per_page=${perPage}&page=${page}`;

    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Ошибка загрузки связей: ${response.getResponseCode()}`);
    }

    const collects = JSON.parse(response.getContentText());
    if (collects.length === 0) break;

    allCollects = allCollects.concat(collects);

    if (collects.length < perPage) break;
    page++;
  }

  return allCollects;
}

/**
 * Обновляет позицию одной связи
 */
function updateCollectPosition(collectId, position, credentials) {
  try {
    const url = `${credentials.baseUrl}/admin/collects/${collectId}.json`;

    const payload = {
      collect: {
        position: position
      }
    };

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
      return true;
    } else {
      console.error(`❌ Ошибка обновления позиции collect ${collectId}: ${response.getResponseCode()} ${response.getContentText()}`);
      return false;
    }
  } catch (e) {
    console.error(`❌ Исключение при обновлении позиции: ${e.message}`);
    return false;
  }
}

/**
 * Проверяет и включает ручную сортировку для категории
 * (В InSales это обычно неявное поведение, но иногда полезно обновить саму коллекцию, 
 * чтобы сбросить кэши или триггернуть пересортировку)
 * 
 * Примечание: В API InSales нет явного поля "sort_type" = "manual". 
 * Обычно сортировка по position работает по умолчанию, если не выбрано иное в админке.
 * Но мы можем попробовать обновить timestamp категории, чтобы сбросить кэш.
 */
function ensureCategoryManualSorting(categoryId, credentials) {
  // Пока оставим пустым или просто залогируем. 
  // Если выяснится, что нужно менять поле sort_order, добавим сюда.
  console.log('ℹ️ Подготовка категории к ручной сортировке...');
}

/**
 * Умная сортировка товаров (равномерное распределение брендов + наличие)
 * 
 * Алгоритм (ОПТИМИЗИРОВАННЫЙ):
 * 1. Определяем размер первой страницы (из настроек или 36).
 * 2. Разделяем товары на "В наличии" и "Нет в наличии".
 * 3. Формируем ПЕРВУЮ СТРАНИЦУ:
 *    - Обязательно по 1 представителю каждого бренда.
 *    - Оставшиеся места заполняем методом Round Robin (для разнообразия).
 * 4. Оставшиеся товары "В наличии" добавляем ГРУППАМИ по бренду (чтобы был запас вариантов).
 * ИСПОЛЬЗУЕТ ОБЩУЮ ЛОГИКУ из 17b_categories_sorting_logic.js
 */
function smartSortProductsByBrand(isInteractive = true) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();

    if (!sheetName.startsWith('Категория — ')) {
      if (isInteractive) {
        SpreadsheetApp.getUi().alert('Ошибка', 'Только для детального листа категории', SpreadsheetApp.getUi().ButtonSet.OK);
      } else {
        console.error('Ошибка: Только для детального листа категории');
      }
      return;
    }

    // Читаем настройку "Товаров на 1 стр" из главного листа
    let itemsPerPage = 36;
    try {
      const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
      if (mainSheet) {
        const settingValue = mainSheet.getRange('P1').getValue();
        if (settingValue && !isNaN(parseInt(settingValue))) {
          itemsPerPage = parseInt(settingValue);
        }
      }
    } catch (e) {
      console.warn('Не удалось прочитать настройку товаров на странице:', e);
    }

    const ui = SpreadsheetApp.getUi();

    if (isInteractive) {
      const confirm = ui.alert(
        'Умная сортировка',
        `Будет выполнена сортировка с учетом брендов и наличия.\nПервая страница: ${itemsPerPage} товаров.\n\nПродолжить?`,
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;
      SpreadsheetApp.getActiveSpreadsheet().toast('Сортируем товары...', '⏳ Обработка', -1);
    } else {
      console.log(`Запуск умной сортировки (limit=${itemsPerPage}) в фоновом режиме...`);
    }

    // 1. Получаем данные товаров
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      if (isInteractive) SpreadsheetApp.getUi().alert('Нет товаров в категории');
      return;
    }

    // Читаем колонки A-G (7 колонок)
    const numCols = 7;
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols);
    const data = range.getValues();

    // 2. Преобразуем данные в объекты для общей логики
    const products = data.map((row, index) => {
      const title = row[0];
      const inStock = row[3] === 'Да';
      const id = row[4];
      const brand = row[6]; // Бренд из колонки G

      return {
        id: id,
        title: title,
        inStock: inStock,
        brand: normalizeBrand(brand, title), // Используем функцию из 17b
        originalPosition: index,
        originalRow: row // Сохраняем весь ряд, чтобы потом восстановить
      };
    });

    // 3. Вычисляем новый порядок (Pure Logic)
    const sortedProducts = calculateSmartSortOrder(products, itemsPerPage);

    // 4. Формируем данные для записи обратно
    const finalSortedData = sortedProducts.map(p => p.originalRow);

    // 5. Записываем обратно в таблицу
    if (finalSortedData.length > 0) {
      range.clearContent();
      sheet.getRange(startRow, 1, finalSortedData.length, numCols).setValues(finalSortedData);
      sheet.getRange(startRow, 6, finalSortedData.length, 1).insertCheckboxes();

      // Обновляем визуальную границу первой страницы
      range.setBorder(false, false, false, false, false, false);
      sheet.getRange(startRow, 1, finalSortedData.length, 7).setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

      const page1Count = Math.min(itemsPerPage, finalSortedData.length);
      if (page1Count > 0) {
        sheet.getRange(startRow, 1, page1Count, 7)
          .setBorder(true, true, true, true, null, null, '#4285f4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

        const limitRow = startRow + page1Count - 1;
        sheet.getRange(startRow, 8, finalSortedData.length, 1).clearContent();
        sheet.getRange(limitRow, 8).setValue(`⬅️ Конец 1-й страницы (${itemsPerPage} шт.)`)
          .setFontColor('#4285f4').setFontWeight('bold');
      }
    }

    // 6. Обновляем InSales
    if (isInteractive) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Готово! Товары распределены.', '✅ Успешно', 5);
      const result = ui.alert(
        'Умная сортировка выполнена',
        `Товары отсортированы на листе.\nОбновить порядок в InSales (только первые ${itemsPerPage} позиций)?`,
        ui.ButtonSet.YES_NO
      );
      if (result === ui.Button.YES) {
        updateProductPositionsFromSheetV2(true, itemsPerPage);
      }
    } else {
      console.log(`Сортировка на листе завершена. Обновляем InSales...`);
      updateProductPositionsFromSheetV2(false, itemsPerPage);
    }

  } catch (error) {
    console.error('❌ Ошибка умной сортировки:', error);
    if (isInteractive) {
      SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
    }
  }
}

/**
 * ПРЯМАЯ СОРТИРОВКА БЕЗ ЛИСТА (через API)
 * Используется, когда детальный лист не создан.
 */
function smartSortCategoryDirectly(categoryId, categoryName, itemsPerPage) {
  try {
    console.log(`🚀 Запуск ПРЯМОЙ сортировки для категории ${categoryName} (ID: ${categoryId})`);

    // 1. Загружаем товары из InSales
    // Используем функцию из 13_categories_search.js
    const apiProducts = loadCategoryProducts(categoryId);

    if (!apiProducts || apiProducts.length === 0) {
      console.warn(`⚠️ Товары не найдены для категории ${categoryId}`);
      return false;
    }

    console.log(`📦 Загружено ${apiProducts.length} товаров. Обрабатываем...`);

    // 2. Преобразуем в формат для сортировки
    const products = apiProducts.map((p, index) => {
      const variant = p.variants && p.variants[0];
      const quantity = variant ? variant.quantity : 0;
      const inStock = quantity > 0;

      return {
        id: p.id,
        title: p.title,
        inStock: inStock,
        brand: normalizeBrand(p.vendor, p.title),
        originalPosition: index
      };
    });

    // 3. Вычисляем новый порядок
    const sortedProducts = calculateSmartSortOrder(products, itemsPerPage);

    // 4. Обновляем позиции в InSales (только первая страница)
    const credentials = getInsalesCredentialsSync();

    // 4.0. Убеждаемся, что включена ручная сортировка
    ensureCollectionIsManualSorted(categoryId, credentials);

    const collects = getAllCollectsForCategory(categoryId, credentials);
    const productToCollectMap = new Map();
    collects.forEach(c => productToCollectMap.set(c.product_id, c.id));

    let updatedCount = 0;
    const limit = Math.min(itemsPerPage, sortedProducts.length);

    // --- PARTY CRASHER LOGIC (DIRECT API) ---
    // We need to ensure that the Top 'limit' positions are occupied ONLY by our 'Active Kings'.
    // Anyone else (Aliens or Sleeping Kings) found in these positions must be evicted.

    const activeKingIds = sortedProducts.slice(0, limit).map(p => p.id);
    const activeKingIdsSet = new Set(activeKingIds.map(id => parseInt(id)));

    console.log(`👑 Active Kings (Direct): ${activeKingIds.length} items.`);
    console.log(`🧹 Checking for Party Crashers (Non-Kings in top ${limit} positions)...`);

    // 1. Evict Party Crashers
    const partyCrashers = collects.filter(c =>
      !activeKingIdsSet.has(c.product_id) && c.position <= limit
    );

    if (partyCrashers.length > 0) {
      console.log(`🚫 Found ${partyCrashers.length} Party Crashers in the top ${limit} positions. Evicting...`);

      for (let i = 0; i < partyCrashers.length; i++) {
        const collect = partyCrashers[i];
        const newPosition = 2000 + i; // Move to safe distance

        try {
          console.log(`   🦵 Evicting Product ${collect.product_id} (Pos ${collect.position}) -> ${newPosition}`);
          updateCollectPosition(collect.id, newPosition, credentials);
          Utilities.sleep(100);
        } catch (e) {
          console.warn(`   ❌ Failed to evict ${collect.product_id}: ${e.message}`);
        }
      }
    } else {
      console.log(`✅ No Party Crashers in the top ${limit}. The floor is clear.`);
    }
    // --- END PARTY CRASHER LOGIC ---

    console.log(`🔄 Обновляем позиции первых ${limit} товаров в InSales...`);

    for (let i = 0; i < limit; i++) {
      const product = sortedProducts[i];
      const collectId = productToCollectMap.get(product.id);
      const newPosition = i + 1;

      if (collectId) {
        // Check if update is needed to save API calls
        const currentCollect = collects.find(c => c.id == collectId);
        if (currentCollect && currentCollect.position === newPosition) {
          console.log(`   ✅ King ${product.id} already on throne (Pos ${newPosition})`);
          continue;
        }

        const success = updateCollectPosition(collectId, newPosition, credentials);
        if (success) updatedCount++;
        Utilities.sleep(200); // Пауза для стабильности
      }
    }

    console.log(`✅ Прямая сортировка завершена. Обновлено: ${updatedCount}`);
    return true;

  } catch (error) {
    console.error(`❌ Ошибка прямой сортировки категории ${categoryId}:`, error);
    return false;
  }
}

/**
 * Updates product positions in InSales based on the order in the sheet.
 * @param {boolean} isInteractive - True if called from UI, false for background.
 * @param {number} itemsPerPage - Number of items to update (usually 1st page limit).
 */
function updateProductPositionsFromSheetV2(isInteractive, itemsPerPage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Default isInteractive to true if undefined (e.g. called from menu)
  if (typeof isInteractive === 'undefined') isInteractive = true;

  try {
    console.log('🚀 [DEBUG] updateProductPositionsFromSheet started (v2)');
    const sheet = ss.getActiveSheet();
    const sheetName = sheet.getName();
    const categoryId = sheet.getRange('B2').getValue(); // Assuming B2 contains category ID

    if (!categoryId) {
      if (isInteractive) ui.alert('Ошибка', 'Не удалось получить ID категории из ячейки B2.', ui.ButtonSet.OK);
      console.error('❌ Не удалось получить ID категории из ячейки B2.');
      return;
    }

    // Explicitly use the sync function
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      if (isInteractive) ui.alert('Ошибка', 'Не удалось получить учетные данные InSales. Проверьте настройки.', ui.ButtonSet.OK);
      console.error('❌ Не удалось получить учетные данные InSales.');
      return;
    }

    // 1. Получаем ID товаров из таблицы
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      if (isInteractive) ui.alert('Нет товаров для обновления позиций.');
      return;
    }

    // Ensure collection is set to manual sorting
    if (!ensureCollectionIsManualSorted(categoryId, credentials)) {
      console.warn('⚠️ Could not ensure manual sorting, proceeding anyway...');
    }

    const productIdsInOrder = sheet.getRange(startRow, 5, lastRow - startRow + 1, 1) // Column E (ID)
      .getValues()
      .flat()
      .filter(id => id && id.toString().trim() !== '');

    if (productIdsInOrder.length === 0) {
      if (isInteractive) ui.alert('Нет ID товаров для обновления позиций.');
      return;
    }

    // Determine limit
    let limit = itemsPerPage;

    // If limit is undefined, try to read from settings
    if (typeof limit === 'undefined' || limit === null) {
      let settingsLimit = 36; // Default
      try {
        const mainSheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
        if (mainSheet) {
          const val = mainSheet.getRange('P1').getValue();
          if (val && !isNaN(parseInt(val))) {
            settingsLimit = parseInt(val);
          }
        }
      } catch (e) {
        console.warn('Error reading itemsPerPage setting', e);
      }

      limit = settingsLimit;

      // If interactive, ask confirmation for First Page only (User preference)
      if (isInteractive) {
        const response = ui.alert(
          'Обновление порядка',
          `Обновить порядок первых ${limit} товаров в InSales?\n\n` +
          `Это обеспечит правильную сортировку на первой странице сайта.`,
          ui.ButtonSet.YES_NO
        );

        if (response !== ui.Button.YES) return;
      }
    }

    // 2. Получаем текущие коллекции для категории
    // Это нужно, чтобы найти collect_id для каждого товара
    const collects = getAllCollectsForCategory(categoryId, credentials);
    if (!collects || collects.length === 0) {
      if (isInteractive) ui.alert('Ошибка', `Не удалось получить коллекции для категории ID ${categoryId}.`, ui.ButtonSet.OK);
      console.error(`❌ Не удалось получить коллекции для категории ID ${categoryId}.`);
      return;
    }

    let updatedCount = 0;
    let errorCount = 0;
    const total = Math.min(limit, productIdsInOrder.length);

    if (isInteractive) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Начинаем обновление ${total} позиций...`, '🔄 Сортировка', -1);
    } else {
      console.log(`Начинаем обновление ${total} позиций для категории ${categoryId}...`);
    }

    // --- UNIFIED SORTING LOGIC ---
    // We need to ensure that the Top 'limit' positions are occupied ONLY by our 'Active Kings'.
    // Anyone else (Aliens or Sleeping Kings) found in these positions must be evicted.

    const activeKingIds = productIdsInOrder.slice(0, total);
    const activeKingIdsSet = new Set(activeKingIds.map(id => parseInt(id)));

    console.log(`👑 Active Kings: ${activeKingIds.length} items.`);
    console.log(`🧹 Checking for Party Crashers (Non-Kings in top ${limit} positions)...`);

    // 1. Evict Party Crashers
    // A "Party Crasher" is any product that is NOT an Active King but has a position <= total.
    // We ONLY evict those who are physically blocking our seats (1 to total).
    // We ignore the mess at positions > total (e.g. 37+), as they don't affect the first page.
    const partyCrashers = collects.filter(c =>
      !activeKingIdsSet.has(c.product_id) && c.position <= total
    );

    if (partyCrashers.length > 0) {
      console.log(`🚫 Found ${partyCrashers.length} Party Crashers in the top ${total} positions. Evicting...`);
      if (isInteractive) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Освобождаем ${partyCrashers.length} мест на первой странице...`, '🧹 Чистка', -1);
      }

      for (let i = 0; i < partyCrashers.length; i++) {
        const collect = partyCrashers[i];
        const newPosition = 2000 + i; // Move to safe distance

        try {
          console.log(`   🦵 Evicting Product ${collect.product_id} (Pos ${collect.position}) -> ${newPosition}`);
          updateCollectPosition(collect.id, newPosition, credentials);
          Utilities.sleep(100);
        } catch (e) {
          console.warn(`   ❌ Failed to evict ${collect.product_id}: ${e.message}`);
        }
      }
    } else {
      console.log(`✅ No Party Crashers in the top ${total}. The floor is clear.`);
    }

    // 2. Crown the Kings
    console.log(`👑 Crowning ${activeKingIds.length} Kings...`);

    for (let i = 0; i < activeKingIds.length; i++) {
      const productId = activeKingIds[i];
      const newPosition = i + 1;

      const collect = collects.find(c => c.product_id == productId);
      if (!collect) {
        console.warn(`⚠️ King Product ID ${productId} not found in category.`);
        errorCount++;
        continue;
      }

      // Force update if position doesn't match
      if (collect.position !== newPosition) {
        try {
          console.log(`   🔄 Moving King ${productId} (Collect ${collect.id}): ${collect.position} -> ${newPosition}`);
          updateCollectPosition(collect.id, newPosition, credentials);
          updatedCount++;
          Utilities.sleep(200);
        } catch (e) {
          console.error(`❌ Error moving King ${productId}:`, e);
          errorCount++;
        }
      } else {
        console.log(`   ✅ King ${productId} already on throne (Pos ${newPosition})`);
      }

      if ((i + 1) % 5 === 0 && isInteractive) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Расставлено: ${i + 1}/${total}.`, '👑 Коронация', -1);
      }
    }

    if (isInteractive) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Готово! Обновлено: ${updatedCount}`, '✅ Успешно', 5);

      const message = `✅ Порядок товаров обновлен!\n\n` +
        `Всего товаров в таблице: ${productIdsInOrder.length}\n` +
        `Обработано (лимит): ${total}\n` +
        `Позиций обновлено: ${updatedCount}\n` +
        `Ошибок: ${errorCount}`;

      ui.alert('Готово', message, ui.ButtonSet.OK);
    } else {
      console.log(`✅ Порядок обновлен. Всего: ${productIdsInOrder.length}, Обработано: ${total}, Обновлено: ${updatedCount}, Ошибок: ${errorCount}`);
    }

  } catch (error) {
    console.error('❌ Ошибка сортировки:', error);
    if (isInteractive) {
      SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
    }
  }
}

/**
 * Сортирует отмеченные категории в списке "Категории — Список"
 */
/**
 * ЗАПУСК ПАКЕТНОЙ СОРТИРОВКИ (UI)
 * Инициирует процесс и запускает воркер
 */
function startBatchSort() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

  if (!listSheet) {
    SpreadsheetApp.getUi().alert('Ошибка', `Лист "${CATEGORY_SHEETS.MAIN_LIST}" не найден`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Читаем настройку itemsPerPage
  let itemsPerPage = 36;
  try {
    const settingValue = listSheet.getRange('P1').getValue();
    if (settingValue && !isNaN(parseInt(settingValue))) {
      itemsPerPage = parseInt(settingValue);
    }
  } catch (e) { }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Пакетная сортировка',
    `Будет запущена умная сортировка (только 1-я страница: ${itemsPerPage} шт.) для всех отмеченных категорий.\n\n` +
    `УСЛОВИЯ:\n` +
    `1. Товаров в наличии > 40 (иначе пропускаем)\n` +
    `2. Если скрипт не успеет, он САМ перезапустится через 1 минуту.\n\n` +
    `Продолжить?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  // Запускаем воркер
  ss.toast('Запускаем процесс...', '🚀 Старт', -1);
  runBatchSortWorker();
}

/**
 * ВОРКЕР ПАКЕТНОЙ СОРТИРОВКИ
 * Работает в фоне, следит за временем и перезапускается
 */
function runBatchSortWorker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
  if (!listSheet) return;

  // Настройки
  let itemsPerPage = 36;
  try {
    const settingValue = listSheet.getRange('P1').getValue();
    if (settingValue && !isNaN(parseInt(settingValue))) itemsPerPage = parseInt(settingValue);
  } catch (e) { }

  const MIN_STOCK_COUNT = 40; // Минимальное кол-во товаров для сортировки
  const MAX_EXECUTION_TIME = 3.5 * 60 * 1000; // 3.5 минуты (безопасный запас для предотвращения Hard Timeout)
  const startTime = new Date().getTime();

  const lastRow = listSheet.getLastRow();
  if (lastRow < 2) return;

  // Читаем данные: A (Check), B (ID), F (Name), I (In Stock)
  // Индексы в массиве (0-based): Check=0, ID=1, Name=5, InStock=8
  // Range берем шире, до I (9 колонок)
  const range = listSheet.getRange(2, 1, lastRow - 1, 9);
  const values = range.getValues();

  let processedCount = 0;
  let skippedCount = 0;
  let hasMoreWork = false;

  for (let i = 0; i < values.length; i++) {
    // 1. Проверка времени
    if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
      console.log('⏳ Время вышло. Планируем перезапуск...');

      try {
        // Удаляем старые триггеры перед созданием нового
        const triggers = ScriptApp.getProjectTriggers();
        for (const trigger of triggers) {
          if (trigger.getHandlerFunction() === 'runBatchSortWorker') {
            ScriptApp.deleteTrigger(trigger);
          }
        }

        // Создаем триггер на запуск через 1 минуту
        ScriptApp.newTrigger('runBatchSortWorker')
          .timeBased()
          .after(60 * 1000)
          .create();

        console.log('✅ Триггер перезапуска создан успешно.');
      } catch (e) {
        console.error('❌ Ошибка создания триггера:', e.message);
        ss.toast('⚠️ Ошибка авто-запуска. Запустите вручную.', 'Ошибка', 10);
      }

      ss.toast('Перерыв на кофе ☕. Продолжу через минуту...', '⏸️ Пауза', -1);
      return; // Выходим, чтобы триггер подхватил эстафету
    }

    const isChecked = values[i][0];

    if (isChecked === true || isChecked === 'true') {
      hasMoreWork = true;
      const categoryId = values[i][1];
      let categoryName = values[i][5];
      const inStock = parseInt(values[i][8]) || 0;

      // 2. Проверка количества товаров
      if (inStock <= MIN_STOCK_COUNT) {
        console.log(`⏩ Пропуск категории "${categoryName}": товаров ${inStock} (минимум ${MIN_STOCK_COUNT})`);
        listSheet.getRange(i + 2, 1).setValue(false); // Снимаем галочку
        skippedCount++;
        SpreadsheetApp.flush();
        continue;
      }

      // 3. Сортировка
      try {
        const cleanName = categoryName.toString().replace(/^[└\s\-|]+/, '').trim();
        const sheetName = `${CATEGORY_SHEETS.DETAIL_PREFIX}${cleanName}`;
        let detailSheet = ss.getSheetByName(sheetName);

        // Поиск листа по ID (если имя отличается)
        if (!detailSheet) {
          const allSheets = ss.getSheets();
          for (const sheet of allSheets) {
            if (sheet.getName().startsWith(CATEGORY_SHEETS.DETAIL_PREFIX)) {
              const sheetId = sheet.getRange('B2').getValue();
              if (sheetId == categoryId) {
                detailSheet = sheet;
                break;
              }
            }
          }
        }

        let success = false;

        if (detailSheet) {
          console.log(`🔄 [ЛИСТ] Обработка: ${cleanName}`);
          ss.toast(`Сортировка: ${cleanName}`, '🔄 В процессе', -1);
          ss.setActiveSheet(detailSheet);
          SpreadsheetApp.flush();
          smartSortProductsByBrand(false);
          success = true;
        } else {
          console.log(`🔄 [API] Прямая сортировка: ${cleanName}`);
          ss.toast(`Сортировка: ${cleanName}`, '🔄 В процессе', -1);
          success = smartSortCategoryDirectly(categoryId, cleanName, itemsPerPage);
        }

        if (success) {
          processedCount++;
          listSheet.getRange(i + 2, 1).setValue(false); // Снимаем галочку
          SpreadsheetApp.flush();
        }

      } catch (e) {
        console.error(`❌ Ошибка при обработке ${categoryName}:`, e);
      }
    }
  }

  // Если мы здесь - значит прошли весь список
  ss.setActiveSheet(listSheet);

  if (!hasMoreWork) {
    ss.toast('Все отмеченные категории обработаны!', '✅ Готово', 5);

    // Удаляем триггеры (на всякий случай, если остались старые)

    try {
      const triggers = ScriptApp.getProjectTriggers();
      for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === 'runBatchSortWorker') {
          ScriptApp.deleteTrigger(trigger);
        }
      }
    } catch (e) {
      console.warn('⚠️ Не удалось удалить триггеры:', e.message);
    }
  }
}

/**
 * HTML диалога добавления товаров (CSV ВЕРСИЯ С ВЫБОРОМ РАЗДЕЛА)
 */
function createAddProductsDialogHTML(categoryId, categoryTitle, existingIds) {
  const existingIdsJson = JSON.stringify(existingIds);

  // Получаем список импортированных разделов
  const importedSections = getImportedSections();
  const sectionsJson = JSON.stringify(importedSections);

  // Получаем конфигурацию основных характеристик для разделов
  const sectionCharsJson = JSON.stringify(SECTION_CHARACTERISTICS);

  // Если нет импортированных разделов - показываем предупреждение
  if (importedSections.length === 0) {
    return `
      <!DOCTYPE html>
      <html>
        <head><base target="_top"></head>
        <body style="font-family: Arial; padding: 20px;">
          <h3 style="color: #d32f2f;">⚠️ Нет импортированных разделов</h3>
          <p>Сначала импортируйте CSV-выгрузку раздела через меню:</p>
          <p><strong>"📁 Категории" → "📤 Импорт каталога из CSV"</strong></p>
          <br>
          <button onclick="google.script.host.close()" style="padding: 10px 20px; background: #1976d2; color: white; border: none; border-radius: 4px; cursor: pointer;">
            Закрыть
          </button>
        </body>
      </html>
    `;
  }

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <!-- VERSION: ${Date.now()}-CACHE-BUST -->
        <style>
          body { font-family: Arial, sans-serif; padding: 15px; margin: 0; background: #f8f9fa; }
          h3 { color: #1976d2; margin-top: 0; margin-bottom: 15px; }

          .info {
            background: #e3f2fd;
            padding: 10px;
            margin-bottom: 15px;
            border-radius: 4px;
            font-size: 13px;
          }

          /* Секция фильтров */
          .filters-section {
            background: white;
            padding: 15px;
            margin-bottom: 15px;
            border-radius: 4px;
            border: 1px solid #ddd;
          }

          .filters-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            cursor: pointer;
          }

          .filters-header h4 {
            margin: 0;
            color: #333;
            font-size: 14px;
          }

          .toggle-icon {
            font-size: 18px;
            transition: transform 0.2s;
          }

          .toggle-icon.open {
            transform: rotate(90deg);
          }

          .filters-content {
            display: none;
          }

          .filters-content.open {
            display: block;
          }

          .filter-row {
            margin-bottom: 12px;
          }

          .filter-label {
            display: block;
            font-size: 12px;
            font-weight: bold;
            color: #555;
            margin-bottom: 4px;
          }

          .filter-input {
            width: 100%;
            padding: 8px;
            font-size: 14px;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-sizing: border-box;
          }

          .filter-input:focus {
            outline: none;
            border-color: #1976d2;
          }

          .checkbox-filter {
            display: flex;
            align-items: center;
            font-size: 14px;
          }

          .checkbox-filter input {
            margin-right: 8px;
          }

          .filter-actions {
            display: flex;
            gap: 10px;
            margin-top: 12px;
          }

          .btn-filter {
            flex: 1;
            padding: 8px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
          }

          .btn-apply-filters {
            background: #1976d2;
            color: white;
          }

          .btn-reset-filters {
            background: #f1f3f4;
            color: #333;
          }

          /* Поиск */
          .search-box {
            margin-bottom: 15px;
          }

          .search-box input {
            width: 100%;
            padding: 10px;
            font-size: 15px;
            border: 2px solid #1976d2;
            border-radius: 4px;
            box-sizing: border-box;
          }

          .search-box input:focus {
            outline: none;
            border-color: #1565c0;
          }

          /* Результаты */
          .results {
            background: white;
            border-radius: 4px;
            border: 1px solid #ddd;
            max-height: 400px;
            overflow-y: auto;
          }

          table {
            width: 100%;
            border-collapse: collapse;
          }

          th {
            position: sticky;
            top: 0;
            background: #1976d2;
            color: white;
            padding: 10px;
            text-align: left;
            font-size: 13px;
            z-index: 10;
          }

          td {
            padding: 8px 10px;
            border-bottom: 1px solid #eee;
            font-size: 13px;
          }

          tr:hover {
            background: #f5f5f5;
          }

          .checkbox-cell {
            width: 40px;
            text-align: center;
          }

          .stock-yes { color: #4caf50; font-weight: bold; }
          .stock-no { color: #f44336; }
          .in-category { color: #999; font-style: italic; }

          .loading {
            text-align: center;
            padding: 40px;
            color: #666;
          }

          .loading-spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #1976d2;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }

          /* Кнопки действий */
          .actions {
            margin-top: 15px;
            text-align: right;
            padding-top: 15px;
            border-top: 1px solid #ddd;
          }

          button {
            padding: 10px 20px;
            margin-left: 10px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
          }

          .btn-cancel {
            background: #f1f3f4;
            color: #333;
          }

          .btn-primary {
            background: #1976d2;
            color: white;
          }

          .btn-primary:disabled {
            background: #ccc;
            cursor: not-allowed;
          }

          .btn-load-more {
            background: #4caf50;
            color: white;
            width: 100%;
            margin-top: 10px;
          }
        </style>
      </head>
      <body>
        <h3>➕ Добавить товары в категорию</h3>

        <div class="info">
          📁 <strong>${categoryTitle}</strong> |
          📦 В категории: <strong>${existingIds.length}</strong> |
          ✅ Выбрано: <strong id="selectedCount">0</strong>
        </div>

        <!-- ВЫБОР РАЗДЕЛА -->
        <div style="background: white; padding: 15px; margin-bottom: 15px; border-radius: 4px; border: 1px solid #ddd;">
          <label style="display: block; font-weight: bold; margin-bottom: 8px; font-size: 14px;">
            📂 Выберите раздел каталога:
          </label>
          <select id="sectionSelect" style="width: 100%; padding: 10px; font-size: 14px; border: 1px solid #ddd; border-radius: 4px;">
            <option value="">-- Выберите раздел --</option>
          </select>
        </div>

        <!-- ПОИСК (ПЕРЕНЕСЕНО НАВЕРХ) -->
        <div class="search-box">
          <input
            type="text"
            id="searchInput"
            placeholder="🔍 Поиск по названию или артикулу..."
            autocomplete="off"
          >
        </div>

        <!-- ФИЛЬТРЫ -->
        <div class="filters-section">
          <div class="filters-header" onclick="toggleFilters()">
            <h4>🔍 Фильтры по характеристикам</h4>
            <span class="toggle-icon" id="toggleIcon">▶</span>
          </div>

          <div class="filters-content" id="filtersContent">
            <div id="dynamicFilters">
              <!-- Фильтры будут добавлены динамически -->
            </div>

            <div class="filter-row">
              <div class="checkbox-filter">
                <input type="checkbox" id="filterInStock">
                <label for="filterInStock">Только товары в наличии</label>
              </div>
            </div>

            <div class="filter-actions">
              <button class="btn-filter btn-apply-filters" onclick="applyFilters()">Применить</button>
              <button class="btn-filter btn-reset-filters" onclick="resetFilters()">Сбросить</button>
            </div>
          </div>
        </div>

        <!-- РЕЗУЛЬТАТЫ -->
        <div class="results">
          <table>
            <thead>
              <tr>
                <th class="checkbox-cell">☑️</th>
                <th>Название товара</th>
                <th style="width: 120px;">Админка</th>
                <th style="width: 100px;">Цена</th>
                <th style="width: 80px;">Наличие</th>
              </tr>
            </thead>
            <tbody id="resultsBody">
              <tr>
                <td colspan="5" class="loading">
                  <strong>📂 Выберите раздел каталога</strong><br>
                  <small>Используйте выпадающий список выше</small>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        <button id="loadMoreBtn" class="btn-load-more" style="display: none;" onclick="loadMoreProducts()">
          📄 Показать ещё 100 товаров
        </button>

        <!-- ДЕЙСТВИЯ -->
        <div class="actions">
          <button class="btn-cancel" onclick="google.script.host.close()">Отмена</button>
          <button id="addBtn" class="btn-primary" onclick="addSelectedProducts()" disabled>
            ➕ Добавить (<span id="addBtnCount">0</span>)
          </button>
        </div>

        <script>
          let allProducts = [];
          let filteredProducts = [];
          let displayedCount = 0;
          const batchSize = 100;
          let selectedProducts = new Set();
          let searchTimeout;
          const existingIds = new Set(${existingIdsJson});
          const sections = ${sectionsJson};
          const sectionCharacteristics = ${sectionCharsJson};
          let currentMetadata = { characteristics: {} };
          let activeFilters = {};
          let selectedSection = '';

          window.onload = function() {
            renderSections();

            document.getElementById('sectionSelect').addEventListener('change', function(e) {
              selectedSection = e.target.value;
              if (selectedSection) {
                loadProductsFromSection(selectedSection);
              } else {
                allProducts = [];
                filteredProducts = [];
                currentMetadata = { characteristics: {} };
                displayResults();
                buildDynamicFilters();
              }
            });

            // ИСПРАВЛЕНО: Поиск срабатывает сразу при вводе
            document.getElementById('searchInput').addEventListener('input', function(e) {
              clearTimeout(searchTimeout);
              searchTimeout = setTimeout(() => {
                console.log('🔍 Поиск активирован, значение:', e.target.value);
                applyFilters();
              }, 300);
            });
          };

          function renderSections() {
            const select = document.getElementById('sectionSelect');
            sections.forEach(section => {
              const option = document.createElement('option');
              option.value = section;
              option.textContent = section;
              select.appendChild(option);
            });
          }

          function loadProductsFromSection(sectionName) {
            document.getElementById('resultsBody').innerHTML =
              '<tr><td colspan="5" class="loading">' +
              '<div class="loading-spinner"></div>' +
              '<strong>Загружаем товары из раздела "' + sectionName + '"...</strong>' +
              '</td></tr>';

            google.script.run
              .withSuccessHandler(function(data) {
                console.log('📦 Загружено товаров:', data.products.length);

                // ОТЛАДКА КОЛОНКИ КРАТНОСТИ
                if (data._debug_kratnost) {
                  console.log('🔍 DEBUG: Колонка "Кратность увеличения, крат"');
                  console.log('  - Индекс колонки в листе:', data._debug_kratnost.index);
                  console.log('  - Заголовок колонки:', data._debug_kratnost.header);
                  console.log('  - Первые 5 значений из ЭТОЙ колонки:', data._debug_kratnost.firstValues);
                }

                // ВАЖНЫЙ ЛОГ: Показываем ВСЕ характеристики
                console.log('🔍 ВСЕ ХАРАКТЕРИСТИКИ - ключи:', Object.keys(data.characteristics));

                // Ищем характеристики с подозрительными значениями типа "1##4"
                for (const [key, values] of Object.entries(data.characteristics)) {
                  if (values && values.length > 0) {
                    const firstValue = values[0].toString();
                    if (firstValue.includes('#' + '#')) {
                      console.log('🚨 НАЙДЕНА характеристика с "#' + '#": "' + key + '"');
                      console.log('   Первые 5 значений:', values.slice(0, 5));
                    }
                  }
                }

                const kratnostValues = data.characteristics['Кратность увеличения, крат'];
                console.log('🔍 КРАТНОСТЬ - массив значений (из метаданных):', kratnostValues);
                console.log('🔍 КРАТНОСТЬ - первые 5 значений (из метаданных):', kratnostValues ? kratnostValues.slice(0, 5) : 'НЕТ ДАННЫХ');

                allProducts = data.products;
                filteredProducts = data.products;
                currentMetadata = { characteristics: data.characteristics };
                buildDynamicFilters();
                applyFilters();
              })
              .withFailureHandler(function(error) {
                console.error('❌ Ошибка загрузки товаров:', error);
                document.getElementById('resultsBody').innerHTML =
                  '<tr><td colspan="5" class="loading">' +
                  '❌ Ошибка: ' + error.message +
                  '</td></tr>';
              })
              .getProductsFromCSVSheet(sectionName);
          }

          function toggleFilters() {
            const content = document.getElementById('filtersContent');
            const icon = document.getElementById('toggleIcon');
            content.classList.toggle('open');
            icon.classList.toggle('open');
          }

          function buildDynamicFilters() {
            const container = document.getElementById('dynamicFilters');
            const characteristics = currentMetadata.characteristics || {};

            // Получаем список основных характеристик для текущего раздела
            const mainChars = sectionCharacteristics[selectedSection] || [];

            let html = '';

            // Показываем только основные характеристики в заданном порядке
            for (const charName of mainChars) {
              if (characteristics[charName] && characteristics[charName].length > 0) {
                const values = characteristics[charName];
                // ИСПРАВЛЕНО: Создаём безопасный ID без специальных символов
                const safeId = 'filter_' + charName.replace(/[^a-zA-Zа-яА-Я0-9]/g, '_');
                html += \`
                  <div class="filter-row">
                    <label class="filter-label">\${charName}</label>
                    <select class="filter-input" id="\${safeId}" data-characteristic="\${charName}">
                      <option value="">Все</option>
                      \${values.map(v => \`<option value="\${v}">\${v}</option>\`).join('')}
                    </select>
                  </div>
                \`;
              }
            }

            // Если нет ни одной основной характеристики, показываем все
            if (html === '') {
              for (const [key, values] of Object.entries(characteristics)) {
                if (values && values.length > 0) {
                  const safeId = 'filter_' + key.replace(/[^a-zA-Zа-яА-Я0-9]/g, '_');
                  html += \`
                    <div class="filter-row">
                      <label class="filter-label">\${key}</label>
                      <select class="filter-input" id="\${safeId}" data-characteristic="\${key}">
                        <option value="">Все</option>
                        \${values.map(v => \`<option value="\${v}">\${v}</option>\`).join('')}
                      </select>
                    </div>
                  \`;
                }
              }
            }

            container.innerHTML = html;
          }

          // Функция удалена - теперь используется loadProductsFromSection

          function applyFilters() {
            const searchQuery = document.getElementById('searchInput').value.toLowerCase();
            const inStockOnly = document.getElementById('filterInStock').checked;

            // Собираем активные фильтры по характеристикам
            // ИСПРАВЛЕНО: Используем data-attribute для получения имени характеристики
            activeFilters = {};
            const filterSelects = document.querySelectorAll('.filter-input[data-characteristic]');
            filterSelects.forEach(select => {
              const charName = select.getAttribute('data-characteristic');
              if (select.value) {
                activeFilters[charName] = select.value;
              }
            });

            console.log('🔍 Применяем фильтры:', {
              searchQuery,
              inStockOnly,
              activeFilters,
              allProductsCount: allProducts.length
            });

            // Фильтруем товары
            let debugCount = 0;
            filteredProducts = allProducts.filter(product => {
              // Поиск по названию ИЛИ артикулу
              if (searchQuery) {
                // ИСПРАВЛЕНО: title и sku уже строки после getProductsFromCSVSheet
                const titleMatch = product.title.toLowerCase().includes(searchQuery);
                const skuMatch = product.sku && product.sku.toLowerCase().includes(searchQuery);

                // Логируем только первые 3 товара для отладки
                if (debugCount < 3) {
                  console.log('🔍 Товар:', product.title, '| titleMatch:', titleMatch, '| skuMatch:', skuMatch);
                  debugCount++;
                }

                if (!titleMatch && !skuMatch) {
                  return false;
                }
              }

              // Фильтр наличия
              if (inStockOnly && !product.in_stock) {
                return false;
              }

              // Фильтры по характеристикам
              for (const [key, value] of Object.entries(activeFilters)) {
                if (!product.characteristics[key] || product.characteristics[key] !== value) {
                  return false;
                }
              }

              return true;
            });

            console.log(\`✅ Отфильтровано: \${filteredProducts.length} из \${allProducts.length}\`);

            displayedCount = 0;
            displayResults();
          }

          function resetFilters() {
            document.getElementById('searchInput').value = '';
            document.getElementById('filterInStock').checked = false;

            // ИСПРАВЛЕНО: Сбрасываем все select-элементы с data-attribute
            const filterSelects = document.querySelectorAll('.filter-input[data-characteristic]');
            filterSelects.forEach(select => {
              select.value = '';
            });

            activeFilters = {};
            applyFilters();
          }

          function displayResults() {
            const tbody = document.getElementById('resultsBody');

            if (filteredProducts.length === 0) {
              tbody.innerHTML = '<tr><td colspan="5" class="loading">😔 Товары не найдены<br><small>Попробуйте изменить фильтры</small></td></tr>';
              document.getElementById('loadMoreBtn').style.display = 'none';
              return;
            }

            const endIndex = Math.min(displayedCount + batchSize, filteredProducts.length);
            const batch = filteredProducts.slice(displayedCount, endIndex);

            const html = batch.map((product, index) => {
              // ОТЛАДКА: Логируем первый товар для проверки ID
              if (index === 0) {
                console.log('🔍 DEBUG первого товара:', {
                  id: product.id,
                  title: product.title,
                  idType: typeof product.id
                });
              }

              const inCategory = existingIds.has(product.id);
              const rowClass = inCategory ? 'in-category' : '';
              const checkbox = inCategory ?
                '<input type="checkbox" disabled title="Уже в категории">' :
                \`<input type="checkbox" value="\${product.id}" onchange="toggleProduct(\${product.id})" \${selectedProducts.has(product.id) ? 'checked' : ''}>\`;

              // ИСПРАВЛЕНО: Ссылка на админку товара вместо характеристик
              const price = product.price ? product.price.toFixed(0) + ' ₽' : '—';
              const adminUrl = 'https://binokl.shop/admin2/products/' + product.id;

              return '<tr class="' + rowClass + '">' +
                '<td class="checkbox-cell">' + checkbox + '</td>' +
                '<td>' + product.title + (inCategory ? '<small>(уже в категории)</small>' : '') + '</td>' +
                '<td style="text-align: center;"><a href="' + adminUrl + '" target="_blank" title="Открыть в админке InSales">🔗 Открыть</a></td>' +
                '<td>' + price + '</td>' +
                '<td class="' + (product.in_stock ? 'stock-yes' : 'stock-no') + '">' +
                  (product.in_stock ? '✅' : '❌') +
                '</td>' +
              '</tr>';
            }).join('');

            if (displayedCount === 0) {
              tbody.innerHTML = html;
            } else {
              tbody.innerHTML += html;
            }

            displayedCount = endIndex;

            if (displayedCount < filteredProducts.length) {
              document.getElementById('loadMoreBtn').style.display = 'block';
              document.getElementById('loadMoreBtn').textContent =
                \`📄 Показать ещё 100 (осталось \${filteredProducts.length - displayedCount})\`;
            } else {
              document.getElementById('loadMoreBtn').style.display = 'none';
            }
          }

          function loadMoreProducts() {
            displayResults();
          }

          function toggleProduct(productId) {
            if (selectedProducts.has(productId)) {
              selectedProducts.delete(productId);
            } else {
              selectedProducts.add(productId);
            }
            updateSelectedCount();
          }

          function updateSelectedCount() {
            const count = selectedProducts.size;
            document.getElementById('selectedCount').textContent = count;
            document.getElementById('addBtnCount').textContent = count;
            document.getElementById('addBtn').disabled = count === 0;
          }

          function addSelectedProducts() {
            if (selectedProducts.size === 0) {
              alert('Выберите товары');
              return;
            }

            const productIds = Array.from(selectedProducts);
            const addBtn = document.getElementById('addBtn');

            addBtn.disabled = true;
            addBtn.innerHTML = '⏳ Добавляем...';

            google.script.run
              .withSuccessHandler(function(result) {
                alert(\`✅ Добавлено: \${result.success}\` + (result.errors > 0 ? \`\\n❌ Ошибок: \${result.errors}\` : ''));
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('❌ Ошибка: ' + error.message);
                addBtn.disabled = false;
                addBtn.innerHTML = '➕ Добавить (<span id="addBtnCount">' + productIds.length + '</span>)';
              })
              .addProductsToCategory(${categoryId}, productIds);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Загружает ВСЕ товары для поиска
 */
function getAllProductsForSearch() {
  try {
    console.log('📦 Загружаем ВСЕ товары для поиска');

    const credentials = getInsalesCredentialsSync();
    if (!credentials) throw new Error('Учетные данные не настроены');

    const allProducts = [];
    let page = 1;
    const perPage = 250;

    while (page <= 20) {
      const url = `${credentials.baseUrl}/admin/products.json?per_page=${perPage}&page=${page}`;

      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) break;

      const products = JSON.parse(response.getContentText());

      if (!products || products.length === 0) break;

      allProducts.push(...products);

      if (products.length < perPage) break;

      page++;
      Utilities.sleep(300);
    }

    console.log(`✅ Загружено товаров: ${allProducts.length}`);

    const simplified = allProducts.map(product => {
      const variant = product.variants && product.variants[0];
      const inStock = variant && variant.quantity > 0;

      return {
        id: product.id,
        title: product.title,
        in_stock: inStock
      };
    });

    return simplified;

  } catch (error) {
    console.error('❌ Ошибка загрузки товаров:', error);
    throw error;
  }
}

/**
 * Полная синхронизация товаров в категории InSales.
 * Сравнивает текущий список в InSales с целевым списком:
 * - Добавляет товары, которых нет в категории
 * - Удаляет товары, которых нет в целевом списке
 * @param {number} categoryId - ID категории InSales
 * @param {number[]} targetProductIds - целевой список ID товаров
 * @returns {{added: number, removed: number, kept: number, errors: number}}
 */
function syncProductsInCategory(categoryId, targetProductIds) {
  var context = 'syncProductsInCategory';
  try {
    console.log('🔄 Синхронизация товаров для категории ' + categoryId + ', целевых: ' + targetProductIds.length);

    var credentials = getInsalesCredentialsSync();
    if (!credentials) throw new Error('Учетные данные не настроены');

    // ШАГ 1: Получаем текущие связи из InSales
    var currentCollects = getAllCollectsForCategory(categoryId, credentials);
    console.log('📋 Текущих товаров в InSales: ' + currentCollects.length);

    // Строим карту: product_id → collect_id
    var currentMap = {};
    for (var i = 0; i < currentCollects.length; i++) {
      currentMap[String(currentCollects[i].product_id)] = currentCollects[i].id;
    }

    // Строим Set целевых ID
    var targetSet = {};
    for (var t = 0; t < targetProductIds.length; t++) {
      targetSet[String(targetProductIds[t])] = true;
    }

    // ШАГ 2: Определяем что добавить и что удалить
    var toAdd = [];
    for (var t2 = 0; t2 < targetProductIds.length; t2++) {
      var pid = String(targetProductIds[t2]);
      if (!currentMap[pid]) {
        toAdd.push(parseInt(pid));
      }
    }

    var toRemove = []; // {collectId, productId}
    var currentProductIds = Object.keys(currentMap);
    for (var c = 0; c < currentProductIds.length; c++) {
      var cpid = currentProductIds[c];
      if (!targetSet[cpid]) {
        toRemove.push({ collectId: currentMap[cpid], productId: parseInt(cpid) });
      }
    }

    var kept = targetProductIds.length - toAdd.length;
    console.log('➕ К добавлению: ' + toAdd.length + ', 🗑️ К удалению: ' + toRemove.length + ', ✅ Остаются: ' + kept);

    var added = 0, removed = 0, errors = 0;

    // ШАГ 3: Удаляем лишние
    for (var r = 0; r < toRemove.length; r++) {
      try {
        var deleteUrl = credentials.baseUrl + '/admin/collects/' + toRemove[r].collectId + '.json';
        var delResp = UrlFetchApp.fetch(deleteUrl, {
          method: 'DELETE',
          headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(credentials.apiKey + ':' + credentials.password),
            'Content-Type': 'application/json'
          },
          muteHttpExceptions: true
        });
        var delCode = delResp.getResponseCode();
        if (delCode === 200 || delCode === 204) {
          removed++;
          console.log('🗑️ Удалён товар ' + toRemove[r].productId);
        } else {
          errors++;
          console.error('❌ Ошибка удаления товара ' + toRemove[r].productId + ': ' + delCode);
        }
        Utilities.sleep(300);
      } catch (e) {
        errors++;
        console.error('❌ Ошибка удаления: ' + e.message);
      }
    }

    // ШАГ 4: Добавляем новые
    for (var a = 0; a < toAdd.length; a++) {
      try {
        var addUrl = credentials.baseUrl + '/admin/collects.json';
        var addResp = UrlFetchApp.fetch(addUrl, {
          method: 'POST',
          headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(credentials.apiKey + ':' + credentials.password),
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify({ collect: { product_id: toAdd[a], collection_id: parseInt(categoryId) } }),
          muteHttpExceptions: true
        });
        var addCode = addResp.getResponseCode();
        if (addCode === 200 || addCode === 201) {
          added++;
          console.log('➕ Добавлен товар ' + toAdd[a]);
        } else if (addCode === 422) {
          // Уже в категории — не считаем ошибкой
          console.log('⚠️ Товар ' + toAdd[a] + ' уже в категории');
          added++;
        } else {
          errors++;
          console.error('❌ Ошибка добавления товара ' + toAdd[a] + ': ' + addCode);
        }
        Utilities.sleep(300);
      } catch (e) {
        errors++;
        console.error('❌ Ошибка добавления: ' + e.message);
      }
    }

    console.log('🔄 Синхронизация завершена: добавлено ' + added + ', удалено ' + removed + ', осталось ' + kept + ', ошибок ' + errors);
    return { added: added, removed: removed, kept: kept, errors: errors };

  } catch (error) {
    console.error('❌ Ошибка синхронизации товаров: ' + error.message);
    throw error;
  }
}

/**
 * Добавляет товары в КОНЕЦ списка
 */
function appendProductsToDetailSheet(products) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;

    console.log(`[DEBUG appendProducts] calculateSheetSections вернул productsStart = ${startRow}`);
    console.log(`[DEBUG appendProducts] Последняя строка листа = ${sheet.getLastRow()}`);

    // Ищем последнюю строку с товаром (где есть ID в колонке E)
    let nextRow = startRow;
    let debugRows = [];
    while (sheet.getRange(nextRow, 5).getValue() !== '') {  // Колонка E - ID товара
      const idValue = sheet.getRange(nextRow, 5).getValue();
      debugRows.push(`Строка ${nextRow}: ID = ${idValue}`);
      nextRow++;
      if (nextRow > 1000) break; // Защита от бесконечного цикла
    }

    console.log(`[DEBUG appendProducts] Проверенные строки с ID:\n${debugRows.join('\n')}`);
    console.log(`📝 Добавляем ${products.length} товаров начиная со строки ${nextRow}`);

    const productRows = products.map(product => {
      const variant = product.variants && product.variants[0];
      const inStock = variant && variant.quantity > 0 ? 'Да' : 'Нет';
      const price = variant ? variant.price : product.price;

      let characteristics = '';
      if (product.characteristics && product.characteristics.length > 0) {
        characteristics = product.characteristics.slice(0, 3).map(ch =>
          `${ch.property_title || ''}: ${ch.title || ch.name || ''}`
        ).join(', ');
      }

      return [
        product.title,                   // A - Название
        variant ? variant.sku : '',      // B - Артикул
        price || '',                     // C - Цена
        inStock,                         // D - В наличии
        product.id,                      // E - ID
        false                            // F - Чекбокс
      ];
    });

    // КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Вставляем новые строки перед записью данных!
    // Это гарантирует, что мы не перезапишем существующие блоки (дополнительные поля, плитки тегов и т.д.)
    console.log(`[DEBUG appendProducts] Вставляем ${productRows.length} новых строк после строки ${nextRow - 1}`);
    sheet.insertRowsAfter(nextRow - 1, productRows.length);

    // Теперь записываем данные в новые пустые строки
    sheet.getRange(nextRow, 1, productRows.length, productRows[0].length).setValues(productRows);
    sheet.getRange(nextRow, 6, productRows.length, 1).insertCheckboxes();  // F - чекбокс
    sheet.getRange(nextRow, 3, productRows.length, 1).setNumberFormat('#,##0.00 ₽');  // C - цена

    console.log(`✅ Добавлено ${productRows.length} товаров в таблицу`);

    updateCategoryStatistics(sheet);

  } catch (error) {
    console.error('❌ Ошибка добавления товаров в таблицу:', error);
    throw error;
  }
}

/**
 * Обновляет статистику категории
 */
function updateCategoryStatistics(sheet) {
  try {
    // ИСПРАВЛЕНО: Используем динамическое позиционирование
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const statsStartRow = sections.statsStart;
    const lastRow = sheet.getLastRow();

    console.log(`[DEBUG updateStats] statsStartRow = ${statsStartRow}, productsStart = ${startRow}`);

    // Если блок статистики не найден, используем старые константы как fallback
    if (!statsStartRow) {
      console.log('[DEBUG updateStats] Блок статистики не найден, пропускаем обновление');
      return;
    }

    if (lastRow < startRow) {
      // Нет товаров - записываем нули
      sheet.getRange(statsStartRow + 1, 2).setValue(0);      // B74: Всего товаров
      sheet.getRange(statsStartRow + 2, 2).setValue(0);      // B75: В наличии
      sheet.getRange(statsStartRow + 3, 2).setValue(0);      // B76: Нет в наличии
      sheet.getRange(statsStartRow + 4, 2).setValue('0%');   // B77: Процент
      return;
    }

    const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();

    let totalCount = 0;
    let inStockCount = 0;

    for (let i = 0; i < data.length; i++) {
      const id = data[i][4];  // E - ID товара
      if (!id || id.toString().trim() === '') break;

      totalCount++;

      const inStock = data[i][3];  // D - В наличии
      if (inStock === 'Да') {
        inStockCount++;
      }
    }

    const outOfStockCount = totalCount - inStockCount;
    const percentInStock = totalCount > 0 ? Math.round(inStockCount / totalCount * 100) + '%' : '0%';

    // Записываем в динамические позиции
    sheet.getRange(statsStartRow + 1, 2).setValue(totalCount);      // Всего товаров
    sheet.getRange(statsStartRow + 2, 2).setValue(inStockCount);    // В наличии
    sheet.getRange(statsStartRow + 3, 2).setValue(outOfStockCount); // Нет в наличии
    sheet.getRange(statsStartRow + 4, 2).setValue(percentInStock);  // Процент

    console.log(`[DEBUG updateStats] Обновлена статистика: ${totalCount} товаров, ${inStockCount} в наличии`);

    console.log(`📊 Статистика обновлена: всего ${totalCount}, в наличии ${inStockCount}`);

  } catch (error) {
    console.error('❌ Ошибка обновления статистики:', error);
  }
}

/**
  console.warn('⚠️ Используется устаревшая функция getInSalesCredentials. Пожалуйста, обновите код на getInsalesCredentialsSync.');
  return getInsalesCredentialsSync();
}
 
 
/**
 * Автоматическая сортировка всех категорий (для запуска по триггеру)
 * Проходит по всем детальным листам категорий и применяет умную сортировку.
 */
function autoSortAllCategories() {
  try {
    console.log('🚀 Запуск автоматической сортировки всех категорий...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let processedCount = 0;

    // Получаем префикс из конфига или используем дефолтный
    const prefix = (typeof CATEGORY_SHEETS !== 'undefined' && CATEGORY_SHEETS.DETAIL_PREFIX)
      ? CATEGORY_SHEETS.DETAIL_PREFIX
      : 'Категория — ';

    for (const sheet of sheets) {
      if (sheet.getName().startsWith(prefix)) {
        console.log(`🔄 Обработка категории: ${sheet.getName()}`);

        // Делаем лист активным, так как smartSortProductsByBrand использует getActiveSheet()
        ss.setActiveSheet(sheet);

        // Запускаем сортировку в неинтерактивном режиме
        smartSortProductsByBrand(false);

        processedCount++;

        // Небольшая пауза между категориями
        Utilities.sleep(1000);
      }
    }

    console.log(`✅ Автоматическая сортировка завершена. Обработано категорий: ${processedCount}`);

  } catch (error) {
    console.error('❌ Критическая ошибка при автоматической сортировке:', error);
  }
}

/**
 * Ensures the collection is set to manual sorting (sort_type = 0)
 */
function ensureCollectionIsManualSorted(categoryId, credentials) {
  try {
    const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;

    // 1. Check current sort type
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      console.error(`❌ Error fetching collection ${categoryId}: ${response.getResponseCode()}`);
      return false;
    }

    const data = JSON.parse(response.getContentText());

    // InSales: sort_type 0 usually means "Manual" (or "By position")
    // Update: User confirmed sort_type 7 is also "Manual" for this shop
    if (data.sort_type == 0 || data.sort_type == 7) {
      console.log(`✅ Collection ${categoryId} is already manually sorted (type: ${data.sort_type}).`);
      return true;
    }

    console.log(`⚠️ Collection ${categoryId} sort_type is ${data.sort_type}. Switching to Manual (0)...`);

    // 2. Update sort type
    const payload = {
      collection: {
        sort_type: "0"
      }
    };

    const updateResponse = UrlFetchApp.fetch(url, {
      method: 'PUT',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (updateResponse.getResponseCode() === 200) {
      console.log(`✅ Collection ${categoryId} switched to Manual sorting.`);
      return true;
    } else if (updateResponse.getResponseCode() === 422) {
      console.warn(`⚠️ Failed to set manual sorting (422): ${updateResponse.getContentText()}`);
      console.warn(`➡️ Proceeding anyway, as position updates might still work with current sort_type.`);
      return true; // Allow proceeding
    } else {
      console.error(`❌ Failed to set manual sorting: ${updateResponse.getResponseCode()} ${updateResponse.getContentText()}`);
      return false;
    }

  } catch (e) {
    console.error(`❌ Exception in ensureCollectionIsManualSorted: ${e.message}`);
    // Return true to avoid blocking the main process on minor errors
    return true;
  }
}