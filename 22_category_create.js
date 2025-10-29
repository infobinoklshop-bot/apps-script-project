/**
 * ========================================
 * СОЗДАНИЕ НОВОЙ КАТЕГОРИИ
 * ========================================
 */

/**
 * Показывает диалог создания новой категории
 */
function showCreateCategoryDialog() {
  const html = HtmlService.createHtmlOutputFromFile('22_category_create_dialog')
    .setWidth(500)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Создание категории');
}

/**
 * Получает список категорий для dropdown родителя
 */
function getCategoriesForParentSelect() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!sheet) {
      console.log('[WARNING] Лист "' + CATEGORY_SHEETS.MAIN_LIST + '" не найден');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const categories = [];

    // Пропускаем заголовок (строка 0)
    for (let i = 1; i < data.length; i++) {
      const id = data[i][MAIN_LIST_COLUMNS.CATEGORY_ID - 1];      // Колонка B - ID
      const title = data[i][MAIN_LIST_COLUMNS.TITLE - 1];         // Колонка F - Название
      const level = data[i][MAIN_LIST_COLUMNS.LEVEL - 1];         // Колонка D - Уровень
      const path = data[i][MAIN_LIST_COLUMNS.HIERARCHY_PATH - 1]; // Колонка E - Путь

      // Пропускаем пустые строки
      if (!id || !title) continue;

      // Убираем визуальные символы из названия
      const cleanTitle = String(title).replace(/[\s└─]/g, '').trim();

      categories.push({
        id: id,
        title: cleanTitle,
        level: level || 0,
        path: path || ''
      });
    }
    
    console.log('[INFO] Загружено категорий для выбора:', categories.length);
    
    return categories;
    
  } catch (error) {
    console.error('[ERROR] Ошибка загрузки категорий:', error);
    return [];
  }
}

/**
 * Создаёт новую категорию с минимальными данными
 */
function createNewCategoryMinimal(data) {
  try {
    console.log('[INFO] Получены данные:', JSON.stringify(data));
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Создаём категорию "' + data.title + '"...',
      '⏳ Создание',
      -1
    );
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные InSales');
    }
    
    const url = generateUrl(data.title);
    
    // Формируем payload с обязательными полями
    const collectionData = {
      title: data.title,
      url: url,
      html_title: data.title,
      is_hidden: false,
      position: 1,
      sort_type: 1
    };

    // parent_id: добавляем ВСЕГДА (API требует)
    if (data.parent_id && data.parent_id !== null && data.parent_id !== '') {
      collectionData.parent_id = parseInt(data.parent_id);
      console.log('[INFO] Категория с родителем:', collectionData.parent_id);
    } else {
      // API InSales не позволяет создавать корневые категории
      // Все новые категории создаются в "Каталог" (ID: 9069711)
      collectionData.parent_id = 9069711;
      console.log('[INFO] Категория без указанного родителя - создаем в "Каталог" (9069711)');
    }
    
    const payload = {
      collection: collectionData
    };
    
    console.log('[INFO] Отправляемый payload:', JSON.stringify(payload, null, 2));
    
    const apiUrl = `${credentials.baseUrl}/admin/collections.json`;
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('[INFO] HTTP Status:', statusCode);
    console.log('[INFO] Response:', responseText);
    
    if (statusCode === 201) {
      const newCategory = JSON.parse(responseText);
      
      console.log('[INFO] ✅ Категория создана! ID:', newCategory.id);
      
      SpreadsheetApp.getActiveSpreadsheet().toast('Категория создана!', '✅ Готово', 3);

      // Перезагружаем список категорий
      loadCategoriesWithHierarchy();
      
      if (data.openDetail) {
        Utilities.sleep(1000);
        createDetailedCategorySheet({
          id: newCategory.id,
          title: newCategory.title,
          url: newCategory.url
        });
      }
      
      return { 
        success: true, 
        categoryId: newCategory.id,
        url: newCategory.url
      };
      
    } else {
      let errorMsg = 'HTTP ' + statusCode;
      try {
        const errorData = JSON.parse(responseText);
        if (errorData.errors) {
          errorMsg += ': ' + JSON.stringify(errorData.errors);
        }
      } catch (e) {
        errorMsg += ': ' + responseText;
      }
      
      throw new Error(errorMsg);
    }
    
  } catch (error) {
    console.error('[ERROR] Ошибка создания категории:', error.message);
    console.error('[ERROR] Stack:', error.stack);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка создания', '❌ Ошибка', 5);
    
    return { success: false, error: error.message };
  }
}

/**
 * Генерирует ЧПУ URL из названия
 */
function generateUrl(title) {
  const translit = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
    'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'h', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch',
    'ы': 'y', 'э': 'e', 'ю': 'yu', 'я': 'ya', ' ': '-'
  };
  
  let url = title.toLowerCase();
  
  for (let char in translit) {
    url = url.replace(new RegExp(char, 'g'), translit[char]);
  }
  
  url = url.replace(/[^a-z0-9-]/g, '-');
  url = url.replace(/-+/g, '-');
  url = url.replace(/^-+|-+$/g, '');
  
  return url;
}

/**
 * ТЕСТОВАЯ ФУНКЦИЯ - запустить вручную для проверки
 */
function testParentIdProcessing() {
  // Тест 1: пустая строка
  const test1 = { title: 'Тест', parent_id: '', openDetail: false };
  console.log('Тест 1 (пустая строка):', createNewCategoryMinimal(test1));

  // Тест 2: null
  const test2 = { title: 'Тест', parent_id: null, openDetail: false };
  console.log('Тест 2 (null):', createNewCategoryMinimal(test2));

  // Тест 3: число
  const test3 = { title: 'Тест', parent_id: 9175197, openDetail: false };
  console.log('Тест 3 (число):', createNewCategoryMinimal(test3));
}

/**
 * ТЕСТ: Создание подкатегории с известным родителем
 */
function testCreateWithKnownParent() {
  // Создаем подкатегорию внутри "Каталог" (ID 9069711)
  const result = createNewCategoryMinimal({
    title: 'ТЕСТ подкатегория',
    parent_id: 9069711,  // ID категории "Каталог"
    openDetail: false
  });

  console.log('Результат создания с родителем:', JSON.stringify(result, null, 2));
  return result;
}

/**
 * ТЕСТ: Проверяем как API возвращает parent_id для корневой категории
 */
function testCheckRootCategoryParentId() {
  try {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные');
    }

    // Загружаем категорию "Каталог" (ID 9069711)
    const url = `${credentials.baseUrl}/admin/collections/9069711.json`;

    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const category = JSON.parse(response.getContentText());

    console.log('=== КОРНЕВАЯ КАТЕГОРИЯ "Каталог" ===');
    console.log('ID:', category.id);
    console.log('Title:', category.title);
    console.log('parent_id:', category.parent_id);
    console.log('parent_id тип:', typeof category.parent_id);
    console.log('parent_id === null:', category.parent_id === null);
    console.log('parent_id === undefined:', category.parent_id === undefined);
    console.log('\nВесь объект (первые 500 символов):');
    console.log(JSON.stringify(category, null, 2).substring(0, 500));

    return category;

  } catch (error) {
    console.error('Ошибка:', error.message);
    throw error;
  }
}

/**
 * ТЕСТ: Создание корневой категории напрямую (БЕЗ parent_id)
 */
function testCreateRootCategory() {
  try {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
      throw new Error('Не удалось получить учетные данные');
    }

    // Пробуем БЕЗ поля parent_id
    const payload = {
      collection: {
        title: 'ТЕСТ Корневая без parent_id',
        url: 'test-kornevaya-bez-parent-id',
        html_title: 'ТЕСТ Корневая без parent_id',
        is_hidden: false,
        position: 1,
        sort_type: 1
        // parent_id НЕ указываем вообще
      }
    };

    console.log('[INFO] Создаем корневую категорию с payload:');
    console.log(JSON.stringify(payload, null, 2));

    const apiUrl = `${credentials.baseUrl}/admin/collections.json`;

    const options = {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log('[INFO] HTTP Status:', statusCode);
    console.log('[INFO] Response:', responseText);

    if (statusCode === 201) {
      const newCategory = JSON.parse(responseText);
      console.log('[INFO] ✅ Корневая категория создана! ID:', newCategory.id);
      return { success: true, category: newCategory };
    } else {
      console.error('[ERROR] Ошибка создания:', statusCode, responseText);
      return { success: false, status: statusCode, error: responseText };
    }

  } catch (error) {
    console.error('[ERROR] Исключение:', error.message);
    return { success: false, error: error.message };
  }
}