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
    const sheet = ss.getSheetByName('Список категорий');
    
    if (!sheet) {
      console.log('[WARNING] Лист "Список категорий" не найден');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const categories = [];
    
    // Пропускаем заголовок (строка 0)
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0];        // Колонка A - ID
      const title = data[i][2];     // Колонка C - Название
      const path = data[i][3];      // Колонка D - Путь
      
      // Пропускаем пустые строки
      if (!id || !title) continue;
      
      // Считаем уровень вложенности по количеству "/" в пути
      let level = 0;
      if (path && typeof path === 'string') {
        level = (path.match(/\//g) || []).length;
      }
      
      categories.push({
        id: id,
        title: title,
        level: level,
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
    
    // КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: формируем payload без parent_id если он null
    const collectionData = {
      title: data.title,
      url: url,
      html_title: data.title,
      is_hidden: false
    };
    
    // Добавляем parent_id ТОЛЬКО если он указан и не null
    if (data.parent_id && data.parent_id !== null && data.parent_id !== '') {
      collectionData.parent_id = parseInt(data.parent_id);
      console.log('[INFO] Категория с родителем:', collectionData.parent_id);
    } else {
      console.log('[INFO] Корневая категория (без parent_id в payload)');
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
      
      loadCategoriesToSheet();
      
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