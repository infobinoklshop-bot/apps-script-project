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
    
    const data = sheet.getRange(startRow, 2, lastRow - startRow + 1, 1).getValues();
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
 * HTML диалога добавления товаров
 */
function createAddProductsDialogHTML(categoryId, categoryTitle, existingIds) {
  const existingIdsJson = JSON.stringify(existingIds);
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
          h3 { color: #1976d2; margin-top: 0; }
          .search-box { margin: 20px 0; }
          .search-box input {
            width: 100%;
            padding: 12px;
            font-size: 16px;
            border: 2px solid #ddd;
            border-radius: 4px;
            box-sizing: border-box;
          }
          .search-box input:focus { outline: none; border-color: #1976d2; }
          .hint { color: #666; font-size: 13px; margin-top: 8px; }
          .results {
            margin: 20px 0;
            max-height: 380px;
            overflow-y: auto;
            border: 1px solid #ddd;
            border-radius: 4px;
          }
          table { width: 100%; border-collapse: collapse; }
          th {
            position: sticky;
            top: 0;
            background: #1976d2;
            color: white;
            padding: 12px;
            text-align: left;
            font-size: 14px;
          }
          td { padding: 10px 12px; border-bottom: 1px solid #eee; font-size: 14px; }
          tr:hover { background: #f5f5f5; }
          .checkbox-cell { width: 40px; text-align: center; }
          .stock-yes { color: #4caf50; font-weight: bold; }
          .stock-no { color: #f44336; }
          .in-category { color: #999; font-style: italic; }
          .actions {
            margin-top: 20px;
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
          .btn-cancel { background: #f1f3f4; color: #333; }
          .btn-primary { background: #1976d2; color: white; }
          .btn-primary:disabled { background: #ccc; cursor: not-allowed; }
          .btn-load-more {
            background: #4caf50;
            color: white;
            width: 100%;
            margin-top: 10px;
          }
          .loading { text-align: center; padding: 40px; color: #666; }
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
          .info {
            background: #e3f2fd;
            padding: 12px;
            margin-bottom: 15px;
            border-radius: 4px;
            font-size: 14px;
          }
        </style>
      </head>
      <body>
        <h3>➕ Добавить товары в категорию</h3>
        
        <div class="info">
          📁 Категория: <strong>${categoryTitle}</strong><br>
          📦 Уже в категории: <strong>${existingIds.length}</strong> товаров<br>
          ✅ Выбрано для добавления: <strong id="selectedCount">0</strong>
        </div>
        
        <div class="search-box">
          <input 
            type="text" 
            id="searchInput" 
            placeholder="Начните вводить название товара..."
            autocomplete="off"
          >
          <div class="hint">
            💡 Минимум 2 символа. Товары, уже находящиеся в категории, будут помечены серым.
          </div>
        </div>
        
        <div class="results">
          <table>
            <thead>
              <tr>
                <th class="checkbox-cell">☑️</th>
                <th>Название товара</th>
                <th style="width: 120px;">Наличие</th>
              </tr>
            </thead>
            <tbody id="resultsBody">
              <tr>
                <td colspan="3" class="loading">
                  <div class="loading-spinner"></div>
                  <strong>Загружаем каталог товаров...</strong><br>
                  <small>Это может занять до 30 секунд</small>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        
        <button id="loadMoreBtn" class="btn-load-more" style="display: none;" onclick="loadMoreProducts()">
          📄 Показать ещё 100 товаров
        </button>
        
        <div class="actions">
          <button class="btn-cancel" onclick="google.script.host.close()">Отмена</button>
          <button id="addBtn" class="btn-primary" onclick="addSelectedProducts()" disabled>
            ➕ Добавить выбранные (<span id="addBtnCount">0</span>)
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
          
          window.onload = function() {
            loadAllProducts();
            
            document.getElementById('searchInput').addEventListener('input', function(e) {
              clearTimeout(searchTimeout);
              searchTimeout = setTimeout(() => performSearch(e.target.value), 500);
            });
          };
          
          function loadAllProducts() {
            google.script.run
              .withSuccessHandler(function(products) {
                allProducts = products;
                document.getElementById('resultsBody').innerHTML = 
                  '<tr><td colspan="3" class="loading">' +
                  '✅ Каталог загружен (' + products.length + ' товаров)<br>' +
                  'Введите текст для поиска' +
                  '</td></tr>';
              })
              .withFailureHandler(function(error) {
                document.getElementById('resultsBody').innerHTML = 
                  '<tr><td colspan="3" class="loading">❌ Ошибка загрузки: ' + error.message + '</td></tr>';
              })
              .getAllProductsForSearch();
          }
          
          function performSearch(query) {
            if (!query || query.length < 2) {
              document.getElementById('resultsBody').innerHTML = 
                '<tr><td colspan="3" class="loading">Введите минимум 2 символа</td></tr>';
              document.getElementById('loadMoreBtn').style.display = 'none';
              return;
            }
            
            const lowerQuery = query.toLowerCase();
            filteredProducts = allProducts.filter(p => 
              p.title.toLowerCase().includes(lowerQuery)
            );
            
            displayedCount = 0;
            displayResults();
          }
          
          function displayResults() {
            const tbody = document.getElementById('resultsBody');
            
            if (filteredProducts.length === 0) {
              tbody.innerHTML = '<tr><td colspan="3" class="loading">😔 Товары не найдены</td></tr>';
              document.getElementById('loadMoreBtn').style.display = 'none';
              return;
            }
            
            const endIndex = Math.min(displayedCount + batchSize, filteredProducts.length);
            const batch = filteredProducts.slice(displayedCount, endIndex);
            
            const html = batch.map(product => {
              const inCategory = existingIds.has(product.id);
              const rowClass = inCategory ? 'in-category' : '';
              const checkbox = inCategory ? 
                '<input type="checkbox" disabled title="Уже в категории">' :
                \`<input type="checkbox" value="\${product.id}" onchange="toggleProduct(\${product.id})" \${selectedProducts.has(product.id) ? 'checked' : ''}>\`;
              
              return \`
                <tr class="\${rowClass}">
                  <td class="checkbox-cell">\${checkbox}</td>
                  <td>\${product.title} \${inCategory ? '(уже в категории)' : ''}</td>
                  <td class="\${product.in_stock ? 'stock-yes' : 'stock-no'}">
                    \${product.in_stock ? '✅ Да' : '❌ Нет'}
                  </td>
                </tr>
              \`;
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
                \`📄 Показать ещё 100 товаров (осталось \${filteredProducts.length - displayedCount})\`;
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
                addBtn.innerHTML = '➕ Добавить выбранные (<span id="addBtnCount">' + productIds.length + '</span>)';
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
 * Добавляет товары через Collects API (РАБОЧИЙ МЕТОД!)
 */
function addProductsToCategory(categoryId, productIds) {
  try {
    console.log(`📦 Добавляем ${productIds.length} товаров в категорию ${categoryId}`);
    
    const credentials = getInsalesCredentialsSync();
    if (!credentials) throw new Error('Учетные данные не настроены');
    
    let successCount = 0;
    let errorCount = 0;
    const addedProducts = [];
    
    for (const productId of productIds) {
      try {
        const collectUrl = `${credentials.baseUrl}/admin/collects.json`;
        
        const payload = {
          collect: {
            product_id: productId,
            collection_id: parseInt(categoryId)
          }
        };
        
        console.log(`📦 Добавляем товар ${productId} в категорию ${categoryId}...`);
        
        const response = UrlFetchApp.fetch(collectUrl, {
          method: 'POST',
          headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        
        const statusCode = response.getResponseCode();
        
        if (statusCode === 200 || statusCode === 201) {
          console.log(`✅ Товар ${productId} добавлен в категорию`);
          successCount++;
          
          // Загружаем данные товара
          try {
            const productUrl = `${credentials.baseUrl}/admin/products/${productId}.json`;
            const productResponse = UrlFetchApp.fetch(productUrl, {
              method: 'GET',
              headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
              },
              muteHttpExceptions: true
            });
            
            if (productResponse.getResponseCode() === 200) {
              const product = JSON.parse(productResponse.getContentText());
              addedProducts.push(product);
            }
          } catch (e) {
            console.warn(`⚠️ Не удалось загрузить данные товара ${productId}`);
          }
          
        } else if (statusCode === 422) {
          console.log(`⚠️ Товар ${productId} уже в категории`);
          successCount++;
        } else {
          console.error(`❌ Ошибка добавления товара ${productId}: ${statusCode}`);
          errorCount++;
        }
        
        Utilities.sleep(500);
        
      } catch (e) {
        console.error(`❌ Ошибка товара ${productId}:`, e.message);
        errorCount++;
      }
    }
    
    console.log(`✅ Итого: добавлено ${successCount}, ошибок ${errorCount}`);
    
    if (addedProducts.length > 0) {
      appendProductsToDetailSheet(addedProducts);
    }
    
    return { success: successCount, errors: errorCount };
    
  } catch (error) {
    console.error('❌ Ошибка добавления товаров:', error);
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