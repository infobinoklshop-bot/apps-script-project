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

          // Загружаем данные товара для добавления в таблицу
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
              const responseData = JSON.parse(productResponse.getContentText());
              const product = responseData.product || responseData;
              addedProducts.push(product);
              console.log(`✅ Данные товара ${productId} загружены для добавления в таблицу`);
            }
          } catch (e) {
            console.error(`❌ Ошибка загрузки товара ${productId}:`, e.message);
          }

        } else if (statusCode === 422) {
          console.log(`⚠️ Товар ${productId} уже в категории InSales`);
          successCount++;

          // Загружаем данные товара для добавления в таблицу
          try {
            const productUrl = `${credentials.baseUrl}/admin/products/${productId}.json`;
            console.log(`[422] Загружаем товар: ${productUrl}`);

            const productResponse = UrlFetchApp.fetch(productUrl, {
              method: 'GET',
              headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
              },
              muteHttpExceptions: true
            });

            const prodStatus = productResponse.getResponseCode();
            console.log(`[422] Product API код: ${prodStatus}`);

            if (prodStatus === 200) {
              const responseData = JSON.parse(productResponse.getContentText());
              console.log(`[422] Response keys: ${Object.keys(responseData).join(', ')}`);
              const product = responseData.product || responseData;
              console.log(`[422] Product keys: ${Object.keys(product).join(', ')}`);
              addedProducts.push(product);
              console.log(`✅ Данные товара ${productId} загружены для добавления в таблицу (всего: ${addedProducts.length})`);
            } else {
              console.error(`[422] ❌ Product API вернул код ${prodStatus}: ${productResponse.getContentText()}`);
            }
          } catch (e) {
            console.error(`❌ Ошибка загрузки товара ${productId}:`, e.message);
            console.error(`Stack: ${e.stack}`);
          }
        } else {
          console.error(`❌ Ошибка добавления товара ${productId}: ${statusCode} - ${response.getContentText()}`);
          errorCount++;
        }
        
        Utilities.sleep(500);
        
      } catch (e) {
        console.error(`❌ Ошибка товара ${productId}:`, e.message);
        errorCount++;
      }
    }
    
    console.log(`✅ Итого: добавлено ${successCount}, ошибок ${errorCount}`);
    console.log(`📋 Товаров для добавления в таблицу: ${addedProducts.length}`);

    if (addedProducts.length > 0) {
      console.log('📝 Добавляем товары в таблицу...');
      appendProductsToDetailSheet(addedProducts);
      console.log('✅ Товары добавлены в таблицу');
    } else {
      console.warn('⚠️ Нет товаров для добавления в таблицу (addedProducts пуст)');
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