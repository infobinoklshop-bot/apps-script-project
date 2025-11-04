/**
 * ========================================
 * МОДУЛЬ КЭШИРОВАНИЯ КАТАЛОГА ТОВАРОВ (CSV)
 * ========================================
 *
 * Импортирует CSV-выгрузки разделов из InSales в отдельные листы
 * для быстрой фильтрации товаров по характеристикам.
 */

/**
 * Показывает диалог импорта CSV файла
 */
function showCSVImportDialog() {
  const html = createCSVImportDialogHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📤 Импорт каталога из CSV');
}

/**
 * Импортирует CSV-выгрузку в отдельный лист
 * @param {string} sectionName - Название раздела (например, "Бинокли")
 * @param {string} csvContent - Содержимое CSV файла
 */
/**
 * Импортирует CSV-выгрузку в отдельный лист
 * Использует надёжный парсинг с автоопределением разделителя
 */
function importCSVToSheet(sectionName, csvContent) {
  const context = 'importCSVToSheet';

  try {
    logInfo(`📥 Импорт CSV для раздела: ${sectionName}`, null, context);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `🗄️ ${sectionName}`;

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Определение формата CSV...',
      '⏳ Импорт',
      -1
    );

    // Определяем разделитель по первой строке
    const sampleLine = csvContent.substring(0, csvContent.indexOf('\n'));

    let delimiter = '\t'; // По умолчанию табуляция
    let maxColumns = 1;

    // Пробуем разные разделители
    const delimiters = ['\t', ';', ',', '|'];
    for (const d of delimiters) {
      const cols = sampleLine.split(d).length;
      if (cols > maxColumns) {
        maxColumns = cols;
        delimiter = d;
      }
    }

    logInfo(`📋 Разделитель: "${delimiter}" (колонок: ${maxColumns})`, null, context);

    // ИСПРАВЛЕНИЕ: Правильный CSV парсинг с поддержкой многострочных полей
    const rows = [];
    let currentRow = [];
    let currentField = '';
    let insideQuotes = false;
    let i = 0;

    console.log(`📋 Начинаем парсинг CSV (${csvContent.length} символов)...`);

    while (i < csvContent.length) {
      const char = csvContent[i];
      const nextChar = csvContent[i + 1];

      if (char === '"') {
        if (insideQuotes && nextChar === '"') {
          // Экранированная кавычка ("") внутри поля
          currentField += '"';
          i += 2;
          continue;
        } else {
          // Начало или конец кавычек
          insideQuotes = !insideQuotes;
          i++;
          continue;
        }
      }

      if (!insideQuotes && char === delimiter) {
        // Конец поля
        currentRow.push(currentField.trim());
        currentField = '';
        i++;
        continue;
      }

      if (!insideQuotes && (char === '\n' || char === '\r')) {
        // Конец строки
        if (currentField.length > 0 || currentRow.length > 0) {
          currentRow.push(currentField.trim());

          if (currentRow.length > 0) {
            rows.push(currentRow);
          }

          currentRow = [];
          currentField = '';
        }

        // Пропускаем \r\n или \n
        if (char === '\r' && nextChar === '\n') {
          i += 2;
        } else {
          i++;
        }
        continue;
      }

      // Обычный символ
      currentField += char;
      i++;
    }

    // Добавляем последнюю строку если есть
    if (currentField.length > 0 || currentRow.length > 0) {
      currentRow.push(currentField.trim());
      if (currentRow.length > 0) {
        rows.push(currentRow);
      }
    }

    console.log(`✅ Парсинг завершён: ${rows.length} строк`);

    if (rows.length === 0) {
      throw new Error('CSV файл пуст');
    }

    logInfo(`📊 Найдено строк: ${rows.length}`, null, context);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Записываем ${rows.length} строк...`,
      '⏳ Импорт',
      -1
    );

    // Удаляем старый лист если есть
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
    }

    // Создаём новый лист
    sheet = ss.insertSheet(sheetName);

    // Вычисляем максимальное количество колонок
    let maxCols = 0;
    for (const row of rows) {
      if (row.length > maxCols) {
        maxCols = row.length;
      }
    }

    // Выравниваем все строки по количеству колонок
    for (let i = 0; i < rows.length; i++) {
      while (rows[i].length < maxCols) {
        rows[i].push('');
      }
    }

    // Исключаем пустые колонки с параметрами (только если ВСЕ значения пустые)
    const columnsToKeep = [];
    for (let col = 0; col < maxCols; col++) {
      const header = rows[0][col];

      // Если это колонка с параметром, проверяем есть ли хотя бы одно заполненное значение
      if (header && header.toString().startsWith('Параметр:')) {
        let hasValue = false;
        for (let row = 1; row < rows.length; row++) {
          if (rows[row][col] && rows[row][col].toString().trim() !== '') {
            hasValue = true;
            break;
          }
        }
        // Если нет ни одного значения - пропускаем эту колонку
        if (!hasValue) {
          continue;
        }
      }

      // Оставляем эту колонку
      columnsToKeep.push(col);
    }

    logInfo(`📋 Исключено пустых колонок: ${maxCols - columnsToKeep.length}`, null, context);

    // Создаём новый массив строк только с нужными колонками
    const filteredRows = [];
    for (let i = 0; i < rows.length; i++) {
      const newRow = [];
      for (const col of columnsToKeep) {
        newRow.push(rows[i][col]);
      }
      filteredRows.push(newRow);
    }

    // Записываем отфильтрованные данные одним запросом
    const finalCols = columnsToKeep.length;
    sheet.getRange(1, 1, filteredRows.length, finalCols).setValues(filteredRows);
    maxCols = finalCols; // Обновляем maxCols для форматирования

    // ИСПРАВЛЕНИЕ: Форсируем текстовый формат для колонок с параметрами
    // чтобы предотвратить автоматическое преобразование "3.5" → Date
    try {
      for (let col = 0; col < columnsToKeep.length; col++) {
        const header = filteredRows[0][col];
        if (header && header.toString().startsWith('Параметр:')) {
          // Устанавливаем текстовый формат для всей колонки (кроме заголовка)
          if (filteredRows.length > 1) {
            sheet.getRange(2, col + 1, filteredRows.length - 1, 1)
              .setNumberFormat('@STRING@');
          }
        }
      }
    } catch (formatError) {
      console.warn('⚠️ Не удалось установить текстовый формат:', formatError.message);
      // Продолжаем без форматирования - это не критично
    }

    // Форматирование заголовков
    sheet.getRange(1, 1, 1, maxCols)
      .setBackground('#1976d2')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    // Закрепляем заголовок
    sheet.setFrozenRows(1);

    // Автоширина колонок (только первые 10)
    const columnsToResize = Math.min(10, maxCols);
    sheet.autoResizeColumns(1, columnsToResize);

    logInfo(`✅ Импорт завершён: ${rows.length - 1} товаров`, null, context);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Импортировано: ${rows.length - 1} товаров`,
      '✅ Готово',
      5
    );

    return {
      success: true,
      sheetName: sheetName,
      rowsImported: rows.length - 1
    };

  } catch (error) {
    logError(`❌ Ошибка импорта CSV для ${sectionName}`, error, context);
    throw error;
  }
}

/**
 * Загружает товары из CSV-листа для фильтрации
 * @param {string} sectionName - Название раздела (например, "Бинокли")
 * @returns {Object} Объект с товарами и метаданными
 */
function getProductsFromCSVSheet(sectionName) {
  const context = 'getProductsFromCSVSheet';

  try {
    // ========== VERSION MARKER ==========
    console.log('🔖 getProductsFromCSVSheet VERSION: 2025-11-03-FIX-v3');
    // ====================================

    logInfo(`📖 Читаем товары из листа: ${sectionName}`, null, context);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `🗄️ ${sectionName}`;
    console.log(`🔍 DEBUG: Ищем лист с именем "${sheetName}"`);
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Лист "${sheetName}" не найден. Импортируйте CSV через меню "📤 Импорт каталога из CSV"`);
    }

    console.log(`✅ Лист найден: "${sheet.getName()}" (ID: ${sheet.getSheetId()})`);
    console.log(`📊 Строк в листе: ${sheet.getLastRow()}, колонок: ${sheet.getLastColumn()}`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      logWarning(`⚠️ Лист ${sheetName} пуст`, null, context);
      return {
        products: [],
        characteristics: {},
        headers: []
      };
    }

    // Читаем все данные одним запросом
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // Первая строка - заголовки

    // ОТЛАДКА: Ищем ВСЕ колонки с "Кратность" в названии
    const kratnostColumns = [];
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString();
      if (header.includes('Кратность') || header.includes('кратност')) {
        kratnostColumns.push({
          index: i,
          name: header,
          charName: header.replace('Параметр:', '').trim() // Имя характеристики
        });
      }
    }
    console.log(`🔍 DEBUG: Найдено колонок с "Кратность": ${kratnostColumns.length}`);
    for (const col of kratnostColumns) {
      console.log(`   Колонка ${col.index}: "${col.name}"`);
      console.log(`   → Имя характеристики: "${col.charName}"`);
      console.log(`   → Длина имени: ${col.charName.length} символов`);
      console.log(`   → Коды символов:`, col.charName.split('').map((c, i) => `[${i}]=${c.charCodeAt(0)}`).join(' '));

      // Показываем первые 5 значений из КАЖДОЙ такой колонки
      const firstValues = [];
      for (let row = 1; row <= Math.min(5, data.length - 1); row++) {
        firstValues.push(data[row][col.index]);
      }
      console.log(`   → Первые 5 значений:`, firstValues);
      console.log(''); // Пустая строка для читаемости
    }

    // Находим индексы нужных колонок
    const columnIndices = findCSVColumns(headers);

    // Находим колонки с характеристиками (начинаются с "Параметр:")
    const characteristicColumns = [];
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      if (header && header.toString().startsWith('Параметр:')) {
        // Извлекаем название характеристики
        const charName = header.toString().replace('Параметр:', '').trim();
        characteristicColumns.push({
          index: i,
          name: charName
        });

        // ОТЛАДКА: Логируем колонку "Кратность увеличения, крат"
        if (charName === 'Кратность увеличения, крат') {
          console.log(`🔍 Колонка "Кратность увеличения, крат" найдена в индексе ${i}`);
          console.log(`🔍 Заголовок колонки: "${header}"`);
          // Показываем первые 5 значений из этой колонки
          const firstValues = [];
          for (let row = 1; row <= Math.min(5, data.length - 1); row++) {
            firstValues.push(data[row][i]);
          }
          console.log(`🔍 Первые 5 значений из колонки ${i}:`, firstValues);
        }
      }
    }

    logInfo(`📋 Найдено колонок с характеристиками: ${characteristicColumns.length}`, null, context);

    // Парсим товары
    const products = [];
    const characteristicsValues = {}; // Для метаданных фильтров

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Извлекаем характеристики
      const characteristics = {};
      for (const charCol of characteristicColumns) {
        const value = row[charCol.index];
        if (value !== null && value !== undefined && value !== '') {
          const valueStr = value.toString().trim();
          if (valueStr) {
            characteristics[charCol.name] = valueStr;

            // Собираем уникальные значения для фильтров
            if (!characteristicsValues[charCol.name]) {
              characteristicsValues[charCol.name] = new Set();
            }

            // ОТЛАДКА: Логируем ВСЕ значения кратности ПЕРЕД добавлением
            if (charCol.name === 'Кратность увеличения, крат') {
              const oldSize = characteristicsValues[charCol.name].size;
              characteristicsValues[charCol.name].add(valueStr);
              const newSize = characteristicsValues[charCol.name].size;

              // Логируем только если это НОВОЕ значение
              if (newSize > oldSize) {
                console.log(`📝 Строка ${i + 1}: НОВОЕ значение "${valueStr}" (уникальных: ${oldSize} → ${newSize})`);
              }
            } else {
              characteristicsValues[charCol.name].add(valueStr);
            }
          }
        }
      }

      // Создаём объект товара
      // ИСПРАВЛЕНО: Конвертируем title и sku в строки для корректной работы фильтров
      const product = {
        id: row[columnIndices.id] || 0,
        title: (row[columnIndices.title] || 'Без названия').toString(),
        sku: (row[columnIndices.sku] || '').toString(),
        price: parseFloat(row[columnIndices.price]) || 0,
        in_stock: row[columnIndices.stock] !== null && row[columnIndices.stock] !== undefined &&
                   parseFloat(row[columnIndices.stock]) > 0,
        characteristics: characteristics
      };

      products.push(product);
    }

    // Конвертируем Set в массивы для метаданных
    const characteristicsArrays = {};
    for (const [key, valuesSet] of Object.entries(characteristicsValues)) {
      // Конвертируем Set в массив и сортируем
      const values = Array.from(valuesSet);

      characteristicsArrays[key] = values.sort((a, b) => {
        // Пытаемся сортировать как числа, если возможно
        const numA = parseFloat(a);
        const numB = parseFloat(b);
        if (!isNaN(numA) && !isNaN(numB)) {
          return numA - numB;
        }
        return a.toString().localeCompare(b.toString(), 'ru');
      });

      // ОТЛАДКА: Логируем массив кратности ПОСЛЕ конвертации
      if (key === 'Кратность увеличения, крат') {
        console.log(`🔍 DEBUG: Массив кратности ПОСЛЕ конвертации Set → Array:`);
        console.log(`   Всего элементов: ${characteristicsArrays[key].length}`);
        console.log(`   Первые 5: ${JSON.stringify(characteristicsArrays[key].slice(0, 5))}`);
        console.log(`   Последние 5: ${JSON.stringify(characteristicsArrays[key].slice(-5))}`);
      }
    }

    logInfo(`✅ Загружено товаров: ${products.length}`, null, context);

    // ОТЛАДКА: Добавляем информацию о колонке "Кратность увеличения, крат"
    let kratnostDebugInfo = null;
    for (const charCol of characteristicColumns) {
      if (charCol.name === 'Кратность увеличения, крат') {
        const firstValues = [];
        for (let row = 1; row <= Math.min(5, data.length - 1); row++) {
          firstValues.push(data[row][charCol.index]);
        }
        kratnostDebugInfo = {
          index: charCol.index,
          header: headers[charCol.index],
          firstValues: firstValues
        };
        break;
      }
    }

    return {
      products: products,
      characteristics: characteristicsArrays,
      headers: headers,
      _debug_kratnost: kratnostDebugInfo  // Отладочная информация
    };

  } catch (error) {
    logError(`❌ Ошибка чтения товаров из ${sectionName}`, error, context);
    throw error;
  }
}

/**
 * Находит индексы нужных колонок в CSV
 * @param {Array} headers - Массив заголовков
 * @returns {Object} Объект с индексами колонок
 */
function findCSVColumns(headers) {
  const indices = {
    id: -1,
    title: -1,
    sku: -1,
    price: -1,
    stock: -1
  };

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toString().toLowerCase();

    // ИСПРАВЛЕНО: ID Товара (product_id) - это то, что нужно для InSales Collects API!
    // ID варианта (variant_id) НЕ ПОДХОДИТ для добавления в категории
    if (header.includes('id товара') || header === 'id товара' || header === 'id') {
      indices.id = i;
    }

    // Название (ИСПРАВЛЕНО: ищем "Название товара или услуги")
    if (header.includes('название товара или услуги') || header === 'название товара или услуги') {
      indices.title = i;
    }
    // Фолбэк на старое название
    else if (indices.title === -1 && (header.includes('название товара') || header === 'название товара')) {
      indices.title = i;
    }

    // Артикул
    if (header.includes('артикул') || header === 'артикул') {
      indices.sku = i;
    }

    // Цена
    if (header.includes('цена продажи') || header === 'цена продажи') {
      indices.price = i;
    }

    // Наличие (остаток)
    if (header.includes('остаток') || header === 'остаток') {
      indices.stock = i;
    }
  }

  return indices;
}

/**
 * HTML диалога импорта CSV
 */
function createCSVImportDialogHTML() {
  const mainSections = [
    'Бинокли',
    'Зрительные трубы',
    'Монокуляры',
    'Прицелы',
    'Тепловизоры',
    'Приборы ночного видения (ПНВ)',
    'Телескопы',
    'Микроскопы',
    'Лупы',
    'Фонари'
  ];

  const sectionsJson = JSON.stringify(mainSections);

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            margin: 0;
            background: #f8f9fa;
          }

          h3 {
            color: #1976d2;
            margin-top: 0;
            margin-bottom: 20px;
          }

          .info {
            background: #e3f2fd;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
            font-size: 14px;
            line-height: 1.6;
          }

          .step {
            background: white;
            padding: 20px;
            margin-bottom: 15px;
            border-radius: 4px;
            border: 1px solid #ddd;
          }

          .step h4 {
            margin-top: 0;
            color: #333;
            font-size: 16px;
          }

          label {
            display: block;
            font-weight: bold;
            margin-bottom: 8px;
            font-size: 14px;
          }

          select, input[type="file"] {
            width: 100%;
            padding: 10px;
            font-size: 14px;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-sizing: border-box;
          }

          select:focus {
            outline: none;
            border-color: #1976d2;
          }

          .file-input-wrapper {
            position: relative;
            margin-top: 10px;
          }

          .actions {
            margin-top: 20px;
            text-align: right;
            padding-top: 20px;
            border-top: 1px solid #ddd;
          }

          button {
            padding: 12px 24px;
            margin-left: 10px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
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

          .status {
            margin-top: 15px;
            padding: 10px;
            border-radius: 4px;
            display: none;
          }

          .status.success {
            background: #d4edda;
            color: #155724;
            display: block;
          }

          .status.error {
            background: #f8d7da;
            color: #721c24;
            display: block;
          }
        </style>
      </head>
      <body>
        <h3>📤 Импорт каталога из CSV</h3>

        <div class="info">
          <strong>Инструкция:</strong><br>
          1️⃣ Выгрузите раздел из админки InSales в CSV<br>
          2️⃣ Выберите название раздела из списка<br>
          3️⃣ Загрузите CSV файл<br>
          4️⃣ Нажмите "Импортировать"<br>
          <br>
          📌 CSV будет импортирован в лист "🗄️ [Название раздела]"
        </div>

        <div class="step">
          <h4>Шаг 1: Выберите раздел</h4>
          <label for="sectionSelect">Раздел каталога:</label>
          <select id="sectionSelect">
            <option value="">-- Выберите раздел --</option>
          </select>
        </div>

        <div class="step">
          <h4>Шаг 2: Загрузите CSV файл</h4>
          <label for="csvFile">CSV файл от InSales:</label>
          <div class="file-input-wrapper">
            <input type="file" id="csvFile" accept=".csv" />
          </div>
        </div>

        <div class="status" id="status"></div>

        <div class="actions">
          <button class="btn-cancel" onclick="google.script.host.close()">Отмена</button>
          <button id="importBtn" class="btn-primary" onclick="startImport()" disabled>
            📤 Импортировать
          </button>
        </div>

        <script>
          const sections = ${sectionsJson};
          let selectedSection = '';
          let csvFile = null;

          window.onload = function() {
            renderSections();

            document.getElementById('sectionSelect').addEventListener('change', function(e) {
              selectedSection = e.target.value;
              updateImportButton();
            });

            document.getElementById('csvFile').addEventListener('change', function(e) {
              csvFile = e.target.files[0];
              updateImportButton();
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

          function updateImportButton() {
            const btn = document.getElementById('importBtn');
            btn.disabled = !selectedSection || !csvFile;
          }

          function startImport() {
            if (!selectedSection || !csvFile) {
              showStatus('Выберите раздел и CSV файл', 'error');
              return;
            }

            const reader = new FileReader();
            const importBtn = document.getElementById('importBtn');

            importBtn.disabled = true;
            importBtn.textContent = '⏳ Импортируем...';

            reader.onload = function(e) {
              const csvContent = e.target.result;

              google.script.run
                .withSuccessHandler(function(result) {
                  showStatus(\`✅ Импортировано: \${result.rowsImported} товаров в лист "\${result.sheetName}"\`, 'success');
                  setTimeout(() => google.script.host.close(), 2000);
                })
                .withFailureHandler(function(error) {
                  showStatus('❌ Ошибка: ' + error.message, 'error');
                  importBtn.disabled = false;
                  importBtn.textContent = '📤 Импортировать';
                })
                .importCSVToSheet(selectedSection, csvContent);
            };

            // Пробуем разные кодировки (InSales может экспортировать в Windows-1251)
            reader.readAsText(csvFile, 'windows-1251');
          }

          function showStatus(message, type) {
            const status = document.getElementById('status');
            status.textContent = message;
            status.className = 'status ' + type;
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Получает список импортированных разделов (листов с кэшем)
 */
function getImportedSections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const sections = [];
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name.startsWith('🗄️ ')) {
      sections.push(name.replace('🗄️ ', ''));
    }
  }

  return sections;
}
