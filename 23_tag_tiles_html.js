/**
 * ========================================
 * МОДУЛЬ: ГЕНЕРАЦИЯ HTML ДЛЯ ПЛИТОК ТЕГОВ
 * ========================================
 */

/**
 * Генерирует полный HTML для обеих плиток (верхняя + нижняя)
 * @param {Object} generationResult - Результат от generateTileAnchors()
 * @returns {Object} { topHTML, bottomHTML, combinedHTML }
 */
function generateTilesHTML(generationResult) {
  try {
    console.log('[INFO] 🎨 Начинаем генерацию HTML для плиток...');

    // Генерируем HTML для верхней плитки
    const topHTML = generateTopTileHTML(generationResult.topTile);
    console.log('[INFO] ✅ Верхняя плитка сгенерирована:', topHTML.length, 'символов');

    // Генерируем HTML для нижней плитки
    const bottomHTML = generateBottomTileHTML(generationResult.bottomTile);
    console.log('[INFO] ✅ Нижняя плитка сгенерирована:', bottomHTML.length, 'символов');

    // Объединяем обе плитки
    const combinedHTML = topHTML + '\n\n' + bottomHTML;

    return {
      topHTML: topHTML,
      bottomHTML: bottomHTML,
      combinedHTML: combinedHTML,
      categoryId: generationResult.categoryId,
      categoryName: generationResult.categoryName
    };

  } catch (error) {
    console.error('[ERROR] Ошибка генерации HTML:', error.message);
    throw error;
  }
}

/**
 * Генерирует HTML для верхней плитки (навигация)
 * @param {Array} anchors - Массив анкоров для верхней плитки
 * @returns {string} HTML код
 */
function generateTopTileHTML(anchors) {
  if (!anchors || anchors.length === 0) {
    return '';
  }

  const cssClasses = TAG_TILES_CONFIG.CSS_CLASSES;

  // Начало контейнера
  let html = `<div class="${cssClasses.TOP_TILE}">\n`;
  html += `  <h3 class="tiles-title">Категории</h3>\n`;
  html += `  <div class="${cssClasses.TILE_GRID}">\n`;

  // Генерируем каждый анкор
  anchors.forEach(anchor => {
    const url = anchor.link || anchor.url || '#';
    const text = anchor.anchor || anchor.text;

    html += `    <a href="${url}" class="${cssClasses.TILE}">${text}</a>\n`;
  });

  // Закрываем контейнер
  html += `  </div>\n`;
  html += `</div>`;

  return html;
}

/**
 * Генерирует HTML для нижней плитки (SEO)
 * @param {Array} anchors - Массив анкоров для нижней плитки
 * @returns {string} HTML код
 */
function generateBottomTileHTML(anchors) {
  if (!anchors || anchors.length === 0) {
    return '';
  }

  const cssClasses = TAG_TILES_CONFIG.CSS_CLASSES;

  // Начало контейнера
  let html = `<div class="${cssClasses.BOTTOM_TILE}">\n`;
  html += `  <h3 class="tiles-title">Популярные товары</h3>\n`;
  html += `  <div class="${cssClasses.TILE_GRID}">\n`;

  // Генерируем каждый анкор
  anchors.forEach(anchor => {
    const url = anchor.link || anchor.url || '#';
    const text = anchor.anchor || anchor.text;

    html += `    <a href="${url}" class="${cssClasses.TILE_SMALL}">${text}</a>\n`;
  });

  // Закрываем контейнер
  html += `  </div>\n`;
  html += `</div>`;

  return html;
}

/**
 * Генерирует CSS стили для плиток
 * Эти стили нужно добавить в тему InSales
 * @returns {string} CSS код
 */
function generateTilesCSS() {
  return `
/* ========================================
   СТИЛИ ДЛЯ ПЛИТОК ТЕГОВ КАТЕГОРИЙ
   ======================================== */

.category-tiles {
  margin: 30px 0;
  padding: 20px;
  background: #f8f9fa;
  border-radius: 8px;
}

.category-tiles--top {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
}

.category-tiles--bottom {
  background: #ffffff;
  border: 1px solid #e1e4e8;
}

.tiles-title {
  font-size: 24px;
  font-weight: 600;
  margin: 0 0 15px 0;
  padding: 0;
  color: inherit;
}

.category-tiles--top .tiles-title {
  color: white;
}

.tiles-grid {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  margin: 0;
  padding: 0;
}

.tile {
  display: inline-block;
  padding: 12px 24px;
  background: rgba(255, 255, 255, 0.2);
  border: 1px solid rgba(255, 255, 255, 0.3);
  border-radius: 6px;
  text-decoration: none;
  color: white;
  font-size: 16px;
  font-weight: 500;
  transition: all 0.3s ease;
  backdrop-filter: blur(10px);
}

.tile:hover {
  background: rgba(255, 255, 255, 0.3);
  border-color: rgba(255, 255, 255, 0.5);
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
}

.tile--small {
  display: inline-block;
  padding: 8px 16px;
  background: #f0f2f5;
  border: 1px solid #e1e4e8;
  border-radius: 4px;
  text-decoration: none;
  color: #24292e;
  font-size: 14px;
  font-weight: 400;
  transition: all 0.2s ease;
}

.tile--small:hover {
  background: #e1e4e8;
  border-color: #d1d5da;
  color: #0366d6;
  transform: translateY(-1px);
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

/* Адаптивность */
@media (max-width: 768px) {
  .category-tiles {
    padding: 15px;
    margin: 20px 0;
  }

  .tiles-title {
    font-size: 20px;
  }

  .tile {
    padding: 10px 18px;
    font-size: 14px;
  }

  .tile--small {
    padding: 6px 12px;
    font-size: 13px;
  }

  .tiles-grid {
    gap: 8px;
  }
}
`.trim();
}

/**
 * Инжектит плитки в описание категории
 * @param {string} existingDescription - Текущее описание категории
 * @param {string} topHTML - HTML верхней плитки
 * @param {string} bottomHTML - HTML нижней плитки
 * @returns {string} Обновленное описание с плитками
 */
function injectTilesIntoDescription(existingDescription, topHTML, bottomHTML) {
  let description = existingDescription || '';

  // Удаляем старые плитки, если есть
  description = removeExistingTiles(description);

  // Добавляем верхнюю плитку в начало
  if (topHTML) {
    description = topHTML + '\n\n' + description;
  }

  // Добавляем нижнюю плитку в конец
  if (bottomHTML) {
    description = description + '\n\n' + bottomHTML;
  }

  return description.trim();
}

/**
 * Удаляет существующие плитки из описания
 * @param {string} html - HTML описания
 * @returns {string} Очищенный HTML
 */
function removeExistingTiles(html) {
  if (!html) return '';

  // Удаляем блоки с классами category-tiles
  let cleaned = html.replace(
    /<div\s+class="category-tiles[^"]*">[\s\S]*?<\/div>/gi,
    ''
  );

  // Удаляем множественные пустые строки
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n');

  return cleaned.trim();
}

/**
 * Создает превью HTML для отображения в листе
 * @param {string} html - HTML код
 * @returns {string} Форматированный превью
 */
function createHTMLPreview(html) {
  if (!html) return 'Нет HTML';

  // Добавляем переносы строк для читаемости
  let preview = html
    .replace(/></g, '>\n<')
    .replace(/\n\s+/g, '\n');

  // Ограничиваем длину
  if (preview.length > 1000) {
    preview = preview.substring(0, 1000) + '\n... (обрезано)';
  }

  return preview;
}

/**
 * ТЕСТОВАЯ ФУНКЦИЯ: Генерирует HTML для тестовых данных
 */
function testHTMLGeneration() {
  console.log('[TEST] === ТЕСТ ГЕНЕРАЦИИ HTML ===');

  // Создаем тестовые данные
  const testData = {
    categoryId: 9175197,
    categoryName: 'Бинокли',
    topTile: [
      { anchor: 'Бинокли для охоты', link: '/collection/dlya-ohoty', category_id: 9071017 },
      { anchor: 'Туристические бинокли', link: '/collection/turisticheskie', category_id: 9175207 },
      { anchor: 'Морские бинокли', link: '/collection/morskie', category_id: 9071623 }
    ],
    bottomTile: [
      { anchor: 'Купить бинокль недорого', link: '/collection/nedorogie', keywords_covered: ['купить', 'недорого'] },
      { anchor: 'Профессиональные бинокли', link: '/collection/professionalnye', keywords_covered: ['профессиональные'] },
      { anchor: 'Бинокли с большим увеличением', link: '/collection/bolshoe-uvelichenie', keywords_covered: ['большое увеличение'] }
    ]
  };

  // Генерируем HTML
  const result = generateTilesHTML(testData);

  console.log('\n--- ВЕРХНЯЯ ПЛИТКА ---');
  console.log(result.topHTML);

  console.log('\n--- НИЖНЯЯ ПЛИТКА ---');
  console.log(result.bottomHTML);

  console.log('\n--- ОБЪЕДИНЕННЫЙ HTML ---');
  console.log(result.combinedHTML);

  console.log('\n--- CSS СТИЛИ ---');
  console.log(generateTilesCSS());

  console.log('\n[TEST] ✅ Тест завершен');

  return result;
}

/**
 * ГЛАВНАЯ ФУНКЦИЯ: Генерирует плитки для активной категории
 * Объединяет генерацию анкоров и HTML
 */
function generateAndApplyTilesForActiveCategory() {
  try {
    const categoryData = getActiveCategoryData();

    if (!categoryData) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Откройте детальный лист категории', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Генерация плиток тегов...', '⏳ Обработка', -1);

    // 1. Генерируем анкоры через AI
    console.log('[INFO] 🤖 Генерируем анкоры...');
    const anchorsResult = generateTileAnchors(categoryData.category_id);

    // 2. Генерируем HTML
    console.log('[INFO] 🎨 Генерируем HTML...');
    const htmlResult = generateTilesHTML(anchorsResult);

    // 3. Сохраняем результат в лист
    const sheet = SpreadsheetApp.getActiveSheet();
    const startRow = DETAIL_SHEET_SECTIONS.TAG_TILES_START;

    sheet.getRange(startRow, 1, 1, 1).setValue('🏷️ ПЛИТКИ ТЕГОВ');
    sheet.getRange(startRow + 2, 1, 1, 1).setValue('Верхняя плитка (Навигация):');
    sheet.getRange(startRow + 3, 1, 1, 2).setValue(htmlResult.topHTML);

    sheet.getRange(startRow + 5, 1, 1, 1).setValue('Нижняя плитка (SEO):');
    sheet.getRange(startRow + 6, 1, 1, 2).setValue(htmlResult.bottomHTML);

    sheet.getRange(startRow + 8, 1, 1, 1).setValue('Объединенный HTML:');
    sheet.getRange(startRow + 9, 1, 1, 2).setValue(htmlResult.combinedHTML);

    SpreadsheetApp.getActiveSpreadsheet().toast('Плитки успешно сгенерированы!', '✅ Готово', 5);

    console.log('[INFO] ✅ Плитки сохранены в лист');

    return htmlResult;

  } catch (error) {
    console.error('[ERROR] Ошибка генерации плиток:', error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка: ' + error.message, '❌ Ошибка', 10);
    throw error;
  }
}
