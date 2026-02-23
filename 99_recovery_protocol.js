/**
 * ========================================
 * 🚑 ПРОТОКОЛ ВОССТАНОВЛЕНИЯ (OPERATION JUMPSTART)
 * ========================================
 * 
 * Этот скрипт объединяет лучшие инструменты автоматизации для
 * экстренного "оживления" категории в глазах поисковиков.
 * 
 * ДЕЙСТВИЯ:
 * 1. 📝 Content Refresh: Переписывает описание через Gemini 2.0
 * 2. 🔗 Structural Boost: Генерирует плитку тегов (Top & Bottom)
 * 3. 🛍️ Commercial Fix: Сортирует товары (Smart Sort)
 */

/**
 * ЗАПУСК ПРОТОКОЛА ДЛЯ АКТИВНОЙ КАТЕГОРИИ
 */
function runRecoveryProtocolForActiveCategory() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();

    // 1. Проверка контекста
    if (!sheetName.startsWith('Категория — ')) {
        ui.alert('Ошибка', 'Запускайте только на детальном листе категории!', ui.ButtonSet.OK);
        return;
    }

    const categoryId = sheet.getRange('B2').getValue();
    const categoryTitle = sheet.getRange('B3').getValue();

    // 2. Подтверждение
    const confirm = ui.alert(
        '🚑 ЗАПУСК OPERATION JUMPSTART',
        `Вы готовы применить протокол восстановления для категории:\n"${categoryTitle}" (ID: ${categoryId})?\n\n` +
        `БУДЕТ ВЫПОЛНЕНО:\n` +
        `1. 📝 Генерация нового описания (Gemini 2.0)\n` +
        `2. 🔗 Генерация плитки тегов (SEO & Nav)\n` +
        `3. 🛍️ Умная сортировка товаров (Top-36)\n\n` +
        `Это займет 1-2 минуты. Продолжить?`,
        ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    try {
        SpreadsheetApp.getActiveSpreadsheet().toast('🚀 Запуск протокола восстановления...', 'Operation Jumpstart', -1);

        // ЛОГИРОВАНИЕ НАЧАЛА
        logPageChange(categoryId, '🚑 ЗАПУСК OPERATION JUMPSTART: Комплексное восстановление');

        // --- ЭТАП 1: КОНТЕНТ (Gemini 2.0) ---
        SpreadsheetApp.getActiveSpreadsheet().toast('📝 Шаг 1/3: Обновление контента...', 'Operation Jumpstart', -1);
        console.log('📝 [Recovery] Запуск генерации описания...');

        // Вызываем функцию из 09_ai_category_descriptions.js
        // Она сама читает лист и пишет в C17
        generateDescriptionWithGemini();

        Utilities.sleep(2000); // Даем время продышаться

        // --- ЭТАП 2: СТРУКТУРА (Tag Tiles) ---
        SpreadsheetApp.getActiveSpreadsheet().toast('🔗 Шаг 2/3: Генерация тегов...', 'Operation Jumpstart', -1);
        console.log('🔗 [Recovery] Запуск генерации тегов...');

        // Вызываем функцию из 23_tag_tiles_generator.js
        // Она возвращает объект, который нужно записать
        const tileResults = generateTileAnchors(categoryId);

        // Записываем результаты (функция из 16_categories_ai_content.js)
        // Нам нужно объединить top и bottom для записи, или использовать специфичную функцию
        // В 16_categories_ai_content.js есть writeTagTilesToDetailSheet(tagTiles)
        // Но generateTileAnchors возвращает {topTile: {anchors: []}, bottomTile: {anchors: []}}
        // Нам нужно преобразовать это в плоский список для writeTagTilesToDetailSheet

        const allTags = [];

        // Добавляем верхние (type = category_link)
        if (tileResults.topTile && tileResults.topTile.anchors) {
            tileResults.topTile.anchors.forEach(a => {
                allTags.push({
                    anchor: a.anchor,
                    target_url: a.category_id ? `ID: ${a.category_id}` : null, // Упрощение, так как writeTagTilesToDetailSheet ждет URL
                    type: 'category_link',
                    relevance: 1.0,
                    target_category_id: a.category_id
                });
            });
        }

        // Добавляем нижние (type = filter)
        if (tileResults.bottomTile && tileResults.bottomTile.anchors) {
            tileResults.bottomTile.anchors.forEach(a => {
                allTags.push({
                    anchor: a.anchor,
                    target_url: a.target_category ? a.target_category.url : 'CREATE_NEW',
                    type: 'filter',
                    relevance: 0.9, // Условная релевантность
                    target_category_id: a.target_category ? a.target_category.id : null
                });
            });
        }

        // Пишем в лист
        if (allTags.length > 0) {
            // Функция из 16_categories_ai_content.js
            writeTagTilesToDetailSheet(allTags);
        }

        Utilities.sleep(2000);

        // --- ЭТАП 3: ТОВАРЫ (Smart Sort) ---
        SpreadsheetApp.getActiveSpreadsheet().toast('🛍️ Шаг 3/3: Сортировка товаров...', 'Operation Jumpstart', -1);
        console.log('🛍️ [Recovery] Запуск умной сортировки...');

        // Вызываем функцию из 17_categories_products.gs.js
        // false = без лишних диалогов (фоновый режим), но мы хотим обновить InSales
        // В 17_categories_products.gs.js: smartSortProductsByBrand(isInteractive)
        // Если false, она делает updateProductPositionsFromSheetV2(false)
        smartSortProductsByBrand(false);

        // --- ФИНАЛ ---
        logPageChange(categoryId, '✅ ЗАВЕРШЕНО OPERATION JUMPSTART');

        SpreadsheetApp.getActiveSpreadsheet().toast('🎉 Протокол успешно выполнен!', '✅ Готово', 10);
        ui.alert('Operation Jumpstart',
            '✅ Протокол восстановления выполнен успешно!\n\n' +
            '1. Описание: Сгенерировано (см. колонку C)\n' +
            '2. Теги: Созданы (см. блок "Плитка тегов")\n' +
            '3. Товары: Отсортированы и обновлены в InSales\n\n' +
            'Рекомендуется проверить результат на сайте через 5-10 минут.',
            ui.ButtonSet.OK
        );

    } catch (error) {
        console.error('❌ [Recovery] Ошибка:', error);
        logError('Ошибка Recovery Protocol', error);
        ui.alert('Ошибка выполнения протокола', error.message, ui.ButtonSet.OK);
    }
}
