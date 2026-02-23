/**
 * ========================================
 * АВТОМАТИЗАЦИЯ ОБНОВЛЕНИЯ КАТЕГОРИЙ
 * ========================================
 */

/**
 * Показывает диалоговое окно обновления категорий
 */
function showCategoryUpdaterDialog() {
    try {
        console.log('🚀 Запуск showCategoryUpdaterDialog');
        const selectedCategories = getSelectedCategoriesFromSheet();
        console.log(`📊 Выбрано категорий: ${selectedCategories.length}`);

        if (selectedCategories.length === 0) {
            SpreadsheetApp.getUi().alert('Пожалуйста, отметьте галочками категории в списке, которые вы хотите обновить.');
            return;
        }

        const template = HtmlService.createTemplateFromFile('99_category_updater_dialog');

        // Больше не передаем данные через шаблон, загрузим их с клиента
        // template.categoriesJson = ... 

        console.log('📝 Шаблон создан, данные будут загружены асинхронно');

        const html = template.evaluate()
            .setWidth(600)
            .setHeight(700)
            .setTitle('Массовое обновление категорий');

        console.log('✅ Шаблон оценен (evaluated)');

        SpreadsheetApp.getUi().showModalDialog(html, 'Массовое обновление категорий');

    } catch (e) {
        console.error('❌ Ошибка в showCategoryUpdaterDialog:', e);
        SpreadsheetApp.getUi().alert('Ошибка при открытии окна: ' + e.message);
    }
}

/**
 * Получает отмеченные категории из листа "Категории — Список"
 */
function getSelectedCategoriesFromSheet() {
    console.log('🔍 Запуск getSelectedCategoriesFromSheet');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!sheet) {
        console.error('❌ Лист "Категории — Список" не найден');
        throw new Error('Лист "Категории — Список" не найден');
    }

    // Проверка констант
    console.log('⚙️ MAIN_LIST_COLUMNS:', JSON.stringify(MAIN_LIST_COLUMNS));
    const checkboxColIndex = MAIN_LIST_COLUMNS.CHECKBOX - 1;
    console.log(`⚙️ Индекс колонки чекбокса: ${checkboxColIndex}`);

    const data = sheet.getDataRange().getValues();
    console.log(`📊 Всего строк данных: ${data.length}`);

    const selectedCategories = [];

    // Пропускаем заголовок (строка 1)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const isChecked = row[checkboxColIndex];

        // Логируем первые 5 проверок для отладки
        if (i < 5) {
            console.log(`Row ${i + 1}: val="${isChecked}", type=${typeof isChecked}`);
        }

        // Более надежная проверка чекбокса (учитывает TRUE, "TRUE", "true")
        if (isChecked === true || String(isChecked).toUpperCase() === 'TRUE') {
            console.log(`✅ Найдена отмеченная строка ${i + 1}: "${isChecked}"`);

            selectedCategories.push({
                id: row[MAIN_LIST_COLUMNS.CATEGORY_ID - 1],
                title: row[MAIN_LIST_COLUMNS.TITLE - 1],
                parentId: row[MAIN_LIST_COLUMNS.PARENT_ID - 1]
            });
        }
    }

    console.log(`🏁 Итого выбрано: ${selectedCategories.length}`);
    return selectedCategories;
}

/**
 * Обрабатывает запрос из диалогового окна
 */
function processCategoryUpdatesFromDialog(categoryIds, settings) {
    try {
        const result = automateCategoryUpdates(categoryIds, settings);
        return {
            success: true,
            message: `Готово! Успешно: ${result.success}, Ошибок: ${result.errors}`
        };
    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Основная функция обновления
 */
function automateCategoryUpdates(targetCategoryIds, settings) {
    const context = "Автоматическое обновление категорий";

    try {
        console.log('🚀 Начинаем автоматическое обновление категорий...');

        // 1. Получаем учетные данные
        const credentials = getInsalesCredentialsSync();
        if (!credentials) {
            throw new Error('Не удалось получить учетные данные InSales');
        }

        // 2. Получаем ID дополнительных полей
        console.log('📥 Загружаем дополнительные поля...');
        const fields = fetchCollectionFields(credentials);

        // Маппинг настроек на названия полей
        const fieldMapping = {};
        if (settings.hideInSlider) fieldMapping["Скрыть категорию в слайдере?"] = "да";
        if (settings.hideInMenu) fieldMapping["Скрыть категорию в меню?"] = "да";
        if (settings.hideInList) fieldMapping["Скрыть категорию в списке?"] = "да";

        const fieldIds = {};

        Object.keys(fieldMapping).forEach(name => {
            const field = fields.find(f => f.title === name || f.name === name);
            if (field) {
                fieldIds[name] = field.id;
                console.log(`✅ Найдено поле "${name}": ID ${field.id}`);
            } else {
                console.warn(`⚠️ Поле "${name}" не найдено!`);
            }
        });

        // 3. Обновляем категории
        let successCount = 0;
        let errorCount = 0;

        targetCategoryIds.forEach(categoryId => {
            try {
                console.log(`🔄 Обновляем категорию ID: ${categoryId}...`);

                // Формируем обновления
                const updates = {
                    collection: {
                        field_values_attributes: []
                    }
                };

                // Настройка "Скрыть/Открыть"
                if (settings.unhideCategory) {
                    updates.collection.is_hidden = false;
                }

                // Настройка родителя
                if (settings.newParentId) {
                    updates.collection.parent_id = settings.newParentId === 'root' ? null : parseInt(settings.newParentId);
                }

                // Загружаем полные данные категории для получения существующих field_values
                const fullCollection = loadFullCategoryData(categoryId); // из 13_categories_search.js

                Object.keys(fieldIds).forEach(fieldName => {
                    const fieldId = fieldIds[fieldName];
                    const targetValue = fieldMapping[fieldName];

                    const existingValue = fullCollection.field_values ? fullCollection.field_values.find(fv => fv.collection_field_id === fieldId) : null;

                    if (existingValue) {
                        updates.collection.field_values_attributes.push({
                            id: existingValue.id,
                            value: targetValue
                        });
                    } else {
                        updates.collection.field_values_attributes.push({
                            collection_field_id: fieldId,
                            value: targetValue
                        });
                    }
                });

                // Отправляем запрос только если есть что обновлять
                if (Object.keys(updates.collection).length > 1 || updates.collection.field_values_attributes.length > 0) {
                    const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
                    const options = {
                        method: 'PUT',
                        headers: {
                            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                            'Content-Type': 'application/json'
                        },
                        payload: JSON.stringify(updates),
                        muteHttpExceptions: true
                    };

                    const response = UrlFetchApp.fetch(url, options);
                    if (response.getResponseCode() === 200 || response.getResponseCode() === 204) {
                        console.log(`✅ Категория ${categoryId} успешно обновлена.`);
                        successCount++;
                        // Обновляем форматирование строки в таблице
                        updateSheetRowFormatting(categoryId);
                    } else {
                        console.error(`❌ Ошибка обновления ${categoryId}: ${response.getResponseCode()} ${response.getContentText()}`);
                        errorCount++;
                    }
                } else {
                    console.log(`ℹ️ Нет изменений для ${categoryId}`);
                    successCount++;
                }

                Utilities.sleep(300); // Пауза

            } catch (e) {
                console.error(`❌ Ошибка при обработке ${categoryId}:`, e);
                errorCount++;
            }
        });

        console.log(`🏁 Обновление завершено. Успешно: ${successCount}, Ошибок: ${errorCount}`);

        return { success: successCount, errors: errorCount };

    } catch (error) {
        console.error('❌ Критическая ошибка:', error);
        throw error;
    }
}

/**
 * Получает список всех категорий для выбора родителя (для диалога)
 */
function getAllCategoriesForSelector() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    const data = sheet.getDataRange().getValues();

    const categories = [];
    for (let i = 1; i < data.length; i++) {
        categories.push({
            id: data[i][MAIN_LIST_COLUMNS.CATEGORY_ID - 1],
            title: data[i][MAIN_LIST_COLUMNS.TITLE - 1]
        });
    }
    return categories;
}

/**
 * Получает список всех дополнительных полей коллекций
 */
function fetchCollectionFields(credentials) {
    const url = `${credentials.baseUrl}/admin/collection_fields.json`;
    const options = {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
        throw new Error(`Ошибка получения полей: ${response.getResponseCode()}`);
    }

    return JSON.parse(response.getContentText());
}

/**
 * Обновляет форматирование строки в таблице после успешного обновления
 * (Убирает серый цвет скрытой категории)
 */
function updateSheetRowFormatting(categoryId) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
        if (!sheet) return;

        const data = sheet.getDataRange().getValues();
        const idColIndex = MAIN_LIST_COLUMNS.CATEGORY_ID - 1;

        // Ищем строку с нужным ID
        for (let i = 1; i < data.length; i++) {
            if (data[i][idColIndex] == categoryId) {
                const rowIndex = i + 1;
                const level = data[i][MAIN_LIST_COLUMNS.LEVEL - 1];
                const inStock = data[i][MAIN_LIST_COLUMNS.IN_STOCK_COUNT - 1];

                // Получаем настройку лимита товаров (P1)
                let itemsPerPage = 36;
                const settingValue = sheet.getRange('P1').getValue();
                if (settingValue && !isNaN(parseInt(settingValue))) {
                    itemsPerPage = parseInt(settingValue);
                }

                let bgColor = '#ffffff';
                let fontColor = '#000000';

                // Логика цветов (как в loader):
                // 1. Если товаров мало -> Красный
                if (inStock < itemsPerPage) {
                    bgColor = '#ffccbc';
                }
                // 2. Иначе - по уровню
                else {
                    if (level === 0) bgColor = '#ede7f6';
                    else if (level === 1) bgColor = '#f3e5f5';
                    else if (level === 2) bgColor = '#fce4ec';
                }

                const range = sheet.getRange(rowIndex, 1, 1, 13);
                range.setBackground(bgColor);
                range.setFontColor(fontColor);

                // Также обновляем чекбокс (снимаем галочку, чтобы показать завершение)
                // sheet.getRange(rowIndex, 1).setValue(false); // Опционально, пока не будем

                console.log(`🎨 Форматирование обновлено для строки ${rowIndex} (ID ${categoryId})`);
                break;
            }
        }
    } catch (e) {
        console.error(`⚠️ Ошибка обновления форматирования для ${categoryId}:`, e);
    }
}
