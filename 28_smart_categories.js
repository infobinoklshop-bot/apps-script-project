/**
 * ========================================
 * МОДУЛЬ: УМНЫЕ КАТЕГОРИИ
 * ========================================
 * 
 * Позволяет наполнять категории InSales товарами на основе правил,
 * заданных в Google Таблице.
 * 
 * СИНТАКСИС ПРАВИЛ (ячейка B10):
 * Каждое правило с новой строки или через точку с запятой (;).
 * 
 * Примеры:
 * Бренд: Levenhuk
 * Цена > 5000
 * Наличие: Да
 * Тип призмы: Porro
 * 
 * Если указано несколько правил, они объединяются через И (AND).
 * То есть товар должен соответствовать ВСЕМ правилам.
 */

// КОНСТАНТЫ
const SMART_CATEGORY_CONFIG = {
    RULES_CELL: 'B10',
    STRICT_MODE_CELL: 'B11',
    SECTION_CELL: 'B12' // Опционально: указать конкретный раздел для ускорения
};

/**
 * Основная функция синхронизации умной категории
 */
const SMART_RULES_FIELD_HANDLE = 'smart_rules';

/**
 * Основная функция синхронизации умной категории
 */
function syncSmartCategory() {
    const context = 'syncSmartCategory';
    const ui = SpreadsheetApp.getUi();

    try {
        const sheet = SpreadsheetApp.getActiveSheet();
        if (!sheet.getName().startsWith('Категория — ')) {
            ui.alert('Ошибка', 'Запускайте только с детального листа категории', ui.ButtonSet.OK);
            return;
        }

        // 1. Читаем настройки и правила
        const categoryId = sheet.getRange('B2').getValue();
        const rulesText = sheet.getRange(SMART_CATEGORY_CONFIG.RULES_CELL).getValue().toString();
        const isStrictMode = sheet.getRange(SMART_CATEGORY_CONFIG.STRICT_MODE_CELL).getValue() === true;

        // Пока не используем, но зарезервируем
        // const targetSection = sheet.getRange(SMART_CATEGORY_CONFIG.SECTION_CELL).getValue(); 

        if (!rulesText || rulesText.trim() === '') {
            ui.alert('Ошибка', `Не заданы правила в ячейке ${SMART_CATEGORY_CONFIG.RULES_CELL}`, ui.ButtonSet.OK);
            return;
        }

        // 1.1. Сохраняем правила в InSales (Additional Field)
        // Делаем это ДО начала длительной операции
        try {
            updateCategorySmartRulesInInSales(categoryId, rulesText);
        } catch (e) {
            console.error('Ошибка сохранения правил в InSales:', e);
            // Не прерываем, но уведомляем
            sheet.toast('⚠️ Не удалось синхронизировать правила с InSales', 'Warning');
        }

        // Парсим правила
        const rules = parseSmartRules(rulesText);

        const confirmMessage = `Будет выполнена синхронизация по правилам:\n` +
            rules.map(r => `• ${r.field} ${r.operator} ${r.value}`).join('\n') +
            `\n\nРежим: ${isStrictMode ? '🔒 СТРОГИЙ (удаление лишних)' : '🔓 Мягкий (только добавление)'}` +
            `\n\nПродолжить?`;

        if (ui.alert('Синхронизация умной категории', confirmMessage, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
            return;
        }

        SpreadsheetApp.getActiveSpreadsheet().toast('Ищем товары по правилам...', '🔍 Поиск', -1);

        // 2. Ищем товары во всех кэшированных разделах
        const allSections = getImportedSections(); // из 25_products_cache.js
        let allMatches = [];

        if (allSections.length === 0) {
            ui.alert('Ошибка', 'Нет кэшированных разделов. Сделайте импорт CSV.', ui.ButtonSet.OK);
            return;
        }

        for (const section of allSections) {
            console.log(`🔍 Сканируем раздел: ${section}`);
            try {
                const result = getProductsFromCSVSheet(section); // из 25_products_cache.js
                const products = result.products;

                // Фильтруем
                const sectionMatches = filterProductsByRules(products, rules);
                console.log(`   ✅ Найдено совпадений: ${sectionMatches.length}`);

                allMatches = allMatches.concat(sectionMatches);
            } catch (e) {
                console.error(`❌ Ошибка чтения раздела ${section}: ${e.message}`);
            }
        }

        console.log(`📊 Всего найдено подходящих товаров: ${allMatches.length}`);

        if (allMatches.length === 0) {
            ui.alert('Результат', 'Подходящих товаров не найдено.', ui.ButtonSet.OK);
            return;
        }

        // 3. Синхронизируем с InSales
        // 3.1. Получаем текущие товары в категории (чтобы не добавлять дубли)
        const currentProductIds = getCurrentCategoryProductIds(sheet); // из 17_categories_products.gs.js
        const currentIdsSet = new Set(currentProductIds);

        // Определяем, кого добавить
        const matchesIds = allMatches.map(p => p.id);
        const toAdd = matchesIds.filter(id => !currentIdsSet.has(id));

        SpreadsheetApp.getActiveSpreadsheet().toast(
            `Найдено: ${allMatches.length}. Новых: ${toAdd.length}.`,
            '🚀 Синхронизация',
            -1
        );

        let addedCount = 0;
        let removedCount = 0;

        // Добавляем новые
        if (toAdd.length > 0) {
            console.log(`➕ Добавляем ${toAdd.length} товаров...`);
            const addResult = addProductsToCategory(categoryId, toAdd); // из 17_categories_products.gs.js
            addedCount = addResult.success;
        }

        // 3.2. Если строгий режим - удаляем лишние
        if (isStrictMode) {
            const matchesSet = new Set(matchesIds);
            const toRemove = currentProductIds.filter(id => !matchesSet.has(id));

            if (toRemove.length > 0) {
                console.log(`✂️ [СТРОГИЙ РЕЖИМ] Удаляем ${toRemove.length} лишних товаров...`);
                const removeResult = removeProductsFromCategoryAPI(categoryId, toRemove); // из 17_categories_products.gs.js
                removedCount = removeResult.success;

                // Также удаляем строки из таблицы для консистентности
                // (Опционально, но желательно для визуального соответствия)
                // Но это сложно сделать надежно без полного перестроения, 
                // поэтому пока оставим только удаление из API. 
                // Пользователь сам обновит таблицу, если надо (или мы можем вызвать очистку).
            }
        }

        ui.alert(
            'Готово',
            `Всего подходящих: ${allMatches.length}\n` +
            `Добавлено: ${addedCount}\n` +
            `Удалено (строгий режим): ${removedCount}\n` +
            `Правила сохранены в InSales (smart_rules).`,
            ui.ButtonSet.OK
        );

    } catch (error) {
        console.error(`❌ Ошибка syncSmartCategory:`, error);
        ui.alert('Ошибка: ' + error.message);
    }
}

/**
 * Открывает диалог конструктора правил
 */
function showSmartRulesBuilderDialog() {
    const html = HtmlService.createHtmlOutputFromFile('28_smart_rules_dialog')
        .setWidth(600)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Конструктор умных правил');
}

/**
 * (Backend) Получает данные для инициализации конструктора
 */
function getInitDataForBuilder() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const currentRules = sheet.getRange(SMART_CATEGORY_CONFIG.RULES_CELL).getValue().toString();
    const sections = getImportedSections();
    return {
        currentRules: currentRules,
        sections: sections
    };
}

/**
 * (Backend) Получает характеристики раздела для конструктора
 */
function getSectionCharacteristics(sectionName) {
    // Получаем кэш. Но нам нужны сами названия характеристик и значения
    // getProductsFromCSVSheet возвращает characteristicsValues (Set)
    const result = getProductsFromCSVSheet(sectionName);

    // Преобразуем Set в Array для передачи в клиент
    const charDict = {};
    for (const [key, valSet] of Object.entries(result.characteristicsValues)) {
        charDict[key] = Array.from(valSet).sort();
    }
    return charDict;
}

/**
 * (Backend) Сохраняет правила в ячейку
 */
function saveRulesToSheet(rulesText) {
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(SMART_CATEGORY_CONFIG.RULES_CELL).setValue(rulesText);
}

/**
 * Сохраняет правила в доп. поле InSales
 */
function updateCategorySmartRulesInInSales(categoryId, rulesText) {
    console.log(`💾 Сохраняем smart_rules в InSales для ${categoryId}`);
    updateCategoryAdditionalField(categoryId, SMART_RULES_FIELD_HANDLE, rulesText);
}

/**
 * Загружает правила из InSales в ячейку (вызывать при открытии листа или вручную)
 */
function loadSmartRulesFromInSales(categoryId, sheet) {
    try {
        if (!sheet) sheet = SpreadsheetApp.getActiveSheet();

        // Получаем поля через API
        const fields = getCategoryAdditionalFields(categoryId);
        // fields обычно массив: [{'field_id':123, 'value':'Text', 'key':'handle' (не всегда)}]
        // В InSales API fields_values не всегда возвращает handle. 
        // Если API возвращает только value и field_id, нам нужно сопоставить
        // Но для простоты предположим, что мы ищем конкретный known field_id или перебираем

        // ВАЖНО: getCategoryAdditionalFields мы реализовали как возврат JSON.
        // Нужно найти поле, соответствующее smart_rules.
        // Но если в ответе нет handle, нам нужно знать ID поля.

        const smartFieldId = getFieldIdByHandle(SMART_RULES_FIELD_HANDLE);
        if (!smartFieldId) return; // Поле не настроено в InSales

        const field = fields.find(f => f.field_id == smartFieldId);
        if (field && field.value) {
            console.log(`📥 Загружены правила из InSales: ${field.value}`);
            sheet.getRange(SMART_CATEGORY_CONFIG.RULES_CELL).setValue(field.value);
        }

    } catch (e) {
        console.error('Ошибка загрузки правил из InSales:', e);
    }
}


/**
 * Массовое наполнение отмеченных категорий из главного листа
 */
function bulkFillSmartCategories() {
    const context = 'bulkFillSmartCategories';
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();

    if (sheet.getName() !== CATEGORY_SHEETS.MAIN_LIST) {
        ui.alert('Ошибка', 'Запускайте с главного листа категорий', ui.ButtonSet.OK);
        return;
    }

    // Получаем данные
    const data = sheet.getDataRange().getValues();
    const selectedCategories = [];

    // Ищем отмеченные чекбоксом
    for (let i = 2; i < data.length; i++) {
        if (data[i][0] === true) { // Checkbox
            selectedCategories.push({
                row: i + 1,
                id: data[i][MAIN_LIST_COLUMNS.CATEGORY_ID - 1],
                title: data[i][MAIN_LIST_COLUMNS.TITLE - 1],
                // Предполагаем, что колонка правил (S) - это 19-я колонка
                // Нужно проверить структуру, либо добавить константу.
                // Пока возьмем значение как есть, если колонки нет - будет пусто.
                // TODO: Убедиться, что колонка есть. Если нет - читать из InSales?
                // Лучше читать из InSales для надежности!
            });
        }
    }

    if (selectedCategories.length === 0) {
        ui.alert('Нет выбранных категорий');
        return;
    }

    if (ui.alert('Массовое наполнение', `Будет обработано категорий: ${selectedCategories.length}. Продолжить?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
        return;
    }

    // Кэшируем характеристики один раз, чтобы не дергать каждый раз?
    // Нет, памяти не хватит. Будем читать по очереди.
    const allSections = getImportedSections();
    if (allSections.length === 0) {
        ui.alert('Нет кэша товаров');
        return;
    }

    let processed = 0;

    for (const cat of selectedCategories) {
        sheet.toast(`Обработка: ${cat.title}`, '⏳ In Progress', -1);
        try {
            // 1. Получаем правила.
            // Сначала пробуем из InSales (самое актуальное)
            const fields = getCategoryAdditionalFields(cat.id);
            const smartFieldId = getFieldIdByHandle(SMART_RULES_FIELD_HANDLE);
            let rulesText = '';

            if (smartFieldId) {
                const field = fields.find(f => f.field_id == smartFieldId);
                if (field) rulesText = field.value;
            }

            if (!rulesText) {
                console.warn(`⚠️ Нет правил для ${cat.title} в InSales. Пропускаем.`);
                // Можно попробовать прочитать из колонки таблицы, если добавим её
                continue;
            }

            // 2. Выполняем поиск и синхронизацию (аналог syncSmartCategory, но без UI)
            // Чтобы не дублировать код, вынесем ядро в отдельную функцию processSmartCategoryCore
            const result = processSmartCategoryCore(cat.id, rulesText, false, allSections); // false = strict mode default off for bulk? Or read from another field?
            // Для безопасности в Bulk режиме Strict Mode лучше выключить или читать настройку.
            // Будем считать Safe Mode (только добавление).

            console.log(`✅ ${cat.title}: Добавлено ${result.added}, Найдено ${result.total}`);
            processed++;

        } catch (e) {
            console.error(`❌ Ошибка ${cat.title}:`, e);
        }
    }

    ui.alert('Готово', `Обработано категорий: ${processed} из ${selectedCategories.length}`, ui.ButtonSet.OK);
}

/**
 * Ядро синхронизации (без UI алертов), возвращает статистику
 */
function processSmartCategoryCore(categoryId, rulesText, isStrictMode, allSections) {
    const rules = parseSmartRules(rulesText);
    let allMatches = [];

    for (const section of allSections) {
        try {
            const result = getProductsFromCSVSheet(section);
            const products = result.products;
            const sectionMatches = filterProductsByRules(products, rules);
            allMatches = allMatches.concat(sectionMatches);
        } catch (e) {
            console.error(`Error scanning section ${section}: ${e.message}`);
        }
    }

    if (allMatches.length === 0) return { added: 0, removed: 0, total: 0 };

    // Получаем текущие позиции (нужен только список ID)
    // getCurrentCategoryProductIds работает с листом. У нас нет листа в bulk mode?
    // Нужно использовать API напрямую!
    // getAllCollectsForCategory из 19_categories_insales.js
    const credentials = getInsalesCredentialsSync();
    const collects = getAllCollectsForCategory(categoryId, credentials);
    const currentProductIds = collects.map(c => c.product_id);
    const currentIdsSet = new Set(currentProductIds);

    const matchesIds = allMatches.map(p => p.id);
    const toAdd = matchesIds.filter(id => !currentIdsSet.has(id));

    let addedCount = 0;
    if (toAdd.length > 0) {
        const addResult = addProductsToCategory(categoryId, toAdd);
        addedCount = addResult.success;
    }

    let removedCount = 0;
    if (isStrictMode) {
        const matchesSet = new Set(matchesIds);
        const toRemove = currentProductIds.filter(id => !matchesSet.has(id));
        if (toRemove.length > 0) {
            const removeResult = removeProductsFromCategoryAPI(categoryId, toRemove);
            removedCount = removeResult.success;
        }
    }

    return { added: addedCount, removed: removedCount, total: allMatches.length };
}

/**
 * Парсит текстовые правила в массив объектов

 * @param {string} text - текст правил
 * @returns {Array} [{field, operator, value}]
 */
function parseSmartRules(text) {
    const rules = [];
    // Разбиваем по переносам строк или точке с запятой
    const lines = text.split(/[\n;]+/).map(l => l.trim()).filter(l => l);

    for (const line of lines) {
        let operator = '=';
        let field = '';
        let value = '';

        // Определяем оператор
        if (line.includes('>=')) operator = '>=';
        else if (line.includes('<=')) operator = '<=';
        else if (line.includes('>')) operator = '>';
        else if (line.includes('<')) operator = '<';
        else if (line.includes(':')) operator = ':'; // тоже равно, но для читаемости
        else if (line.includes('=')) operator = '=';

        // Разбиваем строку
        let parts;
        if (operator === ':') {
            parts = line.split(':');
            operator = '='; // Нормализуем оператор для логики
        } else {
            // Для математических операторов нужно аккуратно разбить
            // Бренд=Levenhuk -> ['Brest', 'Levenhuk']
            // Цена>5000 -> ['Цена', '5000']
            const regex = new RegExp(`^(.+?)(${operator})(.+)$`);
            const match = line.match(regex);
            if (match) {
                parts = [match[1], match[3]]; // match[2] это оператор
            } else {
                // Fallback если сплит не сработал через regex
                parts = line.split(operator);
            }
        }

        if (parts && parts.length >= 2) {
            field = parts[0].trim();
            value = parts.slice(1).join(operator === ':' ? ':' : operator).trim(); // value might contain separator? unlikely with this logic but safe join

            // Удаляем кавычки если есть
            value = value.replace(/^["']|["']$/g, '');

            // Нормализуем поле
            field = normalizeFieldName(field);

            rules.push({ field, operator, value });
        }
    }

    return rules;
}

/**
 * Нормализует название поля для сравнения
 */
function normalizeFieldName(name) {
    name = name.toLowerCase();

    if (name.includes('бренд') || name.includes('brand') || name.includes('производитель')) return 'brand';
    if (name.includes('цена') || name.includes('price')) return 'price';
    if (name.includes('название') || name.includes('title')) return 'title';
    if (name.includes('артикул') || name.includes('sku')) return 'sku';
    if (name.includes('наличие') || name.includes('stock')) return 'stock';

    // Остальное считаем характеристикой как есть
    return name;
}

/**
 * Фильтрует товары по массиву правил
 */
function filterProductsByRules(products, rules) {
    return products.filter(product => {
        // Товар подходит, если он удовлетворяет ВСЕМ правилам (AND)
        for (const rule of rules) {
            const { field, operator, value } = rule;
            let productValue = null;

            // 1. Извлекаем значение из товара
            if (field === 'price') {
                productValue = product.price;
            } else if (field === 'stock') {
                productValue = product.in_stock ? 'да' : 'нет'; // Нормализация для простоты
            } else if (field === 'brand') {
                // Бренд часто в характеристиках или (если есть отдельное поле)
                // В 25_products_cache мы не сохраняли vendor отдельно в объект product явно,
                // но он может быть в характеристиках. 
                // Проверяем характеристики "Бренд" или "Производитель"
                productValue = getCharacteristicValue(product, ['Бренд', 'Производитель']);
            } else if (field === 'title') {
                productValue = product.title;
            } else if (field === 'sku') {
                productValue = product.sku;
            } else {
                // Ищем в характеристиках по точному названию или похожему
                productValue = getCharacteristicValue(product, [field]);
            }

            // Если значения нет, то правило не выполнено
            if (productValue === null || productValue === undefined || productValue === '') {
                return false;
            }

            // 2. Сравниваем
            if (!compareValues(productValue, operator, value)) {
                return false;
            }
        }

        return true;
    });
}

/**
 * Получает значение характеристики из товара с учетом вариантов названия
 */
function getCharacteristicValue(product, possibleNames) {
    if (!product.characteristics) return null;

    for (const name of possibleNames) {
        // Ищем точное совпадение
        if (product.characteristics[name]) return product.characteristics[name];

        // Ищем регистронезависимое
        const key = Object.keys(product.characteristics).find(k => k.toLowerCase() === name.toLowerCase());
        if (key) return product.characteristics[key];
    }
    return null;
}

/**
 * Сравнивает значения
 */
function compareValues(prodVal, op, ruleVal) {
    // Числовое сравнение
    const prodNum = parseFloat(prodVal.toString().replace(/\s/g, '').replace(',', '.'));
    const ruleNum = parseFloat(ruleVal.toString().replace(/\s/g, '').replace(',', '.'));
    const isNumeric = !isNaN(prodNum) && !isNaN(ruleNum);

    if (isNumeric && (op === '>' || op === '<' || op === '>=' || op === '<=')) {
        switch (op) {
            case '>': return prodNum > ruleNum;
            case '<': return prodNum < ruleNum;
            case '>=': return prodNum >= ruleNum;
            case '<=': return prodNum <= ruleNum;
        }
    }

    // Строковое сравнение
    const sProd = String(prodVal).toLowerCase().trim();
    const sRule = String(ruleVal).toLowerCase().trim();

    if (op === '=' || op === ':') {
        return sProd === sRule;
    }

    // Частичное совпадение можно добавить (например "~")

    return sProd == sRule; // Fallback
}
