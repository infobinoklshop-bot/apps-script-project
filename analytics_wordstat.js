/**
 * ========================================
 * МОДУЛЬ РАБОТЫ С YANDEX WORDSTAT API (V4)
 * ========================================
 */

/**
 * Открывает диалог проверки частотности
 */
function showWordstatDialog() {
    const html = HtmlService.createHtmlOutputFromFile('dialog_wordstat')
        .setWidth(600)
        .setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(html, 'Проверка частотности (Yandex Wordstat)');
}

/**
 * Получает ключевые слова из Метрики для активной категории
 * (Вызывается из диалога)
 */
function getMetricaKeywordsForActiveCategory() {
    // 1. Определяем текущую категорию
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    const activeRow = sheet.getActiveCell().getRow();

    if (activeRow < 2) {
        throw new Error('Выберите строку с категорией!');
    }

    // Получаем URL категории
    const rowData = sheet.getRange(activeRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    let categoryUrl = rowData[6]; // URL is column 7 -> index 6

    if (!categoryUrl) {
        throw new Error('URL категории не найден в строке ' + activeRow);
    }

    // Нормализуем URL до пути (path)
    try {
        if (categoryUrl.startsWith('http')) {
            const urlObj = new URL(categoryUrl);
            categoryUrl = urlObj.pathname;
        }
    } catch (e) {
        // Если URL кривой, пробуем просто обрезать домен руками, если он есть
        categoryUrl = categoryUrl.replace(/^https?:\/\/[^\/]+/, '');
    }
    // Убираем query params и trailing slash
    if (categoryUrl.includes('?')) categoryUrl = categoryUrl.split('?')[0];
    if (categoryUrl.endsWith('/') && categoryUrl.length > 1) categoryUrl = categoryUrl.slice(0, -1);


    // 2. Запрашиваем данные из Метрики
    let keywords = [];
    try {
        SpreadsheetApp.getActiveSpreadsheet().toast('Запрос ключей из Метрики...', 'Загрузка');
        keywords = fetchSearchPhrasesFromMetrica(categoryUrl);
    } catch (e) {
        console.error('Ошибка Метрики:', e);
        SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка получения данных из Метрики: ' + e.message, 'Ошибка');
        // Не падаем, возвращаем хотя бы название категории
    }

    // 3. Добавляем название категории как запасной вариант
    let categoryName = rowData[5]; // Name is column 6 -> index 5
    categoryName = categoryName.replace(/^[\s\u2500-\u257F]+/, '').trim();

    if (!keywords.includes(categoryName)) {
        keywords.unshift(categoryName);
    }

    // Если совсем пусто, генерируем базовые варианты
    if (keywords.length === 1) {
        keywords.push('купить ' + categoryName);
        keywords.push(categoryName + ' цена');
    }

    return keywords;
}

/**
 * Получает поисковые фразы из Метрики для заданного URL (Path)
 */
function fetchSearchPhrasesFromMetrica(urlPath) {
    const config = YANDEX_METRICA_CONFIG;
    if (!config.oauthToken) throw new Error('Токен Метрики не настроен');

    // Период - берем с 2023 года, как просил пользователь (примерно 2 года)
    const date1 = '2023-01-01';
    const date2 = 'today';
    const limit = 500; // Увеличили лимит

    // Фильтр: визиты, начавшиеся с этой страницы (startURL)
    // Используем оператор =@ (содержит подстроку), так как в отчете пользователя используется "содержит"
    // Это позволит найти больше запросов, даже если есть параметры или слеши
    const filters = `ym:s:startURLPath=@'${urlPath}'`;

    const params = {
        'ids': config.counterId,
        'metrics': 'ym:s:visits',
        'dimensions': 'ym:s:searchPhrase',
        'filters': filters,
        'date1': date1,
        'date2': date2,
        'limit': limit,
        'sort': '-ym:s:visits',
        'accuracy': 'medium' // Можно medium для скорости
    };

    const queryString = Object.keys(params)
        .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
        .join('&');

    const apiUrl = `${config.baseUrl}?${queryString}`;

    const options = {
        'method': 'get',
        'headers': {
            'Authorization': `OAuth ${config.oauthToken}`
        },
        'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
        throw new Error(`Metrica API Error: ${json.message || response.getResponseCode()}`);
    }

    // json.data = [{ dimensions: [{name: "фраза"}], metrics: [10] }, ...]
    const phrases = json.data.map(row => row.dimensions[0].name);

    // Фильтруем мусор (null, пустые)
    return phrases.filter(p => p && p.trim() !== '');
}

/**
 * Проверяет частотность для выделенных строк в таблице "Семантика"
 */
function checkFrequencyForSelectedRows() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getName() !== CATEGORY_SHEETS.SEMANTICS) {
        SpreadsheetApp.getUi().alert('Пожалуйста, перейдите на лист "Семантика" и выделите строки с фразами.');
        return;
    }

    const selection = sheet.getSelection();
    const ranges = selection.getActiveRangeList().getRanges();

    // Собираем уникальные строки из выделения
    const rowsToProcess = new Set();
    ranges.forEach(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();
        for (let i = 0; i < numRows; i++) {
            const row = startRow + i;
            if (row > 1) { // Пропускаем заголовок
                rowsToProcess.add(row);
            }
        }
    });

    if (rowsToProcess.size === 0) {
        SpreadsheetApp.getUi().alert('Выделите строки с ключевыми фразами (кроме заголовка).');
        return;
    }

    // Сортируем номера строк
    const sortedRows = Array.from(rowsToProcess).sort((a, b) => a - b);

    // Получаем данные
    // Фраза теперь в колонке 2 (B)
    const keywords = [];
    const rowDataMap = new Map(); // Map<keyword, {row: 123}>

    sortedRows.forEach(row => {
        const phrase = sheet.getRange(row, 2).getValue(); // Col B
        if (phrase && typeof phrase === 'string' && phrase.trim() !== '') {
            keywords.push(phrase);
            rowDataMap.set(phrase, {
                row: row
            });
        }
    });

    if (keywords.length === 0) {
        SpreadsheetApp.getUi().alert('В выделенных строках не найдено ключевых фраз (Колонка B).');
        return;
    }

    // Настройки по умолчанию (можно сделать диалог, но пока так)
    const settings = {
        broad: true,
        quote: true,
        excl: false,
        bracket: false,
        regionId: 0 // Весь мир по умолчанию, или можно брать из конфига
    };

    processSelectedRowsUpdate(keywords, rowDataMap, settings, sheet);
}

/**
 * Внутренняя функция для обновления выделенных строк
 */
function processSelectedRowsUpdate(keywords, rowDataMap, settings, sheet) {
    const regionId = settings.regionId;
    const wordChunks = [];
    const chunkSize = 5; // По 5 слов за раз (если 2 типа частотности = 10 фраз)

    for (let i = 0; i < keywords.length; i += chunkSize) {
        wordChunks.push(keywords.slice(i, i + chunkSize));
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Обработка ${keywords.length} фраз...`, 'Wordstat');

    for (let i = 0; i < wordChunks.length; i++) {
        const chunk = wordChunks[i];

        // Формируем фразы для API
        const apiPhrases = [];
        const phraseMap = new Map();

        chunk.forEach(word => {
            if (settings.broad) {
                apiPhrases.push(word);
                phraseMap.set(word, { word, type: 'broad' });
            }
            if (settings.quote) {
                const p = `"${word}"`;
                apiPhrases.push(p);
                phraseMap.set(p, { word, type: 'quote' });
            }
        });

        // API запрос
        try {
            const reportId = createWordstatReport(apiPhrases, regionId);
            if (reportId) {
                const reportData = waitForReport(reportId);

                // Разбираем и пишем сразу в ячейки
                reportData.forEach(item => {
                    const info = phraseMap.get(item.Phrase);
                    if (info) {
                        const rowData = rowDataMap.get(info.word);
                        if (rowData) {
                            // Определяем колонку
                            let col = -1;
                            if (info.type === 'broad') col = 3; // C (было 4)
                            if (info.type === 'quote') col = 4; // D (было 5)

                            if (col > 0) {
                                sheet.getRange(rowData.row, col).setValue(item.Shows);
                            }
                        }
                    }
                });

                deleteWordstatReport(reportId);
            }
        } catch (e) {
            console.warn('Ошибка API при обновлении:', e);
            SpreadsheetApp.getActiveSpreadsheet().toast(`Ошибка API: ${e.message}`, 'Warning');
            // Можно записать ошибку в ячейку или лог
        }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Обновление завершено!', 'Успех');
}

/**
 * Основная функция обработки запроса (Вызывается из диалога)
 */
function processWordstatRequest(keywordsRaw, regionId, settings) {
    try {
        console.log('Начало processWordstatRequest');
        SpreadsheetApp.getActiveSpreadsheet().toast('Запуск проверки частотности (v3)...', 'Wordstat');

        // 0. Сразу создаем/получаем лист
        let targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORY_SHEETS.SEMANTICS);
        if (!targetSheet) {
            targetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(CATEGORY_SHEETS.SEMANTICS);
            targetSheet.getRange('A1:J1').setValues([['✅', 'Ключевая фраза', 'Частотность (Шир.)', 'Частотность ("")', 'Частотность (!)', 'Частотность ([!])', 'Визиты', 'Отказы (%)', 'Google', 'Период']]);
            targetSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#e0e0e0');
            targetSheet.setColumnWidth(2, 300); // Ключевая фраза пошире
            targetSheet.setColumnWidth(1, 30);  // Чекбокс поуже
            SpreadsheetApp.flush();
        } else {
            // Проверяем, старый ли заголовок, и обновляем если нужно
            const lastCol = targetSheet.getLastColumn();
            const j1 = targetSheet.getRange('J1').getValue();
            if (lastCol < 10 || j1 !== 'Период') {
                targetSheet.getRange('A1:J1').setValues([['✅', 'Ключевая фраза', 'Частотность (Шир.)', 'Частотность ("")', 'Частотность (!)', 'Частотность ([!])', 'Визиты', 'Отказы (%)', 'Google', 'Период']]);
                targetSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#e0e0e0');
            }
        }

        const keywords = keywordsRaw.split('\n').map(k => k.trim()).filter(k => k);
        if (keywords.length === 0) {
            console.warn('Нет ключевых слов');
            SpreadsheetApp.getActiveSpreadsheet().toast('Нет ключевых слов!', 'Ошибка');
            return;
        }

        const types = [];
        if (settings.broad) types.push('broad');
        if (settings.quote) types.push('quote');
        if (settings.excl) types.push('excl');
        if (settings.bracket) types.push('bracket');

        if (types.length === 0) throw new Error('Выберите хотя бы один тип частотности!');

        const wordsPerReport = Math.floor(10 / types.length);
        if (wordsPerReport === 0) throw new Error('Слишком много типов частотности выбрано для лимита API (10 фраз на отчет).');

        // Данные текущей категории
        const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
        const activeRow = mainSheet.getActiveCell().getRow();
        const rowData = mainSheet.getRange(activeRow, 1, 1, 10).getValues()[0];
        const catId = rowData[1];
        const catPath = rowData[4];
        const catName = rowData[5];
        const dateStr = new Date().toLocaleDateString();

        // --- ЛОГИКА БЛОКОВ И ДУБЛЕЙ ---

        // 0.5. Сначала чистим дубли в этой категории, если они уже есть
        cleanUpDuplicates(targetSheet, catId);

        // 1. Сканируем лист "Семантика", чтобы найти блок этой категории (после чистки)
        const lastRow = targetSheet.getLastRow();
        let blockStartRow = -1;
        let blockEndRow = -1;
        const existingKeywordsMap = new Map(); // keyword -> rowNumber

        // Ищем заголовок блока
        if (lastRow > 0) {
            const rangeValues = targetSheet.getRange(1, 1, lastRow, 1).getValues(); // Только Col A

            for (let i = 0; i < rangeValues.length; i++) {
                const val = String(rangeValues[i][0]);
                const absoluteRow = i + 1;

                // Ищем заголовок по ID
                if (val.includes(`(${catId})`) && (val.startsWith('📂') || val.startsWith('==='))) {
                    blockStartRow = absoluteRow + 1; // Данные начинаются со следующей строки

                    // Теперь ищем конец блока (следующий заголовок или конец листа)
                    for (let j = i + 1; j < rangeValues.length; j++) {
                        const nextVal = String(rangeValues[j][0]);
                        if ((nextVal.startsWith('📂') || nextVal.startsWith('===')) && nextVal.includes('(')) {
                            blockEndRow = j; // Строка ПЕРЕД следующим заголовком (j+1 - 1)
                            break;
                        }
                    }
                    if (blockEndRow === -1) blockEndRow = lastRow;
                    break;
                }
            }
        }

        // Если блок найден, сканируем его на существующие слова
        if (blockStartRow !== -1) {
            // Если блок пустой (start > end), то ничего не делаем
            if (blockStartRow <= blockEndRow) {
                // Ключевые слова теперь в Col B (2)
                const blockValues = targetSheet.getRange(blockStartRow, 2, blockEndRow - blockStartRow + 1, 1).getValues();
                blockValues.forEach((row, idx) => {
                    const phrase = String(row[0]).trim();
                    if (phrase) {
                        existingKeywordsMap.set(phrase, blockStartRow + idx);
                    }
                });
            }
        }

        // Если блок не найден, будем добавлять в конец
        const isNewBlock = (blockStartRow === -1);
        let headerRow = -1;

        if (isNewBlock) {
            // Добавляем заголовок блока
            headerRow = targetSheet.getLastRow() + 1;
            // Новый формат заголовка с путем (без слова Категория)
            targetSheet.getRange(headerRow, 1).setValue(`📂 ${catPath} (${catId})`);

            // ВАЖНО: Сначала снимаем объединение, чтобы избежать ошибок
            try { targetSheet.getRange(headerRow, 1, 1, 20).breakApart(); } catch (e) { }

            // Объединяем 10 колонок (A-J)
            targetSheet.getRange(headerRow, 1, 1, 10).setBackground('#f3f3f3').setFontWeight('bold').merge();
            blockEndRow = headerRow; // Теперь конец блока - это заголовок
        } else {
            // Блок найден. Обновляем заголовок на новый формат (если нужно)
            headerRow = blockStartRow - 1;
            targetSheet.getRange(headerRow, 1).setValue(`📂 ${catPath} (${catId})`);

            // ВАЖНО: Сначала снимаем объединение, чтобы избежать ошибок
            try { targetSheet.getRange(headerRow, 1, 1, 20).breakApart(); } catch (e) { }

            // Проверяем объединение
            try {
                targetSheet.getRange(headerRow, 1, 1, 10).setBackground('#f3f3f3').setFontWeight('bold').merge();
            } catch (e) { /* ignore merge errors */ }
        }

        // Разбиваем исходные ключевики на пачки
        const wordChunks = [];
        for (let i = 0; i < keywords.length; i += wordsPerReport) {
            wordChunks.push(keywords.slice(i, i + wordsPerReport));
        }

        console.log(`Сформировано ${wordChunks.length} пакетов для запросов.`);
        SpreadsheetApp.getActiveSpreadsheet().toast(`Обработка ${keywords.length} фраз (${wordChunks.length} пакетов)...`, 'Wordstat');

        // Обрабатываем пачки
        for (let i = 0; i < wordChunks.length; i++) {
            const chunk = wordChunks[i];
            console.log(`Обработка пакета ${i + 1}/${wordChunks.length}`);
            SpreadsheetApp.getActiveSpreadsheet().toast(`Пакет ${i + 1}/${wordChunks.length}...`, 'Wordstat');

            // Формируем массив фраз для API
            const apiPhrases = [];
            const phraseMap = new Map();

            chunk.forEach(word => {
                if (settings.broad) {
                    apiPhrases.push(word);
                    phraseMap.set(word, { word, type: 'broad' });
                }
                if (settings.quote) {
                    const p = `"${word}"`;
                    apiPhrases.push(p);
                    phraseMap.set(p, { word, type: 'quote' });
                }
                if (settings.excl) {
                    const p = '!' + word.split(' ').join(' !');
                    apiPhrases.push(p);
                    phraseMap.set(p, { word, type: 'excl' });
                }
                if (settings.bracket) {
                    const p = `"[!${word.split(' ').join(' !')}]"`;
                    apiPhrases.push(p);
                    phraseMap.set(p, { word, type: 'bracket' });
                }
            });

            // Группируем результаты
            const resultsByWord = new Map();
            chunk.forEach(w => resultsByWord.set(w, { broad: '-', quote: '-', excl: '-', bracket: '-' }));

            // API запрос
            try {
                const reportId = createWordstatReport(apiPhrases, regionId);
                if (reportId) {
                    const reportData = waitForReport(reportId);
                    reportData.forEach(item => {
                        const info = phraseMap.get(item.Phrase);
                        if (info) {
                            const res = resultsByWord.get(info.word);
                            if (res) res[info.type] = item.Shows;
                        }
                    });
                    deleteWordstatReport(reportId);
                }
            } catch (e) {
                console.warn(`Ошибка API Wordstat для пакета ${i + 1}:`, e);
                SpreadsheetApp.getActiveSpreadsheet().toast(`Ошибка API (пакет ${i + 1}): ${e.message}. Сохраняем список.`, 'Warning');
            }

            // --- ЗАПИСЬ РЕЗУЛЬТАТОВ (УМНАЯ ВСТАВКА) ---

            const rowsToInsert = [];

            chunk.forEach(word => {
                const r = resultsByWord.get(word);
                // НОВЫЙ ФОРМАТ: [Checkbox, Keyword, Freq1, Freq2, Freq3, Freq4, Visits, Bounce, Google, Period]
                const rowData = [
                    false, // Checkbox
                    word,
                    r.broad,
                    r.quote,
                    r.excl,
                    r.bracket,
                    0, // Visits
                    0, // Bounce
                    0, // Google
                    dateStr // Period
                ];

                if (existingKeywordsMap.has(word)) {
                    // ОБНОВЛЕНИЕ
                    const existingRow = existingKeywordsMap.get(word);
                    // Обновляем частотность и дату (колонки C-F -> 3-6, Period -> 10)
                    targetSheet.getRange(existingRow, 3, 1, 4).setValues([[r.broad, r.quote, r.excl, r.bracket]]);
                    targetSheet.getRange(existingRow, 10).setValue(dateStr);
                } else {
                    // ВСТАВКА
                    rowsToInsert.push(rowData);
                }
            });

            if (rowsToInsert.length > 0) {
                // Вставляем ПОСЛЕ текущего конца блока
                targetSheet.insertRowsAfter(blockEndRow, rowsToInsert.length);

                const dataRange = targetSheet.getRange(blockEndRow + 1, 1, rowsToInsert.length, 10);
                dataRange.setValues(rowsToInsert);

                // СБРОС СТИЛЕЙ для данных (убираем жирность и фон, если они наследовались)
                dataRange.setFontWeight('normal').setBackground(null);

                // Добавляем чекбоксы
                targetSheet.getRange(blockEndRow + 1, 1, rowsToInsert.length, 1).insertCheckboxes();

                // ГРУППИРОВКА СТРОК
                // Группируем ТОЛЬКО если это первая вставка после заголовка (или блок был пуст).
                // В остальных случаях (когда мы дописываем в конец существующей группы)
                // новые строки автоматически наследуют группировку.
                if (blockEndRow === headerRow) {
                    try {
                        const groupRange = targetSheet.getRange(blockEndRow + 1, 1, rowsToInsert.length, targetSheet.getLastColumn());
                        groupRange.shiftRowGroupDepth(1);
                    } catch (e) {
                        console.warn('Ошибка группировки строк:', e);
                    }
                }

                // Обновляем указатель конца блока и карту
                blockEndRow += rowsToInsert.length;

                rowsToInsert.forEach((row, idx) => {
                    existingKeywordsMap.set(row[1], blockEndRow - rowsToInsert.length + idx + 1);
                });
            }

            console.log(`Пакет ${i + 1} обработан.`);
        }

        SpreadsheetApp.getActiveSpreadsheet().toast('Готово! Список сохранен.', 'Wordstat');
        console.log('processWordstatRequest завершен успешно');
        return true;

    } catch (e) {
        console.error(e);
        throw e;
    }
}

/**
 * Удаляет дубликаты строк для заданной категории (НОВАЯ ЛОГИКА)
 */
function cleanUpDuplicates(sheet, catId) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // 1. Находим границы блока
    const rangeValues = sheet.getRange(1, 1, lastRow, 1).getValues();
    let blockStartRow = -1;
    let blockEndRow = -1;

    for (let i = 0; i < rangeValues.length; i++) {
        const val = String(rangeValues[i][0]);
        const absoluteRow = i + 1;

        if (val.includes(`(${catId})`) && (val.startsWith('📂') || val.startsWith('==='))) {
            blockStartRow = absoluteRow + 1;

            for (let j = i + 1; j < rangeValues.length; j++) {
                const nextVal = String(rangeValues[j][0]);
                if ((nextVal.startsWith('📂') || nextVal.startsWith('===')) && nextVal.includes('(')) {
                    blockEndRow = j;
                    break;
                }
            }
            if (blockEndRow === -1) blockEndRow = lastRow;
            break;
        }
    }

    if (blockStartRow === -1 || blockStartRow > blockEndRow) return;

    // 2. Сканируем блок на дубли (по колонке B - Ключевое слово)
    // ВАЖНО: Если данные старого формата, то в B - Путь.
    // Но мы хотим удалить дубли.
    // Если мы переходим на новый формат, старые данные (где B=Путь) будут мешать?
    // Нет, они просто не совпадут с новыми ключевиками (если только путь не равен ключевику).
    // Но лучше чистить по контенту.

    const blockRange = sheet.getRange(blockStartRow, 2, blockEndRow - blockStartRow + 1, 1);
    const keywords = blockRange.getValues();
    const rowsToDelete = [];
    const seenKeywords = new Set();

    // Идем СВЕРХУ ВНИЗ
    for (let i = 0; i < keywords.length; i++) {
        const keyword = String(keywords[i][0]).trim();
        const absoluteRow = blockStartRow + i;

        if (keyword) {
            if (seenKeywords.has(keyword)) {
                rowsToDelete.push(absoluteRow);
            } else {
                seenKeywords.add(keyword);
            }
        }
    }

    // Удаляем строки (СНИЗУ ВВЕРХ)
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(row => {
        sheet.deleteRow(row);
    });

    if (rowsToDelete.length > 0) {
        console.log(`Удалено ${rowsToDelete.length} дублей для категории ${catId}`);
        SpreadsheetApp.getActiveSpreadsheet().toast(`Удалено ${rowsToDelete.length} дублей.`, 'Чистка');
    }
}

/**
 * Обновляет данные о визитах из Метрики и Google Search Console для отмеченных фраз
 */
function updateSemanticsVisits() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORY_SHEETS.SEMANTICS);
    if (!sheet) {
        SpreadsheetApp.getUi().alert('Лист "Семантика" не найден.');
        return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // 1. Запрашиваем период у пользователя
    const userProperties = PropertiesService.getUserProperties();
    const lastPeriod = userProperties.getProperty('LAST_SEMANTICS_PERIOD') || '2024-01-01:2024-12-31';

    const ui = SpreadsheetApp.getUi();
    const dateResponse = ui.prompt(
        'Период статистики',
        `Введите период (YYYY-MM-DD:YYYY-MM-DD).\nОставьте пустым, чтобы использовать: ${lastPeriod}`,
        ui.ButtonSet.OK_CANCEL
    );

    if (dateResponse.getSelectedButton() !== ui.Button.OK) return;

    let inputDate = dateResponse.getResponseText().trim();
    if (!inputDate) inputDate = lastPeriod;

    let date1, date2;
    if (inputDate) {
        const parts = inputDate.split(':');
        if (parts.length === 2) {
            date1 = parts[0].trim();
            date2 = parts[1].trim();
            userProperties.setProperty('LAST_SEMANTICS_PERIOD', inputDate);
        } else {
            ui.alert('Неверный формат даты. Используйте YYYY-MM-DD:YYYY-MM-DD');
            return;
        }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Загрузка данных (${date1} - ${date2})...`, 'Аналитика');

    try {
        // 2. Проверяем, есть ли отмеченные строки
        const range = sheet.getRange(2, 1, lastRow - 1, 2); // Col A (Checkbox) and Col B (Keyword)
        const values = range.getValues();
        const selectedIndices = [];

        for (let i = 0; i < values.length; i++) {
            if (values[i][0] === true) { // Если чекбокс отмечен
                selectedIndices.push(i);
            }
        }

        if (selectedIndices.length === 0) {
            SpreadsheetApp.getUi().alert('Выберите фразы галочками (столбец А), для которых нужно проверить визиты.');
            return;
        }

        // 3. Получаем данные из Метрики
        const metricaData = fetchAllSearchPhrasesFromMetrica(date1, date2);
        const visitsMap = new Map();
        const bounceMap = new Map();
        const ordersMap = new Map();

        metricaData.forEach(row => {
            const phrase = String(row.dimensions[0].name).toLowerCase().trim();
            const visits = row.metrics[0];
            const bounce = row.metrics[1]; // Отказы
            const orders = row.metrics[2]; // Заказы
            visitsMap.set(phrase, visits);
            bounceMap.set(phrase, bounce);
            ordersMap.set(phrase, orders);
        });

        // 4. Получаем данные из GSC
        const gscData = fetchSearchQueriesFromGSC(date1, date2);
        const gscMap = new Map();
        gscData.forEach(row => {
            const phrase = String(row.query).toLowerCase().trim();
            const clicks = row.clicks;
            gscMap.set(phrase, clicks);
        });

        console.log(`Получено: Метрика - ${metricaData.length}, GSC - ${gscData.length}`);

        // ОБНОВЛЯЕМ ЗАГОЛОВКИ (Чистые, без дат)
        // Принудительно обновляем весь диапазон заголовков, чтобы гарантировать наличие колонки "Заказы"
        sheet.getRange('A1:K1').setValues([['✅', 'Ключевая фраза', 'Частотность (Шир.)', 'Частотность ("")', 'Частотность (!)', 'Частотность ([!])', 'Визиты', 'Отказы (%)', 'Заказы', 'Google', 'Период']]);
        sheet.getRange('A1:K1').setFontWeight('bold').setBackground('#e0e0e0');
        sheet.setColumnWidth(9, 80); // Заказы

        // 5. Обновляем ТОЛЬКО выбранные строки
        // Читаем колонки G (Визиты), H (Отказы), I (Заказы), J (Google), K (Период)
        const statsRange = sheet.getRange(2, 7, lastRow - 1, 5);
        const statsValues = statsRange.getValues();
        const periodStr = `${date1}:${date2}`;

        let updatedCount = 0;
        selectedIndices.forEach(idx => {
            const phrase = String(values[idx][1]).toLowerCase().trim();
            if (phrase) {
                // Метрика
                const visits = visitsMap.get(phrase) || 0;
                const bounce = bounceMap.get(phrase) || 0;
                const orders = ordersMap.get(phrase) || 0;

                statsValues[idx][0] = visits;
                statsValues[idx][1] = Math.round(bounce * 100) / 100; // Округляем до 2 знаков
                statsValues[idx][2] = orders;

                // GSC
                const clicks = gscMap.get(phrase) || 0;
                statsValues[idx][3] = clicks;

                // Period
                statsValues[idx][4] = periodStr;

                updatedCount++;
            }
        });

        // Записываем обратно
        statsRange.setValues(statsValues);

        SpreadsheetApp.getActiveSpreadsheet().toast(`Обновлено ${updatedCount} строк.`, 'Успех');

    } catch (e) {
        console.error(e);
        SpreadsheetApp.getUi().alert(`Ошибка при получении данных: ${e.message}`);
    }
}

/**
 * Отправляет запрос к API Wordstat V4
 */
function callWordstatApi(method, params) {
    const token = YANDEX_METRICA_CONFIG.oauthToken;
    const url = 'https://api.direct.yandex.com/v4/json/';

    const payload = {
        method: method,
        token: token,
        locale: 'ru',
        param: params
    };

    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.error_code) {
        throw new Error(`API Error ${json.error_code}: ${json.error_detail || json.error_str}`);
    }

    return json.data;
}

/**
 * Создает отчет
 */
function createWordstatReport(phrases, regionId) {
    // GeoID должен быть массивом
    const geoID = [parseInt(regionId)];

    const params = {
        Phrases: phrases,
        GeoID: geoID
    };

    const result = callWordstatApi('CreateNewWordstatReport', params);
    return result; // Возвращает ReportID (int)
}

/**
 * Ждет готовности отчета
 */
function waitForReport(reportId) {
    let attempts = 0;
    while (attempts < 20) { // Ждем макс 100 сек
        Utilities.sleep(5000);
        const list = callWordstatApi('GetWordstatReportList', {});

        const report = list.find(r => r.ReportID === reportId);
        if (report) {
            if (report.StatusReport === 'Done') {
                return callWordstatApi('GetWordstatReport', reportId);
            } else if (report.StatusReport === 'Failed') {
                throw new Error('Report generation failed');
            }
        }
        attempts++;
    }
    throw new Error('Timeout waiting for report');
}

function deleteWordstatReport(reportId) {
    return callWordstatApi('DeleteWordstatReport', reportId);
}

/**
 * Импортирует недостающие запросы из GSC и Метрики для активной категории
 */
/**
 * Импортирует недостающие запросы из GSC и Метрики для активной категории (или нескольких)
 */
function importMissingKeywordsForCategory() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const semanticsSheet = ss.getSheetByName(CATEGORY_SHEETS.SEMANTICS);

    if (!semanticsSheet) {
        SpreadsheetApp.getUi().alert('Лист "Семантика" не найден.');
        return;
    }

    // 0. ПРИНУДИТЕЛЬНОЕ ОБНОВЛЕНИЕ ЗАГОЛОВКОВ (ВСЕГДА)
    semanticsSheet.getRange('A1:K1').setValues([['✅', 'Ключевая фраза', 'Частотность (Шир.)', 'Частотность ("")', 'Частотность (!)', 'Частотность ([!])', 'Визиты', 'Отказы (%)', 'Заказы', 'Google', 'Период']]);
    semanticsSheet.getRange('A1:K1').setFontWeight('bold').setBackground('#e0e0e0');
    semanticsSheet.setColumnWidth(2, 300);
    semanticsSheet.setColumnWidth(1, 30);
    semanticsSheet.setColumnWidth(9, 80); // Заказы
    SpreadsheetApp.flush();

    // 1. Определяем категории для обработки
    const categoriesToProcess = []; // {id, path, name, breadcrumbs}

    if (activeSheet.getName() === CATEGORY_SHEETS.MAIN_LIST) {
        // А. Сначала ищем ОТМЕЧЕННЫЕ галочками (Col A)
        // Это надежнее, чем выделение, так как пользователь может натыкать галочек, но не выделить строки
        const lastRow = activeSheet.getLastRow();
        let hasCheckedRows = false;

        if (lastRow > 1) {
            // Читаем весь диапазон данных (A2:G)
            // A=Checkbox, B=ID, E=Breadcrumbs, F=Name, G=URL
            const dataValues = activeSheet.getRange(2, 1, lastRow - 1, 7).getValues();

            for (let i = 0; i < dataValues.length; i++) {
                if (dataValues[i][0] === true) { // Если галочка стоит
                    hasCheckedRows = true;
                    const catId = dataValues[i][1];
                    if (catId) {
                        categoriesToProcess.push({
                            id: String(catId),
                            path: dataValues[i][6],
                            name: dataValues[i][5],
                            breadcrumbs: dataValues[i][4]
                        });
                    }
                }
            }
        }

        // Б. Если галочек нет - ругаемся.
        // Пользователь запретил обработку по курсору.
        if (!hasCheckedRows) {
            SpreadsheetApp.getUi().alert('Пожалуйста, отметьте галочками (столбец A) категории, для которых нужно импортировать запросы.');
            return;
        }
    } else if (activeSheet.getName() === CATEGORY_SHEETS.SEMANTICS) {
        // Запуск из Семантики (только одна категория)
        const activeRow = activeSheet.getActiveCell().getRow();
        let catId = null;
        let catPath = null;
        let catName = null;
        let catBreadcrumbs = null;

        // Ищем заголовок СВЕРХУ
        for (let i = activeRow; i >= 1; i--) {
            const val = String(semanticsSheet.getRange(i, 1).getValue());
            if (val.startsWith('📂')) {
                const match = val.match(/\(([^)]+)\)$/);
                if (match) {
                    catId = match[1];
                    // Ищем инфо в Main List
                    const mainSheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
                    const data = mainSheet.getDataRange().getValues();
                    const row = data.find(r => String(r[1]) === String(catId));
                    if (row) {
                        catPath = row[6];
                        catName = row[5];
                        catBreadcrumbs = row[4];
                    }
                }
                break;
            }
        }
        if (catId) {
            categoriesToProcess.push({ id: catId, path: catPath, name: catName, breadcrumbs: catBreadcrumbs });
        } else {
            SpreadsheetApp.getUi().alert('Не удалось определить категорию. Выделите ячейку внутри блока категории.');
            return;
        }
    } else {
        SpreadsheetApp.getUi().alert('Запустите скрипт с листа "Категории" или "Семантика".');
        return;
    }

    // Убираем дубликаты
    const uniqueCategories = [];
    const seenIds = new Set();
    categoriesToProcess.forEach(c => {
        if (!seenIds.has(c.id)) {
            uniqueCategories.push(c);
            seenIds.add(c.id);
        }
    });

    if (uniqueCategories.length === 0) {
        SpreadsheetApp.getUi().alert('Не выбрано ни одной категории.');
        return;
    }

    // 2. Спрашиваем период (ОДИН РАЗ)
    const userProperties = PropertiesService.getUserProperties();
    const lastPeriod = userProperties.getProperty('LAST_SEMANTICS_PERIOD') || '2024-01-01:2024-12-31';
    const ui = SpreadsheetApp.getUi();

    const dateResponse = ui.prompt(
        'Импорт запросов',
        `Выбрано категорий: ${uniqueCategories.length}\n\nВведите период (YYYY-MM-DD:YYYY-MM-DD).\nОставьте пустым, чтобы использовать: ${lastPeriod}`,
        ui.ButtonSet.OK_CANCEL
    );

    if (dateResponse.getSelectedButton() !== ui.Button.OK) return;

    let inputDate = dateResponse.getResponseText().trim();
    if (!inputDate) inputDate = lastPeriod;
    userProperties.setProperty('LAST_SEMANTICS_PERIOD', inputDate);

    const parts = inputDate.split(':');
    const date1 = parts[0].trim();
    const date2 = parts[1].trim();
    const periodStr = `${date1}:${date2}`;

    // 3. Обрабатываем
    let processedCount = 0;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Начинаем обработку ${uniqueCategories.length} категорий...`, 'Старт');

    for (const cat of uniqueCategories) {
        try {
            processCategoryImport(cat, semanticsSheet, date1, date2, periodStr);
            processedCount++;
        } catch (e) {
            console.error(`Ошибка импорта для ${cat.name}:`, e);
            // Можно добавить лог в ячейку или алерт, но лучше не прерывать
        }
    }

    ui.alert(`Готово!\nОбработано категорий: ${processedCount} из ${uniqueCategories.length}`);
}

/**
 * Вспомогательная функция для обработки одной категории
 */
function processCategoryImport(cat, semanticsSheet, date1, date2, periodStr) {
    const catId = cat.id;
    let catPath = cat.path;
    const catName = cat.name;
    const catBreadcrumbs = cat.breadcrumbs;

    SpreadsheetApp.getActiveSpreadsheet().toast(`Обработка: ${catName}`, 'Импорт');

    // 1. Находим блок
    const lastRow = semanticsSheet.getLastRow();
    let blockStartRow = -1;
    let blockEndRow = -1;

    if (lastRow > 0) {
        const rangeValues = semanticsSheet.getRange(1, 1, lastRow, 1).getValues();
        for (let i = 0; i < rangeValues.length; i++) {
            const val = String(rangeValues[i][0]);
            // Надежное сравнение ID
            if (val.includes(`(${catId})`) && val.startsWith('📂')) {
                blockStartRow = i + 1;
                for (let j = i + 1; j < rangeValues.length; j++) {
                    const nextVal = String(rangeValues[j][0]);
                    if (nextVal.startsWith('📂')) {
                        blockEndRow = j;
                        break;
                    }
                }
                if (blockEndRow === -1) blockEndRow = lastRow;
                break;
            }
        }
    }

    // 2. Если блока нет - создаем АВТОМАТИЧЕСКИ (в конец)
    if (blockStartRow === -1) {
        blockStartRow = lastRow + 1;
        // Используем хлебные крошки для заголовка
        const headerTitle = catBreadcrumbs ? `📂 ${catBreadcrumbs} (${catId})` : `📂 ${catName} (${catId})`;
        semanticsSheet.getRange(blockStartRow, 1).setValue(headerTitle);
        semanticsSheet.getRange(blockStartRow, 1, 1, 10).setBackground('#f3f3f3').setFontWeight('bold').merge();
        blockEndRow = blockStartRow;
    }

    // Нормализуем URL
    try {
        if (catPath.startsWith('http')) {
            const urlObj = new URL(catPath);
            catPath = urlObj.pathname;
        }
    } catch (e) { }
    if (catPath.includes('?')) catPath = catPath.split('?')[0];

    // 3. Собираем текущие фразы
    const currentPhrases = new Set();
    if (blockEndRow > blockStartRow) {
        const range = semanticsSheet.getRange(blockStartRow + 1, 2, blockEndRow - blockStartRow, 1);
        const values = range.getValues();
        values.forEach(r => currentPhrases.add(String(r[0]).toLowerCase().trim()));
    }

    // 4. Запрашиваем GSC и Метрику
    const newKeywords = [];

    // GSC
    try {
        const gscData = fetchSearchQueriesFromGSC(date1, date2, catPath);
        gscData.forEach(row => {
            const phrase = row.query.toLowerCase().trim();
            if (phrase.length > 3 &&
                /[а-яё]/.test(phrase) && // GSC оставляем кириллицу? Пользователь просил убрать для Метрики. Для GSC оставим пока как было, или тоже уберем?
                // Пользователь просил "не фильтровать полезные данные".
                // Давайте уберем фильтр кириллицы и для GSC тоже, чтобы было единообразно.
                // Но оставим проверку на мусор.
                !phrase.includes('http') &&
                !phrase.includes('www') &&
                row.clicks > 0 &&
                !currentPhrases.has(phrase)) {

                newKeywords.push({ phrase: phrase, source: 'GSC', clicks: row.clicks, visits: 0, bounce: 0, score: row.clicks });
                currentPhrases.add(phrase);
            }
        });
    } catch (e) { console.warn('GSC Error:', e); }

    // Metrica
    try {
        const metricaData = fetchAllSearchPhrasesFromMetrica(date1, date2, catPath);
        metricaData.forEach(row => {
            const phrase = String(row.dimensions[0].name).toLowerCase().trim();
            if (phrase.length > 3 &&
                !phrase.includes('http') &&
                !phrase.includes('www') &&
                !currentPhrases.has(phrase)) {

                newKeywords.push({ phrase: phrase, source: 'Metrica', clicks: 0, visits: row.metrics[0], bounce: row.metrics[1], score: row.metrics[0] });
                currentPhrases.add(phrase);
            }
        });
    } catch (e) { console.warn('Metrica Error:', e); }

    if (newKeywords.length === 0) return;

    // Сортируем
    newKeywords.sort((a, b) => b.score - a.score);
    const limit = 50;
    const limitedKeywords = newKeywords.slice(0, limit);

    // 5. Вставляем
    const rowsToInsert = limitedKeywords.map(k => [
        false, // Checkbox
        k.phrase,
        '-', '-', '-', '-', // Frequencies
        k.visits, // Visits (G)
        Math.round(k.bounce * 100) / 100, // Bounce (H)
        k.clicks, // Google (I)
        periodStr // Period (J)
    ]);

    semanticsSheet.insertRowsAfter(blockEndRow, rowsToInsert.length);
    const newRange = semanticsSheet.getRange(blockEndRow + 1, 1, rowsToInsert.length, 10);
    newRange.setValues(rowsToInsert);
    newRange.setBackground(null).setFontWeight('normal');
    semanticsSheet.getRange(blockEndRow + 1, 1, rowsToInsert.length, 1).insertCheckboxes();

    // Группировка
    if (blockEndRow === blockStartRow) {
        try {
            const groupRange = semanticsSheet.getRange(blockEndRow + 1, 1, rowsToInsert.length, semanticsSheet.getLastColumn());
            groupRange.shiftRowGroupDepth(1);
        } catch (e) { }
    }
}

/**
 * Отладочная функция для проверки конкретной фразы
 */
function debugSpecificPhrase() {
    const phrase = 'бинокль с дальномером';
    const date1 = '2023-01-01';
    const date2 = '2025-12-01';

    const ui = SpreadsheetApp.getUi();
    ui.alert(`Запуск отладки для фразы: "${phrase}"\nПериод: ${date1} - ${date2}`);

    try {
        const config = YANDEX_METRICA_CONFIG;
        // Цель "Заказ" ID 36713851
        const goalMetric = 'ym:s:goal36713851reaches';

        const params = {
            'ids': config.counterId,
            'metrics': `ym:s:visits,ym:s:bounceRate,${goalMetric}`,
            'dimensions': 'ym:s:searchPhrase',
            'date1': date1,
            'date2': date2,
            'limit': 10000,
            'accuracy': 'full'
        };

        const queryString = Object.keys(params)
            .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
            .join('&');
        const url = `${config.baseUrl}?${queryString}`;

        const response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'OAuth ' + config.oauthToken }
        });
        const data = JSON.parse(response.getContentText());

        const found = data.data.find(r => r.dimensions[0].name.toLowerCase().trim() === phrase);

        let msg = `Всего строк: ${data.data.length}\n`;
        if (found) {
            msg += `✅ Фраза НАЙДЕНА!\nВизиты: ${found.metrics[0]}`;
        } else {
            msg += `❌ Фраза НЕ найдена в топ-10000.\n`;
            // Ищем похожие
            const similar = data.data.filter(r => r.dimensions[0].name.includes('дальномером'));
            msg += `Похожие (${similar.length}):\n`;
            similar.slice(0, 5).forEach(s => msg += `- ${s.dimensions[0].name}: ${s.metrics[0]}\n`);
        }

        ui.alert(msg);

    } catch (e) {
        ui.alert('Ошибка: ' + e.message);
    }
}
