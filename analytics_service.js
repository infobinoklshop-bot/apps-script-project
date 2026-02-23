/**
 * ========================================
 * СЕРВИС АНАЛИТИКИ (YANDEX METRICA)
 * ========================================
 */

/**
 * Получает статистику визитов для всех категорий из Яндекс.Метрики
 * и обновляет колонку "Визиты" в таблице категорий.
 */
function updateCategoryVisitsStats() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    if (!sheet) {
        console.error('Лист категорий не найден');
        return;
    }

    // 0. Обновляем структуру таблицы, чтобы колонки соответствовали конфигу
    updateTableStructureSafe();

    const config = YANDEX_METRICA_CONFIG;
    if (!config.oauthToken || config.oauthToken === 'YOUR_YANDEX_OAUTH_TOKEN') {
        SpreadsheetApp.getUi().alert('Ошибка: Не настроен OAuth токен Яндекс.Метрики в 01_config.gs');
        return;
    }

    // 1. Получаем сохраненный период или вычисляем дефолтный (последние 30 дней)
    const userProperties = PropertiesService.getUserProperties();
    const lastPeriod = userProperties.getProperty('LAST_METRICA_PERIOD');

    let defaultString = lastPeriod;
    if (!defaultString) {
        // Жестко заданный дефолт по просьбе пользователя
        defaultString = '2023-01-01:2025-12-01';
    }

    // 2. Запрашиваем период у пользователя с предзаполнением
    const ui = SpreadsheetApp.getUi();
    const dateResponse = ui.prompt(
        'Период статистики',
        `Введите период (YYYY-MM-DD:YYYY-MM-DD).\nОставьте пустым, чтобы использовать: ${defaultString}`,
        ui.ButtonSet.OK_CANCEL
    );

    // К сожалению, ui.prompt не поддерживает предзаполнение текста в стандартном диалоге Apps Script.
    // Мы можем только показать дефолт в сообщении.
    // Но мы можем использовать HTML Service для кастомного диалога, если пользователь захочет.
    // Пока попробуем простой хак: если пользователь оставил пустым -> берем то, что было в прошлый раз (или дефолт).

    if (dateResponse.getSelectedButton() !== ui.Button.OK) {
        return; // Пользователь отменил
    }

    let startDate, endDate;
    let inputDate = dateResponse.getResponseText().trim();

    // Если пусто - берем дефолт (последний сохраненный или 30 дней)
    if (!inputDate) {
        inputDate = defaultString;
    }

    if (inputDate) {
        const parts = inputDate.split(':');
        if (parts.length === 2) {
            startDate = parts[0].trim();
            endDate = parts[1].trim();
            // Сохраняем выбор пользователя
            userProperties.setProperty('LAST_METRICA_PERIOD', inputDate);
        } else {
            ui.alert('Неверный формат даты. Используйте YYYY-MM-DD:YYYY-MM-DD');
            return;
        }
    }

    console.log(`Запрос статистики за период: ${startDate} - ${endDate}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Загрузка статистики за ${startDate} - ${endDate}...`, 'Яндекс.Метрика');

    // Получаем данные категорий
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(2, 1, lastRow - 1, Object.keys(MAIN_LIST_COLUMNS).length);
    const values = range.getValues();

    // Индексы колонок (в массиве 0-based)
    const urlColIndex = MAIN_LIST_COLUMNS.URL - 1;

    const updates = [];

    try {
        const metricaData = fetchTopPagesFromMetrica(startDate, endDate);

        // 4. Запрашиваем данные из GSC
        let gscStats = new Map();
        if (MAIN_LIST_COLUMNS.GSC_IMPRESSIONS) {
            SpreadsheetApp.getActiveSpreadsheet().toast('Загрузка данных из Google Search Console...', 'Загрузка');
            gscStats = fetchGSCStats(startDate, endDate);
        }

        // 5. Запрашиваем данные из Яндекс.Вебмастера
        let webmasterStats = new Map();
        if (MAIN_LIST_COLUMNS.WEBMASTER_IMPRESSIONS) {
            SpreadsheetApp.getActiveSpreadsheet().toast('Загрузка данных из Яндекс.Вебмастера...', 'Загрузка');
            webmasterStats = fetchWebmasterStats(startDate, endDate);
        }

        // 6. Обработка данных Метрики (как раньше)
        const statsMap = new Map();

        metricaData.forEach(row => {
            let url = row.dimensions[0].name;
            const visits = row.metrics[0];
            const bounceRate = row.metrics[1];
            const orders = row.metrics[2] || 0; // Заказы

            // НОРМАЛИЗАЦИЯ URL
            try {
                if (url.startsWith('http')) {
                    const urlObj = new URL(url);
                    url = urlObj.pathname;
                }
            } catch (e) {
                url = url.replace(/^https?:\/\/[^\/]+/, '');
            }
            if (url.includes('?')) url = url.split('?')[0];
            if (url.endsWith('/') && url.length > 1) url = url.slice(0, -1);

            if (statsMap.has(url)) {
                const current = statsMap.get(url);
                const totalVisits = current.visits + visits;
                const weightedBounce = (current.visits * current.bounceRate + visits * bounceRate) / totalVisits;
                const totalOrders = (current.orders || 0) + orders;
                statsMap.set(url, { visits: totalVisits, bounceRate: weightedBounce, orders: totalOrders });
            } else {
                statsMap.set(url, { visits, bounceRate, orders });
            }
        });

        console.log(`Загружено ${metricaData.length} строк. Уникальных URL: ${statsMap.size}`);

        // 7. Обновляем данные в массиве
        const visitUpdates = [];
        const bounceUpdates = [];
        const orderUpdates = []; // Массив для заказов
        const gscUpdates = [];
        const webmasterUpdates = [];
        const periodUpdates = [];
        const periodString = `${startDate} - ${endDate}`;

        for (let i = 0; i < values.length; i++) {
            let categoryUrl = values[i][urlColIndex];

            if (!categoryUrl) {
                visitUpdates.push([0]);
                bounceUpdates.push(['-']);
                gscUpdates.push([0]);
                webmasterUpdates.push([0]);
                periodUpdates.push([periodString]);
                continue;
            }

            if (categoryUrl.endsWith('/') && categoryUrl.length > 1) categoryUrl = categoryUrl.slice(0, -1);

            // Метрика
            const stats = statsMap.get(categoryUrl);
            if (stats) {
                visitUpdates.push([stats.visits]);
                bounceUpdates.push([stats.bounceRate.toFixed(1) + '%']);
                orderUpdates.push([stats.orders || 0]);
            } else {
                visitUpdates.push([0]);
                bounceUpdates.push(['-']);
                orderUpdates.push([0]);
            }

            // GSC
            const gscImpressions = gscStats.get(categoryUrl) || 0;
            gscUpdates.push([gscImpressions]);

            // Webmaster
            const webmasterImpressions = webmasterStats.get(categoryUrl) || 0;
            webmasterUpdates.push([webmasterImpressions]);

            periodUpdates.push([periodString]);
        }

        // 8. Записываем обновления
        if (visitUpdates.length > 0) {
            sheet.getRange(2, MAIN_LIST_COLUMNS.VISITS, visitUpdates.length, 1).setValues(visitUpdates);

            if (MAIN_LIST_COLUMNS.BOUNCE_RATE) {
                sheet.getRange(2, MAIN_LIST_COLUMNS.BOUNCE_RATE, bounceUpdates.length, 1).setValues(bounceUpdates);
            }

            if (MAIN_LIST_COLUMNS.ORDERS) {
                sheet.getRange(2, MAIN_LIST_COLUMNS.ORDERS, orderUpdates.length, 1).setValues(orderUpdates);
            }

            if (MAIN_LIST_COLUMNS.GSC_IMPRESSIONS) {
                sheet.getRange(2, MAIN_LIST_COLUMNS.GSC_IMPRESSIONS, gscUpdates.length, 1).setValues(gscUpdates);
            }

            if (MAIN_LIST_COLUMNS.WEBMASTER_IMPRESSIONS) {
                sheet.getRange(2, MAIN_LIST_COLUMNS.WEBMASTER_IMPRESSIONS, webmasterUpdates.length, 1).setValues(webmasterUpdates);
            }

            if (MAIN_LIST_COLUMNS.PERIOD) {
                sheet.getRange(2, MAIN_LIST_COLUMNS.PERIOD, periodUpdates.length, 1).setValues(periodUpdates);
            }

            SpreadsheetApp.getActiveSpreadsheet().toast(`Обновлено ${visitUpdates.length} строк.`, 'Успешно');
        }
    } catch (e) {
        console.error('Ошибка при обновлении статистики:', e);

        let errorMessage = e.message;
        if (errorMessage.includes('403') || errorMessage.includes('Access is denied')) {
            errorMessage = 'Ошибка доступа (403). Проверьте права токена.';
        } else if (errorMessage.includes('400')) {
            errorMessage = 'Ошибка запроса (400). Возможно, неверные параметры даты или атрибуции.';
        }

        SpreadsheetApp.getUi().alert('Ошибка: ' + errorMessage);
    } finally {
        // Явно показываем, что все готово, чтобы перебить предыдущие тосты
        SpreadsheetApp.getActiveSpreadsheet().toast('Готово!', 'Status', 3);
        SpreadsheetApp.flush(); // Принудительно обновляем UI
    }
}

/**
 * Получает ТОП популярных страниц за период
 */
function fetchTopPagesFromMetrica(date1, date2) {
    const config = YANDEX_METRICA_CONFIG;
    const limit = 10000; // Максимум строк, чтобы охватить все категории

    // Параметры запроса
    const params = {
        'ids': config.counterId,
        'metrics': 'ym:s:visits,ym:s:bounceRate,ym:s:goal54410371reaches', // Заказы (54410371)
        'dimensions': 'ym:s:startURL', // Группировка по URL входа (или просто URL)
        'date1': date1,
        'date2': date2,
        'limit': limit,
        'accuracy': 'full', // Точные данные, не сэмплирование
        'proposed_accuracy': 'false',
        'attribution': 'lastsign' // Атрибуция: Последний значимый переход (как в отчете)
    };

    // Формируем Query String
    const queryString = Object.keys(params)
        .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
        .join('&');

    const url = `${config.baseUrl}?${queryString}`;

    const options = {
        'method': 'get',
        'headers': {
            'Authorization': `OAuth ${config.oauthToken}`
        },
        'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
        throw new Error(`API Error (${responseCode}): ${responseText}`);
    }

    const data = JSON.parse(responseText);

    // DEBUG: Логируем топ фраз и ищем проблемную
    console.log(`Метрика: получено ${data.data.length} строк. Всего: ${data.total_rows}`);
    const debugPhrase = 'бинокль с дальномером';
    const found = data.data.find(r => r.dimensions[0].name.toLowerCase().trim() === debugPhrase);
    if (found) {
        console.log(`DEBUG: Найдена фраза "${debugPhrase}": ${found.metrics[0]} визитов`);
    } else {
        console.log(`DEBUG: Фраза "${debugPhrase}" НЕ найдена в топ-${limit}!`);
        // Логируем похожие
        const similar = data.data.filter(r => r.dimensions[0].name.toLowerCase().includes('дальномером') && r.dimensions[0].name.toLowerCase().includes('бинокль'));
        console.log(`DEBUG: Похожие фразы (${similar.length}):`);
        similar.slice(0, 5).forEach(s => console.log(`- ${s.dimensions[0].name}: ${s.metrics[0]}`));
    }

    return data.data; // Массив строк отчета
}

/**
 * Получает статистику по поисковым фразам из Метрики
 * @param {string} date1 - Начальная дата
 * @param {string} date2 - Конечная дата
 * @param {string} [urlPath] - Опционально: фильтр по URL посадочной страницы
 */
function fetchAllSearchPhrasesFromMetrica(date1, date2, urlPath) {
    const config = YANDEX_METRICA_CONFIG;
    const limit = 10000;

    const params = {
        'ids': config.counterId,
        'metrics': 'ym:s:visits,ym:s:bounceRate,ym:s:goal54410371reaches', // Заказы (54410371)
        'dimensions': 'ym:s:searchPhrase',
        'date1': date1,
        'date2': date2,
        'limit': limit,
        'accuracy': 'full',
        'proposed_accuracy': 'false'
    };

    if (urlPath) {
        // Фильтр по URL входа
        params['filters'] = `ym:s:startURLPath=@'${urlPath}'`;
    }

    const queryString = Object.keys(params)
        .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
        .join('&');

    const url = `${config.baseUrl}?${queryString}`;

    const options = {
        'method': 'get',
        'headers': {
            'Authorization': `OAuth ${config.oauthToken}`
        },
        'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
        throw new Error(`API Error (${response.getResponseCode()}): ${response.getContentText()}`);
    }

    return JSON.parse(response.getContentText()).data;
}

/**
 * Форматирует дату в YYYY-MM-DD
 */
function formatDateForYandex(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

/**
 * Тестовая функция для проверки подключения
 */
function testYandexConnection() {
    try {
        const config = YANDEX_METRICA_CONFIG;
        if (!config.oauthToken || config.oauthToken === 'YOUR_YANDEX_OAUTH_TOKEN') {
            return 'Токен не настроен';
        }

        // Простой запрос - общая статистика за сегодня
        const url = `${config.baseUrl}?ids=${config.counterId}&metrics=ym:s:visits&date1=today&date2=today`;
        const options = {
            'headers': { 'Authorization': `OAuth ${config.oauthToken}` }
        };

        const response = UrlFetchApp.fetch(url, options);
        if (response.getResponseCode() === 200) {
            return 'OK';
        } else {
            return 'Ошибка: ' + response.getResponseCode();
        }
    } catch (e) {
        return 'Ошибка: ' + e.message;
    }
}

/**
 * DEBUG: Проверка статистики для конкретного URL
 */
function debugUrlStats() {
    debugCheckGoals(); // 1. Сначала выводим список всех целей

    const urlPath = '/collection/morskie';
    const date1 = '2023-01-01';
    const date2 = '2025-12-01';

    console.log(`DEBUG: Проверка URL: ${urlPath}`);

    const config = YANDEX_METRICA_CONFIG;

    // 2. Проверка URL (Last Click)
    console.log(`--- URL CHECK: ${urlPath} ---`);
    const params = {
        'ids': config.counterId,
        'metrics': 'ym:s:visits,ym:s:goal40019707reaches',
        'dimensions': 'ym:s:startURLPath',
        'filters': `ym:s:startURLPath=@'${urlPath}'`,
        'date1': date1,
        'date2': date2,
        'accuracy': 'full',
        'attribution': 'last'
    };

    const queryString = Object.keys(params)
        .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
        .join('&');

    const url = `${config.baseUrl}?${queryString}`;

    try {
        const response = UrlFetchApp.fetch(url, {
            'headers': { 'Authorization': `OAuth ${config.oauthToken}` }
        });
        const data = JSON.parse(response.getContentText());

        // console.log('API Response:', JSON.stringify(data, null, 2)); // Убрали спам

        if (data.data && data.data.length > 0) {
            const row = data.data[0];
            console.log(`✅ URL (Last): ${row.metrics[0]} визитов, ${row.metrics[1]} заказов (40019707)`);
        } else {
            console.log('❌ URL не найден в API (0 визитов)');
        }
    } catch (e) {
        console.error(e);
    }

    // 4. GOAL HUNTER (Проверяем ВСЕ цели для этого URL)
    console.log(`--- GOAL HUNTER: ${urlPath} ---`);

    // ID всех целей из лога
    const goalIds = [
        '40019632', '54410371', '40019707', '54409678',
        '36736793', '36713854', '36713851', '36046234',
        '35866152', '35866182', '35866185', '194386348'
    ];

    // Формируем строку метрик: ym:s:goal<ID>reaches
    const metricsList = goalIds.map(id => `ym:s:goal${id}reaches`).join(',');

    const hunterParams = {
        'ids': config.counterId,
        'metrics': metricsList,
        'dimensions': 'ym:s:startURLPath',
        'filters': `ym:s:startURLPath=@'${urlPath}'`,
        'date1': date1,
        'date2': date2,
        'accuracy': 'full'
    };

    const hunterUrl = `${config.baseUrl}?${Object.keys(hunterParams).map(k => k + '=' + hunterParams[k]).join('&')}`;

    try {
        const response = UrlFetchApp.fetch(hunterUrl, { 'headers': { 'Authorization': `OAuth ${config.oauthToken}` } });
        const data = JSON.parse(response.getContentText());

        if (data.data && data.data.length > 0) {
            const row = data.data[0];
            console.log('Результаты по целям:');
            let found = false;
            goalIds.forEach((id, index) => {
                const val = row.metrics[index];
                console.log(`- Goal ${id}: ${val}`);
                if (val > 0) {
                    found = true;
                    SpreadsheetApp.getActiveSpreadsheet().toast(`Найдена цель ${id}: ${val} заказов!`, 'EUREKA');
                }
            });
            if (!found) {
                SpreadsheetApp.getActiveSpreadsheet().toast('Ни одна цель не дала > 0', 'FAIL');
            }
        } else {
            console.log('Hunter: Данных нет');
        }
    } catch (e) {
        console.error('Hunter Error:', e);
    }
}

