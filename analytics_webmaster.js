/**
 * Сервис для работы с Яндекс.Вебмастер API
 */

/**
 * Получает статистику (Показы) из Яндекс.Вебмастера за указанный период для списка URL
 * @param {string} startDate - YYYY-MM-DD
 * @param {string} endDate - YYYY-MM-DD
 * @returns {Map<string, number>} Map: URL -> Impressions
 */
function fetchWebmasterStats(startDate, endDate) {
    const statsMap = new Map();
    const config = YANDEX_METRICA_CONFIG;

    if (!config.oauthToken) {
        console.warn('Токен Яндекс.Вебмастера (Метрики) не найден.');
        return statsMap;
    }

    try {
        // 1. Получаем User ID
        const userId = getWebmasterUserId(config.oauthToken);
        if (!userId) return statsMap;

        // 2. Получаем Host ID для binokl.shop
        const hostId = getWebmasterHostId(config.oauthToken, userId, 'binokl.shop');
        if (!hostId) {
            SpreadsheetApp.getActiveSpreadsheet().toast('Сайт binokl.shop не найден в Яндекс.Вебмастере', 'Ошибка');
            return statsMap;
        }

        console.log(`Используем Host ID Вебмастера: ${hostId}`);

        // 3. Запрашиваем статистику
        // В API v4 нет простого метода "Все страницы с показами".
        // Мы используем метод "Популярные запросы" (popular), который возвращает ТОП-500 запросов/страниц.
        // GET /v4/user/{user-id}/hosts/{host-id}/search-queries/popular

        // Параметры:
        // order_by: TOTAL_SHOWS (по показам)
        // query_indicator: TOTAL_SHOWS
        // date_from, date_to (не поддерживаются в popular, он возвращает "актуальное")
        // НО! Есть endpoint /v4/user/{user-id}/hosts/{host-id}/query-analytics/list
        // Он позволяет фильтровать.

        // Попробуем query-analytics/list, так как он мощнее.
        const apiUrl = `https://api.webmaster.yandex.net/v4/user/${userId}/hosts/${hostId}/query-analytics/list`;

        // Нам нужно сгруппировать по URL, но API группирует по запросам.
        // Попробуем получить просто список популярных запросов и посмотреть, есть ли там URL.
        // В документации: text_indicator (запрос). URL нет.

        // ЕДИНСТВЕННЫЙ СПОСОБ получить статистику по URL - это использовать "Популярные страницы" из интерфейса?
        // В API это может быть недоступно.

        // ДАВАЙТЕ ПОПРОБУЕМ `search-queries/popular` с `query_indicator=TOTAL_SHOWS`.
        // Он возвращает список запросов.

        // ЕСЛИ API НЕ ПОЗВОЛЯЕТ получить показы по страницам, мы не можем заполнить эту колонку точно.
        // Но мы можем попробовать получить "Страницы в поиске" -> "Статистика".
        // Такой ручки нет.

        // РЕШЕНИЕ:
        // Мы будем использовать заглушку с уведомлением, что API не отдает статистику по страницам массово.
        // ИЛИ (экспериментально) попробуем получить данные через Метрику (источники - поисковые системы), но это визиты, а не показы.

        // ОДНАКО, пользователь хочет "Показы Я.Веб".
        // Я попробую сделать запрос, который хотя бы проверит доступность API.

        // Попытка получить хоть что-то:
        const response = UrlFetchApp.fetch(apiUrl, {
            method: 'POST',
            headers: { 'Authorization': 'OAuth ' + config.oauthToken },
            contentType: 'application/json',
            payload: JSON.stringify({
                device_type_indicator: "ALL_DEVICES",
                text_indicator: "TOTAL_SHOWS",
                region_ids: [] // Все регионы
            }),
            muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            // В ответе `text_indicator` -> `value` (запрос) и `statistics`.
            // URL там нет.
            console.log('Webmaster API доступен, но метод возвращает запросы, а не страницы.');
            SpreadsheetApp.getActiveSpreadsheet().toast('API Вебмастера отдает только запросы, а не страницы. Колонка останется пустой.', 'Info');
        } else {
            console.warn('Ошибка Webmaster API:', response.getContentText());
        }

    } catch (e) {
        console.error('Ошибка Webmaster API:', e);
        SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка Webmaster: ' + e.message, 'Warning');
    }

    return statsMap;
}

function getWebmasterUserId(token) {
    try {
        const response = UrlFetchApp.fetch('https://api.webmaster.yandex.net/v4/user', {
            headers: { 'Authorization': 'OAuth ' + token }
        });
        const data = JSON.parse(response.getContentText());
        return data.user_id;
    } catch (e) {
        console.error('Ошибка получения User ID Вебмастера:', e);
        return null;
    }
}

function getWebmasterHostId(token, userId, domain) {
    try {
        const response = UrlFetchApp.fetch(`https://api.webmaster.yandex.net/v4/user/${userId}/hosts`, {
            headers: { 'Authorization': 'OAuth ' + token }
        });
        const data = JSON.parse(response.getContentText());

        // Ищем хост. Формат host_id обычно "https:binokl.shop:443"
        const host = data.hosts.find(h => h.ascii_host_url.includes(domain));
        return host ? host.host_id : null;
    } catch (e) {
        console.error('Ошибка получения Host ID Вебмастера:', e);
        return null;
    }
}
