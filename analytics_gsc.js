/**
 * Сервис для работы с Google Search Console API (через REST API)
 */

/**
 * Получает статистику (Показы) из GSC за указанный период для списка URL
 * @param {string} startDate - YYYY-MM-DD
 * @param {string} endDate - YYYY-MM-DD
 * @returns {Map<string, number>} Map: URL -> Impressions
 */
function fetchGSCStats(startDate, endDate) {
    const statsMap = new Map();

    try {
        const token = ScriptApp.getOAuthToken();

        // 1. Определяем сайт (Resource URL)
        const sitesUrl = 'https://www.googleapis.com/webmasters/v3/sites';
        const sitesResponse = UrlFetchApp.fetch(sitesUrl, {
            headers: { 'Authorization': 'Bearer ' + token }
        });

        const sitesData = JSON.parse(sitesResponse.getContentText());
        let siteUrl = null;
        const targetDomain = 'binokl.shop';

        if (sitesData.siteEntry) {
            const sitesList = sitesData.siteEntry.map(s => s.siteUrl).join(', ');
            console.log(`Доступные сайты в GSC: ${sitesList}`);

            // Приоритет: URL-prefix (https://...), так как с sc-domain часто проблемы с правами через API
            const urlPrefixSite = sitesData.siteEntry.find(s => s.siteUrl.includes(targetDomain) && s.siteUrl.startsWith('http'));
            const domainSite = sitesData.siteEntry.find(s => s.siteUrl.includes(targetDomain) && s.siteUrl.startsWith('sc-domain:'));

            if (urlPrefixSite) {
                siteUrl = urlPrefixSite.siteUrl;
                console.log(`Выбран URL-prefix ресурс: ${siteUrl}`);
            } else if (domainSite) {
                siteUrl = domainSite.siteUrl;
                console.log(`Выбран Domain ресурс: ${siteUrl}`);
            }
        }

        if (!siteUrl) {
            console.warn(`Сайт ${targetDomain} не найден в GSC. Проверьте доступы.`);
            SpreadsheetApp.getActiveSpreadsheet().toast(`Сайт ${targetDomain} не найден в GSC!`, 'Ошибка');
            return statsMap;
        }

        console.log(`Используем ресурс GSC: ${siteUrl}`);

        // 2. Делаем запрос к API (Search Analytics Query)
        // POST https://www.googleapis.com/webmasters/v3/sites/siteUrl/searchAnalytics/query

        const queryUrl = `https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(siteUrl)}/searchAnalytics/query`;

        const payload = {
            startDate: startDate,
            endDate: endDate,
            dimensions: ['page'],
            rowLimit: 25000,
            startRow: 0
        };

        const response = UrlFetchApp.fetch(queryUrl, {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            throw new Error(`GSC API Error (${response.getResponseCode()}): ${response.getContentText()}`);
        }

        const data = JSON.parse(response.getContentText());

        if (data.rows) {
            console.log(`GSC вернул ${data.rows.length} строк.`);
            // Логируем первые 3 URL для проверки
            data.rows.slice(0, 3).forEach(r => console.log(`Пример GSC URL: ${r.keys[0]} -> ${r.impressions}`));

            data.rows.forEach(row => {
                let url = row.keys[0]; // Полный URL
                const impressions = row.impressions;

                // Нормализация URL
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

                // Суммируем показы
                if (statsMap.has(url)) {
                    statsMap.set(url, statsMap.get(url) + impressions);
                } else {
                    statsMap.set(url, impressions);
                }
            });
        } else {
            console.warn('GSC вернул пустой список строк (data.rows is undefined).');
        }

        console.log(`После обработки: ${statsMap.size} уникальных страниц из GSC.`);

    } catch (e) {
        console.error('Ошибка GSC API:', e);
        SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка GSC: ' + e.message, 'Warning');
    }

    return statsMap;
}

/**
 * Получает статистику по поисковым запросам из GSC за период
 * @param {string} startDate - YYYY-MM-DD
 * @param {string} endDate - YYYY-MM-DD
 * @param {string} [pageUrl] - Опционально: фильтр по URL страницы (должен быть полным или относительным, скрипт попытается подобрать)
 * @returns {Array<{query: string, clicks: number, impressions: number}>}
 */
function fetchSearchQueriesFromGSC(startDate, endDate, pageUrl) {
    const queries = [];

    try {
        const token = ScriptApp.getOAuthToken();
        const targetDomain = 'binokl.shop';

        // 1. Определяем сайт (копипаст логики выбора сайта)
        const sitesUrl = 'https://www.googleapis.com/webmasters/v3/sites';
        const sitesResponse = UrlFetchApp.fetch(sitesUrl, {
            headers: { 'Authorization': 'Bearer ' + token }
        });
        const sitesData = JSON.parse(sitesResponse.getContentText());
        let siteUrl = null;

        if (sitesData.siteEntry) {
            const urlPrefixSite = sitesData.siteEntry.find(s => s.siteUrl.includes(targetDomain) && s.siteUrl.startsWith('http'));
            const domainSite = sitesData.siteEntry.find(s => s.siteUrl.includes(targetDomain) && s.siteUrl.startsWith('sc-domain:'));
            siteUrl = urlPrefixSite ? urlPrefixSite.siteUrl : (domainSite ? domainSite.siteUrl : null);
        }

        if (!siteUrl) return queries;

        // 2. Запрос к API
        const queryUrl = `https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(siteUrl)}/searchAnalytics/query`;

        // Запрашиваем query
        const payload = {
            startDate: startDate,
            endDate: endDate,
            dimensions: ['query'],
            rowLimit: 25000,
            startRow: 0
        };

        // Добавляем фильтр по странице, если передан
        if (pageUrl) {
            // GSC требует полный URL для фильтрации по странице, если ресурс URL-prefix
            // Если ресурс Domain, то тоже полный URL работает лучше.
            let fullPageUrl = pageUrl;
            if (!pageUrl.startsWith('http')) {
                fullPageUrl = `https://${targetDomain}${pageUrl.startsWith('/') ? '' : '/'}${pageUrl}`;
            }

            payload.dimensionFilterGroups = [{
                filters: [{
                    dimension: 'page',
                    operator: 'equals', // СТРОГОЕ соответствие URL
                    expression: fullPageUrl
                }]
            }];
        }

        const response = UrlFetchApp.fetch(queryUrl, {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            throw new Error(`GSC API Error: ${response.getContentText()}`);
        }

        const data = JSON.parse(response.getContentText());
        if (data.rows) {
            data.rows.forEach(row => {
                queries.push({
                    query: row.keys[0],
                    clicks: row.clicks,
                    impressions: row.impressions
                });
            });
        }

    } catch (e) {
        console.error('Ошибка GSC Queries:', e);
        SpreadsheetApp.getActiveSpreadsheet().toast('Ошибка GSC Queries: ' + e.message, 'Warning');
    }

    return queries;
}
