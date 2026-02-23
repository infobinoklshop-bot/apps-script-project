/**
 * DEBUG: Проверка доступных целей и eCommerce
 */
function debugCheckGoals() {
    const config = YANDEX_METRICA_CONFIG;

    console.log('--- DEBUG: GOALS & ECOMMERCE ---');

    // 1. Попытка получить список всех целей (Management API)
    try {
        const goalsUrl = `https://api-metrika.yandex.net/management/v1/counter/${config.counterId}/goals`;
        const response = UrlFetchApp.fetch(goalsUrl, {
            'headers': { 'Authorization': `OAuth ${config.oauthToken}` },
            'muteHttpExceptions': true
        });

        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            console.log(`✅ Получен список целей (${data.goals.length}):`);
            data.goals.forEach(g => {
                console.log(`- [${g.id}] ${g.name} (Type: ${g.type})`);
            });
        } else {
            console.log(`⚠️ Не удалось получить список целей (Код ${response.getResponseCode()}). Возможно, нет прав на Management API.`);
        }
    } catch (e) {
        console.error('Ошибка при запросе списка целей:', e);
    }

    // 2. Проверка eCommerce метрик и старой цели
    console.log('\n--- CHECKING METRICS ---');
    const params = {
        'ids': config.counterId,
        'metrics': 'ym:s:visits,ym:s:ecommercePurchases,ym:s:goal36713851reaches,ym:s:goal40019707reaches',
        'date1': '2023-01-01',
        'date2': '2025-12-01',
        'accuracy': 'full',
        'limit': 1
    };

    const url = `${config.baseUrl}?${Object.keys(params).map(k => k + '=' + params[k]).join('&')}`;

    try {
        const response = UrlFetchApp.fetch(url, {
            'headers': { 'Authorization': `OAuth ${config.oauthToken}` }
        });
        const data = JSON.parse(response.getContentText());

        if (data.data && data.data.length > 0) {
            const r = data.data[0];
            console.log(`Visits: ${r.metrics[0]}`);
            console.log(`eCommerce Purchases: ${r.metrics[1]}`);
            console.log(`Old Goal (36713851): ${r.metrics[2]}`);
            console.log(`Current Goal (40019707): ${r.metrics[3]}`);

            let msg = `eCom: ${r.metrics[1]}, Old: ${r.metrics[2]}, Cur: ${r.metrics[3]}`;
            SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'DEBUG METRICS');
        } else {
            console.log('No data returned');
        }
    } catch (e) {
        console.error('API Error:', e);
    }
}
