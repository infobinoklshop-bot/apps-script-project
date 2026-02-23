/**
 * DEBUG: Проверка данных по цели "Заказ" (40019707)
 */
function debugMetricaGoal() {
    const config = YANDEX_METRICA_CONFIG;
    const goalId = '40019707';
    const urlPath = '/collection/morskie'; // URL из скриншота, где должно быть 14 заказов
    const date1 = '2023-01-01';
    const date2 = '2025-12-01';

    console.log(`DEBUG: Проверка цели ${goalId} для URL ${urlPath}`);

    const params = {
        'ids': config.counterId,
        'metrics': `ym:s:goal${goalId}reaches`,
        'dimensions': 'ym:s:startURLPath',
        'filters': `ym:s:startURLPath=='${urlPath}'`,
        'date1': date1,
        'date2': date2,
        'accuracy': 'full',
        'limit': 100
    };

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

    console.log(`Request URL: ${url}`);

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log(`Response Code: ${responseCode}`);
    console.log(`Response Body: ${responseText}`);

    if (responseCode === 200) {
        const data = JSON.parse(responseText);
        if (data.data && data.data.length > 0) {
            console.log(`✅ Данные найдены: ${data.data[0].metrics[0]} достижений цели`);
        } else {
            console.log('❌ Данные не найдены (пустой массив data)');
        }
    } else {
        console.error('❌ Ошибка API');
    }
}
