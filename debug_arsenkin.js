/**
 * Скрипт для отладки интеграции с Arsenkin
 * Используется для проверки статуса задач и получения сырых ответов API
 */

/**
 * Проверяет статус задачи и выводит сырой JSON ответ
 * @param {string} taskId - ID задачи (например '29325702')
 */
function debugRawCheck(taskId) {
    if (!taskId) {
        taskId = '29325702'; // ID из скриншота по умолчанию
    }

    console.log(`🔍 DEBUG: Checking task ${taskId}...`);

    try {
        // 1. Проверяем конфигурацию
        if (!ARSENKIN_CONFIG || !ARSENKIN_CONFIG.API_TOKEN) {
            console.error('❌ Config Error: ARSENKIN_CONFIG.API_TOKEN is missing');
            return;
        }

        const url = ARSENKIN_CONFIG.BASE_URL + 'check';
        const payload = { task_id: taskId };

        console.log(`URL: ${url}`);
        console.log(`Payload: ${JSON.stringify(payload)}`);

        const options = {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + ARSENKIN_CONFIG.API_TOKEN,
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        console.log(`Response Code: ${responseCode}`);
        console.log(`Response Body (RAW):`);
        console.log(responseText);

        if (responseCode === 200) {
            const json = JSON.parse(responseText);
            console.log(`\nParsed Status: ${json.status}`);
            console.log(`Parsed Progress: ${json.progress}`);
        }

    } catch (e) {
        console.error(`❌ EXCEPTION: ${e.message}`);
        console.error(e.stack);
    }
}

/**
 * Пытается получить результат задачи (если она завершена)
 */
function debugRawGet(taskId) {
    if (!taskId) {
        taskId = '29325702';
    }

    console.log(`🔍 DEBUG: Getting result for task ${taskId}...`);

    try {
        const url = ARSENKIN_CONFIG.BASE_URL + 'get';
        const payload = { task_id: taskId };

        const options = {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + ARSENKIN_CONFIG.API_TOKEN,
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(url, options);
        console.log(`Response Code: ${response.getResponseCode()}`);
        console.log(`Response Body (RAW - first 1000 chars):`);
        console.log(response.getContentText().substring(0, 1000) + '...');

    } catch (e) {
        console.error(`❌ EXCEPTION: ${e.message}`);
    }
}
