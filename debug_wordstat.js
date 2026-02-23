function debugWordstatConnection() {
    const ui = SpreadsheetApp.getUi();
    const config = YANDEX_METRICA_CONFIG;

    if (!config.oauthToken) {
        ui.alert('❌ Ошибка', 'Токен не найден в 01_config.js', ui.ButtonSet.OK);
        return;
    }

    let login = 'Не определен';
    let sandboxStatus = 'Не проверялся';

    try {
        // 1. Пробуем получить информацию о владельце токена через Passport API
        // Это работает, даже если API Директа не включен
        let login = 'Не удалось определить';
        try {
            const passportUrl = 'https://login.yandex.ru/info?format=json';
            const passportResponse = UrlFetchApp.fetch(passportUrl, {
                headers: { 'Authorization': 'OAuth ' + config.oauthToken },
                muteHttpExceptions: true
            });
            const passportData = JSON.parse(passportResponse.getContentText());
            if (passportData.login) {
                login = passportData.login;
                console.log('Token Owner:', login);
            } else {
                console.warn('Passport API error:', passportData);
            }
        } catch (e) {
            console.warn('Passport API fetch error:', e);
        }

        // 2. Пробуем получить информацию через Direct API (GetClientInfo)
        // Это подтвердит, работает ли сам API Директа
        try {
            const clientInfo = callWordstatApi('GetClientInfo', ['Login']);
            console.log('Direct Client Info:', clientInfo);
        } catch (e) {
            console.warn('Direct GetClientInfo failed:', e.message);
        }

        // 3. Пробуем SANDBOX (Песочницу)
        // Если тут сработает - значит токен рабочий, а проблема в статусе боевого аккаунта
        let sandboxStatus = 'Не проверялся';
        try {
            const sandboxUrl = 'https://api-sandbox.direct.yandex.com/v4/json/';
            const payload = {
                method: 'GetWordstatReportList',
                token: config.oauthToken,
                locale: 'ru',
                param: {}
            };
            const options = {
                method: 'post',
                contentType: 'application/json',
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            };
            const sbResp = UrlFetchApp.fetch(sandboxUrl, options);
            const sbJson = JSON.parse(sbResp.getContentText());
            if (sbJson.error_code) {
                sandboxStatus = `Ошибка ${sbJson.error_code}: ${sbJson.error_str}`;
            } else {
                sandboxStatus = '✅ УСПЕШНО (доступ есть)';
            }
        } catch (e) {
            sandboxStatus = 'Ошибка запроса: ' + e.message;
        }
        console.log('Sandbox Status:', sandboxStatus);

        // 4. Пробуем получить список отчетов (Production)
        const response = callWordstatApi('GetWordstatReportList', {});
        console.log('API Response:', response);

        let msg = '✅ Соединение с API Вордстата установлено!\n\n';
        msg += `Владелец токена: ${login}\n`;
        msg += `Песочница: ${sandboxStatus}\n`;
        msg += `Токен: ${config.oauthToken.substring(0, 10)}...\n`;
        msg += `Текущих отчетов: ${response.length}\n`;

        ui.alert('Успех', msg, ui.ButtonSet.OK);

    } catch (e) {
        console.error(e);
        let errorMsg = e.message;
        if (errorMsg.includes('Error 53')) {
            errorMsg += '\n\n🔍 ДИАГНОСТИКА:\n';
            errorMsg += 'Владелец токена: ' + (login || 'Не определен') + '\n';
            errorMsg += 'Песочница: ' + (sandboxStatus || 'Не проверена') + '\n';
            errorMsg += '\nЕсли Песочница работает, а боевой API нет — значит, ваш аккаунт Директа не активен (нужно создать хотя бы одну черновую кампанию).';
        }
        ui.alert('❌ Ошибка API', errorMsg, ui.ButtonSet.OK);
    }
}
