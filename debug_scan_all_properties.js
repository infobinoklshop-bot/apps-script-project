/**
 * Скрипт полного сканирования свойств.
 * 1. Загружает ВСЕ свойства аккаунта, чтобы знать их permalink.
 * 2. Загружает 5 товаров.
 * 3. Выводит ПОЛНУЮ карту данных: какой ID свойства = какой permalink = какое значение.
 * Это покажет, где ИМЕННО лежат бренды.
 */
function debug_scan_all_properties() {
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
        Logger.log('❌ Нет доступов');
        return;
    }

    Logger.log('⏳ 1. Загружаем справочник всех свойств...');
    const propMap = {}; // id -> { title, permalink }

    try {
        let propsUrl = `${credentials.baseUrl}/admin/properties.json`;
        const propsResp = UrlFetchApp.fetch(propsUrl, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
        });

        const allProps = JSON.parse(propsResp.getContentText());
        if (allProps && allProps.length) {
            allProps.forEach(p => {
                propMap[p.id] = { title: p.title, permalink: p.permalink };
            });
            Logger.log(`✅ Загружено ${allProps.length} свойств. Справочник готов.`);
        } else {
            Logger.log('⚠️ Свойства не найдены или ошибка API.');
        }

        Logger.log('⏳ 2. Загружаем 5 товаров для проверки...');
        const prodUrl = `${credentials.baseUrl}/admin/products.json?per_page=5&page=1&fields=id,title,vendor,characteristics`;
        const prodResp = UrlFetchApp.fetch(prodUrl, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
        });

        const products = JSON.parse(prodResp.getContentText());

        Logger.log('==================================================');
        products.forEach(p => {
            Logger.log(`📦 Товар: ${p.title} (ID: ${p.id})`);
            Logger.log(`   🏭 Field 'vendor': "${p.vendor}"`);

            if (p.characteristics && p.characteristics.length) {
                p.characteristics.forEach(c => {
                    const info = propMap[c.property_id] || { title: '???', permalink: '???' };
                    Logger.log(`   🔸 [ID:${c.property_id}] [${info.permalink}] ${info.title}: "${c.title}"`);
                });
            } else {
                Logger.log(`   (Характеристик нет)`);
            }
            Logger.log('--------------------------------------------------');
        });
        Logger.log('==================================================');

    } catch (e) {
        Logger.log('❌ Ошибка: ' + e.message);
    }
}
