/**
 * Скрипт для детальной инспекции данных одного товара.
 * Помогает понять, в каком поле на самом деле лежит Бренд.
 */
function debug_inspect_product_data() {
    const context = 'debug_inspect_product_data';
    const credentials = getInsalesCredentialsSync();

    if (!credentials) {
        Logger.log('❌ Ошибка: Не удалось получить доступы к InSales');
        return;
    }

    Logger.log('🚀 Начинаем инспекцию первого товара...');

    // Загружаем ровно 1 товар с полным набором полей
    const url = `${credentials.baseUrl}/admin/products.json?per_page=1&page=1&fields=id,title,vendor,characteristics,variants`;
    const options = {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const products = JSON.parse(response.getContentText());

        if (!products || products.length === 0) {
            Logger.log('⚠️ Товары не найдены.');
            return;
        }

        const p = products[0];

        Logger.log('--------------------------------------------------');
        Logger.log(`📦 Товар ID: ${p.id}`);
        Logger.log(`🏷️ Название: "${p.title}"`);
        Logger.log(`🏭 Поле vendor (Производитель): "${p.vendor}"`); // Самое важное!
        Logger.log('--------------------------------------------------');
        Logger.log('📋 Характеристики (все):');

        if (p.characteristics && p.characteristics.length > 0) {
            p.characteristics.forEach(c => {
                Logger.log(`   - ID: ${c.property_id} | Title: "${c.title}" | Permalink свойства: (нужен отдельный запрос)`);
            });
        } else {
            Logger.log('   (Характеристик нет)');
        }
        Logger.log('--------------------------------------------------');

        // Также проверим, какие ID свойств мы считаем "Брендом" сейчас
        const brandPropId = getBrandPropertyId_(credentials);
        Logger.log(`🔍 Скрипт сейчас ищет бренд в свойстве с ID: ${brandPropId}`);

    } catch (e) {
        Logger.log(`❌ Ошибка запроса: ${e.message}`);
    }
}
