/**
 * Скрипт для поиска ID свойства, в котором реально лежат бренды.
 * Сканирует товары и ищет в характеристиках известные бренды.
 */
function debug_find_brand_id() {
    const context = 'debug_find_brand_id';
    const credentials = getInsalesCredentialsSync();

    if (!credentials) {
        Logger.log('❌ Ошибка: Не удалось получить доступы к InSales');
        return;
    }

    Logger.log('🚀 Ищем свойство "Бренд"...');

    // Список известных брендов для поиска
    const KNOWN_BRANDS = [
        'Nikon', 'Veber', 'Canon', 'Yukon', 'Bushnell',
        'Levenhuk', 'Bresser', 'Discovery', 'Sturman',
        'Vector Optics', 'Swarovski', 'Zeiss', 'Pulsar'
    ];

    // Загружаем 50 товаров
    const url = `${credentials.baseUrl}/admin/products.json?per_page=50&page=1&fields=id,title,characteristics`;
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

        const usageStats = {}; // { property_id: count }
        const examples = {};   // { property_id: ["Nikon", "Veber"] }

        products.forEach(p => {
            if (!p.characteristics) return;

            p.characteristics.forEach(c => {
                // Проверяем, похоже ли значение на бренд
                // 1. Точное совпадение с известными
                // 2. Или просто значение совпадает с началом названия товара (эвристика)

                let isBrand = false;
                const value = c.title.trim();

                // Проверка по списку
                if (KNOWN_BRANDS.some(b => value.toLowerCase() === b.toLowerCase())) {
                    isBrand = true;
                }

                if (isBrand) {
                    if (!usageStats[c.property_id]) {
                        usageStats[c.property_id] = 0;
                        examples[c.property_id] = new Set();
                    }
                    usageStats[c.property_id]++;
                    examples[c.property_id].add(value);
                }
            });
        });

        Logger.log('📊 Результаты анализа:');
        const ids = Object.keys(usageStats);

        if (ids.length === 0) {
            Logger.log('❌ Не найдено ни одного свойства, содержащего известные бренды.');
            return;
        }

        ids.forEach(id => {
            Logger.log(`✅ ID свойства: ${id}`);
            Logger.log(`   🔸 Количество совпадений: ${usageStats[id]}`);
            Logger.log(`   🔸 Примеры значений: ${Array.from(examples[id]).join(', ')}`);
        });

        Logger.log('--------------------------------------------------');
        Logger.log('💡 Используйте найденный ID (где больше всего совпадений) в основном скрипте.');

    } catch (e) {
        Logger.log(`❌ Ошибка: ${e.message}`);
    }
}
