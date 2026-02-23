function testPopulateWidget() {
    // 1. Настройки
    const categoryId = 21752437; // "С дальномерной сеткой" (которая точно есть)
    const fieldId = 292698;      // ID поля "Товары и категории (популярные)"

    // ID товаров для теста (возьмем несколько реальных ID из вашего магазина)
    // Например, те, что были в логах или просто наугад из "Монокуляров"
    // Лучше взять ID товаров, которые ТОЧНО есть.
    // Давайте сначала получим список товаров из этой категории, чтобы взять их ID.

    const credentials = getInsalesCredentialsSync();
    const headers = {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
    };

    console.log(`🧪 Testing Widget Population for Category ${categoryId}...`);

    // Шаг 1: Получаем товары категории, чтобы взять пару ID для теста
    const productsUrl = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=5`;
    const prodResponse = UrlFetchApp.fetch(productsUrl, { method: 'GET', headers: headers, muteHttpExceptions: true });
    const products = JSON.parse(prodResponse.getContentText());

    if (products.length < 2) {
        console.log("❌ Not enough products in category to test.");
        return;
    }

    const testProductIds = products.map(p => p.id).join(',');
    console.log(`📝 Selected Product IDs for widget: ${testProductIds}`);

    // Шаг 2: Обновляем поле категории
    const updateUrl = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
    const payload = {
        collection: {
            field_values: [
                {
                    field_id: fieldId,
                    value: testProductIds // Записываем ID через запятую
                }
            ]
        }
    };

    const updateOptions = {
        method: 'PUT',
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const updateResponse = UrlFetchApp.fetch(updateUrl, updateOptions);

    if (updateResponse.getResponseCode() === 200) {
        console.log("✅ SUCCESS! Field updated.");
        console.log(`🔗 Check the page: https://binokl.shop/collection/s-dalnomerom`); // Permalink from previous logs
        console.log("The widget 'Товары и категории (популярные)' should now display these products.");
    } else {
        console.log(`❌ Error updating category: ${updateResponse.getResponseCode()}`);
        console.log(updateResponse.getContentText());
    }
}
