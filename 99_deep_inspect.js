function deepInspectCategory() {
    // ID категории "С дальномерной сеткой", которая точно есть (из прошлого лога)
    const categoryId = 21752437;

    console.log(`🕵️ Deep inspection for Category ID: ${categoryId}...`);

    const credentials = getInsalesCredentialsSync();
    const headers = {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
    };

    // 1. Fetch Category Data
    const catUrl = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
    const catResponse = UrlFetchApp.fetch(catUrl, { method: 'GET', headers: headers, muteHttpExceptions: true });
    const category = JSON.parse(catResponse.getContentText());

    console.log(`\n📂 CATEGORY FIELDS:`);
    console.log(`Template: ${category.template}`);
    console.log(`Custom Template: ${category.custom_template}`);

    // 2. Inspect Field Values (Standard Custom Fields)
    if (category.field_values) {
        console.log(`\n🔧 FIELD VALUES (${category.field_values.length}):`);
        category.field_values.forEach(fv => {
            // Try to find handle if possible (sometimes it's nested or just ID)
            console.log(`- ID: ${fv.id}, Field ID: ${fv.collection_field_id}, Value: "${fv.value}"`);
        });
    }

    // 3. Fetch Metafields (Often used for Widgets/Settings)
    // InSales API: /admin/collections/{id}/metafields.json
    const metaUrl = `${credentials.baseUrl}/admin/collections/${categoryId}/metafields.json`;
    const metaResponse = UrlFetchApp.fetch(metaUrl, { method: 'GET', headers: headers, muteHttpExceptions: true });

    if (metaResponse.getResponseCode() === 200) {
        const metafields = JSON.parse(metaResponse.getContentText());
        console.log(`\n🧠 METAFIELDS (${metafields.length || 0}):`);
        if (metafields.length > 0) {
            metafields.forEach(mf => {
                console.log(`- Namespace: ${mf.namespace}, Key: ${mf.key}, Value: ${mf.value}`);
            });
        } else {
            console.log("No metafields found.");
        }
    } else {
        console.log(`\n❌ Error fetching metafields: ${metaResponse.getResponseCode()}`);
    }
}
