function debugInspectFullObjects() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();
    console.log(`🔍 Inspecting Full Objects for Category ${categoryId}...`);

    // 1. Get Collection
    const urlCat = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
    const respCat = UrlFetchApp.fetch(urlCat, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const category = JSON.parse(respCat.getContentText());
    console.log('\n📂 COLLECTION OBJECT:');
    console.log(JSON.stringify(category, null, 2));

    // 2. Get Top 1 Product (Alien)
    const urlProds = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=1`;
    const respProds = UrlFetchApp.fetch(urlProds, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const products = JSON.parse(respProds.getContentText());

    if (products.length === 0) return;
    const alien = products[0];

    console.log('\n👽 ALIEN PRODUCT OBJECT:');
    console.log(JSON.stringify(alien, null, 2));

    // 3. Get Collect for Alien
    const urlCollects = `${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&product_id=${alien.id}`;
    const respCollects = UrlFetchApp.fetch(urlCollects, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const collects = JSON.parse(respCollects.getContentText());

    if (collects.length > 0) {
        console.log('\n🔗 ALIEN COLLECT OBJECT:');
        console.log(JSON.stringify(collects[0], null, 2));
    } else {
        console.log('\n❌ Alien Collect not found via API filter.');
        // Try to find it in all collects
        const allCollects = getAllCollectsForCategory(categoryId, credentials);
        const found = allCollects.find(c => c.product_id == alien.id);
        if (found) {
            console.log('\n🔗 ALIEN COLLECT OBJECT (from full list):');
            console.log(JSON.stringify(found, null, 2));
        }
    }
}
