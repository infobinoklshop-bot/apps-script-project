function debugFindProductRobust() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();

    console.log('🚀 Starting Robust Product Search...');

    const targetNamePart = "БПЦ7"; // Short, distinctive part
    console.log(`\n🔍 Searching for product containing: "${targetNamePart}"...`);

    // Fetch products in batches until found
    let page = 1;
    let foundProduct = null;

    while (true) {
        // Fetch global products (not scoped to collection)
        const url = `${credentials.baseUrl}/admin/products.json?per_page=250&page=${page}`;
        const resp = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const batch = JSON.parse(resp.getContentText());

        if (batch.length === 0) break;

        foundProduct = batch.find(p => p.title.includes(targetNamePart));
        if (foundProduct) break;

        console.log(`   Checked page ${page} (${batch.length} products)...`);
        page++;
        Utilities.sleep(200);
    }

    if (foundProduct) {
        console.log(`\n✅ FOUND PRODUCT:`);
        console.log(`   Title: "${foundProduct.title}"`);
        console.log(`   ID: ${foundProduct.id}`);
        console.log(`   Sort Weight: ${foundProduct.sort_weight}`);

        // Check Collections
        console.log(`\n📂 Checking its collections...`);
        const urlCollects = `${credentials.baseUrl}/admin/collects.json?product_id=${foundProduct.id}`;
        const respCollects = UrlFetchApp.fetch(urlCollects, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const collects = JSON.parse(respCollects.getContentText());

        if (collects.length === 0) {
            console.log(`   ⚠️ This product is in NO collections directly.`);
        } else {
            for (const c of collects) {
                try {
                    const urlCat = `${credentials.baseUrl}/admin/collections/${c.collection_id}.json`;
                    const respCat = UrlFetchApp.fetch(urlCat, {
                        headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
                    });
                    const cat = JSON.parse(respCat.getContentText());
                    console.log(`   -> Collection: "${cat.title}" (ID: ${cat.id}) | Sort: ${cat.sort_type} | Position: ${c.position}`);
                } catch (e) {
                    console.log(`   -> Collection ID: ${c.collection_id} (Error fetching details)`);
                }
            }
        }
    } else {
        console.error(`❌ Could not find any product containing "${targetNamePart}" in the entire shop.`);
    }
}
