function debugAnalyzeManualSort() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    console.log(`🔍 Analyzing Manual Sort for Category ${categoryId}...`);

    // 1. Fetch All Collects (with Pagination)
    let collects = [];
    let page = 1;
    while (true) {
        const urlCollects = `${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&per_page=250&page=${page}`;
        const respCollects = UrlFetchApp.fetch(urlCollects, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const batch = JSON.parse(respCollects.getContentText());
        if (batch.length === 0) break;
        collects = collects.concat(batch);
        page++;
        Utilities.sleep(100);
    }
    console.log(`   Fetched ${collects.length} collects total.`);

    // 2. Search for "Meade" (The Ghost Leader)
    // We need to fetch product details for this check, but let's first check if we can find it by ID if we knew it.
    // Since we don't know the ID, we'll rely on the top sort first.

    // 3. Sort Collects by Position
    collects.sort((a, b) => a.position - b.position);

    // 4. Get Top 20 IDs
    const topCollects = collects.slice(0, 20);
    const topProductIds = topCollects.map(c => c.product_id);

    // 5. Fetch Details for these products
    console.log(`\n🔍 Fetching details for top ${topProductIds.length} products...`);
    let productMap = {};

    for (const pid of topProductIds) {
        try {
            const resp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/products/${pid}.json`, {
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
            });
            const p = JSON.parse(resp.getContentText());
            productMap[p.id] = p.title;
        } catch (e) {
            productMap[pid] = "ERROR_FETCHING";
        }
        Utilities.sleep(100);
    }

    // 6. Display Top 20
    console.log('\n🏆 TOP 20 PRODUCTS (By Collect Position):');
    for (let i = 0; i < topCollects.length; i++) {
        const c = topCollects[i];
        const name = productMap[c.product_id] || `Unknown Product (${c.product_id})`;
        console.log(`   #${i + 1} [Pos: ${c.position}] ${name}`);
        console.log(`       Collect ID: ${c.id} | Product ID: ${c.product_id}`);
    }

    // 7. Check specifically for "Meade" if not in top 20
    // We need to find its ID first.
    console.log('\n🕵️‍♂️ Searching for "Meade"...');
    const meadeSearchUrl = `${credentials.baseUrl}/admin/products.json?q=Meade`;
    const meadeResp = UrlFetchApp.fetch(meadeSearchUrl, {
        headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
    });
    const meadeProds = JSON.parse(meadeResp.getContentText());
    const meade = meadeProds.find(p => p.title.includes("Meade Wilderness 10x42"));

    if (meade) {
        const meadeCollect = collects.find(c => c.product_id == meade.id);
        if (meadeCollect) {
            console.log(`   ✅ Found "Meade Wilderness 10x42" in category!`);
            console.log(`   Position: ${meadeCollect.position}`);
            console.log(`   Collect ID: ${meadeCollect.id}`);
        } else {
            console.log(`   ❌ "Meade Wilderness 10x42" (ID: ${meade.id}) is NOT in this category.`);
        }
    } else {
        console.log(`   ❌ Could not find "Meade Wilderness 10x42" in shop.`);
    }

    // 7. Dump Raw JSON of Top 1
    if (topCollects.length > 0) {
        console.log('\n📄 RAW JSON of Top 1 Collect:');
        console.log(JSON.stringify(topCollects[0], null, 2));
    }
}
