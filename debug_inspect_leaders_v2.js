function debugInspectLeadersV2() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    console.log(`🔍 Inspecting Leaders V2 for Category ${categoryId}...`);

    const targetNames = [
        "БПЦ7 8х30 КОМЗ",
        "MINOX BLU 8x25",
        "БПЦ5 8х30 Байгыш"
    ];

    // 1. Fetch ALL products in category (handling pagination)
    let allProducts = [];
    let page = 1;
    while (true) {
        const url = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=250&page=${page}`;
        const resp = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const batch = JSON.parse(resp.getContentText());
        if (batch.length === 0) break;
        allProducts = allProducts.concat(batch);
        page++;
        Utilities.sleep(200);
    }
    console.log(`   Fetched ${allProducts.length} products total.`);

    // 2. Find Targets
    const foundProducts = [];
    for (const namePart of targetNames) {
        const match = allProducts.find(p => p.title.includes(namePart));
        if (match) {
            foundProducts.push(match);
            console.log(`✅ Found "${namePart}" -> ID: ${match.id} | Title: ${match.title}`);
        } else {
            console.warn(`❌ Could not find product containing "${namePart}"`);
        }
    }

    // 3. Inspect Collects for Found Products
    console.log('\n🕵️‍♂️ DEEP INSPECTION OF LEADERS:');
    for (const p of foundProducts) {
        console.log(`\n📦 Product: ${p.title} (${p.id})`);
        console.log(`   - sort_weight: ${p.sort_weight}`);

        // Fetch Collect
        const urlCollect = `${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&product_id=${p.id}`;
        const respCollect = UrlFetchApp.fetch(urlCollect, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const collects = JSON.parse(respCollect.getContentText());

        if (collects.length > 0) {
            const c = collects[0];
            console.log(`   - Collect ID: ${c.id}`);
            console.log(`   - Collect Position: ${c.position}`);
        } else {
            console.warn(`   ⚠️ No Collect found!`);
        }
    }
}
