function debugInspectManualLeaders() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    console.log(`🔍 Inspecting Manual Leaders for Category ${categoryId}...`);

    // Target Names from Screenshot
    const targetNames = [
        "Бинокль БПЦ7 8х30 КОМЗ",      // Visual Pos 1
        "Бинокль MINOX BLU 8x25",      // Visual Pos 2
        "Бинокль БПЦ5 8х30 Байгыш"     // Visual Pos 3 (Moved here)
    ];

    // 1. Search for these products to get IDs
    const productIds = {};

    for (const name of targetNames) {
        // Search by query
        const encoded = encodeURIComponent(name);
        const urlSearch = `${credentials.baseUrl}/admin/products.json?q=${encoded}`;
        const resp = UrlFetchApp.fetch(urlSearch, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const found = JSON.parse(resp.getContentText());

        if (found.length > 0) {
            // Find exact match if possible, or take first
            const match = found.find(p => p.title.includes(name)) || found[0];
            productIds[name] = match;
            console.log(`✅ Found "${name}" -> ID: ${match.id}`);
        } else {
            console.warn(`❌ Could not find product "${name}"`);
        }
    }

    // 2. Inspect their Collects and Product Data
    console.log('\n🕵️‍♂️ DEEP INSPECTION:');

    for (const name of targetNames) {
        const product = productIds[name];
        if (!product) continue;

        console.log(`\n📦 Product: ${name} (${product.id})`);
        console.log(`   - sort_weight: ${product.sort_weight}`);

        // Fetch Collect for this category
        const urlCollect = `${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&product_id=${product.id}`;
        const respCollect = UrlFetchApp.fetch(urlCollect, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const collects = JSON.parse(respCollect.getContentText());

        if (collects.length > 0) {
            const c = collects[0];
            console.log(`   - Collect ID: ${c.id}`);
            console.log(`   - Collect Position: ${c.position}`);
            console.log(`   - Full Collect JSON: ${JSON.stringify(c)}`);
        } else {
            console.warn(`   ⚠️ No Collect found for this category! (Is it mapped via a parent?)`);
        }
    }
}
