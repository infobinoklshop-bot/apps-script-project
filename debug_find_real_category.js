function debugFindRealCategory() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();

    console.log('🚀 Starting Detective Work...');

    // 1. Find the "Ghost" Product (БПЦ7)
    const targetName = "БПЦ7 8х30 КОМЗ";
    console.log(`\n🔍 Searching for product: "${targetName}"...`);

    const encoded = encodeURIComponent(targetName);
    const urlSearch = `${credentials.baseUrl}/admin/products.json?q=${encoded}`;
    const resp = UrlFetchApp.fetch(urlSearch, {
        headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
    });
    const found = JSON.parse(resp.getContentText());

    let ghostProduct = null;
    if (found.length > 0) {
        ghostProduct = found[0];
        console.log(`✅ Found Product: "${ghostProduct.title}" (ID: ${ghostProduct.id})`);

        // 2. See which collections it belongs to
        console.log(`   Checking its collections...`);
        const urlCollects = `${credentials.baseUrl}/admin/collects.json?product_id=${ghostProduct.id}`;
        const respCollects = UrlFetchApp.fetch(urlCollects, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const collects = JSON.parse(respCollects.getContentText());

        if (collects.length === 0) {
            console.log(`   ⚠️ This product is in NO collections directly.`);
        } else {
            for (const c of collects) {
                // Get Collection Details
                try {
                    const urlCat = `${credentials.baseUrl}/admin/collections/${c.collection_id}.json`;
                    const respCat = UrlFetchApp.fetch(urlCat, {
                        headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
                    });
                    const cat = JSON.parse(respCat.getContentText());
                    console.log(`   -> Found in Collection: "${cat.title}" (ID: ${cat.id}) | Sort Type: ${cat.sort_type}`);
                } catch (e) {
                    console.log(`   -> Found in Collection ID: ${c.collection_id} (Could not fetch details)`);
                }
            }
        }
    } else {
        console.error(`❌ Could not find product "${targetName}" globally.`);
    }

    // 3. List ALL Collections named "Бинокли"
    console.log(`\n🔍 Searching for collections named "Бинокли"...`);
    const urlCats = `${credentials.baseUrl}/admin/collections.json?per_page=250`; // Fetch all (or many)
    const respCats = UrlFetchApp.fetch(urlCats, {
        headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
    });
    const allCats = JSON.parse(respCats.getContentText());

    const matches = allCats.filter(c => c.title.toLowerCase().includes("бинокли"));

    if (matches.length === 0) {
        console.log('   No collections found with "Бинокли" in title.');
    } else {
        for (const c of matches) {
            console.log(`   📂 "${c.title}" (ID: ${c.id}) | Parent: ${c.parent_id} | Sort: ${c.sort_type}`);
        }
    }
}
