function debugFindWorkingSortType() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();

    // 1. Identify 3 products to use for testing
    // We'll take the first 3 IDs from the sheet
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const productIds = [];
    for (let i = 0; i < 3; i++) {
        const id = sheet.getRange(startRow + i, 5).getValue(); // Col E
        if (id) productIds.push(parseInt(id));
    }

    if (productIds.length < 3) {
        console.error('Need at least 3 products in the sheet to test sorting.');
        return;
    }

    console.log(`🎯 Test Products: ${productIds.join(', ')}`);
    console.log(`   Target Order: 1, 2, 3`);

    // 2. Get Collect IDs for these products
    const collects = getAllCollectsForCategory(categoryId, credentials);
    const testCollects = [];
    for (const pid of productIds) {
        const c = collects.find(x => x.product_id == pid);
        if (c) testCollects.push({ collectId: c.id, productId: pid });
    }

    if (testCollects.length < 3) {
        console.error('Could not find collects for test products.');
        return;
    }

    // 3. Iterate through candidates
    // Based on logs, 1-12 are accepted. 7 is current (broken). 1 is likely Price.
    // Let's test likely candidates for "Manual" or "Custom".
    // We'll test ALL accepted integers just to be sure.
    const candidates = [12, 11, 10, 9, 8, 6, 5, 4, 3, 2, 1];

    for (const sortType of candidates) {
        console.log(`\n🧪 Testing sort_type: ${sortType} ...`);

        // A. Set Sort Type
        try {
            const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
            UrlFetchApp.fetch(url, {
                method: 'PUT',
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                },
                payload: JSON.stringify({ collection: { sort_type: sortType } }),
                muteHttpExceptions: true
            });
        } catch (e) {
            console.error(`   ❌ Failed to set sort_type ${sortType}`);
            continue;
        }

        // B. Set Positions (Force them to be 1, 2, 3)
        // We do this EVERY time to ensure they are fresh
        for (let i = 0; i < 3; i++) {
            updateCollectPosition(testCollects[i].collectId, i + 1, credentials);
        }

        Utilities.sleep(1000); // Wait for propagation

        // C. Check ACTUAL Order via Products API
        // This returns products in the order they appear in the category
        const productsUrl = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=5`;
        const response = UrlFetchApp.fetch(productsUrl, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            }
        });
        const products = JSON.parse(response.getContentText());

        // Check if first 3 match our test IDs
        const actualIds = products.map(p => p.id).slice(0, 3);
        console.log(`   👀 Actual Order: ${actualIds.join(', ')}`);

        const isMatch =
            actualIds[0] === productIds[0] &&
            actualIds[1] === productIds[1] &&
            actualIds[2] === productIds[2];

        if (isMatch) {
            console.log(`\n🎉🎉🎉 FOUND IT! sort_type ${sortType} respects manual positioning! 🎉🎉🎉`);
            console.log(`PLEASE TELL THE DEVELOPER: sort_type = ${sortType}`);
            return;
        } else {
            console.log(`   ❌ Mismatch. This sort_type does NOT respect manual positions.`);
        }
    }

    console.log('\n🏁 Tested all candidates. None worked perfectly. Check logs.');
}
