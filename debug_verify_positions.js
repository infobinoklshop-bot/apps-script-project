function debugVerifyPositions() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();

    // 1. Get first 5 products from Sheet
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const testProducts = [];

    console.log('📋 Reading first 5 products from Sheet...');
    for (let i = 0; i < 5; i++) {
        const row = startRow + i;
        const id = sheet.getRange(row, 5).getValue(); // Col E
        const title = sheet.getRange(row, 2).getValue(); // Col B
        if (id) {
            testProducts.push({ id: parseInt(id), title: title, targetPos: i + 1 });
            console.log(`   [${i + 1}] ID: ${id} | ${title}`);
        }
    }

    if (testProducts.length === 0) {
        console.error('No products found in sheet.');
        return;
    }

    // 2. Get Collects
    const collects = getAllCollectsForCategory(categoryId, credentials);

    // 3. Force Update Positions
    console.log('\n🚀 Force updating positions 1-5...');
    for (const p of testProducts) {
        const collect = collects.find(c => c.product_id == p.id);
        if (!collect) {
            console.error(`   ❌ Collect not found for product ${p.id}`);
            continue;
        }

        const success = updateCollectPosition(collect.id, p.targetPos, credentials);
        if (success) {
            console.log(`   ✅ Set ${p.title} to Pos ${p.targetPos}`);
        } else {
            console.error(`   ❌ Failed to set ${p.title}`);
        }
        Utilities.sleep(500); // Slow down to be safe
    }

    console.log('\n⏳ Waiting 2 seconds for propagation...');
    Utilities.sleep(2000);

    // 4. Read back ACTUAL order from API
    console.log('\n👀 Reading actual order from InSales API...');
    const url = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=10`;
    const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const actualProducts = JSON.parse(response.getContentText());

    console.log('\n📊 ACTUAL TOP 10 PRODUCTS IN INSALES:');
    actualProducts.forEach((p, index) => {
        const isExpected = testProducts.find(tp => tp.id === p.id);
        const marker = isExpected ? (isExpected.targetPos === index + 1 ? '✅' : '⚠️') : '👽';

        console.log(`   ${marker} Pos ${index + 1}: [${p.id}] ${p.title}`);

        if (isExpected && isExpected.targetPos !== index + 1) {
            console.log(`      (Expected at ${isExpected.targetPos})`);
        }
        if (!isExpected) {
            console.log(`      (Not in our top 5 list - "Alien")`);
        }
    });

    console.log('\n🏁 Verification Done.');
}
