function debugCreateTestCategory() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();
    const sheet = ss.getActiveSheet(); // Moved up for parent ID logic

    // 0. Get Parent ID from current category
    const currentCategoryId = sheet.getRange('B2').getValue();
    let parentId = null;

    if (currentCategoryId) {
        try {
            const urlCat = `${credentials.baseUrl}/admin/collections/${currentCategoryId}.json`;
            const respCat = UrlFetchApp.fetch(urlCat, {
                method: 'GET',
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                }
            });
            const currentCat = JSON.parse(respCat.getContentText());
            parentId = currentCat.parent_id;
            console.log(`ℹ️ Using Parent ID: ${parentId}`);
        } catch (e) {
            console.warn(`⚠️ Could not fetch current category parent: ${e.message}`);
        }
    }

    console.log('🚀 Starting Test: Create Fresh Category...');

    // 1. Create Category
    const catPayload = {
        collection: {
            title: "TEST_SORTING_DEBUG_" + new Date().getTime(),
            is_hidden: true,
            sort_type: "0" // Try 0 explicitly
        }
    };

    if (parentId) {
        catPayload.collection.parent_id = parentId;
    }

    let categoryId = null;

    try {
        const url = `${credentials.baseUrl}/admin/collections.json`;
        const response = UrlFetchApp.fetch(url, {
            method: 'POST',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(catPayload),
            muteHttpExceptions: true
        });

        const text = response.getContentText();
        if (response.getResponseCode() === 201) {
            const cat = JSON.parse(text);
            categoryId = cat.id;
            console.log(`✅ Created Category ID: ${categoryId}`);
            console.log(`   Sort Type: ${cat.sort_type}`);
        } else {
            console.error(`❌ Failed to create category with sort_type 0: ${text}`);

            // Retry with sort_type 7
            console.log('   Retrying with sort_type 7...');
            catPayload.collection.sort_type = "7";
            const resp2 = UrlFetchApp.fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                },
                payload: JSON.stringify(catPayload),
                muteHttpExceptions: true
            });

            if (resp2.getResponseCode() === 201) {
                const cat = JSON.parse(resp2.getContentText());
                categoryId = cat.id;
                console.log(`✅ Created Category ID: ${categoryId} (with sort_type 7)`);
            } else {
                console.error(`❌ Failed again: ${resp2.getContentText()}`);
                return;
            }
        }
    } catch (e) {
        console.error(`❌ Exception creating category: ${e.message}`);
        return;
    }

    if (!categoryId) return;

    // 2. Add 3 Products
    // We need valid product IDs. Let's grab 3 from the sheet.
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const productIds = [];
    for (let i = 0; i < 3; i++) {
        const id = sheet.getRange(startRow + i, 5).getValue();
        if (id) productIds.push(parseInt(id));
    }

    console.log(`\n📦 Adding products: ${productIds.join(', ')}`);

    const collectIds = [];

    for (const pid of productIds) {
        try {
            const payload = {
                collect: {
                    product_id: pid,
                    collection_id: categoryId
                }
            };
            const resp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects.json`, {
                method: 'POST',
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                },
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            });
            if (resp.getResponseCode() === 201) {
                const c = JSON.parse(resp.getContentText());
                collectIds.push({ id: c.id, pid: pid });
                console.log(`   + Added Product ${pid} (Collect ${c.id})`);
            } else {
                console.error(`   - Failed to add ${pid}: ${resp.getContentText()}`);
            }
        } catch (e) {
            console.error(`   - Exception adding ${pid}: ${e.message}`);
        }
    }

    if (collectIds.length < 3) {
        console.error('Not enough products added.');
        return;
    }

    // 3. Test Sorting
    console.log('\n🔄 Testing Sorting...');
    // Reverse order: 3, 2, 1
    const targetOrder = [productIds[2], productIds[1], productIds[0]];

    for (let i = 0; i < 3; i++) {
        const targetPid = targetOrder[i];
        const collect = collectIds.find(c => c.pid == targetPid);
        updateCollectPosition(collect.id, i + 1, credentials);
        console.log(`   Set ${targetPid} to Pos ${i + 1}`);
    }

    Utilities.sleep(2000);

    // 4. Verify
    const urlProds = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}`;
    const respProds = UrlFetchApp.fetch(urlProds, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const actualProds = JSON.parse(respProds.getContentText());
    const actualIds = actualProds.map(p => p.id);

    console.log(`\n📊 Result:`);
    console.log(`   Expected: ${targetOrder.join(', ')}`);
    console.log(`   Actual:   ${actualIds.join(', ')}`);

    if (JSON.stringify(actualIds) === JSON.stringify(targetOrder)) {
        console.log('   🎉 SUCCESS! Sorting works in new category.');
    } else {
        console.log('   ❌ FAILURE! Sorting ignored in new category.');
    }

    // Cleanup? Maybe leave it for inspection.
    console.log(`\nℹ️ Created Category ID: ${categoryId}. You can delete it later.`);
}
