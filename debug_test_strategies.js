function debugTestStrategies() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();
    const sheet = ss.getActiveSheet();

    // 0. Get Parent ID
    const currentCategoryId = sheet.getRange('B2').getValue();
    let parentId = null;
    if (currentCategoryId) {
        try {
            const urlCat = `${credentials.baseUrl}/admin/collections/${currentCategoryId}.json`;
            const respCat = UrlFetchApp.fetch(urlCat, {
                method: 'GET',
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
            });
            parentId = JSON.parse(respCat.getContentText()).parent_id;
        } catch (e) { }
    }

    console.log('🚀 Starting Strategies Test...');

    // Helper to create category
    function createCat(name) {
        const payload = {
            collection: {
                title: name,
                is_hidden: true,
                sort_type: "7",
                parent_id: parentId
            }
        };
        try {
            const resp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collections.json`, {
                method: 'POST',
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                },
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            });
            if (resp.getResponseCode() === 201) return JSON.parse(resp.getContentText()).id;
        } catch (e) { }
        return null;
    }

    // Get 3 Product IDs
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const pids = [];
    for (let i = 0; i < 3; i++) {
        const id = sheet.getRange(startRow + i, 5).getValue();
        if (id) pids.push(parseInt(id));
    }
    if (pids.length < 3) { console.error('Not enough products'); return; }

    // STRATEGY 1: set_products
    console.log('\n🧪 STRATEGY 1: set_products endpoint');
    const cat1 = createCat("DEBUG_STRATEGY_1_" + new Date().getTime());
    if (cat1) {
        console.log(`   Created Cat ${cat1}`);
        // Try to set products in REVERSE order
        const reversePids = [pids[2], pids[1], pids[0]];
        const urlSet = `${credentials.baseUrl}/admin/collections/${cat1}/set_products.json`;
        const payloadSet = { product_ids: reversePids };

        try {
            const respSet = UrlFetchApp.fetch(urlSet, {
                method: 'PUT', // or POST? Try PUT first
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                },
                payload: JSON.stringify(payloadSet),
                muteHttpExceptions: true
            });
            console.log(`   Response Code: ${respSet.getResponseCode()}`);
            console.log(`   Response Text: ${respSet.getContentText()}`);

            // Verify
            Utilities.sleep(2000);
            const urlCheck = `${credentials.baseUrl}/admin/products.json?collection_id=${cat1}`;
            const respCheck = UrlFetchApp.fetch(urlCheck, {
                method: 'GET',
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
            });
            const actualIds = JSON.parse(respCheck.getContentText()).map(p => p.id);
            console.log(`   Expected: ${reversePids.join(', ')}`);
            console.log(`   Actual:   ${actualIds.join(', ')}`);

            if (JSON.stringify(actualIds) === JSON.stringify(reversePids)) {
                console.log('   🎉 STRATEGY 1 WORKED!');
            } else {
                console.log('   ❌ STRATEGY 1 FAILED.');
            }
        } catch (e) { console.error(e.message); }
    }

    // STRATEGY 2: Delete & Re-add
    console.log('\n🧪 STRATEGY 2: Delete & Re-add');
    const cat2 = createCat("DEBUG_STRATEGY_2_" + new Date().getTime());
    if (cat2) {
        console.log(`   Created Cat ${cat2}`);
        // Add in NORMAL order first
        for (const pid of pids) {
            UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects.json`, {
                method: 'POST',
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`), 'Content-Type': 'application/json' },
                payload: JSON.stringify({ collect: { product_id: pid, collection_id: cat2 } })
            });
        }

        // Now "Re-add" in REVERSE order
        // 1. Delete all collects
        const collectsResp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects.json?collection_id=${cat2}`, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const collects = JSON.parse(collectsResp.getContentText());
        for (const c of collects) {
            UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects/${c.id}.json`, {
                method: 'DELETE',
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
            });
        }
        console.log('   Cleared category.');

        // 2. Add in REVERSE order
        const reversePids = [pids[2], pids[1], pids[0]];
        for (const pid of reversePids) {
            UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects.json`, {
                method: 'POST',
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`), 'Content-Type': 'application/json' },
                payload: JSON.stringify({ collect: { product_id: pid, collection_id: cat2 } })
            });
            Utilities.sleep(200); // Small delay to ensure timestamp diff
        }
        console.log('   Re-added in reverse order.');

        // Verify
        Utilities.sleep(2000);
        const urlCheck = `${credentials.baseUrl}/admin/products.json?collection_id=${cat2}`;
        const respCheck = UrlFetchApp.fetch(urlCheck, {
            method: 'GET',
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const actualIds = JSON.parse(respCheck.getContentText()).map(p => p.id);
        console.log(`   Expected: ${reversePids.join(', ')}`);
        console.log(`   Actual:   ${actualIds.join(', ')}`);

        if (JSON.stringify(actualIds) === JSON.stringify(reversePids)) {
            console.log('   🎉 STRATEGY 2 WORKED!');
        } else {
            console.log('   ❌ STRATEGY 2 FAILED.');
        }
    }
}
