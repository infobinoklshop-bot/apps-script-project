function debugTestControl() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();
    const sheet = ss.getActiveSheet();

    // Get Parent ID
    const currentCategoryId = sheet.getRange('B2').getValue();
    let parentId = null;
    if (currentCategoryId) {
        try {
            const resp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collections/${currentCategoryId}.json`, {
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
            });
            parentId = JSON.parse(resp.getContentText()).parent_id;
        } catch (e) { }
    }

    // Get 3 Products
    const sections = calculateSheetSections(sheet);
    const startRow = sections.productsStart;
    const pids = [];
    for (let i = 0; i < 3; i++) {
        const id = sheet.getRange(startRow + i, 5).getValue();
        if (id) pids.push(parseInt(id));
    }
    if (pids.length < 3) { console.error('Not enough products'); return; }

    console.log('🚀 Starting Control Tests...');

    // TEST 1: Sort by Name (Type 1)
    console.log('\n🧪 TEST 1: Sort by Name (Type 1)');
    const cat1Payload = {
        collection: {
            title: "DEBUG_CONTROL_NAME_" + new Date().getTime(),
            is_hidden: true,
            sort_type: "1",
            parent_id: parentId
        }
    };

    try {
        const resp1 = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collections.json`, {
            method: 'POST',
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`), 'Content-Type': 'application/json' },
            payload: JSON.stringify(cat1Payload)
        });
        const cat1 = JSON.parse(resp1.getContentText()).id;
        console.log(`   Created Cat ${cat1} (Type 1)`);

        // Add products
        for (const pid of pids) {
            UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects.json`, {
                method: 'POST',
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`), 'Content-Type': 'application/json' },
                payload: JSON.stringify({ collect: { product_id: pid, collection_id: cat1 } })
            });
        }

        Utilities.sleep(2000);
        const check1 = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/products.json?collection_id=${cat1}`, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const prods1 = JSON.parse(check1.getContentText());
        const names1 = prods1.map(p => p.title);
        console.log(`   Actual Names: ${names1.join(', ')}`);

        // Check if sorted alphabetically
        const sortedNames = [...names1].sort();
        if (JSON.stringify(names1) === JSON.stringify(sortedNames)) {
            console.log('   ✅ Sort by Name WORKS. (API is functional)');
        } else {
            console.log('   ❌ Sort by Name FAILED. (Global API Issue?)');
        }

    } catch (e) { console.error(e.message); }

    // TEST 2: Manual (Type 7) with Explicit Position
    console.log('\n🧪 TEST 2: Manual (Type 7) with Explicit Position');
    const cat2Payload = {
        collection: {
            title: "DEBUG_CONTROL_POS_" + new Date().getTime(),
            is_hidden: true,
            sort_type: "7",
            parent_id: parentId
        }
    };

    try {
        const resp2 = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collections.json`, {
            method: 'POST',
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`), 'Content-Type': 'application/json' },
            payload: JSON.stringify(cat2Payload)
        });
        const cat2 = JSON.parse(resp2.getContentText()).id;
        console.log(`   Created Cat ${cat2} (Type 7)`);

        // Add products with REVERSE positions (1, 2, 3 assigned to 3rd, 2nd, 1st product)
        // We want P3 at Pos 1, P2 at Pos 2, P1 at Pos 3
        const targetOrder = [pids[2], pids[1], pids[0]];

        for (let i = 0; i < 3; i++) {
            const pid = targetOrder[i];
            const pos = i + 1;
            UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects.json`, {
                method: 'POST',
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`), 'Content-Type': 'application/json' },
                payload: JSON.stringify({
                    collect: {
                        product_id: pid,
                        collection_id: cat2,
                        position: pos
                    }
                })
            });
            console.log(`   Added ${pid} at Pos ${pos}`);
        }

        Utilities.sleep(2000);
        const check2 = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/products.json?collection_id=${cat2}`, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const prods2 = JSON.parse(check2.getContentText());
        const ids2 = prods2.map(p => p.id);

        console.log(`   Expected: ${targetOrder.join(', ')}`);
        console.log(`   Actual:   ${ids2.join(', ')}`);

        if (JSON.stringify(ids2) === JSON.stringify(targetOrder)) {
            console.log('   🎉 SUCCESS! Explicit position works.');
        } else {
            console.log('   ❌ FAILED. Explicit position ignored.');
        }

    } catch (e) { console.error(e.message); }
}
