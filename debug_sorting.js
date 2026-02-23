/**
 * DEBUG SCRIPT FOR PRODUCT SORTING
 * 
 * This script diagnoses the product sorting issue by:
 * 1. Reading the first 5 products from the active sheet.
 * 2. Fetching their CURRENT positions in InSales.
 * 3. Attempting to update their positions.
 * 4. Verifying if the update persisted.
 */

function debugProductSorting() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const ui = SpreadsheetApp.getUi();

    console.log('🚀 [DEBUG] Starting sorting diagnosis...');

    // 1. Get Category ID
    const categoryId = sheet.getRange('B2').getValue();
    if (!categoryId) {
        ui.alert('Error', 'Could not find Category ID in cell B2.', ui.ButtonSet.OK);
        return;
    }
    console.log(`📂 Category ID: ${categoryId}`);

    // 2. Get Credentials
    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
        ui.alert('Error', 'Could not get InSales credentials.', ui.ButtonSet.OK);
        return;
    }

    // 3. Check Collection Sort Type
    const collectionUrl = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
    const collectionResponse = UrlFetchApp.fetch(collectionUrl, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const collectionData = JSON.parse(collectionResponse.getContentText());
    console.log(`⚙️ Current Sort Type: ${collectionData.sort_type} (Expected: 0 for Manual)`);

    if (collectionData.sort_type !== 0) {
        console.warn('⚠️ WARNING: Collection is NOT in Manual sort mode!');
    }

    // 4. Get Products from Sheet
    // We use calculateSheetSections logic manually or rely on it if available
    // Assuming standard layout: find "🛒 ТЕКУЩИЕ ТОВАРЫ"
    const data = sheet.getDataRange().getValues();
    let startRow = -1;
    for (let i = 0; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().includes('🛒 ТЕКУЩИЕ ТОВАРЫ')) {
            startRow = i + 3; // Header + 2 rows offset usually
            break;
        }
    }

    if (startRow === -1) {
        // Fallback to calculateSheetSections if available
        try {
            const sections = calculateSheetSections(sheet);
            startRow = sections.productsStart;
        } catch (e) {
            console.error('Could not find product section start');
            return;
        }
    }

    console.log(`📍 Products start at row: ${startRow}`);

    // Read first 5 products
    // Column E (index 4) is ID
    const productIds = [];
    for (let i = 0; i < 5; i++) {
        const row = startRow + i - 1; // 0-indexed
        const id = sheet.getRange(row + 1, 5).getValue();
        const title = sheet.getRange(row + 1, 1).getValue();
        if (id) {
            productIds.push({ id: id, title: title, targetPosition: i + 1 });
        }
    }

    console.log('📋 Products to check:', JSON.stringify(productIds));

    // 5. Fetch current positions (Collects)
    const collectsUrl = `${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}`;
    const collectsResponse = UrlFetchApp.fetch(collectsUrl, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const collects = JSON.parse(collectsResponse.getContentText());

    console.log(`🔗 Found ${collects.length} collects in InSales.`);

    // Check current positions
    productIds.forEach(p => {
        const collect = collects.find(c => c.product_id == p.id);
        if (collect) {
            console.log(`   Product ${p.id} (${p.title}): Current Position = ${collect.position}, Collect ID = ${collect.id}`);
            p.collectId = collect.id;
        } else {
            console.error(`   ❌ Product ${p.id} not found in this category!`);
        }
    });

    // 6. Attempt Update (Force Manual Sort first)
    if (collectionData.sort_type !== 0) {
        console.log('🔄 Forcing Manual Sort (trying sort_type: 0)...');
        try {
            const sortResponse = UrlFetchApp.fetch(collectionUrl, {
                method: 'PUT',
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                },
                payload: JSON.stringify({ collection: { sort_type: 0 } }),
                muteHttpExceptions: true // Prevent crash
            });

            if (sortResponse.getResponseCode() !== 200) {
                console.warn(`⚠️ Failed to set sort_type to 0. Code: ${sortResponse.getResponseCode()}`);
                console.warn(`Response: ${sortResponse.getContentText()}`);
                console.log('➡️ Proceeding with position updates anyway to test if they work with current sort_type...');
            } else {
                console.log('✅ Successfully set sort_type to 0.');
            }
        } catch (e) {
            console.error('❌ Exception setting sort_type:', e);
        }
    }

    console.log('🔄 Updating positions...');

    for (const p of productIds) {
        if (!p.collectId) continue;

        const updateUrl = `${credentials.baseUrl}/admin/collects/${p.collectId}.json`;
        const payload = { collect: { position: p.targetPosition } };

        console.log(`   Updating Product ${p.id} to position ${p.targetPosition}...`);

        const updateResponse = UrlFetchApp.fetch(updateUrl, {
            method: 'PUT',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        console.log(`   Response: ${updateResponse.getResponseCode()}`);
        if (updateResponse.getResponseCode() !== 200) {
            console.warn(`   Error: ${updateResponse.getContentText()}`);
        }
        Utilities.sleep(500); // Wait a bit
    }

    // 7. Verify Updates
    console.log('🔍 Verifying updates...');
    Utilities.sleep(2000); // Wait for propagation

    const verifyResponse = UrlFetchApp.fetch(collectsUrl, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const verifyCollects = JSON.parse(verifyResponse.getContentText());

    productIds.forEach(p => {
        const collect = verifyCollects.find(c => c.product_id == p.id);
        if (collect) {
            const status = collect.position == p.targetPosition ? '✅ OK' : '❌ FAILED';
            console.log(`   Product ${p.id}: Expected ${p.targetPosition}, Got ${collect.position} -> ${status}`);
        }
    });

    ui.alert('Debug Complete', 'Check the execution logs (View > Execution transcript) for details.', ui.ButtonSet.OK);
}
