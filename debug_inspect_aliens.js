function debugInspectAliens() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();

    console.log(`🔍 Inspecting Category ${categoryId}...`);

    // 1. Get Top 3 Products (The Aliens)
    const url = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=3`;
    const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const aliens = JSON.parse(response.getContentText());

    if (aliens.length === 0) {
        console.log('No products found.');
        return;
    }

    console.log(`\n👽 Found ${aliens.length} Aliens at the top:`);

    // 2. Get their Collects (to see positions)
    const collects = getAllCollectsForCategory(categoryId, credentials);

    const alienCollects = [];

    for (const alien of aliens) {
        const collect = collects.find(c => c.product_id == alien.id);
        if (collect) {
            console.log(`   - [${alien.id}] ${alien.title}`);
            console.log(`     Collect ID: ${collect.id}`);
            console.log(`     Current Position: ${collect.position}`);
            alienCollects.push(collect);
        } else {
            console.log(`   - [${alien.id}] ${alien.title} (Collect NOT found!)`);
        }
    }

    if (alienCollects.length === 0) return;

    // 3. Try to MOVE the first Alien
    const targetAlien = alienCollects[0];
    const newPos = 25;
    console.log(`\n🧪 TEST: Moving Alien "${aliens[0].title}" to Position ${newPos}...`);

    try {
        const updateUrl = `${credentials.baseUrl}/admin/collects/${targetAlien.id}.json`;
        const payload = {
            collect: {
                position: newPos
            }
        };

        UrlFetchApp.fetch(updateUrl, {
            method: 'PUT',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload)
        });

        console.log('   ✅ Move request sent.');

        Utilities.sleep(2000);

        // 4. Verify
        console.log('   👀 Verifying...');
        const verifyResponse = UrlFetchApp.fetch(updateUrl, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            }
        });
        const verifyData = JSON.parse(verifyResponse.getContentText());
        console.log(`   -> New Position in DB: ${verifyData.position}`);

        if (verifyData.position == newPos) {
            console.log('   🎉 SUCCESS! We CAN move Aliens. Manual sorting IS working.');
        } else {
            console.log('   ❌ FAILURE! Position did not change. Category is STUCK.');
        }

    } catch (e) {
        console.error(`   ❌ Error moving alien: ${e.message}`);
    }
}
