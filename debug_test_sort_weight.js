function debugTestSortWeight() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();
    console.log(`🔍 Testing Sort Weight for Category ${categoryId}...`);

    // 1. Find a test subject (The first product)
    const urlList = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=1`;
    const responseList = UrlFetchApp.fetch(urlList, {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        }
    });
    const products = JSON.parse(responseList.getContentText());

    if (products.length === 0) {
        console.error('No products in category.');
        return;
    }

    const subject = products[0];
    console.log(`🧪 Test Subject: [${subject.id}] ${subject.title}`);
    console.log(`   Current Sort Weight: ${subject.sort_weight}`);

    // Helper to check position in list
    function getActualRank(prodId) {
        const url = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=50`;
        const resp = UrlFetchApp.fetch(url, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            }
        });
        const list = JSON.parse(resp.getContentText());
        const idx = list.findIndex(p => p.id == prodId);
        return idx === -1 ? '>50' : idx + 1;
    }

    // TEST: Change Sort Weight
    // Try a very high value (to move it to top or bottom?)
    // Usually higher weight = higher position (or lower, depends on implementation)
    const newWeight = 9999;
    console.log(`\n👉 Setting sort_weight to ${newWeight}...`);

    try {
        const updateUrl = `${credentials.baseUrl}/admin/products/${subject.id}.json`;
        const payload = {
            product: {
                sort_weight: newWeight
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

        console.log('   ✅ Update request sent.');
        Utilities.sleep(2000);

        // Verify change in DB
        const verifyResp = UrlFetchApp.fetch(updateUrl, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            }
        });
        const verifyData = JSON.parse(verifyResp.getContentText());
        console.log(`   -> New sort_weight in DB: ${verifyData.sort_weight}`);

        // Check Rank
        const rank = getActualRank(subject.id);
        console.log(`   Actual Rank in List: ${rank}`);

        if (rank !== 1) {
            console.log('   ℹ️ Rank changed? (It was 1). If it moved down, sort_weight works!');
        } else {
            console.log('   ℹ️ Rank stayed at 1. Trying negative weight...');

            // Try negative
            const negPayload = { product: { sort_weight: -9999 } };
            UrlFetchApp.fetch(updateUrl, {
                method: 'PUT',
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                },
                payload: JSON.stringify(negPayload)
            });
            Utilities.sleep(2000);
            const rank2 = getActualRank(subject.id);
            console.log(`   Actual Rank with -9999: ${rank2}`);
        }

    } catch (e) {
        console.error(`   ❌ Error updating product: ${e.message}`);
    }
}
