function debugTestSortBehavior() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();
    console.log(`🔍 Testing Sort Behavior for Category ${categoryId}...`);

    // 1. Find a test subject (The first product in the current list)
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

    // 2. Get its collect ID
    const collects = getAllCollectsForCategory(categoryId, credentials);
    const collect = collects.find(c => c.product_id == subject.id);

    if (!collect) {
        console.error('Collect not found.');
        return;
    }

    console.log(`   Collect ID: ${collect.id}`);

    // Helper to check position in list
    function getActualRank(prodId) {
        const url = `${credentials.baseUrl}/admin/products.json?collection_id=${categoryId}&per_page=50`; // Check top 50
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

    // TEST A: Move to Position 1
    console.log('\n👉 Moving to Position 1...');
    updateCollectPosition(collect.id, 1, credentials);
    Utilities.sleep(2000);
    let rank = getActualRank(subject.id);
    console.log(`   Actual Rank in List: ${rank}`);

    if (rank !== 1) {
        console.error('   ❌ FAILED. Set to 1, but appeared at ' + rank);
        console.error('   CONCLUSION: Sort Type 7 ignores manual positions.');
        return;
    } else {
        console.log('   ✅ OK. Appeared at Pos 1.');
    }

    // TEST B: Move to Position 10
    console.log('\n👉 Moving to Position 10...');
    updateCollectPosition(collect.id, 10, credentials);
    Utilities.sleep(2000);
    rank = getActualRank(subject.id);
    console.log(`   Actual Rank in List: ${rank}`);

    if (rank === 1) {
        console.error('   ❌ FAILED. Moved to 10, but stayed at 1.');
        console.error('   CONCLUSION: Sort Type 7 ignores manual positions (stuck at top).');
    } else if (rank === 10) {
        console.log('   ✅ SUCCESS! Moved to 10.');
        console.log('   CONCLUSION: Manual sorting IS working correctly.');
    } else {
        console.log(`   ⚠️ Partial Success? Set to 10, appeared at ${rank}. (Maybe hidden items?)`);
        console.log('   CONCLUSION: Manual sorting is active but positions are shifted.');
    }

    // Restore to 1 just in case
    updateCollectPosition(collect.id, 1, credentials);
}
