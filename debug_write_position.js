function debugWritePosition() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentials = getInsalesCredentialsSync();
    const sheet = ss.getActiveSheet();
    const categoryId = 9175197; // Hardcoded confirmed ID

    console.log(`🚀 Starting Write Position Test on Category ${categoryId}...`);

    // 1. Find the current King (Position 1) - likely БПЦ7
    // We know БПЦ7 is 110260922. Let's check its position dynamically to be sure.
    const kingId = 110260922;
    let kingCollectId = null;

    try {
        const resp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&product_id=${kingId}`, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const collects = JSON.parse(resp.getContentText());
        if (collects.length > 0) {
            kingCollectId = collects[0].id;
            console.log(`👑 Current King: ID ${kingId} | Position: ${collects[0].position}`);
        }
    } catch (e) { console.error(e); }

    // 2. Pick a Challenger from the Sheet (e.g. Row 160, Column E)
    // Just picking a random product ID from the sheet to be the new #1
    const sections = calculateSheetSections(sheet);
    const challengerId = sheet.getRange(sections.productsStart + 4, 5).getValue(); // 5th product
    console.log(`⚔️ Challenger: ID ${challengerId}`);

    if (!challengerId) { console.error('No challenger found'); return; }

    // 3. Find Challenger's Collect ID
    let challengerCollectId = null;
    try {
        const resp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects.json?collection_id=${categoryId}&product_id=${challengerId}`, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const collects = JSON.parse(resp.getContentText());
        if (collects.length > 0) {
            challengerCollectId = collects[0].id;
            console.log(`   Challenger Current Position: ${collects[0].position} (Collect ID: ${challengerCollectId})`);
        } else {
            console.error('Challenger not in category!');
            return;
        }
    } catch (e) { console.error(e); return; }

    // 4. EXECUTE THE COUP: Set Challenger to Position 1
    console.log(`\n⚡ Updating Challenger to Position 1...`);
    try {
        const payload = { collect: { position: 1 } };
        const urlUpdate = `${credentials.baseUrl}/admin/collects/${challengerCollectId}.json`;
        const respUpdate = UrlFetchApp.fetch(urlUpdate, {
            method: 'PUT',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            payload: JSON.stringify(payload)
        });
        console.log(`   Update Response Code: ${respUpdate.getResponseCode()}`);
    } catch (e) { console.error(`   Update Failed: ${e.message}`); }

    // 5. Verify Results
    console.log(`\n🔍 Verifying New Positions...`);
    Utilities.sleep(1000);

    // Check Challenger
    try {
        const resp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects/${challengerCollectId}.json`, {
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
        });
        const c = JSON.parse(resp.getContentText());
        console.log(`   ⚔️ Challenger New Position: ${c.position} (Expected: 1)`);
    } catch (e) { }

    // Check Old King
    if (kingCollectId) {
        try {
            const resp = UrlFetchApp.fetch(`${credentials.baseUrl}/admin/collects/${kingCollectId}.json`, {
                headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`) }
            });
            const c = JSON.parse(resp.getContentText());
            console.log(`   👑 Old King New Position: ${c.position} (Expected: >1)`);
        } catch (e) { }
    }
}
