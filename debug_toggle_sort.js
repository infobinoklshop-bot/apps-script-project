function debugToggleSort() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();
    const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;

    console.log(`🔍 Testing Sort Toggle for Category ${categoryId}...`);

    // Helper to set sort
    function setSort(val, name) {
        console.log(`\n👉 Attempting to set sort_type to ${val} (${name})...`);
        try {
            const payload = { collection: { sort_type: val } };
            const response = UrlFetchApp.fetch(url, {
                method: 'PUT',
                headers: {
                    'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                    'Content-Type': 'application/json'
                },
                payload: JSON.stringify(payload),
                muteHttpExceptions: true
            });

            const code = response.getResponseCode();
            const text = response.getContentText();

            if (code === 200) {
                const data = JSON.parse(text);
                console.log(`   ✅ Success! New sort_type in DB: ${data.sort_type}`);
                return true;
            } else {
                console.error(`   ❌ Failed (Code ${code})`);
                console.error(`   Response: ${text}`);
                return false;
            }
        } catch (e) {
            console.error(`   ❌ Exception: ${e.message}`);
            return false;
        }
    }

    // 1. Try to set to PRICE (1)
    // This usually works.
    setSort(1, "Price");

    Utilities.sleep(1000);

    // 2. Try to set to MANUAL (0)
    // This is what we want.
    setSort(0, "Manual 0");

    Utilities.sleep(1000);

    // 3. Try to set to MANUAL string ("manual")
    setSort("manual", "Manual String");

    Utilities.sleep(1000);

    // 4. Try to set to 7 (Current)
    setSort(7, "Type 7");
}
