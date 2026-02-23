function debugSortType() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();
    const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;

    console.log(`🔍 Inspecting Category ${categoryId}...`);

    // 1. Inspect Category
    try {
        const response = UrlFetchApp.fetch(url, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            }
        });
        const data = JSON.parse(response.getContentText());
        console.log(`ℹ️ Category Info:`);
        console.log(`   - Title: ${data.title}`);
        console.log(`   - Is Smart? ${data.is_smart}`);
        console.log(`   - Current Sort Type: ${data.sort_type}`);
        console.log(`   - Parent ID: ${data.parent_id}`);
    } catch (e) {
        console.error(`❌ Failed to fetch category info: ${e.message}`);
    }

    console.log(`\n🧪 Testing sort_type values (0-12)...`);

    for (let val = 0; val <= 12; val++) {
        try {
            const payload = {
                collection: {
                    sort_type: val
                }
            };

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
                const check = JSON.parse(text);
                console.log(`✅ SUCCESS! Value ${val} accepted. DB set to: ${check.sort_type}`);
            } else {
                // console.log(`❌ Failed value ${val}: ${code}`); // Reduce noise
            }

        } catch (e) {
            console.error(`   Exception: ${e.message}`);
        }

        Utilities.sleep(200);
    }

    console.log('\n🏁 Test completed. Please send these logs to the developer.');
}
