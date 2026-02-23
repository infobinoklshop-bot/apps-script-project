/**
 * Debug script to inspect InSales Collection settings
 */
function debugInspectCollection() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('❌ No category ID found in B2');
        return;
    }

    const credentials = getInsalesCredentialsSync();
    if (!credentials) {
        console.error('❌ No credentials');
        return;
    }

    console.log(`🔍 Inspecting Collection ID: ${categoryId}`);

    // 1. Fetch Collection Details
    const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
    try {
        const response = UrlFetchApp.fetch(url, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
        });

        if (response.getResponseCode() === 200) {
            const data = JSON.parse(response.getContentText());
            console.log('✅ Collection Data:', JSON.stringify(data, null, 2));

            // Check for sorting fields
            if (data.sort_type) {
                console.log(`👉 Sort Type: ${data.sort_type}`);
            } else {
                console.log('⚠️ No sort_type field found');
            }

        } else {
            console.error(`❌ Error fetching collection: ${response.getResponseCode()} ${response.getContentText()}`);
        }

    } catch (e) {
        console.error('❌ Exception:', e);
    }
}
