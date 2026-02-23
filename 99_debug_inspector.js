function debugInspectCategoryRaw() {
    const categoryId = 9071624; // ID from CLAUDE.md example (Театральные бинокли) or any valid ID
    // Or better, let's use the active sheet's category if possible, but hardcoding is safer for a quick test.
    // Let's try to get the ID from the currently active sheet if it's a category sheet.

    const sheet = SpreadsheetApp.getActiveSheet();
    let idToFetch = categoryId;

    if (sheet.getName().startsWith('Категория — ')) {
        idToFetch = sheet.getRange('B2').getValue();
    }

    console.log(`🔍 Inspecting Category ID: ${idToFetch}`);

    const credentials = getInsalesCredentialsSync();
    const url = `${credentials.baseUrl}/admin/collections/${idToFetch}.json`;

    const options = {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    console.log(`Response Code: ${response.getResponseCode()}`);
    console.log(`Response Body: ${response.getContentText()}`);
}
