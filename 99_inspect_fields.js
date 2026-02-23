function inspectCollectionFields() {
    console.log("🕵️ Fetching all Collection Fields definitions...");

    const credentials = getInsalesCredentialsSync();
    const url = `${credentials.baseUrl}/admin/collection_fields.json`;

    const options = {
        method: 'GET',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
            'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() !== 200) {
        console.log(`❌ Error fetching fields: ${response.getResponseCode()}`);
        console.log(response.getContentText());
        return;
    }

    const fields = JSON.parse(response.getContentText());

    console.log(`\n✅ Found ${fields.length} fields. Listing them:`);
    console.log("ID | Title | Handle | Type");
    console.log("-".repeat(50));

    fields.forEach(f => {
        console.log(`${f.id} | ${f.title} | ${f.handle} | ${f.kind_id}`);
    });

    console.log("-".repeat(50));
    console.log("Look for 'Набор виджетов' or similar in the Title column.");
}
