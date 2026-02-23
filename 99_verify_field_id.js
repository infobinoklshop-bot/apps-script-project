function verifyFieldId() {
    const credentials = getInsalesCredentialsSync();
    const headers = {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
    };

    console.log("Fetching collection fields...");
    const url = `${credentials.baseUrl}/admin/collection_fields.json`;
    const resp = UrlFetchApp.fetch(url, { headers: headers });
    const fields = JSON.parse(resp.getContentText());

    console.log(`Found ${fields.length} fields.`);
    fields.forEach(f => {
        console.log(`ID: ${f.id} | Handle: ${f.handle} | Title: ${f.title}`);
    });

    const target = fields.find(f => f.handle === 'polular-cat-prod');
    if (target) {
        console.log(`\n✅ TARGET FOUND: ${target.handle} = ID ${target.id}`);
    } else {
        console.log(`\n❌ Target handle 'polular-cat-prod' NOT found.`);
    }
}
