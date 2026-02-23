function findAndInspectCategory() {
    const searchTerms = ["Монокуляры с дальномером", "С дальномером"];
    console.log(`🔍 Searching for categories matching: ${searchTerms.join(', ')}...`);

    const credentials = getInsalesCredentialsSync();
    let page = 1;
    let foundCount = 0;

    while (true) {
        console.log(`📄 Fetching page ${page}...`);
        const url = `${credentials.baseUrl}/admin/collections.json?per_page=250&page=${page}`;

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
            console.log(`❌ Error fetching page ${page}: ${response.getResponseCode()}`);
            break;
        }

        const collections = JSON.parse(response.getContentText());

        if (collections.length === 0) {
            console.log("🏁 End of categories list.");
            break;
        }

        // Search in this batch
        for (const c of collections) {
            if (c.permalink === 's-dalnomerom') {
                console.log(`\n✅ MATCH FOUND!`);
                console.log(`🆔 ID: ${c.id}`);
                console.log(`🏷️ Title: "${c.title}"`);
                console.log(`🔗 Permalink: ${c.permalink}`);
                console.log(`📄 Template: ${c.template || 'null'}`);
                console.log(`📄 Custom Template: ${c.custom_template || 'null'}`);

                foundCount++;
            }
        }

        page++;
        Utilities.sleep(200); // Be nice to API
    }

    if (foundCount === 0) {
        console.log("\n❌ No categories found matching the search terms.");
    } else {
        console.log(`\n✅ Found ${foundCount} potential matches.`);
    }
}
