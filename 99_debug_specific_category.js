function debugCategoryFields() {
    const categoryId = 9071619;
    const credentials = getInsalesCredentialsSync();
    const headers = {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
    };

    console.log(`🔍 Inspecting Category ${categoryId}...`);

    // 1. Fetch via SHOW endpoint (Single resource)
    const showUrl = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;
    const showResp = UrlFetchApp.fetch(showUrl, { headers: headers });
    const showData = JSON.parse(showResp.getContentText());

    console.log("--- SHOW Endpoint Data ---");
    if (showData.field_values) {
        console.log(`Field Values Found: ${showData.field_values.length}`);
        // Log keys of the first item to see structure
        if (showData.field_values.length > 0) {
            console.log("Keys of first item:", Object.keys(showData.field_values[0]));
        }

        showData.field_values.forEach(fv => {
            // Try to find any property that looks like an ID
            const id = fv.field_id || fv.collection_field_id || fv.id;
            console.log(`ID: ${id} | Value: ${fv.value}`);
        });
    } else {
        console.log("No field_values in SHOW response.");
    }

    // 2. Fetch via INDEX endpoint (List resource) - filtering by ID to simulate list
    // Note: InSales doesn't support filtering by ID in list easily, but let's try standard list
    // and see if we can find it in the first page if we were lucky, or just check the structure of one item.
    // Actually, let's just check the SHOW response first, as that's the source of truth. 
    // If SHOW has it, but LIST doesn't, then we know the issue.

    console.log("\n--- Checking Target Field ---");
    const targetFieldId = 292698;
    const target = showData.field_values ? showData.field_values.find(f => f.field_id == targetFieldId) : null;

    if (target) {
        console.log(`✅ Target Field Found! Value: "${target.value}"`);
    } else {
        console.log(`❌ Target Field ${targetFieldId} NOT found in category.`);
    }
}
