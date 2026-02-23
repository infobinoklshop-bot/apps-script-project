function setupShadowTest() {
    const mainCategoryId = 9071619; // "С дальномером" (https://binokl.shop/collection/s-dalnomerom)
    const shadowHandle = "popular";  // Existing category handle

    console.log(`🚀 Linking existing category '${shadowHandle}' to main category...`);

    const credentials = getInsalesCredentialsSync();
    const headers = {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
    };

    // Link Shadow Category to Main Category
    const fieldId = 292698; // polular-cat-prod

    const updateUrl = `${credentials.baseUrl}/admin/collections/${mainCategoryId}.json`;
    const updatePayload = {
        collection: {
            field_values: [
                {
                    field_id: fieldId,
                    value: shadowHandle // Writing 'popular' here
                }
            ]
        }
    };

    const updateResp = UrlFetchApp.fetch(updateUrl, {
        method: 'PUT',
        headers: headers,
        payload: JSON.stringify(updatePayload),
        muteHttpExceptions: true
    });

    if (updateResp.getResponseCode() === 200) {
        console.log("🎉 SETUP COMPLETE!");
        console.log(`Main Category: https://binokl.shop/collection/s-dalnomerom`);
        console.log(`Linked Content: https://binokl.shop/collection/${shadowHandle}`);
        console.log("The widget should now display products from the 'Popular' category.");
    } else {
        console.log("❌ Error linking category.");
        console.log(updateResp.getContentText());
    }
}
