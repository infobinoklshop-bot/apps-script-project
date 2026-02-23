function debugProductFetch() {
    const context = "DEBUG PRODUCT FETCH";
    console.log('🚀 Starting debug fetch...');

    try {
        // 1. Get credentials
        const credentials = getInsalesCredentialsSync();
        if (!credentials) {
            console.error('❌ No credentials found');
            return;
        }

        // 2. Construct URL exactly as in loadAllProductsQuick
        // const url = `${credentials.baseUrl}/admin/products.json?per_page=${perPage}&page=${page}&fields=id,collections_ids,variants&variant_fields=quantity`;
        const url = `${credentials.baseUrl}/admin/products.json?per_page=5&page=1&fields=id,collections_ids,variants`;

        console.log(`📡 Fetching URL: ${url}`);

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

        if (response.getResponseCode() !== 200) {
            console.error(`❌ API Error: ${response.getContentText()}`);
            return;
        }

        const products = JSON.parse(response.getContentText());
        console.log(`📦 Fetched ${products.length} products`);

        if (products.length > 0) {
            const p = products[0];
            console.log('--------------------------------------------------');
            console.log('🔍 FIRST PRODUCT STRUCTURE:');
            console.log(JSON.stringify(p, null, 2));
            console.log('--------------------------------------------------');

            console.log(`ID: ${p.id}`);
            console.log(`Collections: ${JSON.stringify(p.collections_ids)}`);

            if (p.variants) {
                console.log(`Variants count: ${p.variants.length}`);
                if (p.variants.length > 0) {
                    console.log(`First Variant Quantity: ${p.variants[0].quantity}`);
                    console.log(`First Variant Keys: ${Object.keys(p.variants[0]).join(', ')}`);
                }
            } else {
                console.error('❌ NO VARIANTS FIELD FOUND!');
            }
        } else {
            console.warn('⚠️ No products returned');
        }

    } catch (e) {
        console.error('❌ EXCEPTION:', e);
    }
}
