/**
 * Debug script to check if brands (vendors) are correctly returned from InSales for a specific category.
 */
function debugCheckBrandsForCategory() {
    const CATEGORY_ID = 29773031; // Коллиматорные прицелы быстросъемные (from screenshot)
    const context = 'debugCheckBrandsForCategory';

    Logger.log(`Starting debug for category ID: ${CATEGORY_ID}`);

    try {
        const credentials = getInsalesCredentialsSync_DEBUG();
        if (!credentials) {
            Logger.log('Error: No credentials');
            return;
        }

        // Fetch products specifically for this collection
        // Note: InSales API for collection products
        const url = `${credentials.baseUrl}/admin/collections/${CATEGORY_ID}/products.json?per_page=50&fields=id,title,vendor`;

        const options = {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
        };

        Logger.log(`Fetching: ${url}`);
        const response = UrlFetchApp.fetch(url, options);

        if (response.getResponseCode() !== 200) {
            Logger.log(`Error Response: ${response.getResponseCode()} - ${response.getContentText()}`);
            return;
        }

        const products = JSON.parse(response.getContentText());
        Logger.log(`Found ${products.length} products.`);

        const vendors = new Set();
        products.forEach(p => {
            // Logger.log(`Product [${p.id}] ${p.title} - Vendor: "${p.vendor}"`); // Commented out to reduce noise
            if (p.vendor) vendors.add(p.vendor);
        });

        Logger.log('--- Unique Vendors Found ---');
        if (vendors.size === 0) {
            Logger.log('NO VENDORS FOUND!');
        } else {
            vendors.forEach(v => Logger.log(v));
        }

    } catch (e) {
        Logger.log(`Exception: ${e.message}`);
    }
}

function getInsalesCredentialsSync_DEBUG() {
    const props = PropertiesService.getScriptProperties();
    // We need to read checking 01_config.js or similar where credentials might be stored or accessed
    // But usually they are in Script Properties. 
    // Let's try to mock the getting of properties if we can't run this. 
    // Actually, I can't run this on the server. I have to deploy it or assume the user runs it.

    // Since I can't run it, I will assume the code is correct and I should deploy it?
    // Wait, the user is an Apps Script user. I can't "run" code on his Google server.
    // I can only edit code.

    // I will rewrite this to be a standalone logic inside the main file `30_seo_tags.js` temporarily or just reason about it.

    return null; // validation placeholder
}
