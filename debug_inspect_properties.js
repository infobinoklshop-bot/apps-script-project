/**
 * Debug script to inspect properties and characteristics.
 * Helps to identify the correct Property ID for "Brand".
 */
function debugInspectProperties() {
    const context = 'debugInspectProperties';
    const credentials = getInsalesCredentialsSync_DEBUG_(); // Use local or mock
    if (!credentials) return;

    try {
        // 1. Fetch ALL properties to see their IDs and Permalinks
        Logger.log('--- Fetching All Properties ---');
        const propsUrl = `${credentials.baseUrl}/admin/properties.json?per_page=250`; // Get first 250 properties
        const propsOptions = {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            },
            muteHttpExceptions: true
        };
        const propsResponse = UrlFetchApp.fetch(propsUrl, propsOptions);
        const properties = JSON.parse(propsResponse.getContentText());

        // Create a map for easy lookup
        const propMap = {};
        properties.forEach(p => {
            propMap[p.id] = { title: p.title, permalink: p.permalink };
            // Log interesting ones
            if (p.title.toLowerCase().includes('бренд') || p.title.toLowerCase().includes('производитель') || p.permalink.includes('brand') || p.permalink.includes('vendor')) {
                Logger.log(`[Property] ID: ${p.id}, Title: "${p.title}", Permalink: "${p.permalink}"`);
            }
        });

        // 2. Fetch a few products from the target category to see their characteristics
        Logger.log('\n--- Fetching Sample Products ---');
        const CATEGORY_ID = 29773031;
        const prodUrl = `${credentials.baseUrl}/admin/collections/${CATEGORY_ID}/products.json?per_page=5&fields=id,title,characteristics`;
        const prodResponse = UrlFetchApp.fetch(prodUrl, propsOptions);
        const products = JSON.parse(prodResponse.getContentText());

        products.forEach(p => {
            Logger.log(`\nProduct [${p.id}] ${p.title}`);
            if (p.characteristics) {
                p.characteristics.forEach(c => {
                    const propInfo = propMap[c.property_id] || { title: 'UNKNOWN', permalink: '???' };
                    Logger.log(`   - Char: "${c.title}" (PropID: ${c.property_id} -> ${propInfo.title} / ${propInfo.permalink})`);
                });
            } else {
                Logger.log('   No characteristics.');
            }
        });

    } catch (e) {
        Logger.log(`Error: ${e.message}`);
    }
}

// Temporary Copy of Credentials for Debug file (since it can't access other files easily if run standalone in some contexts, 
// though in Apps Script IDE it shares scope, it's safer to rely on global if available or reuse).
// Assuming getInsalesCredentialsSync is available globally.
function getInsalesCredentialsSync_DEBUG_() {
    if (typeof getInsalesCredentialsSync === 'function') {
        return getInsalesCredentialsSync();
    }
    // Fallback if running strangely
    const props = PropertiesService.getScriptProperties();
    return {
        apiKey: props.getProperty('INSALES_API_KEY'),
        password: props.getProperty('INSALES_PASSWORD'),
        baseUrl: `https://${props.getProperty('INSALES_SHOP_URL')}`
    };
}
