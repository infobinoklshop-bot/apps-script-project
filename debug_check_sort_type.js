function debugCheckSortType() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const categoryId = sheet.getRange('B2').getValue();

    if (!categoryId) {
        console.error('Category ID not found');
        return;
    }

    const credentials = getInsalesCredentialsSync();
    const url = `${credentials.baseUrl}/admin/collections/${categoryId}.json`;

    console.log(`🔍 Reading Category ${categoryId}...`);

    try {
        const response = UrlFetchApp.fetch(url, {
            method: 'GET',
            headers: {
                'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
                'Content-Type': 'application/json'
            }
        });
        const data = JSON.parse(response.getContentText());

        console.log(`\n📊 CURRENT SETTINGS:`);
        console.log(`   - Title: "${data.title}"`);
        console.log(`   - Sort Type ID: ${data.sort_type}`);
        console.log(`   - Is Smart?: ${data.is_smart}`);
        console.log(`   - Parent ID: ${data.parent_id}`);

        console.log(`\n💡 INSTRUCTIONS:`);
        console.log(`1. If "Sort Type ID" is NOT what we want (we suspect 0 is Manual),`);
        console.log(`   please go to InSales Admin Panel for this category.`);
        console.log(`2. Manually change the sorting to "Manual" (Ручная) or "By position".`);
        console.log(`3. Save changes in InSales.`);
        console.log(`4. Run this script again.`);
        console.log(`5. Tell the developer the new "Sort Type ID" you see here.`);

    } catch (e) {
        console.error(`❌ Failed to fetch category info: ${e.message}`);
    }
}
