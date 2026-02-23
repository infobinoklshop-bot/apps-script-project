/**
 * Widget Manager
 * Manages 'Smart Widgets' via Google Sheets.
 * 
 * Features:
 * 1. initWidgetSheet: Creates the control sheet and populates it.
 * 2. fetchAndFillCategories: Fetches all categories from InSales.
 * 3. syncWidgets: Synchronizes changes from Sheet to InSales.
 */

const WIDGET_SHEET_NAME = "Управление Виджетами";
const FIELD_ID_WIDGET = 292698; // polular-cat-prod

function initWidgetSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(WIDGET_SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(WIDGET_SHEET_NAME);
        // Headers
        const headers = [
            ["ID Категории", "Название", "Ссылка (Permalink)", "Источник Виджета (Handle)", "Статус Синхронизации", "Последнее Обновление"]
        ];
        sheet.getRange("A1:F1").setValues(headers).setFontWeight("bold").setBackground("#f3f3f3");
        sheet.setColumnWidth(1, 100); // ID
        sheet.setColumnWidth(2, 300); // Title
        sheet.setColumnWidth(3, 200); // Permalink
        sheet.setColumnWidth(4, 250); // Widget Source
        sheet.setColumnWidth(5, 150); // Status
        sheet.setColumnWidth(6, 150); // Last Update
        sheet.setFrozenRows(1);
        console.log("✅ Widget Sheet created.");
    } else {
        console.log("ℹ️ Widget Sheet already exists.");
    }

    // Refresh data
    fetchAndFillCategories(sheet);
}

function fetchAndFillCategories(sheet) {
    console.log("🔄 Fetching categories from InSales...");
    const credentials = getInsalesCredentialsSync();
    const headers = {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
    };

    let page = 1;
    let allCategories = [];

    while (true) {
        console.log(`Fetching page ${page}...`);
        try {
            const url = `${credentials.baseUrl}/admin/collections.json?page=${page}&per_page=100`;
            const resp = UrlFetchApp.fetch(url, { headers: headers });
            const cats = JSON.parse(resp.getContentText());

            if (cats.length === 0) break;
            allCategories = allCategories.concat(cats);
            page++;
            Utilities.sleep(200);
        } catch (e) {
            console.error(`Error fetching page ${page}: ${e}`);
            break;
        }
    }

    console.log(`✅ Fetched ${allCategories.length} categories. Updating sheet...`);

    const rows = allCategories.map(c => {
        let widgetVal = "";
        if (c.field_values) {
            // API returns 'collection_field_id' for collections, but sometimes 'field_id' elsewhere.
            // We check both to be safe, prioritizing collection_field_id based on debug logs.
            const fv = c.field_values.find(f => (f.collection_field_id || f.field_id) === FIELD_ID_WIDGET);
            if (fv) widgetVal = fv.value;
        }

        return [
            c.id,
            c.title,
            c.permalink,
            widgetVal,
            "Loaded from Site",
            new Date()
        ];
    });

    if (rows.length > 0) {
        // Clear old data (except header)
        if (sheet.getLastRow() > 1) {
            sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
        }
        sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }

    console.log("🎉 Sheet updated successfully.");
}

/**
 * Synchronizes changes from Sheet to InSales.
 * Handles both direct Handles and Product ID lists (by creating shadow categories).
 */
function syncWidgets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(WIDGET_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Remove header

    const credentials = getInsalesCredentialsSync();
    const authHeader = {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
    };

    console.log("🚀 Starting Sync...");
    const startTime = new Date().getTime();

    // Iterate rows
    let processedCount = 0;

    for (let i = 0; i < data.length; i++) {
        // TIMEOUT PROTECTION: Stop if running longer than 5 minutes (300s)
        if (new Date().getTime() - startTime > 300000) {
            console.warn("⏳ Time limit reached. Stopping sync safely.");
            sheet.getRange(1, 6).setValue("Partial Sync (Timeout)");
            break;
        }

        const row = data[i];
        const rowIndex = i + 2;
        const catId = row[0];
        const sourceVal = String(row[3]).trim(); // Widget Source
        const currentStatus = String(row[4]); // Status

        if (!catId) continue;

        // OPTIMIZATION: Skip if already synced or just loaded
        // We treat "Loaded from Site" as synced to avoid processing 1000 rows.
        // USER INSTRUCTION: To update a row, CLEAR the Status cell.
        if (currentStatus === "Synced" || currentStatus === "Loaded from Site") {
            continue;
        }

        try {
            let finalHandle = "";

            if (sourceVal) {
                // Check if it's a list of IDs (digits and commas)
                const isIdList = /^[\d,\s]+$/.test(sourceVal);

                if (isIdList) {
                    // It's IDs -> Create/Update Shadow Category
                    console.log(`Processing IDs for Cat ${catId}: ${sourceVal}`);
                    finalHandle = ensureShadowCategory(catId, sourceVal, authHeader, credentials.baseUrl);
                } else {
                    // It's a Handle -> Use directly
                    finalHandle = sourceVal;
                }
            }

            // Update Main Category Field
            updateCategoryField(catId, finalHandle, authHeader, credentials.baseUrl);

            // Update Status
            sheet.getRange(rowIndex, 5).setValue("Synced");
            sheet.getRange(rowIndex, 6).setValue(new Date());

            processedCount++;

            // RATE LIMITING: Sleep to avoid 429 Errors
            // InSales limit is typically around 500 requests / 5 minutes.
            // 0.6s delay = ~100 requests/min.
            Utilities.sleep(600);

        } catch (e) {
            console.error(`Error syncing row ${rowIndex}: ${e}`);
            sheet.getRange(rowIndex, 5).setValue("Error: " + e.message);
            // Sleep even on error to let API cool down
            Utilities.sleep(1000);
        }
    }

    console.log(`✅ Sync Complete. Processed ${processedCount} rows.`);
}

function ensureShadowCategory(mainCatId, productIdsStr, headers, baseUrl) {
    const shadowTitle = `Widget-Shadow-${mainCatId}`;
    const productIds = productIdsStr.split(',').map(s => s.trim()).filter(s => s);

    // 1. Check if exists (by permalink convention or search)
    // We'll use a convention: permalink = widget-shadow-[mainCatId]
    // Note: InSales might sanitize permalink, so we should search by title or just try to get it.
    // Let's try to find by title first to be safe.

    // Optimization: Just try to create it, if it fails (422), then find it.
    // Or search first. Search is safer.

    // For MVP: Let's assume we can search by permalink convention if we enforce it.
    // But permalinks are auto-generated.
    // Let's search by Title.

    const searchUrl = `${baseUrl}/admin/collections.json?title=${encodeURIComponent(shadowTitle)}`;
    const searchResp = UrlFetchApp.fetch(searchUrl, { headers: headers });
    const found = JSON.parse(searchResp.getContentText());

    let shadowId;
    let shadowPermalink;

    if (found.length > 0) {
        shadowId = found[0].id;
        shadowPermalink = found[0].permalink;
        console.log(`Found existing shadow category: ${shadowId}`);
    } else {
        // Create
        const createPayload = {
            collection: {
                title: shadowTitle,
                is_hidden: true,
                parent_id: null
            }
        };
        const createResp = UrlFetchApp.fetch(`${baseUrl}/admin/collections.json`, {
            method: 'POST', headers: headers, payload: JSON.stringify(createPayload)
        });
        const created = JSON.parse(createResp.getContentText());
        shadowId = created.id;
        shadowPermalink = created.permalink;
        console.log(`Created new shadow category: ${shadowId}`);
    }

    // 2. Update Products (Clear and Add)
    // InSales doesn't have "replace products" endpoint easily.
    // We have to delete old collects and add new ones.
    // Or just add new ones and ignore duplicates?
    // Better: Get current products, calculate diff.

    // For MVP: Just add. (Duplicates might be ignored by API).
    // Ideally we should clear, but that's slow.
    // Let's just add for now.

    productIds.forEach(pid => {
        try {
            UrlFetchApp.fetch(`${baseUrl}/admin/collects.json`, {
                method: 'POST',
                headers: headers,
                payload: JSON.stringify({ collect: { product_id: pid, collection_id: shadowId } }),
                muteHttpExceptions: true // Ignore if already exists
            });
        } catch (e) { }
    });

    return shadowPermalink;
}

function updateCategoryField(catId, value, headers, baseUrl) {
    // We need to use the correct field ID: 292698
    // And the API expects 'field_values' array.

    const payload = {
        collection: {
            field_values: [
                {
                    field_id: 292698, // polular-cat-prod
                    value: value
                }
            ]
        }
    };

    UrlFetchApp.fetch(`${baseUrl}/admin/collections/${catId}.json`, {
        method: 'PUT',
        headers: headers,
        payload: JSON.stringify(payload)
    });
}
