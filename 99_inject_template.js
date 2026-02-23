function injectSmartWidgetCode() {
    const templateId = "69f095cb7550ea5b91d5"; // ID from the screenshot
    console.log(`🤖 Auto-Injection initiated for Template ID: ${templateId}`);

    const credentials = getInsalesCredentialsSync();
    const headers = {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
        'Content-Type': 'application/json'
    };

    // 1. Find Active Theme
    const themesUrl = `${credentials.baseUrl}/admin/themes.json`;
    const themesResp = UrlFetchApp.fetch(themesUrl, { method: 'GET', headers: headers, muteHttpExceptions: true });
    const themes = JSON.parse(themesResp.getContentText());
    const activeTheme = themes.find(t => t.is_published);

    if (!activeTheme) {
        console.log("❌ Could not find active theme.");
        return;
    }
    console.log(`✅ Active Theme: ${activeTheme.title} (ID: ${activeTheme.id})`);

    // 2. Find the Asset (Template File)
    const assetsUrl = `${credentials.baseUrl}/admin/themes/${activeTheme.id}/assets.json`;
    const assetsResp = UrlFetchApp.fetch(assetsUrl, { method: 'GET', headers: headers, muteHttpExceptions: true });
    const assets = JSON.parse(assetsResp.getContentText());

    // Look for file containing the ID
    const asset = assets.find(a => a.name.includes(templateId) || a.human_name.includes(templateId));

    if (!asset) {
        console.log("❌ Template file not found in assets list.");
        console.log("Listing some collection templates for debug:");
        assets.filter(a => a.name.includes('collection')).slice(0, 5).forEach(a => console.log(`- ${a.name}`));
        return;
    }

    console.log(`✅ Found Asset: ${asset.name} (ID: ${asset.id})`);

    // 3. Read Asset Content
    // Endpoint: /admin/themes/{theme_id}/assets.json?asset_id={asset_id} OR just get by name logic if needed
    // Usually GET /admin/themes/{theme_id}/assets.json?key={asset.key} or similar.
    // InSales API v1 for assets is a bit tricky, let's try getting by asset ID if available or just the list content if it was included.
    // The list usually doesn't include content. We need to fetch specific asset.
    // Try: GET /admin/themes/:theme_id/assets/content?key=:key

    // Let's try fetching the specific asset content
    // Note: InSales API might use /admin/themes/{theme_id}/assets/{asset_id}.json
    const assetUrl = `${credentials.baseUrl}/admin/themes/${activeTheme.id}/assets/${asset.id}.json`;
    const assetContentResp = UrlFetchApp.fetch(assetUrl, { method: 'GET', headers: headers, muteHttpExceptions: true });
    const assetData = JSON.parse(assetContentResp.getContentText());

    let content = assetData.content || assetData.asset.content;

    if (!content) {
        console.log("❌ Could not read asset content.");
        return;
    }

    console.log(`📄 Read ${content.length} chars.`);

    // 4. Inject Code
    const injectionMarker = "<!-- SMART WIDGET START -->";
    if (content.includes(injectionMarker)) {
        console.log("⚠️ Code already injected. Skipping.");
        return;
    }

    const snippet = `
<!-- SMART WIDGET START -->
{% assign shadow_handle = collection.fields.polular-cat-prod %}
{% if shadow_handle != blank and collections[shadow_handle].products.size > 0 %}
  <div class="container widget-container" style="margin-top: 40px; margin-bottom: 40px;">
    <div class="widget-title-wrapper">
      <h2 class="widget-title">Вам может понравиться</h2>
    </div>
    <div class="row products-grid">
      {% for product in collections[shadow_handle].products limit: 4 %}
        <div class="col-6 col-md-4 col-lg-3">
           {% include 'product_card' %}
        </div>
      {% endfor %}
    </div>
  </div>
{% endif %}
<!-- SMART WIDGET END -->
`;

    // Insert before {% endpaginate %}
    if (content.includes("{% endpaginate %}")) {
        content = content.replace("{% endpaginate %}", snippet + "\n{% endpaginate %}");
    } else {
        // Append to end if no paginate (unlikely for collection)
        content = content + snippet;
    }

    // 5. Update Asset
    const updatePayload = {
        asset: {
            content: content
        }
    };

    const updateResp = UrlFetchApp.fetch(assetUrl, {
        method: 'PUT',
        headers: headers,
        payload: JSON.stringify(updatePayload),
        muteHttpExceptions: true
    });

    if (updateResp.getResponseCode() === 200) {
        console.log("🎉 SUCCESS! Code injected into template.");
    } else {
        console.log(`❌ Error updating asset: ${updateResp.getResponseCode()}`);
        console.log(updateResp.getContentText());
    }
}
