function debugMetricaDiscrepancy() {
    const date1 = '2023-01-01';
    const date2 = '2025-12-01';
    const targetPhrase = 'бинокль с дальномером';

    console.log(`Debugging Metrica data for phrase: "${targetPhrase}" (${date1} - ${date2})`);

    const config = YANDEX_METRICA_CONFIG;
    const params = {
        'ids': config.counterId,
        'metrics': 'ym:s:visits',
        'dimensions': 'ym:s:searchPhrase',
        'date1': date1,
        'date2': date2,
        'limit': 10000,
        'accuracy': 'full',
        'proposed_accuracy': 'false'
    };

    // 1. Fetch ALL data (like the main script)
    const queryString = Object.keys(params)
        .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
        .join('&');
    const url = `${config.baseUrl}?${queryString}`;

    try {
        const response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'OAuth ' + config.oauthToken }
        });
        const data = JSON.parse(response.getContentText());

        console.log(`Total rows fetched: ${data.data.length}`);

        // Find the specific phrase
        const matches = data.data.filter(row => {
            const phrase = row.dimensions[0].name.toLowerCase().trim();
            return phrase.includes('бинокль') && phrase.includes('дальномером');
        });

        console.log(`Found ${matches.length} matches containing "бинокль" and "дальномером":`);
        matches.forEach(m => {
            console.log(`- "${m.dimensions[0].name}": ${m.metrics[0]} visits`);
        });

        const exactMatch = data.data.find(row => row.dimensions[0].name.toLowerCase().trim() === targetPhrase);
        if (exactMatch) {
            console.log(`\nEXACT MATCH FOUND: "${exactMatch.dimensions[0].name}" -> ${exactMatch.metrics[0]} visits`);
        } else {
            console.log(`\n❌ EXACT MATCH NOT FOUND for "${targetPhrase}"`);
        }

    } catch (e) {
        console.error('API Error:', e);
    }
}
