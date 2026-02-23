function listGeminiModels() {
    const apiKey = 'AIzaSyD7xd_gvG7XvMOR97Iw9aZbTYgs1q5RTRM';
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;

    try {
        const response = UrlFetchApp.fetch(url);
        const data = JSON.parse(response.getContentText());

        let message = "Доступные модели:\n";
        if (data.models) {
            data.models.forEach(m => {
                if (m.supportedGenerationMethods && m.supportedGenerationMethods.includes('generateContent')) {
                    message += `- ${m.name}\n`;
                }
            });
        } else {
            message += "Модели не найдены (странный ответ).";
        }

        Logger.log(message);
        SpreadsheetApp.getUi().alert(message);

    } catch (e) {
        SpreadsheetApp.getUi().alert("Ошибка при получении списка моделей:\n" + e.toString());
    }
}
