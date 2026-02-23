/**
 * ========================================
 * МОДУЛЬ: ИНТЕГРАЦИЯ С GOOGLE GEMINI API
 * ========================================
 */

/**
 * Генерирует контент для категории через Gemini
 * Возвращает объект с полями: seo_title, meta_description, h1, body_html
 */
function generateCategoryContentWithGemini_V2(categoryData, products, feedback) {
    const context = `Генерация Gemini для категории ${categoryData.title}`;

    try {
        logInfo('🤖 Запуск генерации через Gemini V2...', null, context);
        if (feedback) {
            logInfo(`💬 С учетом замечаний: "${feedback}"`, null, context);
        }

        // 1. Подготовка промпта
        const systemInstruction = `Ты - профессиональный SEO-копирайтер и контент-менеджер интернет-магазина оптических приборов Binokl.shop.
Твоя задача - создать продающий, SEO-оптимизированный контент для категории товаров.
Ты должен вернуть ответ СТРОГО в формате JSON.`;

        const userPrompt = buildGeminiPrompt(categoryData, products, feedback);

        // 2. Вызов API
        const responseJson = callGeminiAPI_V2(userPrompt, systemInstruction);

        if (!responseJson) {
            throw new Error('Пустой ответ от Gemini API');
        }

        // 3. Парсинг ответа
        let contentData;
        try {
            // Очищаем от markdown форматирования (```json ... ```)
            const cleanJson = responseJson.replace(/^```json\s*/, '').replace(/\s*```$/, '');
            contentData = JSON.parse(cleanJson);
        } catch (e) {
            logError('Ошибка парсинга JSON от Gemini', e, context);
            console.log('Raw response:', responseJson);
            throw new Error('Некорректный JSON формат ответа Gemini');
        }

        // 4. Валидация полей
        if (!contentData.body_html || !contentData.seo_title) {
            throw new Error('В ответе Gemini отсутствуют обязательные поля');
        }

        logInfo('✅ Генерация Gemini успешно завершена', null, context);
        return contentData;

    } catch (error) {
        logError('❌ Ошибка генерации Gemini', error, context);
        throw error;
    }
}

/**
 * Генерирует баннеры для категории (Desktop и Mobile)
 */
function generateCategoryBanners(categoryData) {
    const context = `Генерация баннеров для ${categoryData.title}`;
    logInfo('🎨 Запуск генерации баннеров...', null, context);

    try {
        // 1. Промпт для Desktop (Горизонтальный)
        // Стиль: Минимализм, светло-серый фон, продукт справа, место под текст слева.
        const desktopPrompt = `Professional e-commerce banner background for category "${categoryData.title}".
        Subject: High-end ${categoryData.title} (optical instrument).
        Style: Minimalist, clean, premium product photography.
        Background: Soft light gray (#f5f5f5) or pure white studio background.
        Composition: Wide 16:9 landscape. Place the subject slightly to the right. Leave empty negative space on the left for text overlay.
        Lighting: Soft, diffused studio lighting, no harsh shadows.
        Details: Photorealistic, 8k resolution, sharp focus.
        Strictly NO text, NO watermarks, NO graphics on the image.`;

        // 2. Промпт для Mobile (Квадратный/Вертикальный)
        // Стиль: Крупный план продукта по центру на светлом фоне.
        const mobilePrompt = `Professional e-commerce mobile banner for category "${categoryData.title}".
        Subject: High-end ${categoryData.title} (optical instrument).
        Style: Minimalist, clean, premium product photography.
        Background: Soft light gray (#f5f5f5) or pure white studio background.
        Composition: Square 1:1 or Vertical 4:5. Subject centered.
        Lighting: Soft, diffused studio lighting.
        Details: Photorealistic, 8k resolution, sharp focus.
        Strictly NO text, NO watermarks.`;

        // 3. Генерация Desktop
        logInfo('🖼️ Генерация Desktop баннера...', null, context);
        const desktopImageBase64 = callGeminiImageAPI(desktopPrompt);
        const desktopUrl = saveImageToDrive(desktopImageBase64, `Banner_Desktop_${categoryData.title}_${Date.now()}.png`);

        // 4. Генерация Mobile
        logInfo('📱 Генерация Mobile баннера...', null, context);
        const mobileImageBase64 = callGeminiImageAPI(mobilePrompt);
        const mobileUrl = saveImageToDrive(mobileImageBase64, `Banner_Mobile_${categoryData.title}_${Date.now()}.png`);

        logInfo('✅ Баннеры успешно созданы', null, context);

        return {
            desktopUrl: desktopUrl,
            mobileUrl: mobileUrl
        };

    } catch (error) {
        logError('❌ Ошибка генерации баннеров', error, context);
        throw error;
    }
}

/**
 * Вызывает Gemini API для генерации изображения
 * Возвращает Base64 строку изображения
 */
function callGeminiImageAPI(prompt) {
    const apiKey = GOOGLE_GEMINI_IMAGE_CONFIG.apiKey;
    const model = GOOGLE_GEMINI_IMAGE_CONFIG.model;
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const payload = {
        contents: [{
            parts: [{ text: prompt }]
        }]
    };

    const options = {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    // Retry logic for 503 errors
    const maxRetries = 3;
    let attempt = 0;

    while (attempt < maxRetries) {
        try {
            const response = UrlFetchApp.fetch(url, options);
            const responseCode = response.getResponseCode();
            const responseText = response.getContentText();

            // Success
            if (responseCode === 200) {
                const data = JSON.parse(responseText);
                if (data.candidates &&
                    data.candidates[0].content &&
                    data.candidates[0].content.parts &&
                    data.candidates[0].content.parts[0].inlineData) {
                    return data.candidates[0].content.parts[0].inlineData.data;
                }
                throw new Error('Не удалось извлечь изображение из ответа API.');
            }

            // Handle 503 Overloaded
            if (responseCode === 503) {
                console.warn(`Gemini Image API 503 (Attempt ${attempt + 1}/${maxRetries}): Model overloaded. Retrying...`);
                attempt++;
                if (attempt < maxRetries) {
                    Utilities.sleep(2000 * attempt); // Exponential backoff: 2s, 4s
                    continue;
                }
            }

            throw new Error(`Gemini Image API Error (${responseCode}): ${responseText}`);

        } catch (e) {
            // If it's the last attempt or a non-retriable error, throw it
            if (attempt >= maxRetries || !e.message.includes('503')) {
                console.error('Gemini Image API Request Failed:', e);
                throw e;
            }
        }
    }
}

/**
 * Сохраняет Base64 изображение в Google Drive
 */
function saveImageToDrive(base64Data, filename) {
    try {
        const folderName = GOOGLE_GEMINI_IMAGE_CONFIG.driveFolder;
        let folder;

        // Ищем папку
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) {
            folder = folders.next();
        } else {
            folder = DriveApp.createFolder(folderName);
        }

        // Декодируем Base64
        const decodedBlob = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decodedBlob, 'image/png', filename);

        // Создаем файл
        const file = folder.createFile(blob);

        // Делаем файл доступным по ссылке (опционально, для просмотра)
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        return file.getUrl(); // Или file.getDownloadUrl()

    } catch (e) {
        console.error('Error saving to Drive:', e);
        throw new Error('Ошибка сохранения файла в Google Drive: ' + e.message);
    }
}

/**
 * Формирует промпт для Gemini
 */
function buildGeminiPrompt(categoryData, products, feedback) {
    const productsList = products.slice(0, 15).map(p =>
        `- ${p.title} (Цена: ${p.price}₽, Бренд: ${p.brand || 'Н/Д'})`
    ).join('\n');

    let feedbackSection = '';
    if (feedback && feedback.trim() !== '') {
        feedbackSection = `
ВАЖНЫЕ ЗАМЕЧАНИЯ ОТ ПОЛЬЗОВАТЕЛЯ (ОБЯЗАТЕЛЬНО УЧЕСТЬ):
!!! ${feedback} !!!
`;
    }

    return `
ДАННЫЕ О КАТЕГОРИИ:
- Название: "${categoryData.title}"
- Текущий H1: "${categoryData.h1 || categoryData.title}"
- Количество товаров: ${products.length}
- Примеры товаров:
${productsList}
${feedbackSection}
ТРЕБОВАНИЯ К КОНТЕНТУ:
1. **SEO Title**: Привлекательный заголовок для поисковой выдачи (до 60 символов). Должен содержать ключевые слова.
2. **Meta Description**: Кликабельное описание для сниппета (до 160 символов). С призывом к действию.
3. **H1**: Главный заголовок страницы. Должен быть кратким и точным.
4. **Описание (body_html)**:
   - Объем: 2000-3000 знаков.
   - Формат: Чистый HTML (без <html>, <body>, только контент).
   - Структура:
     - Вводный абзац (боли клиента, решение).
     - Преимущества товаров этой категории.
     - Почему стоит купить в Binokl.shop (гарантия, доставка, эксперты).
     - Советы по выбору.
     - Призыв к действию.
   - Использовать теги: <h2>, <p>, <ul>, <li>, <strong>.
   - Стиль: Экспертный, но доступный. Без воды и штампов ("высокое качество", "широкий ассортимент" - заменять на конкретику).

ФОРМАТ ОТВЕТА (JSON):
{
  "seo_title": "...",
  "meta_description": "...",
  "h1": "...",
  "body_html": "..."
}
`;
}

/**
 * Выполняет запрос к Gemini API
 */
function callGeminiAPI_V2(prompt, systemInstruction) {
    // DEBUG ALERT
    try {
        // SpreadsheetApp.getUi().alert("DEBUG: Code Updated! Using model: " + GOOGLE_GEMINI_V2_CONFIG.model);
    } catch (e) { }

    console.log('DEBUG: GOOGLE_GEMINI_V2_CONFIG:', JSON.stringify(GOOGLE_GEMINI_V2_CONFIG));

    const apiKey = GOOGLE_GEMINI_V2_CONFIG.apiKey;
    const model = GOOGLE_GEMINI_V2_CONFIG.model;

    console.log('DEBUG: Using model:', model);
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const payload = {
        contents: [{
            parts: [{ text: prompt }]
        }],
        systemInstruction: {
            parts: [{ text: systemInstruction }]
        },
        generationConfig: {
            temperature: GOOGLE_GEMINI_V2_CONFIG.temperature,
            maxOutputTokens: GOOGLE_GEMINI_V2_CONFIG.maxOutputTokens,
            responseMimeType: "application/json" // Force JSON output
        }
    };

    const options = {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        if (responseCode !== 200) {
            throw new Error(`Gemini API Error (${responseCode}): ${responseText}`);
        }

        const data = JSON.parse(responseText);

        if (!data.candidates || data.candidates.length === 0) {
            throw new Error('No candidates returned from Gemini');
        }

        return data.candidates[0].content.parts[0].text;

    } catch (e) {
        console.error('Gemini API Request Failed:', e);
        throw e;
    }
}
