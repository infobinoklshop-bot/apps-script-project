/**
 * Скрипт для удаления категории с переносом товаров и созданием редиректа
 */

function deleteSelectedCategory() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    if (!sheet) {
        SpreadsheetApp.getUi().alert('Лист "Категории — Список" не найден!');
        return;
    }

    // 1. Поиск выбранной категории
    const data = sheet.getDataRange().getValues();
    let selectedRowIndex = -1;
    let selectedCount = 0;

    // Пропускаем заголовок (строка 1)
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === true) { // Checkbox column A
            selectedRowIndex = i;
            selectedCount++;
        }
    }

    if (selectedCount === 0) {
        SpreadsheetApp.getUi().alert('Пожалуйста, выберите категорию (галочка в столбце A).');
        return;
    }

    if (selectedCount > 1) {
        SpreadsheetApp.getUi().alert('Пожалуйста, выберите ТОЛЬКО ОДНУ категорию для удаления.');
        return;
    }

    // Данные удаляемой категории
    const rowData = data[selectedRowIndex];
    const categoryId = rowData[MAIN_LIST_COLUMNS.CATEGORY_ID - 1];
    const parentId = rowData[MAIN_LIST_COLUMNS.PARENT_ID - 1];
    const categoryUrl = rowData[MAIN_LIST_COLUMNS.URL - 1];
    const categoryName = rowData[MAIN_LIST_COLUMNS.TITLE - 1];

    if (!categoryId) {
        SpreadsheetApp.getUi().alert('Не удалось определить ID категории.');
        return;
    }

    // 2. Поиск родительской категории
    let parentUrl = '';
    if (parentId) {
        // Ищем строку с таким ID
        const parentRow = data.find(row => row[MAIN_LIST_COLUMNS.CATEGORY_ID - 1] == parentId);
        if (parentRow) {
            parentUrl = parentRow[MAIN_LIST_COLUMNS.URL - 1];
        } else {
            SpreadsheetApp.getUi().alert(`Родительская категория с ID ${parentId} не найдена в таблице.`);
            return;
        }
    } else {
        SpreadsheetApp.getUi().alert('У выбранной категории нет родителя (это корневая категория). Перенос товаров невозможен.');
        return;
    }

    // 3. Подтверждение пользователя
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        'Подтверждение удаления',
        `Вы уверены, что хотите удалить категорию "${categoryName}" (ID: ${categoryId})?\n\n` +
        `Действия:\n` +
        `1. Товары будут перенесены в родительскую категорию (ID: ${parentId}).\n` +
        `2. Будет создан 301 редирект: ${categoryUrl} -> ${parentUrl}\n` +
        `3. Категория будет удалена из InSales и из таблицы.`,
        ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
        return;
    }

    // 4. Выполнение операций
    const credentials = getInsalesCredentialsSync();
    if (!credentials) return;

    try {
        SpreadsheetApp.getActiveSpreadsheet().toast('Начинаем перенос товаров...', 'Удаление');

        // 4.1 Перенос товаров
        const movedCount = moveProductsToParent(credentials, categoryId, parentId);

        SpreadsheetApp.getActiveSpreadsheet().toast(`Перенесено товаров: ${movedCount}. Создаем редирект...`, 'Удаление');

        // 4.2 Создание редиректа
        createRedirect(credentials, categoryUrl, parentUrl);

        SpreadsheetApp.getActiveSpreadsheet().toast('Редирект создан. Удаляем категорию...', 'Удаление');

        // 4.3 Удаление категории
        deleteCollectionInSales(credentials, categoryId);

        // 4.4 Удаление строки из таблицы
        // selectedRowIndex - это индекс в массиве data (0-based). В таблице строки 1-based.
        // data[0] это строка 1. data[selectedRowIndex] это строка selectedRowIndex + 1.
        sheet.deleteRow(selectedRowIndex + 1);

        SpreadsheetApp.getActiveSpreadsheet().toast(`Категория "${categoryName}" успешно удалена!`, 'Готово');
        SpreadsheetApp.flush();

        ui.alert(`Успешно!\n\nКатегория удалена.\nТоваров перенесено: ${movedCount}\nРедирект создан.`);

    } catch (e) {
        console.error('Ошибка при удалении категории:', e);
        ui.alert('Произошла ошибка: ' + e.message);
    }
}

/**
 * Переносит товары из одной категории в другую (меняет category_id)
 */
function moveProductsToParent(credentials, sourceCategoryId, targetCategoryId) {
    let page = 1;
    let movedCount = 0;
    let hasMore = true;

    while (hasMore) {
        // Получаем товары категории
        const url = `${credentials.baseUrl}/admin/products.json?collection_id=${sourceCategoryId}&per_page=250&page=${page}`;
        const response = UrlFetchApp.fetch(url, {
            method: 'GET',
            headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(credentials.apiKey + ':' + credentials.password) },
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            throw new Error(`Ошибка получения товаров: ${response.getContentText()}`);
        }

        const products = JSON.parse(response.getContentText());
        if (products.length === 0) {
            hasMore = false;
            break;
        }

        // Обновляем каждый товар
        for (const product of products) {
            // Проверяем, является ли удаляемая категория ГЛАВНОЙ для товара
            if (product.category_id == sourceCategoryId) {
                updateProductCategory(credentials, product.id, targetCategoryId);
                movedCount++;
            }
            // Если товар просто лежит в этой категории как в дополнительной, 
            // то при удалении категории связь пропадет сама. Ничего делать не нужно.
        }

        page++;
        Utilities.sleep(200); // Пауза чтобы не спамить API
    }

    return movedCount;
}

/**
 * Обновляет category_id товара
 */
function updateProductCategory(credentials, productId, newCategoryId) {
    const url = `${credentials.baseUrl}/admin/products/${productId}.json`;
    const payload = {
        product: {
            category_id: newCategoryId
        }
    };

    const response = UrlFetchApp.fetch(url, {
        method: 'PUT',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(credentials.apiKey + ':' + credentials.password),
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
        console.error(`Ошибка обновления товара ${productId}: ${response.getContentText()}`);
        // Не прерываем весь процесс из-за одного товара, но логируем
    }

    Utilities.sleep(100); // Лимит запросов
}

/**
 * Создает 301 редирект
 */
function createRedirect(credentials, oldUrl, newUrl) {
    const url = `${credentials.baseUrl}/admin/redirects.json`;
    const payload = {
        redirect: {
            old_url: oldUrl,
            new_url: newUrl,
            status_code: 301
        }
    };

    const response = UrlFetchApp.fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': 'Basic ' + Utilities.base64Encode(credentials.apiKey + ':' + credentials.password),
            'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 201 && response.getResponseCode() !== 422) { // 422 если уже есть
        throw new Error(`Ошибка создания редиректа: ${response.getContentText()}`);
    }
}

/**
 * Удаляет категорию (коллекцию)
 */
function deleteCollectionInSales(credentials, collectionId) {
    const url = `${credentials.baseUrl}/admin/collections/${collectionId}.json`;

    const response = UrlFetchApp.fetch(url, {
        method: 'DELETE',
        headers: { 'Authorization': 'Basic ' + Utilities.base64Encode(credentials.apiKey + ':' + credentials.password) },
        muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
        throw new Error(`Ошибка удаления категории: ${response.getContentText()}`);
    }
}
