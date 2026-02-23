/**
 * Обработчик меню "Создать баннеры (Gemini)"
 */
function onGenerateBanners() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const ui = SpreadsheetApp.getUi();

    // Проверяем, что мы на листе "Категория — ..."
    if (!sheet.getName().startsWith('Категория — ')) {
        ui.alert('Ошибка', 'Запустите скрипт на листе категории (название должно начинаться с "Категория — ")', ui.ButtonSet.OK);
        return;
    }

    try {
        // Читаем данные категории
        // 1. Пробуем взять из названия листа (самый надежный способ)
        let categoryTitle = '';
        const sheetName = sheet.getName();
        if (sheetName.startsWith('Категория — ')) {
            categoryTitle = sheetName.substring('Категория — '.length).trim();
        }

        // 2. Если не вышло, пробуем из ячейки B1
        if (!categoryTitle) {
            const b1Value = sheet.getRange('B1').getValue();
            if (b1Value) {
                categoryTitle = b1Value.toString().replace('ИНФОРМАЦИЯ О КАТЕГОРИИ: ', '').trim();
            }
        }

        if (!categoryTitle) {
            throw new Error('Не удалось определить название категории ни из имени листа, ни из ячейки B1');
        }

        ui.alert('Генерация баннеров', `Запускаем создание баннеров для "${categoryTitle}".\nЭто может занять 10-20 секунд.\nРезультат будет сохранен в Google Drive и добавлен в ячейки B50/B51.`, ui.ButtonSet.OK);

        const categoryData = {
            title: categoryTitle
        };

        // Вызываем функцию генерации
        const result = generateCategoryBanners(categoryData);

        // Записываем результаты
        // B50 - Баннер 1 изображение
        // B51 - Баннер 1 мобильное изображение
        sheet.getRange('B50').setValue(result.desktopUrl);
        sheet.getRange('B51').setValue(result.mobileUrl);

        // Подсвечиваем
        sheet.getRange('B50:B51').setBackground('#e8f5e9');

        ui.alert('Готово!', 'Баннеры успешно созданы и добавлены.', ui.ButtonSet.OK);

    } catch (error) {
        console.error(error);
        ui.alert('Ошибка', 'Не удалось создать баннеры:\n' + error.message, ui.ButtonSet.OK);
    }
}
