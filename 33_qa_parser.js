function importQaFromSheet() {
    const context = 'importQaFromSheet';
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
        const importSheet = ss.getSheetByName(CATEGORY_SHEETS.IMPORT_QA);
        if (!importSheet) {
            ui.alert('Ошибка', `Лист "${CATEGORY_SHEETS.IMPORT_QA}" не найден.`, ui.ButtonSet.OK);
            return;
        }

        // Получаем данные с листа импорта
        const lastRow = importSheet.getLastRow();
        if (lastRow < 2) {
            ui.alert('Ошибка', 'Лист импорта пуст (нет данных со 2-й строки).', ui.ButtonSet.OK);
            return;
        }

        const importData = importSheet.getRange(2, 1, lastRow - 1, 3).getValues();

        // Группируем вопросы по названию категории
        const qaByCategory = {};
        for (let i = 0; i < importData.length; i++) {
            const row = importData[i];
            const categoryName = String(row[0] || '').trim();
            const question = String(row[1] || '').trim();
            const answer = String(row[2] || '').trim();

            if (!categoryName || !question || !answer) continue;

            if (!qaByCategory[categoryName]) {
                qaByCategory[categoryName] = [];
            }

            qaByCategory[categoryName].push(`В: ${question}\nО: ${answer}`);
        }

        // Получаем отмеченные строки на листе SEO-теги
        const selectedRows = getSelectedSeoTagRows_();
        if (selectedRows.length === 0) {
            ui.alert('Нет данных', 'Отметьте чекбоксами строки категорий на листе SEO-теги для того, чтобы импортировать в них Q&A.', ui.ButtonSet.OK);
            return;
        }

        // Подтверждение импорта
        const confirm = ui.alert('Импорт Q&A',
            `С листа "${CATEGORY_SHEETS.IMPORT_QA}" загружены вопросы для ${Object.keys(qaByCategory).length} категорий.\n\n` +
            `Данные будут добавлены в выбранные строки (${selectedRows.length} шт.) на листе SEO-тегов, если совпадут имена категорий (столбец Page Name).\n\n` +
            `Продолжить?`,
            ui.ButtonSet.YES_NO);

        if (confirm !== ui.Button.YES) return;

        const seoSheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
        let updatedCount = 0;

        // Идем по выбранным строкам и обновляем их
        for (const rowObj of selectedRows) {
            // Убираем возможные лишние пробелы из названия категории для точного сравнения
            const pageName = String(rowObj.pageName || '').trim();

            if (qaByCategory[pageName]) {
                const qaText = qaByCategory[pageName].join('\n\n');
                seoSheet.getRange(rowObj.row, SEO_TAGS_CONFIG.MASS_COLUMNS.QA).setValue(qaText);
                updatedCount++;
            }
        }

        logInfo(`📥 Q&A импортированы. Обновлено категорий: ${updatedCount}`, null, context);
        ui.alert('Готово', `Вопросы и ответы успешно загружены.\nОбновлено строк: ${updatedCount}`, ui.ButtonSet.OK);

    } catch (error) {
        logError('❌ Ошибка импорта Q&A с листа', error, context);
        ui.alert('Ошибка', 'Импорт Q&A: ' + error.message, ui.ButtonSet.OK);
    }
}
