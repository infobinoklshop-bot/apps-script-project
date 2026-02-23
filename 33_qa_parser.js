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
        let currentCategory = '';

        for (let i = 0; i < importData.length; i++) {
            const row = importData[i];
            const colA = String(row[0] || '').trim();
            const colB = String(row[1] || '').trim();

            if (!colA) continue; // пустая строка

            // Если в колонке B пусто, значит в колонке A находится заголовок (название категории)
            if (!colB) {
                currentCategory = colA;
                if (!qaByCategory[currentCategory]) {
                    qaByCategory[currentCategory] = [];
                }
            } else if (currentCategory) {
                // Если колонка B не пустая, то это вопрос (A) и ответ (B)
                qaByCategory[currentCategory].push(`В: ${colA}\nО: ${colB}`);
            }
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

        // Функция для нечеткого сравнения строк (обрезка окончаний)
        function isMatch(name1, name2) {
            const getStems = s => s.toLowerCase().replace(/[^a-zа-яё0-9]/g, ' ').split(' ')
                .filter(w => w.length > 2)
                .map(w => w.length > 4 ? w.slice(0, -2) : w);

            const stems1 = getStems(name1);
            const stems2 = getStems(name2);

            if (stems1.length === 0 || stems2.length === 0) return false;

            // Проверяем, содержит ли одно название все корни другого
            const s2in1 = stems2.every(s2 => stems1.some(s1 => s1.includes(s2) || s2.includes(s1)));
            const s1in2 = stems1.every(s1 => stems2.some(s2 => s2.includes(s1) || s1.includes(s2)));

            return s2in1 || s1in2;
        }

        // Идем по выбранным строкам и обновляем их
        for (const rowObj of selectedRows) {
            const pageName = String(rowObj.pageName || '').trim();
            if (!pageName) continue;

            let matchedCategory = null;
            // 1. Пытаемся найти точное совпадение (без учета регистра)
            const exactMatch = Object.keys(qaByCategory).find(c => c.toLowerCase() === pageName.toLowerCase());

            if (exactMatch) {
                matchedCategory = exactMatch;
            } else {
                // 2. Нечеткое совпадение (стемминг)
                for (const cat of Object.keys(qaByCategory)) {
                    if (isMatch(pageName, cat)) {
                        matchedCategory = cat;
                        break;
                    }
                }
            }

            if (matchedCategory && qaByCategory[matchedCategory].length > 0) {
                const qaText = qaByCategory[matchedCategory].join('\n\n');
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
