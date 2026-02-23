/**
 * ========================================
 * МОДУЛЬ: UI HELPERS
 * Вспомогательные функции для интерфейса
 * ========================================
 */

/**
 * Ставит галочки (TRUE) для всех выделенных строк в столбце A
 * Работает только на листе "Категории — Список"
 */
function checkSelectedRows() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (sheet.getName() !== CATEGORY_SHEETS.MAIN_LIST) {
        SpreadsheetApp.getUi().alert('Эта функция работает только на листе "Категории — Список"');
        return;
    }

    const selection = sheet.getSelection();
    const ranges = selection.getActiveRangeList().getRanges();

    // Оптимизация: собираем все строки, чтобы не дергать API
    // Но проще просто пройтись по диапазонам

    ranges.forEach(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();

        // Пропускаем заголовок
        if (startRow === 1 && numRows === 1) return;

        const actualStartRow = startRow === 1 ? 2 : startRow;
        const actualNumRows = startRow === 1 ? numRows - 1 : numRows;

        if (actualNumRows > 0) {
            sheet.getRange(actualStartRow, 1, actualNumRows, 1).setValue(true);
        }
    });
}

/**
 * Снимает галочки (FALSE) для всех выделенных строк в столбце A
 */
function uncheckSelectedRows() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (sheet.getName() !== CATEGORY_SHEETS.MAIN_LIST) {
        SpreadsheetApp.getUi().alert('Эта функция работает только на листе "Категории — Список"');
        return;
    }

    const selection = sheet.getSelection();
    const ranges = selection.getActiveRangeList().getRanges();

    ranges.forEach(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();

        if (startRow === 1 && numRows === 1) return;

        const actualStartRow = startRow === 1 ? 2 : startRow;
        const actualNumRows = startRow === 1 ? numRows - 1 : numRows;

        if (actualNumRows > 0) {
            sheet.getRange(actualStartRow, 1, actualNumRows, 1).setValue(false);
        }
    });
}

/**
 * Ставит галочки (TRUE) для всех выделенных строк в столбце A
 * Работает только на листе "SEO-теги"
 */
function checkSelectedRowsSeoTags() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (sheet.getName() !== SEO_TAGS_CONFIG.MASS_SHEET_NAME) {
        SpreadsheetApp.getUi().alert('Эта функция работает только на листе "SEO-теги"');
        return;
    }

    const selection = sheet.getSelection();
    const ranges = selection.getActiveRangeList().getRanges();
    const dataStartRow = SEO_TAGS_CONFIG.MASS_DATA_START_ROW;

    ranges.forEach(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();

        if (startRow < dataStartRow && startRow + numRows <= dataStartRow) return;

        const actualStartRow = Math.max(startRow, dataStartRow);
        const actualNumRows = numRows - (actualStartRow - startRow);

        if (actualNumRows > 0) {
            sheet.getRange(actualStartRow, 1, actualNumRows, 1).setValue(true);
        }
    });
}

/**
 * Снимает галочки (FALSE) для всех выделенных строк в столбце A
 * Работает только на листе "SEO-теги"
 */
function uncheckSelectedRowsSeoTags() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (sheet.getName() !== SEO_TAGS_CONFIG.MASS_SHEET_NAME) {
        SpreadsheetApp.getUi().alert('Эта функция работает только на листе "SEO-теги"');
        return;
    }

    const selection = sheet.getSelection();
    const ranges = selection.getActiveRangeList().getRanges();
    const dataStartRow = SEO_TAGS_CONFIG.MASS_DATA_START_ROW;

    ranges.forEach(range => {
        const startRow = range.getRow();
        const numRows = range.getNumRows();

        if (startRow < dataStartRow && startRow + numRows <= dataStartRow) return;

        const actualStartRow = Math.max(startRow, dataStartRow);
        const actualNumRows = numRows - (actualStartRow - startRow);

        if (actualNumRows > 0) {
            sheet.getRange(actualStartRow, 1, actualNumRows, 1).setValue(false);
        }
    });
}
