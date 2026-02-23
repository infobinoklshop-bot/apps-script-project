function debugCheckCheckboxes() {
    // Updated debug script
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Категории — Список');

    if (!sheet) {
        console.error('Лист "Категории — Список" не найден');
        return;
    }

    const data = sheet.getDataRange().getValues();
    console.log(`Всего строк: ${data.length}`);

    let checkedCount = 0;

    // Проверяем первые 100 строк
    for (let i = 1; i < Math.min(data.length, 100); i++) {
        const row = data[i];
        const val = row[0]; // Column A
        const type = typeof val;

        if (val === true || String(val).toUpperCase() === 'TRUE') {
            console.log(`Row ${i + 1}: [CHECKED] Value: "${val}" (Type: ${type})`);
            checkedCount++;
        } else if (val !== false && val !== '') {
            console.log(`Row ${i + 1}: [UNCHECKED] Value: "${val}" (Type: ${type})`);
        }
    }

    console.log(`Найдено отмеченных (в первых 100 строках): ${checkedCount}`);
}
