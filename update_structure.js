/**
 * ========================================
 * ОБНОВЛЕНИЕ СТРУКТУРЫ ТАБЛИЦЫ
 * ========================================
 */

/**
 * Безопасно обновляет структуру таблицы, добавляя новые колонки
 * Не удаляет существующие данные!
 */
function updateTableStructureSafe() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CATEGORY_SHEETS.MAIN_LIST);

    if (!sheet) {
        SpreadsheetApp.getUi().alert('Лист категорий не найден!');
        return;
    }

    // 3. Очищаем старое место настройки "Товаров на стр", если оно там есть (O1/P1/Q1/R1)
    const o1Value = sheet.getRange('O1').getValue();
    if (o1Value === 'Товаров на 1 стр:' || o1Value === 'Период') {
        // Очистка не обязательна, мы просто перезапишем заголовки
    }

    // 4. Принудительно ставим заголовки в N1, O1, P1, Q1, R1, S1
    const n1Range = sheet.getRange('N1');
    const o1Range = sheet.getRange('O1');
    const p1Range = sheet.getRange('P1');
    const q1Range = sheet.getRange('Q1');
    const r1Range = sheet.getRange('R1');
    const s1Range = sheet.getRange('S1');

    if (n1Range.getValue() !== 'Трафик') {
        n1Range.setValue('Трафик');
        sheet.setColumnWidth(14, 100);
        formatHeader(n1Range);
    }

    if (o1Range.getValue() !== 'Отказы') {
        o1Range.setValue('Отказы');
        sheet.setColumnWidth(15, 80);
        formatHeader(o1Range);
    }

    if (p1Range.getValue() !== 'Заказы') {
        p1Range.setValue('Заказы');
        sheet.setColumnWidth(16, 80);
        formatHeader(p1Range);
    }

    if (q1Range.getValue() !== 'Показы GSC') {
        q1Range.setValue('Показы GSC');
        sheet.setColumnWidth(17, 100);
        formatHeader(q1Range);
    }

    if (r1Range.getValue() !== 'Показы Я.Веб') {
        r1Range.setValue('Показы Я.Веб');
        sheet.setColumnWidth(18, 100);
        formatHeader(r1Range);
    }

    if (s1Range.getValue() !== 'Период') {
        s1Range.setValue('Период');
        sheet.setColumnWidth(19, 200);
        formatHeader(s1Range);
    }

    // Переносим настройку "Товаров на стр" еще правее (в T1/U1)
    const t1Value = sheet.getRange('T1').getValue();
    if (t1Value !== 'Товаров на 1 стр:') {
        sheet.getRange('T1').setValue('Товаров на 1 стр:').setFontWeight('bold').setHorizontalAlignment('right');
        sheet.getRange('U1').setValue(36).setHorizontalAlignment('center').setBackground('#fff3cd');
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Структура обновлена: добавлена колонка Заказы!', 'Успешно');
}

function formatHeader(range) {
    range.setBackground('#4285f4')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
}

/**
 * Исправляет дублирование колонок в листе категорий
 */
function fixCategorySheetColumns() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORY_SHEETS.MAIN_LIST);
    if (!sheet) return;

    // Проверяем N1 (Визиты -> Трафик)
    const n1 = sheet.getRange('N1');
    if (n1.getValue() === 'Визиты') {
        n1.setValue('Трафик');
        formatHeader(n1);
    }

    // Проверяем O1 и P1
    const o1 = sheet.getRange('O1').getValue();

    if (o1 === 'Товаров на 1 стр:') {
        // Очищаем O и P
        sheet.getRange('O1:P1').clearContent(); // Только заголовки, чтобы не стереть данные если они там есть (хотя там 36)
        // А данные? Если там 36, то это не данные категории.
        // Если это колонка O (15), то это должен быть BOUNCE_RATE.
        // Если там "Товаров на 1 стр", то это мусор.
        sheet.getRange('O:P').clearContent();

        // Восстанавливаем правильные заголовки
        sheet.getRange('O1').setValue('Отказы');
        sheet.getRange('P1').setValue('Показы GSC');

        // Форматируем
        formatHeader(sheet.getRange('O1'));
        formatHeader(sheet.getRange('P1'));

        console.log('Очищены дублирующиеся колонки O и P и восстановлены заголовки');
    }

    // Проверяем, есть ли они в S
    const s1 = sheet.getRange('S1').getValue();
    if (s1 !== 'Товаров на 1 стр:') {
        sheet.getRange('S1').setValue('Товаров на 1 стр:').setFontWeight('bold').setHorizontalAlignment('right');
        sheet.getRange('T1').setValue(36).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#fff3cd');
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Колонки исправлены!', 'Успех');
}
