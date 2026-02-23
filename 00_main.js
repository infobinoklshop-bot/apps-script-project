/**
 * ========================================
 * ГЛАВНЫЙ ФАЙЛ - ИНИЦИАЛИЗАЦИЯ
 * ========================================
 */

/**
 * Создает меню при открытии таблицы
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    const mainMenu = ui.createMenu('🔧 InSales Manager');

    // Добавляем меню категорий
    addFullCategoryMenu(mainMenu);

    // Дополнительные пункты меню
    mainMenu
      .addSeparator()
      .addItem('⚙️ Настройки', 'showMainSettings')
      .addItem('📋 Логи', 'showLogsSheet')
      .addItem('ℹ️ О приложении', 'showAboutDialog');

    mainMenu.addToUi();

    console.log('✅ Меню InSales Manager загружено'); // ИЗМЕНЕНО: вместо logInfo

  } catch (error) {
    console.error('Ошибка создания меню:', error);
    SpreadsheetApp.getUi().alert('Ошибка загрузки меню: ' + error.message);
  }
}

/**
 * Показывает диалог настроек
 */
function showMainSettings() {
  showAPIKeysSetup();

  const message = `⚙️ НАСТРОЙКИ ПРИЛОЖЕНИЯ\n\n` +
    `Доступные настройки:\n\n` +
    `📁 Категории:\n` +
    `• Семантика → Настроить API - для сбора ключевых слов\n` +
    `• Позиции → Настроить маркерный запрос - для отслеживания позиций\n\n` +
    `🔑 Script Properties (для разработчиков):\n` +
    `• insalesApiKey - API ключ InSales\n` +
    `• insalesPassword - Пароль InSales\n` +
    `• insalesShop - Домен магазина\n` +
    `• openaiApiKey - API ключ OpenAI\n` +
    `• semantics_api_service - Сервис семантики (serpstat/ahrefs/semrush)\n` +
    `• semantics_api_key - API ключ сервиса семантики`;

  ui.alert('Настройки', message, ui.ButtonSet.OK);
}

/**
 * Показывает лист логов
 */
function showLogsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Логи');

    if (!logSheet) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Лист логов не найден',
        'Лист с логами еще не создан. Создать сейчас?',
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        logSheet = ss.insertSheet('Логи');

        // Настраиваем лист логов
        const headers = ['Дата', 'Время', 'Уровень', 'Сообщение', 'Контекст', 'Детали'];
        logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        logSheet.getRange(1, 1, 1, headers.length)
          .setBackground('#1976d2')
          .setFontColor('#ffffff')
          .setFontWeight('bold');

        logSheet.setColumnWidth(1, 100);
        logSheet.setColumnWidth(2, 80);
        logSheet.setColumnWidth(3, 80);
        logSheet.setColumnWidth(4, 400);
        logSheet.setColumnWidth(5, 200);
        logSheet.setColumnWidth(6, 300);

        logSheet.setFrozenRows(1);
      } else {
        return;
      }
    }

    ss.setActiveSheet(logSheet);

  } catch (error) {
    console.error('❌ Ошибка показа логов', error);
  }
}

/**
 * Показывает информацию о приложении
 */
function showAboutDialog() {
  const ui = SpreadsheetApp.getUi();

  const message = `📦 InSales Manager v2.0\n\n` +
    `Комплексное решение для управления категориями InSales\n\n` +
    `✨ Возможности:\n\n` +
    `📁 Управление категориями:\n` +
    `• Загрузка с иерархией\n` +
    `• Детальные листы для каждой категории\n` +
    `• Массовое обновление\n\n` +
    `🔑 Семантическое ядро:\n` +
    `• Сбор ключевых слов (Serpstat/Ahrefs/SEMrush)\n` +
    `• LSI и тематические слова через AI\n` +
    `• Автоматическое распределение по элементам страницы\n\n` +
    `🤖 AI-генерация:\n` +
    `• SEO теги (Title, Description, H1)\n` +
    `• Описания категорий\n` +
    `• Плитка тегов с релевантностью\n` +
    `• Анализ текущего состояния\n\n` +
    `📊 Отслеживание позиций:\n` +
    `• Снятие позиций по маркерным запросам\n` +
    `• История изменений на странице\n` +
    `• Связь изменений с динамикой позиций\n\n` +
    `🛒 Управление товарами:\n` +
    `• Подбор товаров с фильтрами\n` +
    `• Массовое добавление в категории\n\n` +
    `Разработано для Binokl.shop 🔭`;

  ui.alert('О приложении', message, ui.ButtonSet.OK);
}
