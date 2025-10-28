/**
 * ========================================
 * МОДУЛЬ: ЛОГИРОВАНИЕ
 * ========================================
 */

/**
 * Логирует информационное сообщение
 */
function logInfo(message, data = null, context = '') {
  const timestamp = new Date().toLocaleString('ru-RU');
  const logMessage = `[INFO] ${timestamp} ${context ? '| ' + context + ' | ' : ''}${message}`;
  
  console.log(logMessage);
  if (data) console.log(data);
  
  // Опционально: записываем в лист логов
  writeToLogSheet('INFO', message, context, data);
}

/**
 * Логирует предупреждение
 */
function logWarning(message, data = null, context = '') {
  const timestamp = new Date().toLocaleString('ru-RU');
  const logMessage = `[WARNING] ${timestamp} ${context ? '| ' + context + ' | ' : ''}${message}`;
  
  console.warn(logMessage);
  if (data) console.warn(data);
  
  writeToLogSheet('WARNING', message, context, data);
}

/**
 * Логирует ошибку
 */
function logError(message, error = null, context = '') {
  const timestamp = new Date().toLocaleString('ru-RU');
  const logMessage = `[ERROR] ${timestamp} ${context ? '| ' + context + ' | ' : ''}${message}`;
  
  console.error(logMessage);
  if (error) {
    console.error(error);
  }
  
  const errorDetails = error ? (error.message || error.toString()) : '';
  writeToLogSheet('ERROR', message, context, errorDetails);
}

/**
 * Записывает лог в лист (опционально)
 */
function writeToLogSheet(level, message, context, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Логи');
    
    // Если листа нет - не создаем автоматически, чтобы не замедлять
    if (!logSheet) {
      return;
    }
    
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy');
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    
    const detailsStr = details ? (typeof details === 'object' ? JSON.stringify(details) : String(details)) : '';
    
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
      dateStr,
      timeStr,
      level,
      message,
      context || '',
      detailsStr
    ]]);
    
    // Цветовое кодирование
    const row = logSheet.getRange(lastRow + 1, 1, 1, 6);
    if (level === 'ERROR') {
      row.setBackground('#ffebee');
    } else if (level === 'WARNING') {
      row.setBackground('#fff9c4');
    } else {
      row.setBackground('#ffffff');
    }
    
    // Ограничиваем количество строк (последние 1000)
    if (lastRow > 1001) {
      logSheet.deleteRows(2, lastRow - 1000);
    }
    
  } catch (e) {
    // Не логируем ошибку логирования, чтобы избежать рекурсии
    console.error('Ошибка записи в лист логов:', e);
  }
}

/**
 * Создает лист логов если его нет
 */
function createLogSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Логи');
    
    if (logSheet) {
      SpreadsheetApp.getUi().alert('Лист "Логи" уже существует');
      return;
    }
    
    logSheet = ss.insertSheet('Логи');
    
    // Настраиваем заголовки
    const headers = ['Дата', 'Время', 'Уровень', 'Сообщение', 'Контекст', 'Детали'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const headerRange = logSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1976d2')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    
    logSheet.setColumnWidth(1, 100);  // Дата
    logSheet.setColumnWidth(2, 80);   // Время
    logSheet.setColumnWidth(3, 100);  // Уровень
    logSheet.setColumnWidth(4, 400);  // Сообщение
    logSheet.setColumnWidth(5, 200);  // Контекст
    logSheet.setColumnWidth(6, 300);  // Детали
    
    logSheet.setFrozenRows(1);
    
    SpreadsheetApp.getUi().alert('Лист логов создан успешно!');
    
    logInfo('✅ Лист логов создан');
    
  } catch (error) {
    console.error('Ошибка создания листа логов:', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}

/**
 * Очищает старые логи
 */
function clearOldLogs(daysToKeep = 30) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Логи');
    
    if (!logSheet) {
      SpreadsheetApp.getUi().alert('Лист логов не найден');
      return;
    }
    
    const data = logSheet.getDataRange().getValues();
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    let deleteCount = 0;
    
    // Идем с конца, чтобы не сбивать индексы при удалении
    for (let i = data.length - 1; i >= 1; i--) {
      const dateStr = data[i][0];
      if (!dateStr) continue;
      
      const rowDate = new Date(dateStr.split('.').reverse().join('-'));
      
      if (rowDate < cutoffDate) {
        logSheet.deleteRow(i + 1);
        deleteCount++;
      }
    }
    
    SpreadsheetApp.getUi().alert(`Удалено ${deleteCount} старых записей логов`);
    logInfo(`🗑️ Очищено старых логов: ${deleteCount}`);
    
  } catch (error) {
    console.error('Ошибка очистки логов:', error);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
  }
}