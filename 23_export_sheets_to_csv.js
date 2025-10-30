/**
 * Экспорт всех листов таблицы в CSV файлы
 * С поддержкой больших листов (разбивка на части)
 */

const MAX_PROPERTY_SIZE = 9000; // Лимит Script Properties (~9KB)

function exportAllSheetsToCsv() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const results = [];
    
    // Очищаем старые данные
    const props = PropertiesService.getScriptProperties();
    const allKeys = props.getKeys();
    allKeys.forEach(key => {
      if (key.startsWith('CSV_')) {
        props.deleteProperty(key);
      }
    });
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      console.log(`[INFO] Экспортирую лист: ${sheetName}`);
      
      const data = sheet.getDataRange().getValues();
      const csv = convertToCsv(data);
      
      // Сохраняем с разбивкой если нужно
      const saved = saveCsvToProperties(sheetName, csv);
      
      results.push({
        sheet: sheetName,
        rows: data.length,
        size: csv.length,
        parts: saved.parts,
        saved: saved.success
      });
      
      console.log(`[INFO] Лист ${sheetName}: ${data.length} строк, ${csv.length} символов, частей: ${saved.parts}`);
    });
    
    // Сохраняем метаданные
    const metadata = {
      exportDate: new Date().toISOString(),
      sheets: results,
      totalSheets: sheets.length
    };
    
    props.setProperty('CSV_METADATA', JSON.stringify(metadata));
    
    console.log('[SUCCESS] Экспорт завершён!');
    console.log(`Экспортировано листов: ${sheets.length}`);
    
    return {
      success: true,
      message: `Экспортировано ${sheets.length} листов`,
      sheets: results
    };
    
  } catch (error) {
    console.error('[ERROR] Ошибка экспорта:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Конвертация данных в CSV
 */
function convertToCsv(data) {
  return data.map(row => {
    return row.map(cell => {
      let value = String(cell);
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        value = '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    }).join(',');
  }).join('\n');
}

/**
 * Сохранение CSV с разбивкой на части если нужно
 */
function saveCsvToProperties(sheetName, csv) {
  const props = PropertiesService.getScriptProperties();
  const baseKey = `CSV_${sheetName.replace(/[^a-zA-Z0-9]/g, '_')}`;
  
  // Если маленький - сохраняем целиком
  if (csv.length <= MAX_PROPERTY_SIZE) {
    props.setProperty(baseKey, csv);
    return { success: true, parts: 1 };
  }
  
  // Если большой - разбиваем на части
  const parts = Math.ceil(csv.length / MAX_PROPERTY_SIZE);
  
  for (let i = 0; i < parts; i++) {
    const start = i * MAX_PROPERTY_SIZE;
    const end = Math.min((i + 1) * MAX_PROPERTY_SIZE, csv.length);
    const chunk = csv.substring(start, end);
    props.setProperty(`${baseKey}_part${i}`, chunk);
  }
  
  // Сохраняем информацию о частях
  props.setProperty(`${baseKey}_info`, JSON.stringify({ parts: parts, totalSize: csv.length }));
  
  return { success: true, parts: parts };
}

/**
 * Получить CSV данные конкретного листа (собирает части если нужно)
 */
function getCsvForSheet(sheetName) {
  const props = PropertiesService.getScriptProperties();
  const baseKey = `CSV_${sheetName.replace(/[^a-zA-Z0-9]/g, '_')}`;
  
  // Проверяем есть ли информация о частях
  const info = props.getProperty(`${baseKey}_info`);
  
  if (!info) {
    // Маленький файл - возвращаем целиком
    return props.getProperty(baseKey);
  }
  
  // Большой файл - собираем части
  const { parts } = JSON.parse(info);
  let csv = '';
  
  for (let i = 0; i < parts; i++) {
    const chunk = props.getProperty(`${baseKey}_part${i}`);
    if (chunk) {
      csv += chunk;
    }
  }
  
  return csv;
}

/**
 * Получить метаданные экспорта
 */
function getExportMetadata() {
  const metadata = PropertiesService.getScriptProperties()
    .getProperty('CSV_METADATA');
  return metadata ? JSON.parse(metadata) : null;
}

/**
 * Сохранить CSV в Google Drive (альтернатива для больших файлов)
 */
function exportSheetsToGoogleDrive() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const folder = DriveApp.getRootFolder().createFolder('CSV_Exports_' + new Date().getTime());
    const results = [];
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      console.log(`[INFO] Экспортирую лист в Drive: ${sheetName}`);
      
      const data = sheet.getDataRange().getValues();
      const csv = convertToCsv(data);
      
      // Создаём файл в Drive
      const file = folder.createFile(`${sheetName}.csv`, csv, MimeType.CSV);
      
      results.push({
        sheet: sheetName,
        fileId: file.getId(),
        url: file.getUrl(),
        size: csv.length
      });
      
      console.log(`[INFO] Сохранено: ${sheetName} → ${file.getUrl()}`);
    });
    
    console.log('[SUCCESS] Все файлы сохранены в Drive!');
    console.log(`Папка: ${folder.getUrl()}`);
    
    return {
      success: true,
      folderUrl: folder.getUrl(),
      files: results
    };
    
  } catch (error) {
    console.error('[ERROR] Ошибка экспорта в Drive:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ТЕСТОВАЯ ФУНКЦИЯ
 */
function testExport() {
  console.log('=== ТЕСТ ЭКСПОРТА ===');
  const result = exportAllSheetsToCsv();
  console.log('Результат:', JSON.stringify(result, null, 2));
  
  const metadata = getExportMetadata();
  console.log('Метаданные:', JSON.stringify(metadata, null, 2));
}

/**
 * ТЕСТ ЭКСПОРТА В GOOGLE DRIVE (рекомендуется для больших таблиц)
 */
function testExportToDrive() {
  console.log('=== ТЕСТ ЭКСПОРТА В DRIVE ===');
  const result = exportSheetsToGoogleDrive();
  console.log('Результат:', JSON.stringify(result, null, 2));
}