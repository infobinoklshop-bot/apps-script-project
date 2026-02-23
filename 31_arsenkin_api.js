/**
 * ========================================
 * МОДУЛЬ: ИНТЕГРАЦИЯ С ARSENKIN API
 * ========================================
 *
 * Импорт SEO-данных из сервиса Arsenkin (arsenkin.ru):
 * 1. Парсинг ключевых фраз (wordstat type=2) → столбец G листа "SEO-теги"
 * 2. Сбор частотности (wordstat type=1) для существующих ключей → обновление столбца G
 * 3. Title/Description конкурентов (инструмент "check-h") → столбцы H, I
 *
 * Arsenkin API — асинхронный (очередь задач):
 *   /set   — отправить задачу → получить task_id
 *   /check — проверить статус задачи
 *   /get   — получить результаты
 *
 * Из-за 6-минутного таймаута Google Apps Script используется
 * многофазный подход с сохранением прогресса в Script Properties.
 */

// ========================================
// СЕКЦИЯ 1: БАЗОВЫЕ API-ФУНКЦИИ
// ========================================

/**
 * Выполняет запрос к Arsenkin API с retry и обработкой ошибок
 * @param {string} endpoint - Endpoint (/set, /check, /get, /info)
 * @param {Object} payload - Тело запроса
 * @returns {Object} Ответ API
 */
function arsenkinApiRequest_(endpoint, payload) {
  const context = 'arsenkinApiRequest_';
  const token = ARSENKIN_CONFIG.API_TOKEN;

  if (!token) {
    throw new Error('Токен Arsenkin API не настроен. Укажите ARSENKIN_CONFIG.API_TOKEN в 01_config.js');
  }

  const url = ARSENKIN_CONFIG.BASE_URL + endpoint;
  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const retryDelays = ARSENKIN_CONFIG.RETRY_DELAYS;

  for (let attempt = 0; attempt < ARSENKIN_CONFIG.MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        return JSON.parse(responseText);
      }

      // Rate limit — ждём и повторяем
      if (responseCode === 429) {
        const delay = retryDelays[attempt] * 2;
        logWarning(`⏳ Rate limit (429), попытка ${attempt + 1}/${ARSENKIN_CONFIG.MAX_RETRIES}, ожидание ${delay}мс`, null, context);
        Utilities.sleep(delay);
        continue;
      }

      // Серверная ошибка — ждём и повторяем
      if (responseCode === 503 || responseCode === 502) {
        const delay = retryDelays[attempt];
        logWarning(`⏳ Сервер недоступен (${responseCode}), попытка ${attempt + 1}/${ARSENKIN_CONFIG.MAX_RETRIES}`, null, context);
        Utilities.sleep(delay);
        continue;
      }

      // Другие ошибки — сразу выбрасываем
      console.error(`[DEBUG] Arsenkin Error Response Body:`, responseText);
      throw new Error(`Arsenkin API: HTTP ${responseCode} — ${responseText.substring(0, 200)}`);

    } catch (error) {
      if (error.message.startsWith('Arsenkin API:')) throw error;
      if (attempt < ARSENKIN_CONFIG.MAX_RETRIES - 1) {
        logWarning(`⚠️ Ошибка запроса, попытка ${attempt + 1}: ${error.message}`, null, context);
        Utilities.sleep(retryDelays[attempt]);
      } else {
        throw new Error(`Arsenkin API: все ${ARSENKIN_CONFIG.MAX_RETRIES} попыток исчерпаны — ${error.message}`);
      }
    }
  }
}

/**
 * Отправляет задачу в Arsenkin
 * @param {string} toolsName - Имя инструмента (semantics, check-h)
 * @param {Object} data - Параметры задачи
 * @returns {string} ID задачи
 */
function arsenkinSubmitTask_(toolsName, data) {
  const context = 'arsenkinSubmitTask_';

  const payload = {
    tools_name: toolsName,
    data: data
  };

  logInfo(`📤 Отправка задачи "${toolsName}"`, { queries: JSON.stringify(data.queries || []).substring(0, 200) }, context);

  // --- DEBUG LOGGING START ---
  console.log(`[DEBUG] Arsenkin Payload (${toolsName}):`, JSON.stringify(payload, null, 2));
  // --- DEBUG LOGGING END ---

  const result = arsenkinApiRequest_('set', payload);

  if (result && result.task_id) {
    logInfo(`✅ Задача создана: ${result.task_id} (${result.cost || '?'} лим.)`, null, context);
    return result.task_id;
  }

  if (result && result.id) {
    return result.id;
  }

  throw new Error('Arsenkin API не вернул task_id: ' + JSON.stringify(result).substring(0, 300));
}

/**
 * Проверяет статус задачи
 * @param {string} taskId - ID задачи
 * @returns {Object} {status: 'pending'|'done'|'error', ...}
 */
function arsenkinCheckTask_(taskId) {
  return arsenkinApiRequest_('check', { task_id: taskId });
}

/**
 * Получает результат завершённой задачи
 * @param {string} taskId - ID задачи
 * @returns {Object} Данные результата
 */
function arsenkinGetResult_(taskId) {
  return arsenkinApiRequest_('get', { task_id: taskId });
}

// ========================================
// СЕКЦИЯ 2: ПОЛУЧЕНИЕ ВЫБРАННЫХ СТРОК
// ========================================

/**
 * Получает строки с отмеченными чекбоксами из листа "SEO-теги"
 * @returns {Array<Object>} [{row, id, url, h1}, ...]
 */
function getSelectedSeoTagRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Лист "${SEO_TAGS_CONFIG.MASS_SHEET_NAME}" не найден. Создайте его через меню.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < SEO_TAGS_CONFIG.MASS_DATA_START_ROW) {
    throw new Error('Лист "SEO-теги" пуст. Загрузите категории через меню.');
  }

  const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
  const dataRange = sheet.getRange(SEO_TAGS_CONFIG.MASS_DATA_START_ROW, 1, lastRow - 1, cols.RESULT_DESC);
  const values = dataRange.getValues();

  const selected = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[cols.CHECKBOX - 1] === true) {
      selected.push({
        row: i + SEO_TAGS_CONFIG.MASS_DATA_START_ROW,  // Номер строки в таблице
        id: row[cols.ID - 1],
        h1: row[cols.PAGE_NAME - 1],
        url: row[cols.URL - 1]
      });
    }
  }

  return selected;
}

// ========================================
// СЕКЦИЯ 3: ПАРСИНГ КЛЮЧЕВЫХ ФРАЗ (WORDSTAT)
// ========================================

/**
 * Парсинг ключевых фраз из Яндекс Wordstat через Arsenkin (вызывается из меню)
 * Использует H1 из столбца C как запрос, собирает связанные фразы с частотностью.
 * Результат: "фраза (частота), фраза2 (частота2), ..." — отсортировано по убыванию.
 */
/**
 * @param {boolean} [silent=false] — если true, пропускает UI-диалоги (для пайплайна)
 * @returns {{complete: boolean, pendingCount: number}|undefined}
 */
function importKeywordsFromArsenkin(silent) {
  const context = 'importKeywordsFromArsenkin';
  const ui = silent ? null : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressKey = ARSENKIN_CONFIG.PROGRESS_KEYS.WORDSTAT_PHRASES;

  try {
    if (!ARSENKIN_CONFIG.API_TOKEN) {
      if (!silent) ui.alert('Ошибка', 'Токен Arsenkin API не настроен.\nУкажите его в ARSENKIN_CONFIG.API_TOKEN в файле 01_config.js', ui.ButtonSet.OK);
      return { complete: false, pendingCount: 0 };
    }

    // Проверяем незавершённый прогресс
    const savedProgress = loadProgressChunked_(progressKey);
    let progress;

    if (savedProgress && savedProgress.tasks && savedProgress.tasks.length > 0) {
      const pendingCount = savedProgress.tasks.filter(t => t.status === 'pending').length;
      if (pendingCount > 0) {
        if (silent) {
          progress = savedProgress; // Автоматически продолжаем в режиме пайплайна
        } else {
          const resume = ui.alert('Незавершённый импорт',
            `Найден незавершённый парсинг ключевиков:\n` +
            `Всего задач: ${savedProgress.tasks.length}\n` +
            `Ожидают: ${pendingCount}\n\n` +
            `Продолжить с места остановки?`,
            ui.ButtonSet.YES_NO);

          if (resume === ui.Button.YES) {
            progress = savedProgress;
          } else {
            clearProgressChunked_(progressKey);
            progress = null;
          }
        }
      } else {
        clearProgressChunked_(progressKey);
        progress = null;
      }
    }

    // Фаза отправки
    if (!progress) {
      const selected = getSelectedSeoTagRows_();
      if (selected.length === 0) {
        if (!silent) ui.alert('Нет данных', 'Отметьте чекбоксами строки для парсинга ключевиков на листе "SEO-теги"', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      // Фильтруем строки без H1 (столбец C)
      const validRows = selected.filter(r => r.h1 && r.h1.toString().trim() !== '');
      const skippedCount = selected.length - validRows.length;

      if (validRows.length === 0) {
        if (!silent) ui.alert('Нет запросов', 'Все выбранные строки не содержат H1 (столбец C)', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      if (!silent) {
        const confirm = ui.alert('Парсинг ключевиков (Wordstat → Арсенкин)',
          `Будет отправлено ${validRows.length} задач в Arsenkin.\n` +
          (skippedCount > 0 ? `Пропущено (нет H1): ${skippedCount}\n` : '') +
          `Инструмент: Парсинг фраз Яндекс Wordstat\n` +
          `Запрос: H1 из столбца C\n\n` +
          `Асинхронная обработка: если задачи не успеют завершиться за 5 мин,\n` +
          `запустите функцию повторно для получения результатов.\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);

        if (confirm !== ui.Button.YES) return;
      }

      progress = { tasks: [], startedAt: new Date().toISOString() };
      const cfg = ARSENKIN_CONFIG.WORDSTAT_PHRASES;

      for (let i = 0; i < validRows.length; i++) {
        const row = validRows[i];
        ss.toast(`Отправка задач: ${i + 1}/${validRows.length}...`, '📤 Arsenkin', 3);

        try {
          const query = row.h1.toString().trim();

          const taskData = {
            type: cfg.type,
            queries: [query],
            device: cfg.device,
            region: cfg.region,
            minus_words: cfg.minus_words,
            is_clear_minus: cfg.is_clear_minus,
            is_right: cfg.is_right,
            is_clear: cfg.is_clear
          };

          const taskId = arsenkinSubmitTask_(cfg.tools_name, taskData);

          progress.tasks.push({
            taskId: taskId,
            row: row.row,
            query: query,
            status: 'pending',
            result: null
          });

        } catch (error) {
          logError(`❌ Ошибка отправки задачи для строки ${row.row}: ${error.message}`, error, context);
          progress.tasks.push({
            taskId: null,
            row: row.row,
            query: row.h1,
            status: 'error',
            error: error.message
          });
        }

        if (i < validRows.length - 1) {
          Utilities.sleep(ARSENKIN_CONFIG.DELAY_BETWEEN_SUBMITS_MS);
        }
      }

      saveProgressChunked_(progressKey, progress);
      logInfo(`📤 Отправлено ${progress.tasks.filter(t => t.status === 'pending').length} задач wordstat`, null, context);
    }

    // Фаза опроса
    arsenkinPollAndWriteResults_(progress, progressKey, 'wordstat-phrases', silent);

    // Возвращаем статус для пайплайна
    const updatedProgress = loadProgressChunked_(progressKey);
    if (!updatedProgress) {
      return { complete: true, pendingCount: 0 };
    }
    const remaining = updatedProgress.tasks ? updatedProgress.tasks.filter(t => t.status === 'pending').length : 0;
    return { complete: remaining === 0, pendingCount: remaining };

  } catch (error) {
    logError('❌ Ошибка парсинга ключевиков', error, context);
    if (!silent) ui.alert('Ошибка', 'Парсинг ключевиков: ' + error.message, ui.ButtonSet.OK);
    return { complete: false, pendingCount: 0 };
  }
}

// ========================================
// СЕКЦИЯ 4: ИМПОРТ ТЕГОВ КОНКУРЕНТОВ
// ========================================

/**
 * Импорт Title/Description конкурентов из Arsenkin (вызывается из меню)
 */
/**
 * @param {boolean} [silent=false] — если true, пропускает UI-диалоги (для пайплайна)
 * @returns {{complete: boolean, pendingCount: number}|undefined}
 */
function importCompetitorTagsFromArsenkin(silent) {
  const context = 'importCompetitorTagsFromArsenkin';
  const ui = silent ? null : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressKey = ARSENKIN_CONFIG.PROGRESS_KEYS.CHECK_H;

  try {
    if (!ARSENKIN_CONFIG.API_TOKEN) {
      if (!silent) ui.alert('Ошибка', 'Токен Arsenkin API не настроен.\nУкажите его в ARSENKIN_CONFIG.API_TOKEN в файле 01_config.js', ui.ButtonSet.OK);
      return { complete: false, pendingCount: 0 };
    }

    const savedProgress = loadProgressChunked_(progressKey);
    let progress;

    if (savedProgress && savedProgress.tasks && savedProgress.tasks.length > 0) {
      const pendingCount = savedProgress.tasks.filter(t => t.status === 'pending').length;
      if (pendingCount > 0) {
        if (silent) {
          progress = savedProgress;
        } else {
          const resume = ui.alert('Незавершённый импорт',
            `Найден незавершённый импорт тегов конкурентов:\n` +
            `Всего задач: ${savedProgress.tasks.length}\n` +
            `Ожидают: ${pendingCount}\n\n` +
            `Продолжить с места остановки?`,
            ui.ButtonSet.YES_NO);

          if (resume === ui.Button.YES) {
            progress = savedProgress;
          } else {
            clearProgressChunked_(progressKey);
            progress = null;
          }
        }
      } else {
        clearProgressChunked_(progressKey);
        progress = null;
      }
    }

    if (!progress) {
      const selected = getSelectedSeoTagRows_();
      if (selected.length === 0) {
        if (!silent) ui.alert('Нет данных', 'Отметьте чекбоксами строки для импорта тегов конкурентов на листе "SEO-теги"', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      const validRows = selected.filter(r => r.h1 && r.h1.toString().trim() !== '');
      if (validRows.length === 0) {
        if (!silent) ui.alert('Нет запросов', 'Все выбранные строки не содержат H1/маркерный запрос (столбец C)', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      if (!silent) {
        const skippedCount = selected.length - validRows.length;
        const confirm = ui.alert('Импорт тегов конкурентов (Арсенкин)',
          `Будет отправлено ${validRows.length} задач в Arsenkin.\n` +
          (skippedCount > 0 ? `Пропущено (нет H1): ${skippedCount}\n` : '') +
          `Поиск: Яндекс, регион Москва, глубина ТОП-${ARSENKIN_CONFIG.CHECK_H.depth}\n\n` +
          `Асинхронная обработка: если задачи не успеют завершиться за 5 мин,\n` +
          `запустите функцию повторно для получения результатов.\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);

        if (confirm !== ui.Button.YES) return;
      }

      progress = { tasks: [], startedAt: new Date().toISOString() };

      for (let i = 0; i < validRows.length; i++) {
        const row = validRows[i];
        ss.toast(`Отправка задач: ${i + 1}/${validRows.length}...`, '📤 Arsenkin', 3);

        try {
          const query = row.h1.toString().trim();
          const taskData = {
            pause: ARSENKIN_CONFIG.CHECK_H.pause,
            foreign: false,
            mode: ARSENKIN_CONFIG.CHECK_H.mode,
            queries: [query],
            se: ARSENKIN_CONFIG.CHECK_H.se,
            region: ARSENKIN_CONFIG.CHECK_H.region,
            depth: ARSENKIN_CONFIG.CHECK_H.depth
          };

          const taskId = arsenkinSubmitTask_(ARSENKIN_CONFIG.CHECK_H.tools_name, taskData);
          progress.tasks.push({ taskId, row: row.row, query, status: 'pending', result: null });
        } catch (error) {
          logError(`❌ Ошибка отправки задачи для строки ${row.row}: ${error.message}`, error, context);
          progress.tasks.push({ taskId: null, row: row.row, query: row.h1, status: 'error', error: error.message });
        }

        if (i < validRows.length - 1) {
          Utilities.sleep(ARSENKIN_CONFIG.DELAY_BETWEEN_SUBMITS_MS);
        }
      }

      saveProgressChunked_(progressKey, progress);
      logInfo(`📤 Отправлено ${progress.tasks.filter(t => t.status === 'pending').length} задач`, null, context);
    }

    arsenkinPollAndWriteResults_(progress, progressKey, 'check-h', silent);

    const updatedProgress = loadProgressChunked_(progressKey);
    if (!updatedProgress) return { complete: true, pendingCount: 0 };
    const remaining = updatedProgress.tasks ? updatedProgress.tasks.filter(t => t.status === 'pending').length : 0;
    return { complete: remaining === 0, pendingCount: remaining };

  } catch (error) {
    logError('❌ Ошибка импорта тегов конкурентов', error, context);
    if (!silent) ui.alert('Ошибка', 'Импорт тегов конкурентов: ' + error.message, ui.ButtonSet.OK);
    return { complete: false, pendingCount: 0 };
  }
}

// ========================================
// СЕКЦИЯ 4b: СБОР ЧАСТОТНОСТИ ДЛЯ СУЩЕСТВУЮЩИХ КЛЮЧЕЙ
// ========================================

/**
 * Сбор частотности из Яндекс Wordstat для ключей, уже находящихся в столбце G.
 * Берёт ключевики из ячейки (импортированные из Метрики или вручную),
 * отправляет в Arsenkin wordstat type=1, получает частотность,
 * обновляет ячейку в формате "ключ (частота)" и сортирует по убыванию.
 */
/**
 * @param {boolean} [silent=false] — если true, пропускает UI-диалоги (для пайплайна)
 * @returns {{complete: boolean, pendingCount: number}|undefined}
 */
function collectFrequencyFromArsenkin(silent) {
  const context = 'collectFrequencyFromArsenkin';
  const ui = silent ? null : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressKey = ARSENKIN_CONFIG.PROGRESS_KEYS.WORDSTAT_FREQ;

  try {
    if (!ARSENKIN_CONFIG.API_TOKEN) {
      if (!silent) ui.alert('Ошибка', 'Токен Arsenkin API не настроен.\nУкажите его в ARSENKIN_CONFIG.API_TOKEN в файле 01_config.js', ui.ButtonSet.OK);
      return { complete: false, pendingCount: 0 };
    }

    const savedProgress = loadProgressChunked_(progressKey);
    let progress;

    if (savedProgress && savedProgress.tasks && savedProgress.tasks.length > 0) {
      const pendingCount = savedProgress.tasks.filter(t => t.status === 'pending').length;
      if (pendingCount > 0) {
        if (silent) {
          progress = savedProgress;
        } else {
          const resume = ui.alert('Незавершённый сбор',
            `Найден незавершённый сбор частотности:\n` +
            `Всего задач: ${savedProgress.tasks.length}\n` +
            `Ожидают: ${pendingCount}\n\n` +
            `Продолжить с места остановки?`,
            ui.ButtonSet.YES_NO);

          if (resume === ui.Button.YES) {
            progress = savedProgress;
          } else {
            clearProgressChunked_(progressKey);
            progress = null;
          }
        }
      } else {
        clearProgressChunked_(progressKey);
        progress = null;
      }
    }

    if (!progress) {
      const selected = getSelectedSeoTagRows_();
      if (selected.length === 0) {
        if (!silent) ui.alert('Нет данных', 'Отметьте чекбоксами строки для сбора частотности на листе "SEO-теги"', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
      const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

      const rowsWithKeywords = [];
      for (const row of selected) {
        const cellValue = (sheet.getRange(row.row, cols.SEMANTIC_CORE).getValue() || '').toString().trim();
        if (!cellValue) continue;
        const phrases = extractPhrasesWithoutFrequency_(cellValue);
        if (phrases.length === 0) continue;
        rowsWithKeywords.push({ row: row.row, phrases, originalValue: cellValue });
      }

      if (rowsWithKeywords.length === 0) {
        if (!silent) ui.alert('Нет ключевиков', 'В выбранных строках столбец G (Семантическое ядро) пуст.\nСначала импортируйте ключевики.', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      if (!silent) {
        const totalPhrases = rowsWithKeywords.reduce((sum, r) => sum + r.phrases.length, 0);
        const confirm = ui.alert('Сбор частотности (Wordstat → Арсенкин)',
          `Строк с ключевиками: ${rowsWithKeywords.length}\n` +
          `Всего фраз для проверки: ${totalPhrases}\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;
      }

      progress = { tasks: [], startedAt: new Date().toISOString() };
      const cfg = ARSENKIN_CONFIG.WORDSTAT_FREQUENCY;

      for (let i = 0; i < rowsWithKeywords.length; i++) {
        const rowData = rowsWithKeywords[i];
        ss.toast(`Отправка задач: ${i + 1}/${rowsWithKeywords.length}...`, '📤 Частотность', 3);

        try {
          const taskData = { type: cfg.type, queries: rowData.phrases, device: cfg.device, regions: cfg.regions, ws: cfg.ws };
          const taskId = arsenkinSubmitTask_(cfg.tools_name, taskData);
          progress.tasks.push({ taskId, row: rowData.row, phrases: rowData.phrases, status: 'pending', result: null });
        } catch (error) {
          logError(`❌ Ошибка отправки задачи частотности для строки ${rowData.row}: ${error.message}`, error, context);
          progress.tasks.push({ taskId: null, row: rowData.row, phrases: rowData.phrases, status: 'error', error: error.message });
        }

        if (i < rowsWithKeywords.length - 1) {
          Utilities.sleep(ARSENKIN_CONFIG.DELAY_BETWEEN_SUBMITS_MS);
        }
      }

      saveProgressChunked_(progressKey, progress);
      logInfo(`📤 Отправлено ${progress.tasks.filter(t => t.status === 'pending').length} задач частотности`, null, context);
    }

    arsenkinPollAndWriteResults_(progress, progressKey, 'wordstat-freq', silent);

    const updatedProgress = loadProgressChunked_(progressKey);
    if (!updatedProgress) return { complete: true, pendingCount: 0 };
    const remaining = updatedProgress.tasks ? updatedProgress.tasks.filter(t => t.status === 'pending').length : 0;
    return { complete: remaining === 0, pendingCount: remaining };

  } catch (error) {
    logError('❌ Ошибка сбора частотности', error, context);
    if (!silent) ui.alert('Ошибка', 'Сбор частотности: ' + error.message, ui.ButtonSet.OK);
    return { complete: false, pendingCount: 0 };
  }
}

/**
 * Извлекает чистые фразы из содержимого ячейки, убирая частотность в скобках.
 * "бинокль (1500), театральный бинокль (800)" → ["бинокль", "театральный бинокль"]
 * "бинокль\nтеатральный бинокль" → ["бинокль", "театральный бинокль"]
 */
function extractPhrasesWithoutFrequency_(cellValue) {
  // Разделяем по запятой или переносу строки
  const items = cellValue.split(/[,\n]/).map(s => s.trim()).filter(s => s);

  return items.map(item => {
    // Убираем частотность в скобках: "бинокль (1500)" → "бинокль"
    return item.replace(/\s*\(\d+\)\s*$/, '').trim();
  }).filter(s => s && !s.startsWith('ОШИБКА'));
}

// ========================================
// СЕКЦИЯ 5: ОПРОС СТАТУСА И ЗАПИСЬ РЕЗУЛЬТАТОВ
// ========================================

/**
 * Опрашивает статус задач Arsenkin и записывает результаты в таблицу
 * @param {Object} progress - Объект прогресса с массивом tasks
 * @param {string} progressKey - Ключ для Script Properties
 * @param {string} toolType - 'semantics' или 'check-h'
 * @param {boolean} [silent=false] - Если true, не показывает UI-диалоги
 */
function arsenkinPollAndWriteResults_(progress, progressKey, toolType, silent) {
  const context = 'arsenkinPollAndWriteResults_';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
  const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
  const startTime = Date.now();

  let completedInThisRun = 0;

  while (Date.now() - startTime < ARSENKIN_CONFIG.EXECUTION_SAFETY_MS) {
    const pendingTasks = progress.tasks.filter(t => t.status === 'pending');

    if (pendingTasks.length === 0) break;

    ss.toast(
      `Ожидание: ${pendingTasks.length} задач, готово: ${progress.tasks.filter(t => t.status === 'done').length}`,
      '⏳ Arsenkin', 5
    );

    // Проверяем каждую pending-задачу
    for (const task of pendingTasks) {
      if (!task.taskId) continue;
      if (Date.now() - startTime >= ARSENKIN_CONFIG.EXECUTION_SAFETY_MS) break;

      try {
        const checkResult = arsenkinCheckTask_(task.taskId);

        // Определяем статус задачи
        // Arsenkin API: status="process" (в работе), status="finish" (готово)
        const isDone = checkResult && (
          checkResult.status === 'finish' ||      // Arsenkin: задача завершена
          checkResult.status === 'done' ||
          checkResult.status === 'completed' ||
          checkResult.progress === 100 ||
          checkResult.progress === '100'
        );

        const isError = checkResult && (
          checkResult.status === 'error' ||
          checkResult.status === 'fail' ||
          (checkResult.error && checkResult.code !== 'TASK_STATUS')
        );

        if (isDone) {
          // Задача завершена — получаем результат через /get
          const resultData = arsenkinGetResult_(task.taskId);
          task.status = 'done';
          task.result = resultData;

          writeArsenkinResult_(sheet, task, toolType, cols);
          completedInThisRun++;

          logInfo(`✅ Задача ${task.taskId} завершена (строка ${task.row})`, null, context);

        } else if (isError) {
          task.status = 'error';
          task.error = checkResult.error || checkResult.msg || 'Ошибка обработки задачи в Arsenkin';
          markRowError_(sheet, task.row, task.error, toolType, cols);
          logError(`❌ Задача ${task.taskId} завершена с ошибкой`, { error: task.error }, context);
        }
        // Иначе — всё ещё pending, проверим на следующей итерации

      } catch (error) {
        logWarning(`⚠️ Ошибка проверки задачи ${task.taskId}: ${error.message}`, null, context);
      }

      Utilities.sleep(500); // Мини-пауза между проверками
    }

    // Если все задачи обработаны — выходим
    if (progress.tasks.filter(t => t.status === 'pending').length === 0) break;

    // Ждём перед следующим циклом опроса
    Utilities.sleep(ARSENKIN_CONFIG.POLL_INTERVAL_MS);
  }

  // Подводим итоги
  const done = progress.tasks.filter(t => t.status === 'done').length;
  const errors = progress.tasks.filter(t => t.status === 'error').length;
  const pending = progress.tasks.filter(t => t.status === 'pending').length;
  const total = progress.tasks.length;

  if (pending > 0) {
    // Сохраняем прогресс для продолжения
    saveProgressChunked_(progressKey, progress);

    const toolNameMap = { 'wordstat-phrases': 'ключевиков', 'wordstat-freq': 'частотности', 'check-h': 'тегов конкурентов', 'search-highlights': 'LSI (подсветки)', 'paa': 'Q&A (PAA)' };
    const toolName = toolNameMap[toolType] || toolType;
    ss.toast(
      `Завершено: ${done}/${total}. Ожидают: ${pending}. Запустите функцию повторно.`,
      `⏸️ Импорт ${toolName}`, 10
    );

    if (!silent) {
      SpreadsheetApp.getUi().alert('Таймаут',
        `Время выполнения исчерпано.\n\n` +
        `Завершено в этом запуске: ${completedInThisRun}\n` +
        `Всего завершено: ${done}/${total}\n` +
        `Ошибок: ${errors}\n` +
        `Ожидают: ${pending}\n\n` +
        `Запустите функцию повторно для получения оставшихся результатов.`,
        SpreadsheetApp.getUi().ButtonSet.OK);
    }

  } else {
    // Все задачи обработаны — очищаем прогресс
    clearProgressChunked_(progressKey);

    const toolNameMap = { 'wordstat-phrases': 'ключевиков', 'wordstat-freq': 'частотности', 'check-h': 'тегов конкурентов', 'search-highlights': 'LSI (подсветки)', 'paa': 'Q&A (PAA)' };
    const toolName = toolNameMap[toolType] || toolType;
    ss.toast(`Готово! Успешно: ${done}, ошибок: ${errors}`, `✅ Импорт ${toolName}`, 5);

    if (!silent) {
      SpreadsheetApp.getUi().alert('Импорт завершён',
        `Импорт ${toolName} завершён.\n\n` +
        `Успешно: ${done}\n` +
        `Ошибок: ${errors}\n` +
        `Всего: ${total}`,
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Записывает результат Arsenkin в соответствующие ячейки таблицы.
 * Дополняет существующие данные, а не перезаписывает.
 */
function writeArsenkinResult_(sheet, task, toolType, cols) {
  const context = 'writeArsenkinResult_';

  try {
    if (toolType === 'wordstat-phrases') {
      // Парсинг фраз Wordstat → столбец G
      const keywords = parseWordstatPhrasesResult_(task.result);
      const cell = sheet.getRange(task.row, cols.SEMANTIC_CORE);
      const existing = (cell.getValue() || '').toString().trim();

      // Объединяем с существующими, дедупликация по чистой фразе
      const merged = mergeKeywordsWithFrequency_(existing, keywords);
      cell.setValue(merged);
      cell.setBackground('#e8f5e9');

    } else if (toolType === 'wordstat-freq') {
      // Сбор частотности → обновление столбца G
      const freqMap = parseWordstatFrequencyResult_(task.result);
      const cell = sheet.getRange(task.row, cols.SEMANTIC_CORE);
      const existing = (cell.getValue() || '').toString().trim();

      const updated = updateFrequenciesAndSort_(existing, freqMap);
      cell.setValue(updated);
      cell.setBackground('#e8f5e9');

    } else if (toolType === 'check-h') {
      const formatted = parseCompetitorTagsResult_(task.result);
      const cell = sheet.getRange(task.row, cols.COMPETITORS_DESC);
      cell.setValue(formatted);
      cell.setBackground('#e8f5e9');

    } else if (toolType === 'search-highlights') {
      // Подсветки → столбец H (LSI)
      const lsiWords = parseSearchHighlightsResult_(task.result);
      const cell = sheet.getRange(task.row, cols.LSI);
      const existing = (cell.getValue() || '').toString().trim();
      const merged = mergeWithoutDuplicates_(existing, lsiWords, '\n');
      cell.setValue(merged);
      cell.setBackground('#e8f5e9');

    } else if (toolType === 'paa') {
      // People Also Ask → столбец I (QA)
      const qaText = parsePaaResult_(task.result);
      const cell = sheet.getRange(task.row, cols.QA);
      cell.setValue(qaText);
      cell.setBackground('#e8f5e9');
    }

  } catch (error) {
    logError(`❌ Ошибка записи результата в строку ${task.row}`, error, context);
    markRowError_(sheet, task.row, error.message, toolType, cols);
  }
}

/**
 * Отмечает строку как ошибочную (красный фон)
 */
function markRowError_(sheet, row, errorMsg, toolType, cols) {
  if (toolType === 'wordstat-phrases' || toolType === 'wordstat-freq') {
    sheet.getRange(row, cols.SEMANTIC_CORE).setValue('ОШИБКА: ' + errorMsg);
    sheet.getRange(row, cols.SEMANTIC_CORE).setBackground('#ffebee');
  } else if (toolType === 'check-h') {
    sheet.getRange(row, cols.COMPETITORS_TITLE).setValue('ОШИБКА: ' + errorMsg);
    sheet.getRange(row, cols.COMPETITORS_TITLE).setBackground('#ffebee');
    sheet.getRange(row, cols.COMPETITORS_DESC).setBackground('#ffebee');
  } else if (toolType === 'clustering') {
    sheet.getRange(row, cols.SEMANTIC_CORE).setValue('ОШИБКА: ' + errorMsg);
    sheet.getRange(row, cols.SEMANTIC_CORE).setBackground('#ffebee');
  } else if (toolType === 'search-highlights') {
    sheet.getRange(row, cols.LSI).setValue('ОШИБКА: ' + errorMsg);
    sheet.getRange(row, cols.LSI).setBackground('#ffebee');
  } else if (toolType === 'paa') {
    sheet.getRange(row, cols.QA).setValue('ОШИБКА: ' + errorMsg);
    sheet.getRange(row, cols.QA).setBackground('#ffebee');
  }
}

/**
 * Объединяет существующие и новые данные, удаляя точные дубли.
 * Сравнение нормализованное: trim + lowercase.
 * @param {string} existing - Текущее содержимое ячейки
 * @param {string} newData - Новые данные для добавления
 * @param {string} separator - Разделитель между элементами (', ' или '\n')
 * @returns {string} Объединённые данные без дублей
 */
function mergeWithoutDuplicates_(existing, newData, separator) {
  if (!newData || newData.trim() === '') return existing;
  if (!existing || existing.trim() === '') return newData;

  // Разбиваем на отдельные элементы
  const existingItems = existing.split(separator).map(s => s.trim()).filter(s => s);
  const newItems = newData.split(separator).map(s => s.trim()).filter(s => s);

  // Набор нормализованных существующих элементов для быстрого поиска
  const existingNorm = new Set(existingItems.map(s => s.toLowerCase()));

  // Добавляем только те новые, которых ещё нет
  const toAdd = newItems.filter(item => !existingNorm.has(item.toLowerCase()));

  if (toAdd.length === 0) return existing;

  return existing + separator + toAdd.join(separator);
}

// ========================================
// СЕКЦИЯ 6: ПАРСЕРЫ РЕЗУЛЬТАТОВ
// ========================================

/**
 * Парсит результат wordstat type=2 (парсинг фраз).
 * Возвращает строку "фраза (частота), фраза2 (частота2), ..."
 * отсортированную по убыванию частотности.
 * @param {Object} apiResult - Ответ Arsenkin API
 * @returns {string} Ключевые слова с частотами
 */
function parseWordstatPhrasesResult_(apiResult) {
  const context = 'parseWordstatPhrasesResult_';

  if (!apiResult) return '';

  try {
    // Arsenkin /get возвращает:
    // { code: "TASK_RESULT", task_id: ..., result: { data: [{ freq: 104, left: { "фраза": частота, ... } }] } }
    const result = apiResult.result || apiResult;
    const dataArray = result.data || result;

    let items = [];

    if (Array.isArray(dataArray)) {
      // Основной формат: data = [{ freq, left: { phrase: freq } }]
      for (const entry of dataArray) {
        if (entry.left && typeof entry.left === 'object') {
          // left — объект { "фраза": частота }
          for (const [phrase, freq] of Object.entries(entry.left)) {
            items.push({
              phrase: phrase.trim(),
              frequency: parseInt(freq) || 0
            });
          }
        }
        // Также проверяем right (правая колонка Wordstat)
        if (entry.right && typeof entry.right === 'object') {
          for (const [phrase, freq] of Object.entries(entry.right)) {
            items.push({
              phrase: phrase.trim(),
              frequency: parseInt(freq) || 0
            });
          }
        }
      }
    } else if (typeof dataArray === 'string') {
      // CSV формат (fallback)
      items = parseWordstatCsv_(dataArray);
    } else if (dataArray && typeof dataArray === 'object') {
      // Может быть объект { "фраза": частота }
      for (const [phrase, freq] of Object.entries(dataArray)) {
        if (typeof freq === 'number') {
          items.push({ phrase: phrase.trim(), frequency: freq });
        }
      }
    }

    if (items.length === 0) {
      logWarning('⚠️ Парсер не нашёл ключевых фраз', {
        resultKeys: Object.keys(result || {}).join(','),
        dataType: typeof dataArray,
        isArray: Array.isArray(dataArray),
        sample: JSON.stringify(dataArray).substring(0, 500)
      }, context);
      return '';
    }

    // Дедупликация по фразе (lowercase), оставляем максимальную частоту
    const merged = new Map();
    for (const item of items) {
      const key = item.phrase.toLowerCase();
      if (!merged.has(key) || item.frequency > merged.get(key).frequency) {
        merged.set(key, item);
      }
    }

    // Сортировка по частотности (убывание)
    const sorted = Array.from(merged.values()).sort((a, b) => b.frequency - a.frequency);

    return sorted
      .filter(item => item.phrase)
      .map(item => item.frequency > 0 ? `${item.phrase} (${item.frequency})` : item.phrase)
      .join('\n');

  } catch (error) {
    logError('❌ Ошибка парсинга фраз wordstat', error, context);
    return 'Ошибка парсинга: ' + error.message;
  }
}

/**
 * Парсит CSV-ответ Arsenkin wordstat.
 * Формат: "фраза\tчастотность\n..." или "фраза;частотность\n..."
 * @param {string} csvText - CSV текст
 * @returns {Array<{phrase: string, frequency: number}>}
 */
function parseWordstatCsv_(csvText) {
  const lines = csvText.split('\n').map(l => l.trim()).filter(l => l);
  const items = [];

  for (const line of lines) {
    // Пробуем разные разделители: tab, точка с запятой, запятая
    let parts;
    if (line.includes('\t')) {
      parts = line.split('\t');
    } else if (line.includes(';')) {
      parts = line.split(';');
    } else {
      // Если нет разделителя — это просто фраза
      items.push({ phrase: line, frequency: 0 });
      continue;
    }

    const phrase = (parts[0] || '').trim();
    const freq = parseInt(parts[1] || '0') || 0;

    if (phrase && !phrase.match(/^(keyword|фраза|запрос)/i)) {
      items.push({ phrase: phrase, frequency: freq });
    }
  }

  return items;
}

/**
 * Парсит результат wordstat type=1 (сбор частотности).
 * Возвращает Map: фраза (lowercase) → частотность.
 * @param {Object} apiResult - Ответ Arsenkin API
 * @returns {Map<string, number>} Карта частотностей
 */
function parseWordstatFrequencyResult_(apiResult) {
  const context = 'parseWordstatFrequencyResult_';
  const freqMap = new Map();

  if (!apiResult) return freqMap;

  try {
    // Arsenkin /get для type=1 возвращает вложенную структуру:
    // { code, result: { queries, regions, result: { "фраза": {"213": {"base": freq}} } } }
    // Фразы в apiResult.result.result (двойная вложенность)

    // Навигация по структуре:
    // apiResult = { code, task_id, result: { type, task_id, data: { queries, regions, result: { "фраза": {"213": {"base": freq}} } } } }
    const outerResult = apiResult.result || apiResult;
    const dataObj = (outerResult && outerResult.data) || {};

    // Ищем объект с фразами — пробуем все пути
    let phrasesData = null;

    // Путь 1: result.data.result (основной для type=1)
    if (dataObj && dataObj.result && typeof dataObj.result === 'object') {
      phrasesData = dataObj.result;
    }
    // Путь 2: result.result (альтернативный)
    else if (outerResult && outerResult.result && typeof outerResult.result === 'object') {
      phrasesData = outerResult.result;
    }

    // Если phrasesData найден и это объект с фразами
    if (phrasesData && typeof phrasesData === 'object' && !Array.isArray(phrasesData)) {
      // Формат type=1: { "фраза": {"regionId": {"base": freq}} }
      for (const [phrase, regionData] of Object.entries(phrasesData)) {
        if (typeof regionData !== 'object' || regionData === null) continue;

        let freq = 0;
        if (regionData['213'] && typeof regionData['213'] === 'object' && regionData['213'].base !== undefined) {
          freq = parseInt(regionData['213'].base) || 0;
        } else if (regionData['0'] && typeof regionData['0'] === 'object' && regionData['0'].base !== undefined) {
          freq = parseInt(regionData['0'].base) || 0;
        } else {
          // Берём первый доступный регион
          for (const regionVal of Object.values(regionData)) {
            if (typeof regionVal === 'object' && regionVal !== null && regionVal.base !== undefined) {
              freq = parseInt(regionVal.base) || 0;
              break;
            }
          }
        }

        if (phrase.trim()) {
          freqMap.set(phrase.toLowerCase().trim(), freq);
        }
      }
    }

    // Путь 2: result.data (формат type=2)
    if (freqMap.size === 0) {
      const dataArray = (outerResult && outerResult.data) || apiResult.data;
      if (dataArray && Array.isArray(dataArray)) {
        for (const entry of dataArray) {
          if (entry.left && typeof entry.left === 'object') {
            for (const [phrase, freq] of Object.entries(entry.left)) {
              freqMap.set(phrase.toLowerCase().trim(), parseInt(freq) || 0);
            }
          }
        }
      }
    }

    // Путь 3: фразы прямо в apiResult.result (если result = { "фраза": {"213": {...}} })
    if (freqMap.size === 0 && outerResult && typeof outerResult === 'object' && !Array.isArray(outerResult)) {
      // Проверяем, похоже ли на региональную структуру
      const firstKey = Object.keys(outerResult)[0];
      const firstVal = outerResult[firstKey];
      if (firstVal && typeof firstVal === 'object' && (firstVal['213'] || firstVal['0'])) {
        logInfo('📊 Парсер: Путь 3 — фразы прямо в result', null, context);
        for (const [phrase, regionData] of Object.entries(outerResult)) {
          if (typeof regionData !== 'object' || regionData === null) continue;

          let freq = 0;
          if (regionData['213'] && typeof regionData['213'] === 'object') {
            freq = parseInt(regionData['213'].base) || 0;
          } else if (regionData['0'] && typeof regionData['0'] === 'object') {
            freq = parseInt(regionData['0'].base) || 0;
          }

          if (phrase.trim()) {
            freqMap.set(phrase.toLowerCase().trim(), freq);
          }
        }
      }
    }

    logInfo(`📊 Частотность: ${freqMap.size} фраз`, null, context);
    return freqMap;

  } catch (error) {
    logError('❌ Ошибка парсинга частотности', error, context);
    return freqMap;
  }
}

/**
 * Объединяет новые ключевики (с частотами) с существующими.
 * Дедупликация по чистой фразе (без частоты).
 * Сортировка по частотности (убывание).
 * @param {string} existing - Текущее содержимое ячейки
 * @param {string} newKeywords - Новые ключевики "фраза (частота), ..."
 * @returns {string} Объединённый и отсортированный список
 */
function mergeKeywordsWithFrequency_(existing, newKeywords) {
  if (!newKeywords || newKeywords.trim() === '') return existing;
  if (!existing || existing.trim() === '') return newKeywords;

  // Парсим оба списка в структуры {phrase, frequency}
  const existingItems = parseKeywordsString_(existing);
  const newItems = parseKeywordsString_(newKeywords);

  // Объединяем: новые данные обновляют частотность существующих
  const merged = new Map();
  for (const item of existingItems) {
    merged.set(item.phrase.toLowerCase(), item);
  }
  for (const item of newItems) {
    const key = item.phrase.toLowerCase();
    if (merged.has(key)) {
      // Обновляем частотность если новая больше
      const old = merged.get(key);
      if (item.frequency > old.frequency) {
        merged.set(key, { phrase: item.phrase, frequency: item.frequency });
      }
    } else {
      merged.set(key, item);
    }
  }

  // Сортируем по убыванию частотности
  const sorted = Array.from(merged.values()).sort((a, b) => b.frequency - a.frequency);

  return sorted
    .map(item => item.frequency > 0 ? `${item.phrase} (${item.frequency})` : item.phrase)
    .join('\n');
}

/**
 * Парсит строку ключевиков в массив {phrase, frequency}.
 * Поддерживает форматы: "фраза (123), фраза2 (456)" и "фраза\nфраза2"
 */
function parseKeywordsString_(text) {
  const items = text.split(/[,\n]/).map(s => s.trim()).filter(s => s);
  return items.map(item => {
    const match = item.match(/^(.+?)\s*\((\d+)\)\s*$/);
    if (match) {
      return { phrase: match[1].trim(), frequency: parseInt(match[2]) || 0 };
    }
    return { phrase: item.trim(), frequency: 0 };
  }).filter(item => item.phrase && !item.phrase.startsWith('ОШИБКА'));
}

/**
 * Обновляет частотности в существующем списке ключевиков и пересортирует.
 * @param {string} existing - Текущее содержимое ячейки (может быть без частот)
 * @param {Map<string, number>} freqMap - Карта частотностей
 * @returns {string} Обновлённый и отсортированный список
 */
function updateFrequenciesAndSort_(existing, freqMap) {
  if (!existing || existing.trim() === '') return '';
  if (freqMap.size === 0) return existing;

  const items = parseKeywordsString_(existing);

  // Обновляем частотности
  for (const item of items) {
    const key = item.phrase.toLowerCase();
    if (freqMap.has(key)) {
      item.frequency = freqMap.get(key);
    }
  }

  // Сортируем по убыванию
  items.sort((a, b) => b.frequency - a.frequency);

  return items
    .map(item => item.frequency > 0 ? `${item.phrase} (${item.frequency})` : item.phrase)
    .join('\n');
}

/**
 * Извлекает Title, Description и заголовки конкурентов из результата "check-h".
 * Всё форматируется в одну строку с блоками по конкурентам.
 *
 * Формат Arsenkin API /get для check-h:
 * { code: "TASK_RESULT", result: { result: [
 *   false,  // позиция без данных
 *   { url, title, description, headers: [{1: "level", value: "text"}, ...] },
 *   ...
 * ]}}
 *
 * @param {Object} apiResult - Ответ Arsenkin API
 * @returns {string} Форматированный текст для ячейки
 */
function parseCompetitorTagsResult_(apiResult) {
  const context = 'parseCompetitorTagsResult_';

  if (!apiResult) return '';

  try {
    // Arsenkin check-h: apiResult.result = объект-обёртка с ключом result
    // apiResult.result.result = массив конкурентов (с false на позициях без данных)
    let items = null;

    const outerResult = apiResult.result;

    // Путь 1: apiResult.result.result — основной формат check-h
    if (outerResult && typeof outerResult === 'object' && !Array.isArray(outerResult) && Array.isArray(outerResult.result)) {
      items = outerResult.result;
    }
    // Путь 2: apiResult.result — уже массив
    else if (Array.isArray(outerResult)) {
      items = outerResult;
    }
    // Путь 3: apiResult.data
    else if (Array.isArray(apiResult.data)) {
      items = apiResult.data;
    }

    if (!items || !Array.isArray(items)) {
      logWarning('⚠️ check-h: не удалось найти массив конкурентов', {
        keys: Object.keys(apiResult || {}).join(','),
        resultType: typeof outerResult,
        resultKeys: outerResult ? Object.keys(outerResult).join(',') : 'null',
        sample: JSON.stringify(apiResult).substring(0, 500)
      }, context);
      return '';
    }

    // Фильтруем false/null — позиции без данных
    const competitors = items.filter(item => item && typeof item === 'object');

    logInfo(`📊 check-h: ${competitors.length} конкурентов из ${items.length} позиций`, null, context);

    const blocks = [];

    for (let i = 0; i < competitors.length; i++) {
      const item = competitors[i];

      const lines = [];

      if (item.title) {
        lines.push('Title: ' + item.title);
      }
      if (item.description) {
        lines.push('Description: ' + item.description);
      }

      // Заголовки: только H1 (H2-H6 не нужны)
      // Формат API: {"i": "1", "value": "text"} — h.i = уровень H
      if (item.headers && Array.isArray(item.headers)) {
        for (const h of item.headers) {
          if (!h || !h.value) continue;
          if (h.i === '1' || h.i === 1) {
            lines.push('H1: ' + h.value);
          }
        }
      }

      if (lines.length > 0) {
        blocks.push(lines.join('\n'));
      }
    }

    return blocks.join('\n\n');

  } catch (error) {
    logError('❌ Ошибка парсинга тегов конкурентов', error, context);
    return 'Ошибка парсинга: ' + error.message;
  }
}

// ========================================
// СЕКЦИЯ 7: СТАТУС И УТИЛИТЫ
// ========================================

/**
 * Показывает текущий статус импорта из Arsenkin (вызывается из меню)
 */
function showArsenkinImportStatus() {
  const ui = SpreadsheetApp.getUi();

  const phrasesProgress = loadProgressChunked_(ARSENKIN_CONFIG.PROGRESS_KEYS.WORDSTAT_PHRASES);
  const freqProgress = loadProgressChunked_(ARSENKIN_CONFIG.PROGRESS_KEYS.WORDSTAT_FREQ);
  const chProgress = loadProgressChunked_(ARSENKIN_CONFIG.PROGRESS_KEYS.CHECK_H);

  let message = '📊 Статус импорта Arsenkin\n\n';

  if (phrasesProgress && phrasesProgress.tasks) {
    const done = phrasesProgress.tasks.filter(t => t.status === 'done').length;
    const pending = phrasesProgress.tasks.filter(t => t.status === 'pending').length;
    const errors = phrasesProgress.tasks.filter(t => t.status === 'error').length;
    message += `🔍 Парсинг ключевиков (Wordstat):\n`;
    message += `   Всего: ${phrasesProgress.tasks.length}\n`;
    message += `   Готово: ${done}\n`;
    message += `   Ожидают: ${pending}\n`;
    message += `   Ошибок: ${errors}\n`;
    message += `   Начато: ${phrasesProgress.startedAt}\n\n`;
  } else {
    message += `🔍 Парсинг ключевиков: нет активных задач\n\n`;
  }

  if (freqProgress && freqProgress.tasks) {
    const done = freqProgress.tasks.filter(t => t.status === 'done').length;
    const pending = freqProgress.tasks.filter(t => t.status === 'pending').length;
    const errors = freqProgress.tasks.filter(t => t.status === 'error').length;
    message += `📈 Сбор частотности:\n`;
    message += `   Всего: ${freqProgress.tasks.length}\n`;
    message += `   Готово: ${done}\n`;
    message += `   Ожидают: ${pending}\n`;
    message += `   Ошибок: ${errors}\n`;
    message += `   Начато: ${freqProgress.startedAt}\n\n`;
  } else {
    message += `📈 Сбор частотности: нет активных задач\n\n`;
  }

  if (chProgress && chProgress.tasks) {
    const done = chProgress.tasks.filter(t => t.status === 'done').length;
    const pending = chProgress.tasks.filter(t => t.status === 'pending').length;
    const errors = chProgress.tasks.filter(t => t.status === 'error').length;
    message += `📊 Теги конкурентов:\n`;
    message += `   Всего: ${chProgress.tasks.length}\n`;
    message += `   Готово: ${done}\n`;
    message += `   Ожидают: ${pending}\n`;
    message += `   Ошибок: ${errors}\n`;
    message += `   Начато: ${chProgress.startedAt}\n`;
  } else {
    message += `📊 Теги конкурентов: нет активных задач\n`;
  }

  ui.alert('Статус импорта', message, ui.ButtonSet.OK);
}

/**
 * Очищает сохранённый прогресс импорта Arsenkin
 */
function clearArsenkinProgress() {
  clearProgressChunked_(ARSENKIN_CONFIG.PROGRESS_KEYS.WORDSTAT_PHRASES);
  clearProgressChunked_(ARSENKIN_CONFIG.PROGRESS_KEYS.WORDSTAT_FREQ);
  clearProgressChunked_(ARSENKIN_CONFIG.PROGRESS_KEYS.CHECK_H);
  clearProgressChunked_(ARSENKIN_CONFIG.PROGRESS_KEYS.CLUSTERING);

  SpreadsheetApp.getUi().alert('Готово', 'Прогресс импорта Arsenkin очищен', SpreadsheetApp.getUi().ButtonSet.OK);
}


// ========================================
// СЕКЦИЯ 6: КЛАСТЕРИЗАЦИЯ
// ========================================

/**
 * Кластеризация запросов через Arsenkin
 * 1. Берет текущие ключевики из столбца G (Semantic Core)
 * 2. Отправляет на кластеризацию
 * 3. Распределяет новые группы по новым строкам
 */
/**
 * @param {boolean} [silent=false] — если true, пропускает UI-диалоги (для пайплайна)
 * @returns {{complete: boolean, pendingCount: number}|undefined}
 */
function clusterCategoriesWithArsenkin(silent) {
  const context = 'clusterCategoriesWithArsenkin';
  const ui = silent ? null : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressKey = ARSENKIN_CONFIG.PROGRESS_KEYS.CLUSTERING;

  try {
    if (!ARSENKIN_CONFIG.API_TOKEN) {
      if (!silent) ui.alert('Ошибка', 'Токен Arsenkin API не настроен.', ui.ButtonSet.OK);
      return { complete: false, pendingCount: 0 };
    }

    const savedProgress = loadProgressChunked_(progressKey);
    console.log(`[DEBUG] Starting clusterCategoriesWithArsenkin. Saved progress: ${savedProgress ? 'YES' : 'NO'}`);
    let progress;

    if (savedProgress && savedProgress.tasks && savedProgress.tasks.length > 0) {
      const pendingCount = savedProgress.tasks.filter(t => t.status === 'pending').length;
      if (pendingCount > 0) {
        if (silent) {
          progress = savedProgress;
        } else {
          const resume = ui.alert('Незавершённая кластеризация',
            `Найдено незавершённых задач: ${pendingCount}\nПродолжить?`,
            ui.ButtonSet.YES_NO);

          if (resume === ui.Button.YES) {
            progress = savedProgress;
          } else {
            clearProgressChunked_(progressKey);
            progress = null;
          }
        }
      } else {
        clearProgressChunked_(progressKey);
        progress = null;
      }
    }

    if (!progress) {
      const selected = getSelectedSeoTagRows_();
      if (selected.length === 0) {
        if (!silent) ui.alert('Нет отмеченных строк', 'Выделите строки на листе "SEO-теги".', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
      const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
      const validRows = [];

      for (const row of selected) {
        const rawKeywords = (sheet.getRange(row.row, cols.SEMANTIC_CORE).getValue() || '').toString();
        const phrases = extractPhrases_(rawKeywords);

        if (phrases.length > 2) {
          validRows.push({ row: row.row, phrases, originalH1: row.h1 });
        }
      }

      if (validRows.length === 0) {
        if (!silent) ui.alert('Нет данных', 'В выбранных строках слишком мало ключевиков (нужно > 2) в столбце G.', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      if (!silent) {
        const confirm = ui.alert('Кластеризация (Arsenkin)',
          `Будет создано ${validRows.length} задач.\n` +
          `Результатом будут НОВЫЕ СТРОКИ в таблице для выделенных групп.\n` +
          `Главная группа останется в исходной строке.\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;
      }

      progress = { tasks: [], startedAt: new Date().toISOString() };
      const cfg = ARSENKIN_CONFIG.CLUSTERING;

      for (let i = 0; i < validRows.length; i++) {
        const rowData = validRows[i];
        ss.toast(`Отправка: ${i + 1}/${validRows.length}...`, '📤 Кластеризация', 3);

        try {
          const taskData = {
            region: cfg.region,
            threshold: cfg.threshold,
            group: cfg.group,
            count: cfg.count,
            se: cfg.se,
            ws: ['base'],
            stoplist: [],
            main: false,
            depth: 10,
            queries: rowData.phrases.map(p => p.replace(/\s*\(\d+\)$/, '')),
            name: rowData.originalH1
          };

          const taskId = arsenkinSubmitTask_(cfg.tools_name, taskData);
          progress.tasks.push({ taskId, row: rowData.row, originalH1: rowData.originalH1, status: 'pending', result: null });
        } catch (error) {
          logError(`❌ Ошибка отправки кластеризации для строки ${rowData.row}: ${error.message}`, error, context);
          markRowError_(sheet, rowData.row, error.message, 'clustering', cols);
        }

        Utilities.sleep(ARSENKIN_CONFIG.DELAY_BETWEEN_SUBMITS_MS);
      }

      saveProgressChunked_(progressKey, progress);
    }

    arsenkinPollAndProcessClustering_(progress, progressKey, silent);

    const updatedProgress = loadProgressChunked_(progressKey);
    if (!updatedProgress) return { complete: true, pendingCount: 0 };
    const remaining = updatedProgress.tasks ? updatedProgress.tasks.filter(t => t.status === 'pending').length : 0;
    return { complete: remaining === 0, pendingCount: remaining };

  } catch (error) {
    logError('❌ Ошибка кластеризации', error, context);
    if (!silent) ui.alert('Ошибка: ' + error.message);
    return { complete: false, pendingCount: 0 };
  }
}

/**
 * Опрос и обработка результатов кластеризации (с вставкой строк)
 */
function arsenkinPollAndProcessClustering_(progress, progressKey, silent) {
  const context = 'arsenkinPollAndProcessClustering_';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
  const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;
  const startTime = Date.now();

  let insertedRowsTotal = 0;

  // Сортируем задачи ОТ ПОСЛЕДНЕЙ К ПЕРВОЙ по номеру строки, 
  // чтобы вставка строк не ломала индексы предыдущих задач.
  progress.tasks.sort((a, b) => b.row - a.row);

  while (Date.now() - startTime < ARSENKIN_CONFIG.EXECUTION_SAFETY_MS) {
    const pendingTasks = progress.tasks.filter(t => t.status === 'pending');
    if (pendingTasks.length === 0) break;

    // Safety timeout check (30 minutes)
    const MAX_EXECUTION_TIME_MS = 30 * 60 * 1000;
    const taskStartTime = new Date(progress.startedAt).getTime();

    if (Date.now() - taskStartTime > MAX_EXECUTION_TIME_MS) {
      logError('❌ Clustering timed out (exceeded 30 minutes)', null, context);

      // Marks all pending tasks as error
      pendingTasks.forEach(t => {
        t.status = 'error';
        t.error = 'Timeout: Execution exceeded 30 minutes';
        try {
          markRowError_(sheet, t.row, t.error, 'clustering', cols);
        } catch (e) {
          console.error(`Error marking row ${t.row}: ${e.message}`);
        }
      });

      break; // Exit loop
    }

    ss.toast(`Ожидание: ${pendingTasks.length}, готово: ${progress.tasks.filter(t => t.status === 'done').length}`, '⏳ Кластеризация', 5);

    // Берем ПЕРВУЮ (самую нижнюю) задачу из отсортированного списка
    const task = pendingTasks[0];

    try {
      const checkResult = arsenkinCheckTask_(task.taskId);

      // --- DEBUG LOGGING ---
      console.log(`[DEBUG] Polling task ${task.taskId}: status=${checkResult ? checkResult.status : 'null'}, progress=${checkResult ? checkResult.progress : '?'}, msg=${checkResult ? checkResult.msg : ''}`);
      // --- DEBUG LOGGING ---

      const isDone = checkResult && (
        checkResult.status === 'finish' || checkResult.status === 'done' || checkResult.status === 'completed' || checkResult.progress === 100
      );

      if (isDone) {
        const resultData = arsenkinGetResult_(task.taskId);
        task.status = 'done';
        task.result = resultData;

        // Вставляем строки и обновляем данные
        const newRowsCount = processClusteringResult_(sheet, task, cols);
        insertedRowsTotal += newRowsCount;

        logInfo(`✅ Кластеризация строки ${task.row}: +${newRowsCount} новых строк`, null, context);

      } else if (checkResult && (
        checkResult.status === 'error' ||
        checkResult.status === 'fail' ||
        (checkResult.msg && checkResult.msg.includes('Такого не должно было случиться'))
      )) {
        task.status = 'error';
        task.error = checkResult.error || checkResult.msg || 'Ошибка API';
        logError(`❌ Ошибка задачи ${task.taskId}: ${task.error}`, null, context);
        markRowError_(sheet, task.row, task.error, 'clustering', cols);
      }

    } catch (e) {
      logWarning(`Ошибка проверки задачи ${task.taskId}: ${e.message}`, null, context);
    }

    Utilities.sleep(ARSENKIN_CONFIG.POLL_INTERVAL_MS);
  }

  const pendingCount = progress.tasks.filter(t => t.status === 'pending').length;

  if (pendingCount > 0) {
    saveProgressChunked_(progressKey, progress);
    ss.toast(`Частично завершено. Ожидает: ${pendingCount}`, '⚠️ Таймаут', 10);
  } else {
    clearProgressChunked_(progressKey);
    ss.toast(`Все задачи обработаны! Добавлено строк: ${insertedRowsTotal}`, '✅ Готово', 5);
    if (!silent) {
      SpreadsheetApp.getUi().alert('Кластеризация завершена',
        `Добавлено новых групп (строк): ${insertedRowsTotal}`,
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Обработка результата кластеризации:
 * 1. Парсинг ответа
 * 2. Обновление исходной строки (главный кластер)
 * 3. Вставка новых строк (остальные кластеры)
 * @returns {number} Количество вставленных строк
 */
function processClusteringResult_(sheet, task, cols) {
  const data = task.result;

  // --- DEBUG LOGGING START ---
  console.log(`[DEBUG] Arsenkin Result (Task ${task.taskId}):`, JSON.stringify(data, null, 2));
  // --- DEBUG LOGGING END ---

  if (!data || !data.clusters) {
    // Попробуем найти clusters внутри data (вдруг формат другой)
    if (data.data && data.data.clusters) {
      // data = data.data; // sometimes wrapped
    } else {
      // throw new Error('Некорректный формат ответа от Arsenkin (нет clusters)');
    }
  }

  // Новая логика парсинга для формата data.clustering
  let clusters = [];

  // Handle various nesting levels: data.result.clustering (API full response) or data.clustering (direct result)
  const clusteringData = (data.result && data.result.clustering) || data.clustering || (data.data && data.data.clustering);

  // --- DEBUG DEEP DIVING ---
  console.log(`[DEBUG] processClusteringResult_ keys(data): ${Object.keys(data).join(', ')}`);
  if (data.result) console.log(`[DEBUG] keys(data.result): ${Object.keys(data.result).join(', ')}`);
  if (clusteringData) {
    console.log(`[DEBUG] clusteringData found. Type: ${typeof clusteringData}, Keys: ${Object.keys(clusteringData).join(', ')}`);
  } else {
    console.error(`[DEBUG] clusteringData NOT FOUND in data`);
  }
  // --- DEBUG DEEP DIVING ---

  if (clusteringData) {
    // Helper to extract clusters from an object
    const extractClusters = (obj) => {
      const result = [];
      Object.keys(obj).forEach(key => {
        const value = obj[key];
        // If it has 'words' property, it is a cluster
        if (value && value.words) {
          result.push({ name: key, data: value });
        }
        // If not, and it's an object, try to look inside (wrapper like "single")
        // But exclude 'words', 'topurl' etc strings/numbers just in case (though value.words check handles it)
        else if (value && typeof value === 'object') {
          Object.keys(value).forEach(subKey => {
            const subValue = value[subKey];
            if (subValue && subValue.words) {
              result.push({ name: subKey, data: subValue });
            }
          });
        }
      });
      return result;
    };

    const rawClusters = extractClusters(clusteringData);

    clusters = rawClusters.map(c => {
      const wordsObj = c.data.words || {};
      const phrases = Object.keys(wordsObj);

      // Сортируем фразы по частоте (ws), если она есть
      phrases.sort((a, b) => {
        const wsA = wordsObj[a] ? (wordsObj[a].ws || 0) : 0;
        const wsB = wordsObj[b] ? (wordsObj[b].ws || 0) : 0;
        return wsB - wsA;
      });

      // LOG keys and FULL structure for the first cluster to hunt for URLs
      if (phrases.length > 0 && c === rawClusters[0]) {
        console.log(`[Arsenkin] FULL DATA for first cluster '${c.name}':`, JSON.stringify(c.data));
        // Check if stats contains what we need
        if (data.result && data.result.stats) {
          console.log(`[Arsenkin] STATS DATA:`, JSON.stringify(data.result.stats));
        }
      }

      return {
        name: c.name,
        queries: phrases,
        // Собираем topurl со всех фраз кластера
        sites: (function () {
          // 1. Попытка взять явный список (если вдруг появится)
          if (c.data.sites && c.data.sites.length) return c.data.sites;
          if (c.data.urls && c.data.urls.length) return c.data.urls;

          // 2. Агрегация topurl из words
          if (c.data.words) {
            const allTopUrls = Object.values(c.data.words)
              .map(w => w.topurl)
              .filter(url => url && url !== '-' && url.trim() !== '');
            const uniqueUrls = [...new Set(allTopUrls)];
            if (uniqueUrls.length > 0) return uniqueUrls;
          }

          // 3. Фолбэк на общий topurl
          return c.data.topurl && c.data.topurl !== '-' ? [c.data.topurl] : [];
        })()
      };
    });
  } else {
    // Старая логика (резерв)
    clusters = data.clusters || (data.data && data.data.clusters) || [];
  }

  if (clusters.length === 0) {
    console.warn('[Arsenkin] Empty clusters result');
    return 0;
  }

  // 1. Определяем главный кластер (содержащий исходный H1)
  let mainClusterIndex = -1;
  const originalH1 = task.name ? task.name.toLowerCase().trim() : '';

  // Ищем кластер, где есть фраза, содержащая H1 (частичное вхождение)
  if (originalH1) {
    mainClusterIndex = clusters.findIndex(c =>
      c.queries.some(q => q.toLowerCase().includes(originalH1) || originalH1.includes(q.toLowerCase()))
    );
  }

  // Если не нашли по вхождению, берем самый большой кластер
  if (mainClusterIndex === -1) {
    // Сортируем по размеру, берем первый
    // (Но лучше оставить исходный порядок, если Arsenkin сортирует по релевантности)
    mainClusterIndex = 0;
  }

  const mainCluster = clusters[mainClusterIndex];

  // 2. Остальные кластеры (исключая главный)
  const otherClusters = [];
  const singletons = []; // For "Некластеризовано"

  clusters.forEach((c, index) => {
    if (index === mainClusterIndex) return;

    // Если в кластере > 1 фразы, это полноценная группа -> новая строка
    if (c.queries.length > 1) {
      otherClusters.push(c);
    } else {
      // Если фраза одна - собираем в "Некластеризовано"
      singletons.push(c);
    }
  });

  // Все одиночки собираем в ОДНУ группу "Некластеризовано"
  if (singletons.length > 0) {
    const unclusteredQueries = singletons.flatMap(c => c.queries);
    // Дедупликация
    const uniqueQueries = [...new Set(unclusteredQueries)];

    const unclusteredGroup = {
      name: "Некластеризовано",
      queries: uniqueQueries,
      sites: []
    };
    otherClusters.push(unclusteredGroup);
  }

  updateRowWithClusterData_(sheet, task.row, mainCluster, cols, true);

  // 2. Вставляем новые строки
  if (otherClusters.length > 0) {
    sheet.insertRowsAfter(task.row, otherClusters.length);

    for (let i = 0; i < otherClusters.length; i++) {
      const cluster = otherClusters[i];
      const newRowIndex = task.row + 1 + i;

      sheet.getRange(newRowIndex, cols.CHECKBOX).insertCheckboxes().uncheck();
      sheet.getRange(newRowIndex, cols.PAGE_NAME).setValue(cluster.name || cluster.queries[0]);

      updateRowWithClusterData_(sheet, newRowIndex, cluster, cols, false);

      // Подсветка новой строки
      sheet.getRange(newRowIndex, 1, 1, Object.keys(cols).length).setBackground('#f1f8e9');
    }
  }
  return otherClusters.length;
}

/**
 * Заполняет строку данными кластера
 */
function updateRowWithClusterData_(sheet, row, cluster, cols, isOriginalRow) {
  // 1. Семантическое ядро (H)
  const keywords = cluster.queries ? cluster.queries.join('\n') : '';
  sheet.getRange(row, cols.SEMANTIC_CORE).setValue(keywords);

  // 2. Конкуренты (I)
  let competitorsText = '';
  if (cluster.sites && Array.isArray(cluster.sites)) {
    competitorsText = cluster.sites.map(s => typeof s === 'object' ? s.url : s).join('\n');
  } else if (cluster.competitors && Array.isArray(cluster.competitors)) {
    competitorsText = cluster.competitors.join('\n');
  }

  sheet.getRange(row, cols.COMPETITORS_TITLE).setValue(competitorsText);
}

/**
 * Extracts phrases without frequency from a string.
 * Removes (123) and newlines
 * @param {string} rawText - Raw text from cell
 * @return {Array<string>} - Array of clean phrases
 */
/**
 * Extracts phrases from a string, preserving frequency.
 * Keeps "keyword (123)" format if present.
 * @param {string} rawText - Raw text from cell
 * @return {Array<string>} - Array of phrases
 */
function extractPhrases_(rawText) {
  if (!rawText) return [];

  return rawText.toString().split('\n')
    .map(function (line) {
      // Just trim, keep frequency
      return line.trim();
    })
    .filter(function (line) { return line.length > 0; });
}

// ========================================
// СЕКЦИЯ 8: ИМПОРТ LSI (ПОДСВЕТКИ)
// ========================================

/**
 * Импорт LSI-слов (подсветки из поисковой выдачи) через Arsenkin.
 * Берёт ключевые фразы из столбца G, отправляет в инструмент "sp",
 * записывает результат (тематические слова) в столбец H (LSI).
 *
 * @param {boolean} [silent=false] — если true, пропускает UI-диалоги (для пайплайна)
 * @returns {{complete: boolean, pendingCount: number}|undefined}
 */
function importLsiFromArsenkin(silent) {
  const context = 'importLsiFromArsenkin';
  const ui = silent ? null : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressKey = ARSENKIN_CONFIG.PROGRESS_KEYS.SEARCH_HIGHLIGHTS;

  try {
    if (!ARSENKIN_CONFIG.API_TOKEN) {
      if (!silent) ui.alert('Ошибка', 'Токен Arsenkin API не настроен.', ui.ButtonSet.OK);
      return { complete: false, pendingCount: 0 };
    }

    // Проверяем незавершённый прогресс
    const savedProgress = loadProgressChunked_(progressKey);
    let progress;

    if (savedProgress && savedProgress.tasks && savedProgress.tasks.length > 0) {
      const pendingCount = savedProgress.tasks.filter(t => t.status === 'pending').length;
      if (pendingCount > 0) {
        if (silent) {
          progress = savedProgress;
        } else {
          const resume = ui.alert('Незавершённый импорт LSI',
            `Найден незавершённый импорт подсветок:\n` +
            `Всего задач: ${savedProgress.tasks.length}\n` +
            `Ожидают: ${pendingCount}\n\n` +
            `Продолжить с места остановки?`,
            ui.ButtonSet.YES_NO);
          if (resume === ui.Button.YES) {
            progress = savedProgress;
          } else {
            clearProgressChunked_(progressKey);
            progress = null;
          }
        }
      } else {
        clearProgressChunked_(progressKey);
        progress = null;
      }
    }

    // Фаза отправки
    if (!progress) {
      const selected = getSelectedSeoTagRows_();
      if (selected.length === 0) {
        if (!silent) ui.alert('Нет данных', 'Отметьте чекбоксами строки для импорта LSI', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      // Получаем ключевые фразы из столбца G для каждой строки
      const sheet = ss.getSheetByName(SEO_TAGS_CONFIG.MASS_SHEET_NAME);
      const cols = SEO_TAGS_CONFIG.MASS_COLUMNS;

      const validRows = [];
      for (const row of selected) {
        const cellValue = (sheet.getRange(row.row, cols.SEMANTIC_CORE).getValue() || '').toString().trim();
        if (cellValue) {
          // Извлекаем фразы (без частотности) для запроса подсветок
          const phrases = cellValue.split(/[\n,]/)
            .map(p => p.replace(/\s*\(\d+\)\s*$/, '').trim())
            .filter(p => p.length > 0);
          if (phrases.length > 0) {
            validRows.push({ row: row.row, h1: row.h1, queries: phrases });
          }
        }
      }

      if (validRows.length === 0) {
        if (!silent) ui.alert('Нет данных', 'Выбранные строки не содержат ключевых фраз в столбце G', ui.ButtonSet.OK);
        return { complete: true, pendingCount: 0 };
      }

      if (!silent) {
        const confirm = ui.alert('Импорт LSI (подсветки → Арсенкин)',
          `Будет отправлено ${validRows.length} задач в Arsenkin.\n` +
          `Инструмент: Парсинг подсветок (sp)\n` +
          `Источник: ключевые фразы из столбца G\n` +
          `Результат → столбец H (LSI)\n\n` +
          `Продолжить?`,
          ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;
      }

      progress = { tasks: [], startedAt: new Date().toISOString() };
      const cfg = ARSENKIN_CONFIG.SEARCH_HIGHLIGHTS;

      for (let i = 0; i < validRows.length; i++) {
        const row = validRows[i];
        ss.toast(`Отправка LSI: ${i + 1}/${validRows.length}...`, '📤 Arsenkin', 3);

        try {
          const taskData = {
            queries: row.queries,
            se: cfg.se,
            region: cfg.region,
            depth: cfg.depth
          };

          const taskId = arsenkinSubmitTask_(cfg.tools_name, taskData);

          progress.tasks.push({
            taskId: taskId,
            row: row.row,
            query: row.h1,
            status: 'pending',
            result: null
          });

        } catch (error) {
          logError(`❌ Ошибка отправки LSI для строки ${row.row}: ${error.message}`, error, context);
          progress.tasks.push({
            taskId: null,
            row: row.row,
            query: row.h1,
            status: 'error',
            error: error.message
          });
        }

        if (i < validRows.length - 1) {
          Utilities.sleep(ARSENKIN_CONFIG.DELAY_BETWEEN_SUBMITS_MS);
        }
      }

      saveProgressChunked_(progressKey, progress);
      logInfo(`📤 Отправлено ${progress.tasks.filter(t => t.status === 'pending').length} задач LSI`, null, context);
    }

    // Фаза опроса
    arsenkinPollAndWriteResults_(progress, progressKey, 'search-highlights', silent);

    // Возвращаем статус для пайплайна
    const updatedProgress = loadProgressChunked_(progressKey);
    if (!updatedProgress) {
      return { complete: true, pendingCount: 0 };
    }
    const remaining = updatedProgress.tasks ? updatedProgress.tasks.filter(t => t.status === 'pending').length : 0;
    return { complete: remaining === 0, pendingCount: remaining };

  } catch (error) {
    logError('❌ Ошибка импорта LSI', error, context);
    if (!silent) ui.alert('Ошибка', 'Импорт LSI: ' + error.message, ui.ButtonSet.OK);
    return { complete: false, pendingCount: 0 };
  }
}

/**
 * Парсит результат инструмента "Подсветки" (sp).
 * Возвращает строку с LSI-словами через запятую.
 * @param {Object} apiResult - Результат /get от Arsenkin
 * @returns {string} - LSI-слова через запятую
 */
function parseSearchHighlightsResult_(apiResult) {
  const context = 'parseSearchHighlightsResult_';

  try {
    if (!apiResult) return '';

    const result = apiResult.result || apiResult;

    // Собираем слова с частотностью: Map<word, freq>
    const wordFreqMap = new Map();

    // 1. Подсвеченные слова: result.hlWords.normal — { "слово": частота }
    let hlCount = 0;
    if (result && result.hlWords && result.hlWords.normal && typeof result.hlWords.normal === 'object') {
      for (const [word, freq] of Object.entries(result.hlWords.normal)) {
        const w = word.toLowerCase().trim();
        if (w) {
          wordFreqMap.set(w, Math.max(wordFreqMap.get(w) || 0, parseInt(freq) || 0));
          hlCount++;
        }
      }
    }

    // 1b. Лемматизированные слова из запросов: result.queryWords — { "слово": частота }
    // Веб-интерфейс Арсенкина включает их в "Слова, задающие тематику"
    let qwCount = 0;
    if (result && result.queryWords && typeof result.queryWords === 'object' && !Array.isArray(result.queryWords)) {
      for (const [word, freq] of Object.entries(result.queryWords)) {
        const w = word.toLowerCase().trim();
        if (w) {
          wordFreqMap.set(w, Math.max(wordFreqMap.get(w) || 0, parseInt(freq) || 0));
          qwCount++;
        }
      }
      logInfo(`[DEBUG] result.queryWords: ${qwCount} слов, первые: ${Object.keys(result.queryWords).slice(0, 5).join(', ')}`, null, context);
    }

    // 2. LSI (тематические слова) — основной блок данных
    let lsiCount = 0;
    if (result && result.LSI) {
      // === DEBUG: Полная диагностика result.LSI ===
      logInfo(`[DEBUG LSI] тип: ${typeof result.LSI}, isArray: ${Array.isArray(result.LSI)}`, null, context);
      const lsiStr = JSON.stringify(result.LSI);
      logInfo(`[DEBUG LSI] JSON (первые 1500): ${lsiStr.substring(0, 1500)}`, null, context);
      if (lsiStr.length > 1500) {
        logInfo(`[DEBUG LSI] JSON (1500-3000): ${lsiStr.substring(1500, 3000)}`, null, context);
      }
      if (Array.isArray(result.LSI)) {
        logInfo(`[DEBUG LSI] длина массива: ${result.LSI.length}`, null, context);
        if (result.LSI.length > 0) {
          logInfo(`[DEBUG LSI] [0] тип: ${typeof result.LSI[0]}, значение: ${JSON.stringify(result.LSI[0]).substring(0, 300)}`, null, context);
          if (result.LSI.length > 1) {
            logInfo(`[DEBUG LSI] [1] тип: ${typeof result.LSI[1]}, значение: ${JSON.stringify(result.LSI[1]).substring(0, 300)}`, null, context);
          }
        }
      } else if (typeof result.LSI === 'object') {
        const keys = Object.keys(result.LSI);
        logInfo(`[DEBUG LSI] ключей: ${keys.length}, первые 10: ${keys.slice(0, 10).join(', ')}`, null, context);
        if (keys.length > 0) {
          logInfo(`[DEBUG LSI] значение первого ключа "${keys[0]}": ${JSON.stringify(result.LSI[keys[0]])}`, null, context);
        }
      }
      // === Конец DEBUG ===

      const addLsiWord = (word, freq) => {
        const w = word.toLowerCase().trim();
        if (w) {
          wordFreqMap.set(w, Math.max(wordFreqMap.get(w) || 0, parseInt(freq) || 0));
          lsiCount++;
        }
      };

      if (Array.isArray(result.LSI)) {
        for (const item of result.LSI) {
          if (typeof item === 'string') {
            addLsiWord(item, 0);
          } else if (item && typeof item === 'object') {
            const w = item.word || item.text || item.name || '';
            if (w) {
              addLsiWord(w, item.count || item.freq || item.cnt || 0);
            } else {
              for (const [key, val] of Object.entries(item)) {
                if (typeof val === 'number') addLsiWord(key, val);
              }
            }
          }
        }
      } else if (typeof result.LSI === 'object') {
        for (const [word, freq] of Object.entries(result.LSI)) {
          addLsiWord(word, freq);
        }
      }

      logInfo(`[DEBUG LSI] Извлечено тематических слов: ${lsiCount}`, null, context);
    } else {
      logInfo(`[DEBUG LSI] result.LSI ОТСУТСТВУЕТ! Ключи result: ${result ? Object.keys(result).join(', ') : 'null'}`, null, context);
    }

    // Сортировка по убыванию частотности, затем по алфавиту
    const sorted = [...wordFreqMap.entries()]
      .filter(([w]) => w.length > 0)
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));

    // Формат: одно слово на строку, с частотностью если > 0
    const lines = sorted.map(([word, freq]) => freq > 0 ? `${word} (${freq})` : word);

    logInfo(`📊 LSI: найдено ${sorted.length} слов (подсветки: ${hlCount}, слова запросов: ${qwCount}, тематические: ${lsiCount})`, null, context);
    return lines.join('\n');

  } catch (error) {
    logError('❌ Ошибка парсинга подсветок', error, context);
    return '';
  }
}

// Removed PAA processing
