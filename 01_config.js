/**
 * ========================================
 * КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
 * ========================================
 * 
 * Все ключи и настройки в одном месте
 */

// ========================================
// НАСТРОЙКИ INSALES
// ========================================

const INSALES_CONFIG = {
  // API ключи InSales
  apiKey: '63798fce31db4378484c6d6a8afd8674',
  password: 'd3e999e5e5f48ff73415e1ad561f927f',
  shop: 'binokl.shop',

  // ИСПРАВЛЕНО: правильный домен админки InSales
  adminDomain: 'myshop-on665.myinsales.ru',
  adminPath: '/admin2',  // ВАЖНО: admin2, не admin!

  get baseUrl() {
    return `https://${this.shop}`;
  },

  // Новый геттер для ссылок на админку
  get adminBaseUrl() {
    return `https://${this.adminDomain}${this.adminPath}`;
  }
};

// ========================================
// ГЛОБАЛЬНЫЕ ФУНКЦИИ АВТОРИЗАЦИИ
// ========================================

/**
 * Получает учетные данные InSales (Синхронная версия)
 * Централизованная функция для всего проекта
 */
function getInsalesCredentialsSync() {
  try {
    const config = getInsalesConfig();

    if (!config || !config.apiKey || !config.password || !config.shop) {
      throw new Error('Учетные данные InSales не настроены в 01_config.gs');
    }

    return {
      apiKey: config.apiKey,
      password: config.password,
      shop: config.shop,
      baseUrl: config.baseUrl
    };
  } catch (error) {
    console.error('❌ Ошибка получения учетных данных:', error);
    return null;
  }
}

/**
 * УСТАРЕВШАЯ ФУНКЦИЯ - Оставлена для совместимости
 * Если где-то остался старый вызов getInSalesCredentials(), он попадет сюда
 */
function getInSalesCredentials() {
  console.warn('⚠️ Используется устаревшая функция getInSalesCredentials. Пожалуйста, обновите код на getInsalesCredentialsSync.');
  return getInsalesCredentialsSync();
}

// ========================================
// НАСТРОЙКИ OPENAI
// ========================================

const OPENAI_CONFIG = {
  // ВАЖНО: Замените на ваш реальный ключ OpenAI
  apiKey: 'YOUR_OPENAI_API_KEY_HERE',

  // Модель для генерации (можно менять)
  model: 'gpt-4o-mini',

  // Температура генерации (0.0 - 1.0)
  // Меньше = более предсказуемо, больше = более креативно
  temperature: 0.7,

  // Максимум токенов в ответе
  maxTokens: 4000
};

// ========================================
// НАСТРОЙКИ GEMINI (GOOGLE)
// ========================================

// ========================================
// НАСТРОЙКИ GEMINI (GOOGLE)
// ========================================

const GOOGLE_GEMINI_V2_CONFIG = {
  apiKey: 'AIzaSyBCP2-mHQvJjHB6J_4STCJi592hqx8m3Ts',
  model: 'gemini-3-pro-preview',
  temperature: 0.7,
  maxOutputTokens: 8192
};

const GOOGLE_GEMINI_IMAGE_CONFIG = {
  apiKey: 'AIzaSyBCP2-mHQvJjHB6J_4STCJi592hqx8m3Ts', // Используем тот же ключ
  model: 'gemini-3-pro-image-preview',
  driveFolder: 'AI_Generated_Banners'
};

const CATEGORY_DESCRIPTION_ASSISTANT_ID = 'asst_qVFYH8Q5qgzMKsvOOzwnAuHN';

// ========================================
// НАСТРОЙКИ YANDEX METRICA
// ========================================

const YANDEX_METRICA_CONFIG = {
  // ID счетчика (из вашего скриншота)
  counterId: '41613924',

  // OAuth токен (нужно получить)
  // ВАЖНО: Токен должен иметь права на Метрику (ym:r:read) И Вебмастер (webmaster:host:information, webmaster:report:analytics)
  // Если Вебмастер выдает ошибку 403, обновите токен, добавив права в https://oauth.yandex.ru/
  oauthToken: 'y0__xDkppDPARj4jzwgneL2xRXiGf55eP6f3E44-SD7tOQGB29BSw',

  // Базовый URL API
  baseUrl: 'https://api-metrika.yandex.net/stat/v1/data'
};

// ========================================
// НАСТРОЙКИ СЕМАНТИКИ (ОПЦИОНАЛЬНО)
// ========================================

const SEMANTICS_CONFIG = {
  // Выберите сервис: 'serpstat', 'ahrefs', 'semrush' или null
  service: null,  // Если null - ручной ввод позиций

  // API ключ для выбранного сервиса
  apiKey: '',

  // Регион для проверки (для Serpstat)
  region: 'g_ru'  // Google Россия
};

// ========================================
// НАСТРОЙКИ ПРОВЕРКИ ПОЗИЦИЙ
// ========================================

const POSITIONS_CONFIG = {
  // Автоматическая проверка позиций (если true - нужен API)
  autoCheck: false,

  // Сервис для проверки позиций: null, 'serpstat', 'serankings'
  service: null,

  // API ключ для сервиса проверки позиций
  apiKey: '',

  // Глубина проверки (ТОП-10, ТОП-30, ТОП-100)
  depth: 100,

  // Поисковая система по умолчанию
  defaultSearchEngine: 'yandex',  // 'yandex' или 'google'

  // Регионы для Вордстата
  WORDSTAT_REGIONS: {
    '0': 'Весь мир',
    '225': 'Россия',
    '213': 'Москва',
    '1': 'Москва и область',
    '2': 'Санкт-Петербург',
    '10174': 'Санкт-Петербург и область'
  }
};

// ========================================
// НАСТРОЙКИ ПРИЛОЖЕНИЯ
// ========================================

const APP_CONFIG = {
  // Название приложения
  name: 'InSales Manager',
  version: '2.0',

  // Максимальное количество товаров для загрузки за раз
  maxProductsPerPage: 250,

  // Максимальное количество категорий для загрузки
  maxCategoriesPages: 20,

  // Задержка между API запросами (мс)
  apiDelay: 500,

  // Количество повторных попыток при ошибке API
  apiRetries: 3,

  // Логирование в лист
  enableSheetLogging: true,

  // Максимальное количество записей в логе
  maxLogRecords: 1000
};

// ========================================
// НАЗВАНИЯ ЛИСТОВ
// ========================================

const CATEGORY_SHEETS = {
  MAIN_LIST: 'Категории — Список',
  DETAIL_PREFIX: 'Категория — ',
  KEYWORDS: 'Ключевые слова',
  SEMANTICS: 'Семантика',
  LSI_WORDS: 'LSI и тематика',
  IMPORT_QA: 'Импорт Q&A',
  PRODUCTS_CATALOG: 'Каталог для подбора',
  POSITIONS_HISTORY: 'История позиций'
};

// ========================================
// КОЛОНКИ ГЛАВНОГО ЛИСТА
// ========================================

const MAIN_LIST_COLUMNS = {
  CHECKBOX: 1,              // A
  CATEGORY_ID: 2,           // B
  PARENT_ID: 3,             // C
  LEVEL: 4,                 // D
  HIERARCHY_PATH: 5,        // E
  TITLE: 6,                 // F
  URL: 7,                   // G
  PRODUCTS_COUNT: 8,        // H
  IN_STOCK_COUNT: 9,        // I
  SEO_STATUS: 10,           // J
  AI_STATUS: 11,            // K
  LAST_UPDATED: 12,         // L
  ADMIN_LINK: 13,           // M
  VISITS: 14,               // N - Визиты из Яндекс.Метрики
  BOUNCE_RATE: 15,          // O - Отказы
  ORDERS: 16,               // P - Заказы (Цель 54410371 - Оформление заказа)
  GSC_IMPRESSIONS: 17,      // Q - Показы из Google Search Console (GSC)

  // Показы из Яндекс.Вебмастера (ВРЕМЕННО ОТКЛЮЧЕНО из-за проблем с OAuth)
  WEBMASTER_IMPRESSIONS: null, // Было 18 (R)

  // Период сбора данных
  PERIOD: 19                // S - Период сбора статистики
};

// ========================================
// МАРКЕРЫ СЕКЦИЙ ДЕТАЛЬНОГО ЛИСТА
// ========================================
// Используются для поиска секций по заголовкам (текстовый поиск)

const DETAIL_SHEET_SECTION_MARKERS = {
  // Заголовки секций (emoji-маркеры для поиска)
  INFO_HEADER: '📁 ИНФОРМАЦИЯ О КАТЕГОРИИ',
  SEO_HEADER: '🎯 SEO ДАННЫЕ',
  KEYWORDS_HEADER: '📝 КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПЛИТОК ТЕГОВ',
  STATS_HEADER: '📊 СТАТИСТИКА ТОВАРОВ',
  PRODUCTS_HEADER: '🛒 ТЕКУЩИЕ ТОВАРЫ В КАТЕГОРИИ',
  FIELDS_HEADER: '⚙️ ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ КАТЕГОРИИ',
  UPPER_TILE_HEADER: '🏷️ ПЛИТКА ТЕГОВ - ВЕРХНЯЯ',
  LOWER_TILE_HEADER: '🏷️ ПЛИТКА ТЕГОВ - НИЖНЯЯ',

  // Под-маркеры внутри секций
  HTML_CODE_MARKER: 'HTML код (финальный):',
  INSTRUCTION_TEXT: 'Инструкция:',

  // Архитектурные константы (НИКОГДА НЕ МЕНЯЮТСЯ!)
  CATEGORY_ID_CELL: 'B2',
  PRODUCTS_COUNT_CELL: 'B33', // Всего товаров (используется для расчетов)

  // Фиксированные строки (только для секций с известной структурой)
  INFO_BLOCK_START: 1,
  INFO_BLOCK_END: 11,
  SEO_BLOCK_START: 12,
  SEO_BLOCK_END: 21,
  KEYWORDS_TABLE_START: 22,  // ИСПРАВЛЕНО: было 18, фактически 22

  // Смещения (только когда текстовый поиск невозможен)
  HEADER_TO_DATA_OFFSET: 2, // Заголовок → Заголовки столбцов → Данные
  SECTION_SPACING: 3 // Пустые строки между секциями
};

// ========================================
// СЕКЦИИ ДЕТАЛЬНОГО ЛИСТА (УСТАРЕВШИЕ)
// ========================================
// ВНИМАНИЕ: Эти константы устарели и будут удалены в будущих версиях
// Используйте DETAIL_SHEET_SECTION_MARKERS и findSectionByMarker() вместо них

const DETAIL_SHEET_SECTIONS = {
  INFO_START: 1,
  CATEGORY_ID_CELL: 'B2',
  TITLE_CELL: 'B3',
  URL_CELL: 'B4',
  PARENT_PATH_CELL: 'B5',

  SEO_START: 12,
  SEO_TITLE_CELL: 'B13',
  SEO_DESCRIPTION_CELL: 'B14',
  SEO_H1_CELL: 'B15',
  SEO_KEYWORDS_CELL: 'B16',

  // ============================================
  // СЕКЦИЯ ПЛИТОК ТЕГОВ (после SEO, до товаров)
  // ============================================

  // Таблица ключевых слов для плиток (ручной ввод)
  TAG_KEYWORDS_START: 22,  // ИСПРАВЛЕНО: было 18, фактически 22
  TAG_KEYWORDS_HEADER_ROW: 23,  // ИСПРАВЛЕНО: было 19
  TAG_KEYWORDS_DATA_START: 24,  // ИСПРАВЛЕНО: было 20
  TAG_KEYWORDS_COLUMNS: {
    CHECKBOX: 1,           // A: Чекбокс для включения в генерацию
    KEYWORD: 2,            // B: Ключевое слово
    TILE_TYPE: 3,          // C: Тип плитки (Верхняя/Нижняя)
    ANCHOR_TEXT: 4,        // D: Текст анкора
    CATEGORY_LINK: 5,      // E: URL или ID категории
    CATEGORY_STATUS: 6,    // F: Статус (Существует/Создать/Не указана)
    PARENT_CATEGORY: 7,    // G: ID родительской категории (для новых)
    PICKER_LINK: 8         // H: Ссылка для вызова диалога подбора категории
  },

  // Сгенерированные плитки (динамическое размещение)
  // Эти значения рассчитываются функцией getTagTilesPositions()
  TAG_TILES_UPPER_GENERATED: null,  // Рассчитывается динамически
  TAG_TILES_LOWER_GENERATED: null,  // Рассчитывается динамически

  // Список товаров (динамическое размещение)
  CURRENT_PRODUCTS_START: 27,  // УСТАРЕЛО - используйте getProductsStartRow()
  CURRENT_PRODUCTS_COLUMNS: {
    CHECKBOX: 1,
    PRODUCT_ID: 2,
    ARTICLE: 3,
    TITLE: 4,
    IN_STOCK: 5,
    PRICE: 6,
    BRAND: 7,
    CHARACTERISTICS: 8
  },

  PRODUCT_PICKER_START: 100,
  AI_GENERATION_START: 200
};

// ========================================
// ENDPOINTS INSALES API
// ========================================

const CATEGORY_API_ENDPOINTS = {
  COLLECTIONS: '/admin/collections.json',
  COLLECTION_BY_ID: '/admin/collections/{id}.json',
  COLLECTION_PRODUCTS: '/admin/products.json?collection_id={id}',
  UPDATE_COLLECTION: '/admin/collections/{id}.json',
  PRODUCTS: '/admin/products.json',
  COLLECTS: '/admin/collects.json'
};

// ========================================
// API СЕМАНТИКИ
// ========================================

const SEMANTICS_API = {
  SERPSTAT: {
    endpoint: 'https://api.serpstat.com/v4/',
    methods: {
      keywords: 'keywords_suggestions',
      search_volume: 'search_analytics_overview'
    }
  },

  AHREFS: {
    endpoint: 'https://apiv2.ahrefs.com/',
    methods: {
      keywords: 'keywords-explorer/v2/keywords'
    }
  },

  SEMRUSH: {
    endpoint: 'https://api.semrush.com/',
    methods: {
      keywords: 'phraseorganic'
    }
  }
};

// ========================================
// КОНФИГУРАЦИЯ ПЛИТОК ТЕГОВ
// ========================================

const TAG_TILES_CONFIG = {
  // Пороги плотности ключевых слов (TF в процентах)
  DENSITY_THRESHOLDS: {
    MAIN_KEYWORD: { min: 2.0, max: 4.0 },      // Главный ключевик: 2-4%
    ADDITIONAL: { min: 1.0, max: 2.0 },        // Дополнительные: 1-2%
    LSI: { min: 0.5, max: 1.0 },               // LSI слова: 0.5-1%
    SPAM_THRESHOLD: 5.0                         // Порог переспама: >5%
  },

  // Лимиты количества анкоров
  TOP_TILE_LIMIT: { min: 5, max: 8 },          // Верхняя плитка: 5-8 анкоров
  BOTTOM_TILE_LIMIT: { min: 15, max: 30 },     // Нижняя плитка: 15-30 анкоров

  // Настройки AI для генерации анкоров
  AI_PROVIDER: 'openai',  // 'gemini', 'claude' или 'openai'
  AI_MODEL: 'gpt-4o-mini',
  AI_TEMPERATURE: 0.7,
  AI_MAX_TOKENS: 4000,

  // Названия листов
  SHEET_NAME: 'Плитки тегов',
  DENSITY_ANALYSIS_SHEET: 'Анализ плотности',

  // HTML классы
  CSS_CLASSES: {
    TOP_TILE: 'category-tiles category-tiles--top',
    BOTTOM_TILE: 'category-tiles category-tiles--bottom',
    TILE_GRID: 'tiles-grid',
    TILE: 'tile',
    TILE_SMALL: 'tile tile--small'
  }
};

// Колонки листа "Плитки тегов"
const TAG_TILES_COLUMNS = {
  CHECKBOX: 1,          // A - Чекбокс для выбора
  CATEGORY_ID: 2,       // B - ID категории
  CATEGORY_NAME: 3,     // C - Название категории
  POSITION: 4,          // D - Позиция ('Верх' или 'Низ')
  ANCHOR: 5,            // E - Текст анкора
  KEYWORDS: 6,          // F - Покрываемые ключевики
  LINK: 7,              // G - Ссылка на категорию
  STATUS: 8,            // H - Статус (Создан/Применен/Требует обновления)
  HTML: 9,              // I - HTML код анкора
  CREATED_AT: 10,       // J - Дата создания
  APPLIED_AT: 11        // K - Дата применения
};

// Колонки листа "Анализ плотности"
const DENSITY_ANALYSIS_COLUMNS = {
  CATEGORY_ID: 1,           // A - ID категории
  CATEGORY_NAME: 2,         // B - Название категории
  KEYWORD: 3,               // C - Ключевое слово
  TYPE: 4,                  // D - Тип (Целевой/LSI/Тематика)
  OCCURRENCES: 5,           // E - Количество вхождений
  TOTAL_WORDS: 6,           // F - Всего слов на странице
  TF_PERCENT: 7,            // G - TF в процентах
  TARGET_MIN: 8,            // H - Целевой минимум
  TARGET_MAX: 9,            // I - Целевой максимум
  STATUS: 10,               // J - Статус (Достаточно/Недостаточно/Переспам)
  USED_IN_TILE: 11,         // K - Использовано в плитке (Да/Нет)
  ANCHOR_TEXT: 12,          // L - Текст анкора где использовано
  ANALYZED_AT: 13           // M - Дата анализа
};

// ========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ========================================

/**
 * Получает конфигурацию InSales
 */
function getInsalesConfig() {
  return INSALES_CONFIG;
}

/**
 * Получает конфигурацию OpenAI
 */
function getOpenAIConfig() {
  // СНАЧАЛА пробуем взять из Script Properties
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKeyFromProps = scriptProps.getProperty('OPENAI_API_KEY');

  // Если есть в Properties - используем его
  if (apiKeyFromProps) {
    return {
      apiKey: apiKeyFromProps,
      model: OPENAI_CONFIG.model || 'gpt-4',
      temperature: OPENAI_CONFIG.temperature || 0.7,
      maxTokens: OPENAI_CONFIG.maxTokens || 4000
    };
  }

  // Иначе берем из константы
  return OPENAI_CONFIG;
}

/**
 * Получает настройки семантики
 */
function getSemanticsConfig() {
  return SEMANTICS_CONFIG;
}

/**
 * Получает настройки позиций
 */
function getPositionsConfig() {
  return POSITIONS_CONFIG;
}

/**
 * Проверяет, настроен ли OpenAI
 */
function isOpenAIConfigured() {
  return OPENAI_CONFIG.apiKey &&
    OPENAI_CONFIG.apiKey.length > 20 &&
    OPENAI_CONFIG.apiKey.startsWith('YOUR_OPENAI_API_KEY_HERE');
}

/**
 * Проверяет, настроен ли InSales
 */
function isInsalesConfigured() {
  return INSALES_CONFIG.apiKey &&
    INSALES_CONFIG.password &&
    INSALES_CONFIG.shop;
}

/**
 * Показывает диалог настройки API ключей
 */
function showAPIKeysSetup() {
  const ui = SpreadsheetApp.getUi();

  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { 
            font-family: Arial, sans-serif; 
            padding: 20px;
            max-width: 600px;
          }
          h3 { color: #1976d2; margin-top: 0; }
          .section { 
            background: #f5f5f5; 
            padding: 15px; 
            margin: 15px 0;
            border-radius: 4px;
          }
          .key-info {
            background: white;
            padding: 10px;
            margin: 10px 0;
            border-left: 4px solid #4caf50;
            font-family: monospace;
            word-break: break-all;
          }
          .warning {
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 10px;
            margin: 10px 0;
          }
          .success {
            background: #d4edda;
            border-left: 4px solid #28a745;
            padding: 10px;
            margin: 10px 0;
          }
          .error {
            background: #f8d7da;
            border-left: 4px solid #dc3545;
            padding: 10px;
            margin: 10px 0;
          }
          button {
            background: #1976d2;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 4px;
            cursor: pointer;
            margin: 5px;
          }
          button:hover {
            background: #1565c0;
          }
        </style>
      </head>
      <body>
        <h3>⚙️ Настройка API ключей</h3>
        
        <div class="section">
          <h4>InSales API</h4>
          <div class="success">✅ Настроено в файле конфигурации</div>
          <div class="key-info">
            Магазин: ${INSALES_CONFIG.shop}<br>
            API Key: ${INSALES_CONFIG.apiKey.substring(0, 20)}...
          </div>
          <button onclick="testInsales()">Проверить подключение</button>
        </div>
        
        <div class="section">
          <h4>OpenAI API</h4>
          ${isOpenAIConfigured() ?
      `<div class="success">✅ Настроено</div>
             <div class="key-info">API Key: ${OPENAI_CONFIG.apiKey.substring(0, 20)}...</div>` :
      `<div class="error">❌ Не настроено</div>
             <div class="warning">
               Откройте файл <strong>01_config.gs</strong> и замените:<br><br>
               <code>apiKey: 'sk-proj-ВАШ_КЛЮЧ_OPENAI_СЮДА'</code><br><br>
               на ваш реальный ключ OpenAI
             </div>`
    }
          <button onclick="testOpenAI()" ${!isOpenAIConfigured() ? 'disabled' : ''}>
            Проверить подключение
          </button>
        </div>
        
        <div class="section">
          <h4>Семантика (опционально)</h4>
          ${SEMANTICS_CONFIG.service ?
      `<div class="success">✅ Настроено: ${SEMANTICS_CONFIG.service}</div>` :
      `<div class="warning">⚠️ Не настроено - будет использоваться только OpenAI для LSI</div>`
    }
        </div>
        
        <div class="section">
          <h4>Проверка позиций</h4>
          ${POSITIONS_CONFIG.service ?
      `<div class="success">✅ Настроено: ${POSITIONS_CONFIG.service}</div>` :
      `<div class="warning">⚠️ Не настроено - будет ручной ввод позиций</div>`
    }
        </div>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd;">
          <button onclick="google.script.host.close()">Закрыть</button>
        </div>
        
        <script>
          function testInsales() {
            google.script.run
              .withSuccessHandler(function(result) {
                alert(result ? '✅ InSales подключен успешно!' : '❌ Ошибка подключения к InSales');
              })
              .testInsalesConnection();
          }
          
          function testOpenAI() {
            google.script.run
              .withSuccessHandler(function(result) {
                alert(result ? '✅ OpenAI подключен успешно!' : '❌ Ошибка подключения к OpenAI');
              })
              .testOpenAIConnection();
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(650)
    .setHeight(600);

  ui.showModalDialog(htmlOutput, 'Настройка API');
}
// ========================================
// КОНФИГУРАЦИЯ ХАРАКТЕРИСТИК ДЛЯ РАЗДЕЛОВ
// ========================================

/**
 * Основные характеристики для фильтрации товаров в каждом разделе
 * Показываются в интерфейсе добавления товаров в категории
 */
const SECTION_CHARACTERISTICS = {
  'Бинокли': [
    'Кратность увеличения, крат',
    'Диаметр объектива, мм',
    'Поле зрения на 1000 метров',
    'Оптическое покрытие',
    'Марка стекла',
    'Призменная схема',
    'Бренд',
    'Серия',
    'Тип кратности',
    'Способ фокусировки',
    'Фокусное расстояние, м',
    'Защита от влаги и пыли',
    'Газовое заполнение',
    'Вес, г'
  ],
  'Зрительные трубы': [
    'Кратность увеличения, крат',
    'Диаметр объектива, мм',
    'Поле зрения на 1000 метров',
    'Оптическое покрытие',
    'Марка стекла',
    'Бренд',
    'Серия',
    'Тип кратности',
    'Способ фокусировки',
    'Фокусное расстояние, м',
    'Защита от влаги и пыли',
    'Газовое заполнение',
    'Вес, г',
    'Угол наклона окуляра, градус'
  ],
  'Монокуляры': [
    'Кратность увеличения, крат',
    'Диаметр объектива, мм',
    'Поле зрения на 1000 метров',
    'Оптическое покрытие',
    'Марка стекла',
    'Призменная схема',
    'Бренд',
    'Серия',
    'Тип кратности',
    'Способ фокусировки',
    'Фокусное расстояние, м',
    'Защита от влаги и пыли',
    'Газовое заполнение',
    'Вес, г'
  ],
  'Прицелы': [
    'Кратность увеличения, крат',
    'Диаметр объектива, мм',
    'Тип прицельной сетки',
    'Оптическое покрытие',
    'Марка стекла',
    'Бренд',
    'Серия',
    'Тип кратности',
    'Диаметр трубы прицела, мм',
    'Вынос выходного зрачка, мм',
    'Защита от влаги и пыли',
    'Газовое заполнение',
    'Вес, г',
    'Подсветка сетки'
  ],
  'Тепловизоры': [
    'Разрешение матрицы',
    'Кратность увеличения, крат',
    'Дальность обнаружения, м',
    'Разрешение дисплея',
    'Бренд',
    'Серия',
    'Тип кратности',
    'Чувствительность температуры, мК',
    'Частота обновления изображения, Гц',
    'Защита от влаги и пыли',
    'Вес, г',
    'Время автономной работы, ч'
  ],
  'Приборы ночного видения (ПНВ)': [
    'Поколение ЭОП',
    'Кратность увеличения, крат',
    'Диаметр объектива, мм',
    'Дальность обнаружения, м',
    'Бренд',
    'Серия',
    'Тип ПНВ',
    'ИК-подсветка',
    'Защита от влаги и пыли',
    'Вес, г',
    'Время автономной работы, ч'
  ],
  'Телескопы': [
    'Диаметр объектива, мм',
    'Фокусное расстояние, мм',
    'Тип монтировки',
    'Оптическая схема',
    'Бренд',
    'Серия',
    'Максимальное увеличение',
    'Светосила',
    'Масса телескопа, кг'
  ],
  'Микроскопы': [
    'Максимальное увеличение',
    'Минимальное увеличение',
    'Тип микроскопа',
    'Бренд',
    'Серия',
    'Тип подсветки',
    'Назначение'
  ],
  'Лупы': [
    'Кратность увеличения, крат',
    'Диаметр линзы, мм',
    'Бренд',
    'Серия',
    'Тип лупы',
    'Подсветка'
  ],
  'Фонари': [
    'Световой поток, лм',
    'Дальность луча, м',
    'Бренд',
    'Серия',
    'Тип элемента питания',
    'Время автономной работы, ч',
    'Защита от влаги и пыли',
    'Вес, г',
    'Режимы работы'
  ]
};

// ========================================
// НАСТРОЙКИ SEO-ТЕГОВ
// ========================================

const SEO_TAGS_CONFIG = {
  // Температуры по этапам
  TEMPERATURES: {
    SINGLE: 0.5,      // Режим "сразу" - баланс
    STAGE_1: 0.3,     // Аналитик - точность
    STAGE_2: 0.4,     // Копирайтер - дисциплина + вариативность LSI
    STAGE_3: 0.1      // Редактор - консистентность
  },

  // Retry для API
  MAX_RETRIES: 3,
  RETRY_DELAYS: [2000, 4000, 8000], // Exponential backoff: 2s, 4s, 8s

  // Задержки между запросами
  DELAY_BETWEEN_API_CALLS: 700,    // мс между запросами к API
  DELAY_BETWEEN_ROWS: 1000,        // мс между строками

  // Ячейки детального листа категории
  DETAIL_SHEET_CELLS: {
    CATEGORY_ID: 'B2',          // ID категории (ключ для поиска)
    CATEGORY_NAME: 'B3',        // Название категории
    SEO_TITLE: 'B13',           // Текущий Title (результат сюда же)
    META_DESC: 'B14',           // Текущий Description (результат сюда же)
    H1: 'B15',                  // H1 заголовок
    KEYWORDS: 'B16',            // Ключевые слова
    DESCRIPTION_START: 'B17',   // Начало описания
    GEMINI_FEEDBACK: 'D11'      // Замечания для Gemini
  },

  // Лист для массовой обработки
  MASS_SHEET_NAME: 'SEO-теги',
  MASS_DATA_START_ROW: 2,       // Первая строка с данными

  // Индексы столбцов листа "SEO-теги" (1-based)
  // Соответствует реальной структуре существующего листа
  MASS_COLUMNS: {
    CHECKBOX: 1,           // A: ☑️
    ID: 2,                 // B: ID категории
    PAGE_NAME: 3,          // C: Название страницы
    URL: 4,                // D: URL
    ADMIN_LINK: 5,         // E: Админка
    PAGE_CONTENT: 6,       // F: Содержание страницы
    SEMANTIC_CORE: 7,      // G: Семантическое ядро
    LSI: 8,                // H: LSI
    QA: 9,                 // I: Вопрос - Ответ
    COMPETITORS_TITLE: 10, // J: Title конкурентов
    COMPETITORS_DESC: 11,  // K: Description конкурентов
    USP: 12,               // L: Наше УТП
    PROMPT_SINGLE: 13,     // M: Общий запрос
    PROMPT_STAGE_1: 14,    // N: Запрос: Этап 1
    PROMPT_STAGE_2: 15,    // O: Запрос: Этап 2
    PROMPT_STAGE_3: 16,    // P: Запрос: Этап 3
    RESULT_TITLE: 17,      // Q: Title
    RESULT_DESC: 18,       // R: Description
    SHORT_DESC: 19,        // S: Краткое описание
    PRODUCT_CRITERIA: 20,  // T: Запрос: подбор товаров
    PRODUCTS_RESULT: 21    // U: Товары
  },

  // Настройки подбора товаров
  EXPORT_SHEET_NAME: 'Выгрузка',       // Лист с CSV-выгрузкой InSales
  PRODUCTS_BATCH_SIZE: 30,              // Товаров в одном запросе к AI (100 = MAX_TOKENS ошибка)

  // Пайплайн автоматической генерации
  PIPELINE_PROGRESS_KEY: 'seo_pipeline_progress',
  LOAD_CHECKED_PROGRESS_KEY: 'seo_load_checked_progress',
  MANDATORY_COLUMNS: [3, 6],            // C (PAGE_NAME), F (PAGE_CONTENT) — обязательные данные
  MINUS_WORDS_SHEET_NAME: 'Минус-слова' // Лист со списком минус-слов (столбец A)
};

// ========================================
// НАСТРОЙКИ ARSENKIN API
// ========================================

const ARSENKIN_CONFIG = {
  API_TOKEN: '668ad513b50296dd81721ee6ced7b517',
  BASE_URL: 'https://arsenkin.ru/api/tools/',

  // Таймауты и задержки
  DELAY_BETWEEN_SUBMITS_MS: 2000,   // Между отправками задач (мс)
  POLL_INTERVAL_MS: 10000,           // Интервал проверки статуса (мс)
  EXECUTION_SAFETY_MS: 330000,       // 5.5 мин (запас 30 сек до таймаута GAS)

  // Retry
  MAX_RETRIES: 3,
  RETRY_DELAYS: [2000, 4000, 8000],

  // Настройки инструмента "Парсинг фраз Wordstat" (подбор ключевиков)
  WORDSTAT_PHRASES: {
    tools_name: 'wordstat',
    type: 2,                // 2 = парсинг фраз
    device: '',             // Все устройства
    region: 213,            // Москва
    minus_words: [],        // Минус-слова (пусто)
    is_clear_minus: true,   // Удалять фразы с минус-словами
    is_right: false,        // Не собирать правую колонку
    is_clear: true          // Удалить '+' из ключевых фраз
  },

  // Настройки инструмента "Сбор частотности Wordstat" (для существующих ключей)
  WORDSTAT_FREQUENCY: {
    tools_name: 'wordstat',
    type: 1,                // 1 = сбор частотности
    device: '',             // Все устройства
    regions: [213],         // Москва
    ws: ['base']            // Базовая частотность (широкое соответствие)
  },

  // Настройки инструмента "Парсинг тегов Title/Description"
  CHECK_H: {
    tools_name: 'check-h',
    mode: 'key',            // По ключевому запросу
    se: 1,                  // 1 = Яндекс
    region: 213,            // Москва
    depth: 10,              // ТОП-10
    pause: 1
  },

  // Ключи для Script Properties (сохранение прогресса)
  PROGRESS_KEYS: {
    WORDSTAT_PHRASES: 'arsenkin_wordstat_phrases_progress',
    WORDSTAT_FREQ: 'arsenkin_wordstat_freq_progress',
    CHECK_H: 'arsenkin_check_h_progress',
    CLUSTERING: 'arsenkin_clustering_progress',
    SEARCH_HIGHLIGHTS: 'arsenkin_highlights_progress'
  },

  // Настройки инструмента "Кластеризация"
  CLUSTERING: {
    tools_name: 'clustering',
    se: 1,                  // 1 = Яндекс, 2 = Google
    region: 213,            // Москва
    threshold: 3,           // Порог кластеризации (стандартный - 3)
    group: 'hard',          // Метод (hard/soft)
    count: 3                // Степень группировки (как на скриншоте пользователя)
  },

  // Настройки инструмента "Парсинг подсветок" (LSI-слова из выдачи)
  SEARCH_HIGHLIGHTS: {
    tools_name: 'sp',
    se: 1,                  // 1 = Яндекс
    region: 213,            // Москва
    depth: 50               // Глубина проверки
  }
};

// ========================================
// ПРОМПТЫ ДЛЯ SEO-ТЕГОВ
// ========================================

const SEO_TAGS_PROMPTS = {
  // Системная инструкция для Gemini (общая для всех режимов)
  SYSTEM: `Ты - профессиональный SEO-специалист интернет-магазина оптических приборов Binokl.shop.
Твоя задача - создавать Title и Description для страниц категорий.
Формат ответа: Title: [текст] Description: [текст]`,

  DESCRIPTION_SYSTEM: `Ты - профессиональный SEO-копирайтер интернет-магазина оптических приборов Binokl.shop.
Твоя задача - создавать качественные, привлекательные и оптимизированные описания для страниц категорий.
Отвечай ТОЛЬКО текстом самого описания без каких-либо вводных слов, заголовков вроде "Вот ваше описание:" или "Description:".`,

  // Системные инструкции для этапов поэтапной генерации
  STAGE_1_SYSTEM: `Ты — ведущий SEO-аналитик и эксперт по поисковому продвижению в Google и Яндекс.`,
  STAGE_2_SYSTEM: `Ты — экспертный SEO-копирайтер и контент-маркетолог, мастер LSI-оптимизации.`,
  STAGE_3_SYSTEM: `Ты — специалист по финальному оформлению поисковой выдачи.`

  // Промпты для этапов задаются ТОЛЬКО в ячейках листа "SEO-теги" (столбцы M, N, O, P).
  // Если ячейка пуста — генерация прерывается с ошибкой.
};
