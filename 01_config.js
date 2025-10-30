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

const CATEGORY_DESCRIPTION_ASSISTANT_ID = 'asst_qVFYH8Q5qgzMKsvOOzwnAuHN';

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
  defaultSearchEngine: 'yandex'  // 'yandex' или 'google'
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
  LSI_WORDS: 'LSI и тематика',
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
  ADMIN_LINK: 13            // M
};

// ========================================
// СЕКЦИИ ДЕТАЛЬНОГО ЛИСТА
// ========================================

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
  
  CURRENT_PRODUCTS_START: 27,
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
  AI_GENERATION_START: 200,
  TAG_TILES_START: 300
};

// ========================================
// ENDPOINTS INSALES API
// ========================================

const CATEGORY_API_ENDPOINTS = {
  COLLECTIONS: '/admin/collections.json',
  COLLECTION_BY_ID: '/admin/collections/{id}.json',
  COLLECTION_PRODUCTS: '/admin/products.json?collection_id={id}',
  UPDATE_COLLECTION: '/admin/collections/{id}.json',
  PRODUCTS: '/admin/products.json'
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