# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script application for managing InSales e-commerce categories with AI-powered SEO optimization. Operates through a Google Sheets interface with bidirectional sync to InSales API.

**Core Technologies:**
- Google Apps Script (server-side JavaScript, V8 runtime)
- InSales API (Basic Auth)
- OpenAI GPT-4 & Claude API (content generation)
- Google Sheets (UI & data persistence)

**Production Site:** https://binokl.shop (binoculars e-commerce)

## Development Workflow

### Local Development
```bash
# Edit files locally
code ~/Documents/AppsScriptProject

# Deploy to Apps Script
clasp push

# Sync to GitHub
git add . && git commit -m "Description" && git push
```

### From Apps Script Editor
```bash
# After editing at https://script.google.com
clasp pull
git add . && git commit -m "Quick fix via Apps Script" && git push
```

### Testing
- No automated test suite
- Test functions embedded in modules (e.g., `testDensityAnalysis()`, `testGenerateAnchors()`)
- Manual testing via Apps Script editor or spreadsheet menu
- Always test on a single category before batch operations

## Architecture

### File Naming Convention

Files use **numeric prefixes** that define load order and functional grouping:

```
00_main.js              - Entry point, menu initialization
01_config.js            - All configuration constants (497 lines)
05_logging.js           - Logging utilities
09_*.js                 - Helper modules (InSales, AI, descriptions)
11_*.js                 - Data structure setup
12_*.js                 - Category loading from InSales
13_*.js                 - Search & UI (1,390 lines - LARGEST MODULE)
14_*.js                 - Keywords & semantic analysis
15_*.js                 - AI category analysis
16_*.js                 - AI content generation (SEO)
17_*.js                 - Product management
18_*.js                 - Position tracking (SERP)
19_*.js                 - InSales API update functions
20_*.js                 - Menu structure (762 lines)
21_*.js                 - Send category changes to InSales (773 lines)
22_*.js                 - Category creation
23_*.js                 - Tag tiles generation (analyzer, generator, HTML)
24_*.js                 - Manual tag tiles workflow (ACTIVE - 990 lines)
```

**Critical dependencies:** Config (01) must load before all other modules. Core modules (05-12) must load before feature modules (13-24).

### Configuration Architecture

**Everything lives in [01_config.js](01_config.js:1)**:

- `INSALES_CONFIG` - API credentials, shop domain
- `OPENAI_CONFIG` - API key, model, parameters
- `CATEGORY_SHEETS` - Sheet name constants
- `MAIN_LIST_COLUMNS` - Column indices for main category list
- `DETAIL_SHEET_SECTIONS` - Cell references for category detail sheets (e.g., B2 = Category ID)
- `TAG_TILES_CONFIG` - Density thresholds, anchor limits, AI settings
- `CATEGORY_API_ENDPOINTS` - InSales API endpoint templates

**Secrets Management:**
- Development: Hardcoded in `01_config.js`
- Production: Script Properties (`PropertiesService.getScriptProperties()`)
- Fallback pattern: Check Script Properties first, then fall back to config constants

**KNOWN ISSUE:** `getInsalesCredentialsSync()` is duplicated across 3 files:
- [09_insales_helpers.js](09_insales_helpers.js:1)
- [19_categories_update_insales.js](19_categories_update_insales.js:1)
- [21_category_update.js](21_category_update.js:1)

This creates maintenance risk. Consider consolidating into single source.

### Data Structures

**Google Sheets as Database:**

#### 1. "Категории — Список" - Main category list with hierarchy

- Tree visualization using indentation (└─ prefix)
- Columns defined in `MAIN_LIST_COLUMNS` (A-M)
- Row 1: Headers (frozen), Row 3+: Data

#### 2. "Категория — [Name]" - Detail sheets (one per category)

**CRITICAL: This is the actual production structure from [data/Категории - Категория — Театральные(5).csv](data/Категории - Категория — Театральные(5).csv:178)**

```
Rows 1-11:   📁 ИНФОРМАЦИЯ О КАТЕГОРИИ
  Row 1:  Header: "📁 ИНФОРМАЦИЯ О КАТЕГОРИИ"
  Row 2:  ID категории: [value in B2] - ARCHITECTURAL CONSTANT - NEVER MOVE
  Row 3:  Название: [category name]
  Row 4:  URL: [/collection/...]
  Row 5:  Путь в иерархии: [parent category path]
  Row 6:  (empty)
  Row 7:  Маркерный запрос: [optional marker query]
  Row 8:  Админка InSales: [admin URL link]
  Row 9-11: (empty)

Rows 12-21:  🎯 SEO ДАННЫЕ
  Row 12: Header: "🎯 SEO ДАННЫЕ"
  Row 13: SEO Title: [title value in B13]
  Row 14: Meta Description: [description in B14]
  Row 15: H1 заголовок: [H1 in B15]
  Row 16: Ключевые слова: [keywords in B16]
  Row 17: Описание категории: [HTML description spanning multiple columns B17-H17]
          Note: Can span rows 17-21 for long HTML content

Rows 22+:    📝 КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПЛИТОК ТЕГОВ (ручной ввод)
  Row 22: Header: "📝 КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПЛИТОК ТЕГОВ (ручной ввод)"
  Row 23: Column headers:
          A: ☑️ (checkbox)
          B: Ключевое слово
          C: Тип плитки (dropdown: Верхняя/Нижняя)
          D: Текст анкора
          E: URL/ID категории
          F: Статус категории (auto-filled after validation)
          G: ID родителя (новая) (for creating new categories)
  Row 24+: Manual keyword input rows (user fills these)

  ARCHITECTURAL CONSTANT: Row 22 is FIXED - tag keywords table ALWAYS starts here

Rows 32+:    📊 СТАТИСТИКА ТОВАРОВ
  Row 32: Header: "📊 СТАТИСТИКА ТОВАРОВ"
  Row 33: Всего товаров: [count in B33] - CRITICAL for dynamic positioning
  Row 34: В наличии: [count in B34]
  Row 35: Нет в наличии: [count in B35]
  Row 36: Процент наличия: [percentage in B36]
  Row 37-38: (empty)

Rows 39+:    🛒 ТЕКУЩИЕ ТОВАРЫ В КАТЕГОРИИ
  Row 39: Header: "🛒 ТЕКУЩИЕ ТОВАРЫ В КАТЕГОРИИ"
  Row 40: Column headers:
          A: Название
          B: Артикул
          C: Цена
          D: В наличии
          E: ID
          F: ☑️ (checkbox for selection)
  Row 41+: Product rows (count from B33)
          Example: Row 41-103 for 63 products

Rows 107+:   ⚙️ ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ КАТЕГОРИИ (field_values from InSales)
  Row 107: Header: "⚙️ ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ КАТЕГОРИИ"
  Row 108: (empty)
  Row 109: H1: [value in B109]
  Row 110: Доставка и оплата: [value in B110]
  Row 111: Блок ссылок сверху: [HTML - UPPER tag tile from InSales]
  Row 112-113: (continuation of upper tile HTML)
  Row 114: Блок ссылок: [HTML - LOWER tag tile from InSales]
  Row 115: (continuation of lower tile HTML)
  Row 116: Иконка для меню: [value]
  Row 117: Параметры товаров (через точку с запятой): [value]
  Row 118: Товары и категории (популярные): [value]
  Row 119: Видеообзор: [value]
  Row 120: Раздел в меню: [value]
  Row 121: ID товаров через запятую: [value]
  Row 122: Название для региональных страниц: [value]
  Row 123: Показывать Официального дилера?: [value]
  Row 124-126: Баннер 1 (изображение, мобильное, ссылка)
  Row 127-129: Баннер 2 (изображение, мобильное, ссылка)
  Row 130-132: Вертикальный баннер (изображение, мобильное, ссылка)
  Row 133: Скрыть категорию в слайдере?: [value]
  Row 134: Скрыть категорию в меню?: [value]
  Row 135: Скрыть категорию в списке?: [value]
  Row 136: noindex: [value]
  Row 137-139: (empty)

  IMPORTANT: Position calculated dynamically:
  extraFieldsStartRow = productsStartRow + products + 5
  (where products = parseInt(B33))

Rows 140+:   🏷️ ПЛИТКА ТЕГОВ - ВЕРХНЯЯ (над описанием категории)
  Row 140: Header: "🏷️ ПЛИТКА ТЕГОВ - ВЕРХНЯЯ (над описанием категории)"
  Row 141: Инструкция: "Слева - существующие теги из InSales. Справа - новые..."
  Row 142: (empty)
  Row 143: Table headers:
           A: БЫЛО: Текст ссылки
           B: БЫЛО: URL
           C: БЫЛО: ☑️ Включить
           D: (empty separator)
           E: СТАЛО: Текст анкора
           F: СТАЛО: URL
           G: СТАЛО: ID категории
           H: СТАЛО: Примечание
  Row 144+: Generated upper tile data (comparison between existing and new)
            Example rows:
            - Row 144: Existing link "Лорнеты (с ручкой)" vs new generated tag
            - Row 145-146: Empty rows or additional tags
  Row XX:   "HTML код (финальный):" [final HTML in merged cells B-H]
            CRITICAL: Located by TEXT SEARCH, not calculation
            Search for cell containing "HTML код" AND "финальный"
            within bounds [upperTileRow, lowerTileRow)

Rows 155+:   🏷️ ПЛИТКА ТЕГОВ - НИЖНЯЯ (под описанием категории)
  Row 155: Header: "🏷️ ПЛИТКА ТЕГОВ - НИЖНЯЯ (под описанием категории)"
  Row 156: Инструкция: "Слева - существующие теги из InSales. Справа - новые..."
  Row 157: (empty)
  Row 158: Table headers (same as upper tile):
           A: БЫЛО: Текст ссылки
           B: БЫЛО: URL
           C: БЫЛО: ☑️ Включить
           D: (empty separator)
           E: СТАЛО: Текст анкора
           F: СТАЛО: URL
           G: СТАЛО: ID категории
           H: СТАЛО: Примечание
  Row 159+: Generated lower tile data
            Example rows:
            - Row 159: "Бинокли-лорнеты (с ручкой)" + new generated
            - Row 160: "Маленькие"
            - Row 161: "Недорогие"
            - Row 162: "Складные"
            - Row 163-167: More existing tags
  Row YY:   "HTML код (финальный):" [final HTML in merged cells B-H]
            CRITICAL: Located by TEXT SEARCH
            Search from lowerTileRow onwards
```

**Fixed Cell References (ARCHITECTURAL CONSTANTS - NEVER CHANGE):**
- **B2:** Category ID
- **Row 22:** Start of tag keywords table
- **B33:** Total products count (used for dynamic positioning calculations)

**Dynamic Positioning (calculated by `calculateSheetSections()`):**
- Products section start (Row 39+)
- field_values section (Row 107+ = productsStartRow + products + 5)
- Upper tag tile section (Row 140+)
- Lower tag tile section (Row 155+)
- "HTML код (финальный)" fields - Located by TEXT SEARCH within section boundaries

**CRITICAL INSIGHTS from production data:**

1. **БЫЛО vs СТАЛО pattern**: System compares existing InSales tags (БЫЛО = "was") with newly generated tags (СТАЛО = "became")
   - Existing tags from InSales field_values: "Блок ссылок сверху" (row 111) and "Блок ссылок" (row 114)
   - Generated tags in tile sections: rows 144+ and 159+
   - User can manually select which tags to keep via checkboxes

2. **HTML code fields are merged cells**: B through H columns, with descriptive label in column A

3. **Type coercion critical**: B33 (products count) must be parsed with `parseInt()` to prevent string concatenation bugs

#### 3. Other Sheets

- **"Ключевые слова"** - Keywords with frequency, competition, type
- **"LSI и тематика"** - LSI words and semantic relationships
- **"Плитки тегов"** - Generated tag tiles (legacy/experimental)
- **"Анализ плотности"** - Keyword density analysis results
- **"Логи"** - Auto-generated logs from logging system

### InSales API Integration

**Authentication Pattern:**
```javascript
const credentials = getInsalesCredentialsSync();
const headers = {
  'Authorization': 'Basic ' + Utilities.base64Encode(`${credentials.apiKey}:${credentials.password}`),
  'Content-Type': 'application/json'
};
```

**Key Endpoints:**
- `GET /admin/collections.json` - List all categories
- `GET /admin/collections/{id}.json` - Get category details
- `PUT /admin/collections/{id}.json` - Update category
- `POST /admin/collections.json` - Create category (ALWAYS requires parent_id)

**Critical API Behaviors:**

1. **parent_id is REQUIRED** when creating categories. For root-level, use `parent_id: 9069711` (existing "Каталог" category).
   - InSales doesn't accept `parent_id: null` or `parent_id: 0`
   - Fixed in commit bc4b8fd after HTTP 422 errors

2. **field_values structure:**
   ```javascript
   {
     "collection": {
       "field_values": [
         {
           "field_id": 123456,
           "values": ["Текстовое значение"],
           "field": { "handle": "h1" }  // Russian handle
         }
       ]
     }
   }
   ```

   **Common InSales field_values handles (Russian, hardcoded in codebase):**
   - `h1` - H1 заголовок (row 109)
   - `dostavka-i-oplata` - Доставка и оплата (row 110)
   - `blok-ssylok-sverhu` - Блок ссылок сверху (upper tag tile, row 111)
   - `blok-ssylok` - Блок ссылок (lower tag tile, row 114)
   - `ikonka-dlya-menyu` - Иконка для меню (row 116)
   - `parametry-tovarov-cherez-tochku-s-zapyatoy` - Параметры товаров (row 117)
   - `tovary-i-kategorii-populyarnye` - Товары и категории (row 118)
   - `videoobzor` - Видеообзор (row 119)
   - `razdel-v-menyu` - Раздел в меню (row 120)
   - `id-tovarov-cherez-zapyatuyu` - ID товаров через запятую (row 121)
   - `nazvaniye-dlya-regionalnykh-stranits` - Название для региональных страниц (row 122)
   - `pokazyvat-ofitsialnogo-dilera` - Показывать Официального дилера? (row 123)
   - Banner and display settings (rows 124-136)

   CRITICAL: Must read from row 107+ to get all field_values
   Position calculated: `productsStartRow + parseInt(B33) + 5`

3. **Type coercion bug (commit 706d279)**:
   - Products count from B33 must be parsed with `parseInt()` to avoid string concatenation
   - Bug example: `"37" + "Nижняя" + "520"` → `"37Нижняя520"` (breaks InSales API)
   - Fix: `const products = parseInt(productsValue) || 0;`

4. **Rate limiting:**
   - 500ms delay between requests (`APP_CONFIG.apiDelay`)
   - Retry logic: 3 attempts (`APP_CONFIG.apiRetries`)

### AI Integration Patterns

**Three AI Providers:**

1. **OpenAI Chat Completions API** ([16_categories_ai_content.js](16_categories_ai_content.js:1))
   - For quick SEO generation (Title, Description, H1)
   - Model: `gpt-4o-mini`
   - Temperature: 0.7

2. **OpenAI Assistants API** ([09_ai_category_descriptions.js](09_ai_category_descriptions.js:1))
   - For complex description workflows with context
   - Assistant ID stored in config
   - Supports rewrite vs. new generation modes

3. **Claude API** ([23_tag_tiles_generator.js](23_tag_tiles_generator.js:1))
   - For tag tile anchor generation (switched from Gemini in commit fa93f2f)
   - Expects JSON responses with structured anchor data
   - API key stored in Script Properties as `CLAUDE_API_KEY`
   - Better results than Gemini for natural language anchor generation

### Dynamic Sheet Positioning Pattern

**Problem:** In early versions, tag tiles were placed at fixed rows (916, 933), causing conflicts when products list grew.

**Solution:** Dynamic positioning system in [24_tag_tiles_manual.js](24_tag_tiles_manual.js:990):

```javascript
// calculateSheetSections() automatically calculates:
1. Find last row with tag keywords (row 22+)
2. Place upper tile after keywords + 3 rows offset
3. Place lower tile after upper tile + its size
4. Products section shifts automatically after all tiles
```

**RECENT INSTABILITY (Oct 2025):**
- calculateSheetSections() has been modified 4 times in 24 hours
- Commits: 1a8adfb, 8ca5465 (dynamic positioning fixes)
- Issues fixed:
  - HTML fields written to wrong rows (nasloeniye - overlapping)
  - Upper tile capturing lower tile tags
- **Root cause:** Calculating row positions instead of searching for text markers

**Current Approach (STABLE since commit 8ca5465):**
```javascript
// Search for "HTML код (финальный)" text within block boundaries
let upperHTMLRow = null;
for (let row = upperTileRow; row < lowerTileRow; row++) {
  const cellValue = sheet.getRange(row, 1).getValue();
  if (cellValue && cellValue.toString().includes('HTML код') &&
      cellValue.toString().includes('финальный')) {
    upperHTMLRow = row;
    break;
  }
}
```

**When adding new sections:**
- Use `calculateSheetSections()` to get current positions
- Never hardcode row numbers for dynamic sections
- Only B2 (category ID), row 22 (keywords start), and B33 (products count) are fixed
- Search for section markers by text, not by calculated position
- Use bounds checking: upper tile searches [upperTileRow, lowerTileRow), lower tile searches from lowerTileRow onwards

### Error Handling Pattern

**Consistent throughout codebase:**
```javascript
function operationName() {
  const context = 'Operation name';
  try {
    logInfo('📌 Starting', null, context);
    // ... operation ...
    logInfo('✅ Success', null, context);
    return result;
  } catch (error) {
    logError('❌ Error', error, context);
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    throw error;
  }
}
```

All logs go to "Логи" sheet if logging is enabled.

## Key Workflows

### Category Management Flow
```
onOpen() → Menu Creation
↓
User: "Load Categories"
↓
loadCategoriesWithHierarchy() [12_categories_loader.js:1]
  → loadAllCollectionsFromInSales() (API call)
  → buildCategoryHierarchy() (tree structure)
  → addProductCountsToCategories() (API calls)
  → writeCategoriesToMainList() (sheet write)
↓
User: Search/Select Category
↓
showCategorySearchDialog() [13_categories_search.js:1]
  → openDetailedCategorySheet() (creates detail sheet)
↓
User: Modify category in detail sheet
↓
sendCategoryChangesToInSales() [21_category_update.js:1]
  → PUT /admin/collections/{id}.json
  → Update status in main list
```

### Tag Tiles Generation (ACTIVE FEATURE)

**Two Approaches:**

#### A. Manual Workflow (CURRENT - ACTIVE) - [24_tag_tiles_manual.js](24_tag_tiles_manual.js:990)

**Status:** Fully implemented and in active use

**Workflow:**
1. **Initialize table** - `initializeTagKeywordsTable()` creates input table at row 22+
2. **Manual input** - User enters keywords, tile types (Upper/Lower), anchor text, category IDs
3. **Validate categories** - `validateTagKeywords()` checks if categories exist via InSales API
4. **Create new categories** - `createCategoriesForTags()` creates missing categories if needed
5. **Generate HTML** - `generateTilesFromManualData()` creates styled HTML tiles
6. **Preview** - `showTilesPreviewManual()` shows preview before applying
7. **Write to sheet** - `saveGeneratedTilesToSheet()` writes СТАЛО columns and HTML
8. **Sync to InSales** - User clicks "Отправить изменения в InSales" → updates field_values (blok-ssylok-sverhu, blok-ssylok)

**Key Features:**
- Full control over keywords and anchor text
- Choose existing categories or create new ones
- Dynamic section positioning (no conflicts with products)
- Preview before applying
- Integrated with InSales API for category creation
- Compares generated tiles (СТАЛО) vs existing InSales tiles (БЫЛО)
- Manual checkbox selection for which tags to include

**Recent Fixes (commits 1a8adfb, 8ca5465, 706d279):**
- Fixed HTML field positioning using text search (commit 8ca5465)
- Fixed upper tile capturing lower tile data (commit 8ca5465)
- Fixed products count type coercion bug (commit 706d279)
- Fixed dynamic positioning to prevent nasloeniye (commit 1a8adfb)

#### B. Automated Workflow (EXPERIMENTAL) - [23_tag_tiles_*.js](23_tag_tiles_analyzer.js:1)

**Status:** Analyzer complete, generator complete, integration pending

**Workflow:**
1. **Density Analysis** - Calculates TF (Term Frequency), identifies keywords needing SEO boost
2. **AI Anchor Generation** - Uses Claude API to generate natural anchor text
3. **HTML Generation** - Converts anchors to styled HTML tiles

**Note:** Manual workflow is prioritized due to better control over SEO strategy.

See [COMMANDS.md](COMMANDS.md:342) for detailed manual workflow instructions.

## Common Development Tasks

### Adding a New Feature Module

1. **Choose numeric prefix** based on dependencies (e.g., 25_*.js for next feature)
2. **Add configuration** to [01_config.js](01_config.js:497) first
3. **Use consistent naming:**
   - `loadXxx()` - Fetch from API
   - `getXxx()` - Retrieve from sheets
   - `updateXxx()` / `setXxx()` - Write data
   - `showXxx()` - UI dialogs
   - `generateXxx()` - AI/content creation
4. **Add to menu** in [20_categories_menu.js](20_categories_menu.js:762)
5. **Include logging** with context strings
6. **Test function** named `testXxx()`

### Working with InSales API

**Always:**
- Use `getInsalesCredentialsSync()` for auth headers
- Set `muteHttpExceptions: true` to handle errors gracefully
- Check `response.getResponseCode()` before parsing JSON
- Add delay between requests: `Utilities.sleep(APP_CONFIG.apiDelay)`
- Log requests with category ID for debugging

**parent_id handling:**
```javascript
// CORRECT - Always include parent_id
if (data.parent_id && data.parent_id !== null && data.parent_id !== '') {
  collectionData.parent_id = parseInt(data.parent_id);
} else {
  collectionData.parent_id = 9069711; // Default to "Каталог"
}
```

**field_values handling:**
```javascript
// CRITICAL: Products count must be parsed as integer
const productsValue = sheet.getRange('B33').getValue();
const products = parseInt(productsValue) || 0; // Prevents string concatenation
const productsStartRow = calculateSheetSections(sheet).productsStart;
const extraFieldsStartRow = productsStartRow + products + 5; // Row after products
```

### Modifying Sheet Structures

**Before changing:**
1. Check all references to column constants in `01_config.js`
2. Search codebase for hardcoded row numbers (e.g., "row 29", "B2", "B33")
3. Update `MAIN_LIST_COLUMNS` or `DETAIL_SHEET_SECTIONS` constants
4. Test on a single category before batch operations

**DO NOT:**
- Hardcode column indices - always use constants
- Hardcode sheet names - use `CATEGORY_SHEETS` constants
- Change cell B2 in detail sheets (category ID reference)
- Change row 22 (tag keywords table start)
- Change cell B33 (products count - used for dynamic positioning)
- Manually position products section (use `calculateSheetSections()`)
- Calculate HTML field positions (search by text "HTML код (финальный)" instead)

### Working with AI APIs

**OpenAI:**
- API key in Script Properties or `OPENAI_CONFIG`
- Use `response_format: { type: 'json_object' }` for structured responses
- Handle rate limits with retries

**Claude:**
- API key in Script Properties as `CLAUDE_API_KEY`
- Switched from Gemini to Claude for better anchor generation results (commit fa93f2f)
- Use detailed prompts with JSON schema in prompt text
- Expects structured JSON responses

**Note:** Old Gemini integration still exists in codebase but Claude is preferred for new development.

## Important Constraints

### Google Apps Script Limitations

- **No async/await** - All operations are synchronous
- **6-minute execution timeout** - Paginate large operations
- **Quota limits** - UrlFetchApp: 20,000 calls/day
- **No native JSON parsing in templates** - Pass pre-processed data to HTML

**Workarounds:**
- Use `Utilities.sleep(ms)` for delays
- Script Properties for state persistence between runs
- Toast notifications for progress feedback
- Break large operations into smaller chunks

### Critical Business Rules

1. **Category parent_id:** NEVER create categories with null parent_id. Always use 9069711 as fallback.

2. **Sheet cell references:**
   - Category ID is ALWAYS at B2
   - Tag keywords table ALWAYS starts at row 22
   - Products count ALWAYS at B33 (used for dynamic positioning)
   - Products section uses DYNAMIC positioning (calculated by `calculateSheetSections()`)
   - field_values section uses DYNAMIC positioning (productsStartRow + products count + 5)
   - HTML fields located by TEXT SEARCH, not calculation

3. **Hierarchical display:** Main list shows tree structure with indentation. Must preserve visual hierarchy when updating.

4. **Keyword density thresholds:**
   - Main keywords: 2-4%
   - Additional: 1-2%
   - LSI: 0.5-1%
   - Spam threshold: >5%

5. **Tag tiles limits:**
   - Upper tile (navigation): 5-8 anchors
   - Lower tile (SEO): 15-30 anchors

6. **Type coercion:** Always use `parseInt()` when reading numeric values from sheets that will be used in arithmetic operations.

## Debugging

### Logging
- Use `logInfo()`, `logWarning()`, `logError()` from [05_logging.js](05_logging.js:1)
- Logs saved to "Логи" sheet (auto-created)
- Include context string for filtering: `logInfo('Message', data, 'ModuleName')`

### Common Issues

**"parent_id: Не заполнены обязательные поля"**
- InSales requires parent_id for all category creation
- Solution: Always set `parent_id: 9069711` for root categories
- Fixed in commit bc4b8fd

**"Невозможно преобразовать '37Нижняя520' в int"**
- String concatenation instead of arithmetic
- Solution: Use `parseInt()` when reading numeric values
- Fixed in commit 706d279 for products count

**HTML fields written to wrong rows (nasloeniye - overlapping)**
- Calculated position doesn't match actual structure
- Solution: Search for "HTML код (финальный)" text within block boundaries
- Fixed in commit 8ca5465

**Upper tile includes tags from lower tile**
- Reading too many rows without bounds checking
- Solution: Scan rows dynamically until "HTML код" marker, within block boundaries
- Fixed in commit 8ca5465

**"Категории не найдены" in parent dropdown**
- Sheet name mismatch: Use `CATEGORY_SHEETS.MAIN_LIST` constant
- Column mismatch: Use `MAIN_LIST_COLUMNS` constants

**Script timeout**
- Operation taking >6 minutes
- Solution: Reduce batch size, add pagination, show progress toasts

**API rate limit**
- Too many requests to InSales/OpenAI/Claude
- Solution: Increase `APP_CONFIG.apiDelay`, implement exponential backoff

## Architecture Issues & Technical Debt

### Known Issues

1. **Duplicate getInsalesCredentialsSync() function**
   - Exists in 3 files: 09_insales_helpers.js, 19_categories_update_insales.js, 21_category_update.js
   - Creates maintenance risk
   - Recommendation: Consolidate into single source in 09_insales_helpers.js

2. **Inconsistent secrets management**
   - Some modules check Script Properties first, others use config directly
   - Recommendation: Standardize on fallback pattern everywhere

3. **Module size imbalance**
   - 13_categories_search.js is 1,390 lines (too large)
   - Contains both UI and business logic
   - Recommendation: Split into separate UI and logic modules

4. **Recent instability in dynamic positioning (Oct 2025)**
   - calculateSheetSections() modified 4 times in 24 hours
   - Root cause: Calculating row positions instead of searching for text markers
   - Now stabilized with text-based field location (commit 8ca5465)

5. **Hardcoded Russian field names**
   - InSales field_values handles are hardcoded strings throughout codebase
   - If InSales changes handles, need to update in multiple files
   - Recommendation: Create constants in 01_config.js for all field handles

6. **No automated testing**
   - All testing is manual
   - Recent bugs (commits 1a8adfb, 8ca5465, 706d279) could have been caught by tests
   - Recommendation: Add automated tests for critical functions

7. **CSV Import with Multiline Fields (FIXED Nov 2025)**
   - InSales exports CSV with multiline fields enclosed in quotes
   - Primitive `line.split()` parsing broke on these fields, creating "garbage" rows
   - **Solution:** Implemented RFC 4180-compliant CSV parser in [25_products_cache.js](25_products_cache.js:63-133)
   - Now handles: multiline fields, escaped quotes (`""`), different delimiters
   - See [FIX_SUMMARY.md](FIX_SUMMARY.md:1) for detailed explanation

8. **Google Sheets Auto-conversion of Decimal Values to Dates (FIXED Nov 2025)**
   - Google Sheets automatically converts decimal values like "3.5", "8.5", "12.5" to dates (month.day format)
   - This corrupted filter values for characteristics like "Кратность увеличения, крат" (Magnification multiplier)
   - **Solution:** Force text format (`@STRING@`) for all parameter columns after CSV import in [25_products_cache.js](25_products_cache.js:212-228)
   - Applied immediately after `setValues()` to override Google Sheets auto-detection
   - Wrapped in try-catch to make formatting non-critical if it fails

## Reference

**Documentation:**
- [COMMANDS.md](COMMANDS.md:342) - User guide for tag tiles workflow
- This file (CLAUDE.md) - Technical reference for developers

**Key Files:**
- [01_config.js](01_config.js:497) - ALL configuration constants (start here)
- [13_categories_search.js](13_categories_search.js:1390) - Search UI & detail sheet creation (largest module)
- [17_categories_products.gs.js](17_categories_products.gs.js:1097) - Product selection dialog with CSV filtering
- [21_category_update.js](21_category_update.js:773) - Main update logic for InSales sync
- [22_category_create.js](22_category_create.js:1) - Category creation with parent_id handling
- [24_tag_tiles_manual.js](24_tag_tiles_manual.js:990) - Manual tag tiles workflow (active development)
- [23_tag_tiles_html.js](23_tag_tiles_html.js:1) - HTML/CSS generation for tag tiles
- [25_products_cache.js](25_products_cache.js:256) - CSV import with RFC 4180 parser and text formatting

**InSales Documentation:**
- Base URL: https://binokl.shop
- Admin: https://binokl.shop/admin2 (NOT myshop-on665.myinsales.ru - that domain doesn't work for direct product links)
- API uses Basic Auth with base64-encoded credentials
- Product edit URLs: `https://binokl.shop/admin2/products/{id}` (no `/edit` suffix needed)

## Recent Development History

**Latest Features (commits fe9b1ca - present):**
- ✅ Manual tag tiles workflow with full control over keywords and categories
- ✅ Dynamic section positioning (no conflicts with products)
- ✅ Category creation via InSales API from tag tiles interface
- ✅ Preview functionality for tag tiles before applying
- ✅ HTML/CSS generation for styled tag tiles
- ✅ БЫЛО vs СТАЛО comparison (existing vs generated tags)

**Critical Fixes (Oct-Nov 2025):**
- 706d279: Fixed string concatenation bug in products count (InSales API error)
- 8ca5465: Fixed HTML field positioning using text search instead of calculation
- 1a8adfb: Fixed dynamic positioning to prevent data bleeding between sections
- bc4b8fd: Fixed HTTP 422 error when creating root categories (parent_id now always set)
- fa93f2f: Switched from Gemini to Claude API for better anchor generation
- Nov 2025: Fixed Google Sheets auto-converting decimal values ("3.5") to dates by forcing `@STRING@` format
- Nov 2025: Fixed admin product links - changed domain from `myshop-on665.myinsales.ru` to `binokl.shop`

**In Progress:**
- Manual tag tiles workflow is ACTIVE and prioritized
- Automated density analysis exists but not integrated with manual workflow
- Future: Merge manual control with AI suggestions

**Known Issues:**
- None critical - manual workflow is stable after recent fixes
- Automated workflow integration pending based on user feedback

## Production Data Example

See [data/Категории - Категория — Театральные(5).csv](data/Категории - Категория — Театральные(5).csv:178) for actual production detail sheet structure.

Key insights from production data:
- Real category: "Театральные бинокли" (Theater binoculars) - ID 9071624
- 63 products total, 25 in stock, 38 out of stock (40% availability)
- Upper tile has 2 tags, lower tile has 9 tags (8 existing + 1 generated)
- Shows БЫЛО (existing from InSales) vs СТАЛО (newly generated) comparison
- Demonstrates checkbox selection for manual control over tag inclusion
- field_values section starts at row 107 (after 63 products + 5 rows offset)
