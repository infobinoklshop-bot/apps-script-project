# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script application for managing InSales e-commerce categories with AI-powered SEO optimization. Operates through a Google Sheets interface with bidirectional sync to InSales API.

**Core Technologies:**
- Google Apps Script (server-side JavaScript)
- InSales API (Basic Auth)
- OpenAI GPT-4 & Google Gemini (content generation)
- Google Sheets (UI & data persistence)

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
13_*.js                 - Search & UI (largest: 1,169 lines)
14_*.js                 - Keywords & semantic analysis
15_*.js                 - AI category analysis
16_*.js                 - AI content generation (SEO)
17_*.js                 - Product management
18_*.js                 - Position tracking (SERP)
19_*.js                 - InSales API update functions
20_*.js                 - Menu structure
21_*.js                 - Send category changes to InSales (762 lines)
22_*.js                 - Category creation
23_*.js                 - Export & tag tiles generation
```

**Critical dependencies:** Config (01) must load before all other modules. Core modules (05-12) must load before feature modules (13-23).

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

### Data Structures

**Google Sheets as Database:**

1. **"Категории — Список"** - Main category list with hierarchy
   - Tree visualization using indentation (└─ prefix)
   - Columns defined in `MAIN_LIST_COLUMNS` (A-M)
   - Row 1: Headers (frozen), Row 3+: Data

2. **"Категория — [Name]"** - Detail sheets (one per category)
   - B2: Category ID (ALWAYS at this cell)
   - Rows 12-25: SEO section
   - Row 29+: Products section
   - Row 300+: Tag tiles section (TAG_TILES_START)

3. **"Ключевые слова"** - Keywords with frequency, competition, type
4. **"LSI и тематика"** - LSI words and semantic relationships
5. **"Плитки тегов"** - Generated tag tiles (new feature in development)
6. **"Анализ плотности"** - Keyword density analysis results

**IMPORTANT:** Cell references in detail sheets are hardcoded (e.g., B2 for category ID, row 29 for products). Do not modify these locations without updating all references.

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
- **parent_id is REQUIRED** when creating categories. For root-level, use `parent_id: 9069711` (existing "Каталог" category).
- InSales doesn't accept `parent_id: null` or `parent_id: 0`
- Rate limiting: 500ms delay between requests (`APP_CONFIG.apiDelay`)
- Retry logic: 3 attempts (`APP_CONFIG.apiRetries`)

### AI Integration Patterns

**Two OpenAI Approaches:**

1. **Chat Completions API** ([16_categories_ai_content.js](16_categories_ai_content.js:1)) - For quick SEO generation
2. **Assistants API** ([09_ai_category_descriptions.js](09_ai_category_descriptions.js:1)) - For complex description workflows

**Gemini API** ([23_tag_tiles_generator.js](23_tag_tiles_generator.js:1)) - For tag tile anchor generation:
- Model: `gemini-1.5-pro`
- Temperature: 0.7
- Expects JSON responses with structured anchor data
- API key stored in Script Properties as `GEMINI_API_KEY`

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

### Tag Tiles Generation (New Feature - IN DEVELOPMENT)

**Status:** Density analyzer complete, anchor generator complete, HTML generator pending

**Workflow:**
1. **Density Analysis** ([23_tag_tiles_analyzer.js](23_tag_tiles_analyzer.js:1))
   - Loads category HTML, meta, products from InSales
   - Calculates TF (Term Frequency): `(occurrences / total_words) * 100`
   - Compares against thresholds (Main: 2-4%, Additional: 1-2%, LSI: 0.5-1%)
   - Saves to "Анализ плотности" sheet

2. **Anchor Generation** ([23_tag_tiles_generator.js](23_tag_tiles_generator.js:1))
   - **Upper tile** (5-8 anchors): Navigation to child subcategories
   - **Lower tile** (15-30 anchors): SEO keywords with insufficient density
   - Uses Gemini API with structured prompts
   - Returns JSON with anchor text, target category, keywords covered

3. **HTML Generation** (PENDING) - Convert anchors to styled HTML tiles
4. **Review Interface** (PENDING) - Tables in detail sheets for user approval
5. **Apply to InSales** (PENDING) - Update category descriptions via API

See [docs/TAG_TILES_WORKFLOW.md](docs/TAG_TILES_WORKFLOW.md:1) for complete specifications.

## Common Development Tasks

### Adding a New Feature Module

1. **Choose numeric prefix** based on dependencies (e.g., 24_*.js for next feature)
2. **Add configuration** to [01_config.js](01_config.js:1) first
3. **Use consistent naming:**
   - `loadXxx()` - Fetch from API
   - `getXxx()` - Retrieve from sheets
   - `updateXxx()` / `setXxx()` - Write data
   - `showXxx()` - UI dialogs
   - `generateXxx()` - AI/content creation
4. **Add to menu** in [20_categories_menu.js](20_categories_menu.js:1)
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

### Modifying Sheet Structures

**Before changing:**
1. Check all references to column constants in `01_config.js`
2. Search codebase for hardcoded row numbers (e.g., "row 29", "B2")
3. Update `MAIN_LIST_COLUMNS` or `DETAIL_SHEET_SECTIONS` constants
4. Test on a single category before batch operations

**DO NOT:**
- Hardcode column indices - always use constants
- Hardcode sheet names - use `CATEGORY_SHEETS` constants
- Change cell B2 in detail sheets (category ID reference)
- Change row 29 (product section start)

### Working with AI APIs

**OpenAI:**
- API key in Script Properties or `OPENAI_CONFIG`
- Use `response_format: { type: 'json_object' }` for structured responses
- Handle rate limits with retries

**Gemini:**
- API key in Script Properties as `GEMINI_API_KEY`
- Setup docs: [docs/SETUP_GEMINI_API.md](docs/SETUP_GEMINI_API.md:1)
- Use detailed prompts with JSON schema in prompt text
- Parse response from `candidates[0].content.parts[0].text`

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

2. **Sheet cell references:** Category ID is ALWAYS at B2 in detail sheets. Products ALWAYS start at row 29. These are architectural constraints.

3. **Hierarchical display:** Main list shows tree structure with indentation. Must preserve visual hierarchy when updating.

4. **Keyword density thresholds:**
   - Main keywords: 2-4%
   - Additional: 1-2%
   - LSI: 0.5-1%
   - Spam threshold: >5%

5. **Tag tiles limits:**
   - Upper tile (navigation): 5-8 anchors
   - Lower tile (SEO): 15-30 anchors

## Debugging

### Logging
- Use `logInfo()`, `logWarning()`, `logError()` from [05_logging.js](05_logging.js:1)
- Logs saved to "Логи" sheet (auto-created)
- Include context string for filtering: `logInfo('Message', data, 'ModuleName')`

### Common Issues

**"parent_id: Не заполнены обязательные поля"**
- InSales requires parent_id for all category creation
- Solution: Always set `parent_id: 9069711` for root categories

**"Категории не найдены" in parent dropdown**
- Sheet name mismatch: Use `CATEGORY_SHEETS.MAIN_LIST` constant
- Column mismatch: Use `MAIN_LIST_COLUMNS` constants

**Script timeout**
- Operation taking >6 minutes
- Solution: Reduce batch size, add pagination, show progress toasts

**API rate limit**
- Too many requests to InSales/OpenAI/Gemini
- Solution: Increase `APP_CONFIG.apiDelay`, implement exponential backoff

## Reference

**Documentation:**
- [COMMANDS.md](COMMANDS.md:1) - Quick command reference for git, clasp, development workflow
- [docs/TAG_TILES_WORKFLOW.md](docs/TAG_TILES_WORKFLOW.md:1) - Tag tiles feature specification
- [docs/SETUP_GEMINI_API.md](docs/SETUP_GEMINI_API.md:1) - Gemini API setup instructions

**Key Files:**
- [01_config.js](01_config.js:1) - ALL configuration constants (start here)
- [13_categories_search.js](13_categories_search.js:1) - Search UI & detail sheet creation (largest module)
- [21_category_update.js](21_category_update.js:1) - Main update logic for InSales sync
- [22_category_create.js](22_category_create.js:1) - Category creation with parent_id handling

**InSales Documentation:**
- Base URL: https://binokl.shop
- Admin: https://myshop-on665.myinsales.ru/admin2
- API uses Basic Auth with base64-encoded credentials
