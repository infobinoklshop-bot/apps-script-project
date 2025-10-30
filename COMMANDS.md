# 📚 Полная инструкция по работе с проектом

## 🎯 Схема работы

### Вариант 1: Редактирую код локально в VS Code (основной)

1. **Открываю проект:**
```bash
   cd ~/Documents/AppsScriptProject
   code .
```

2. **Редактирую файлы** в VS Code
   - Нахожу нужный файл слева
   - Открываю, редактирую
   - Сохраняю (Cmd+S)

3. **Загружаю в Apps Script:**
```bash
   clasp push
```

4. **Сохраняю в Git (историю версий):**
```bash
   git add .
   git commit -m "Описание что изменил"
   git push
```

5. **Проверяю** работу в веб-приложении или Apps Script

---

### Вариант 2: Редактирую в Apps Script (быстрые правки)

1. Открываю https://script.google.com
2. Редактирую код
3. Сохраняю (Ctrl+S)
4. Синхронизирую обратно:
```bash
   cd ~/Documents/AppsScriptProject
   clasp pull
   git add .
   git commit -m "Быстрое исправление через Apps Script"
   git push
```

---

### Вариант 3: Использую Claude Code для помощи

1. **Открываю Claude Code:** https://claude.ai/code
2. **Подключаю GitHub:** кнопка в интерфейсе
3. **Описываю проблему:** "У меня ошибка HTTP 422..."
4. Claude Code анализирует код и предлагает исправления
5. **Создаю Pull Request** (кнопка "Create PR")
6. **Применяю изменения:**
```bash
   git checkout main
   git pull
   clasp push
```

---

## 🚀 Быстрые команды

### Открыть проект
```bash
cd ~/Documents/AppsScriptProject
code .
```

### Деплой (всё сразу)
```bash
clasp push && git add . && git commit -m "Update" && git push
```

### Получить изменения с GitHub
```bash
git pull
clasp push
```

### Посмотреть что изменилось
```bash
git status
```

### Откатить изменения (если что-то сломалось)
```bash
git checkout .
```

### Посмотреть историю изменений
```bash
git log --oneline
```

---

## 🆘 Частые проблемы

### Ошибка при git push
```bash
git pull
git push
```

### Забыл что менял
```bash
git status
git diff
```

### Нужно отменить последний commit
```bash
git reset --soft HEAD~1
```

---

## 📂 Структура проекта

- `00_main.js` - главный файл
- `01_config.js` - конфигурация
- `22_category_create.js` - создание категорий
- `.clasp.json` - настройки clasp
- `COMMANDS.md` - эта шпаргалка

---

## 🔗 Полезные ссылки

- GitHub: https://github.com/infobinoklshop-bot/apps-script-project
- Apps Script: https://script.google.com
- Claude Code: https://claude.ai/code

---

## 💡 Советы

1. Всегда делайте `git commit` с понятным описанием
2. После каждой работы делайте `git push`
3. Используйте Claude Code для анализа ошибок
4. Держите эту шпаргалку открытой в VS Code

---

## 📖 Открыть эту шпаргалку
```bash
code ~/Documents/AppsScriptProject/COMMANDS.md
```

или
```bash
open ~/Documents/AppsScriptProject/COMMANDS.md
```---

## 📊 Обновление данных из Google Sheets

### Быстрое обновление:

1. Откройте Apps Script: https://script.google.com
2. Запустите функцию **testExportToDrive**
3. Скопируйте ссылку на папку из логов
4. Откройте папку → Выделите все (Cmd+A) → Скачать
5. Распакуйте и переместите CSV файлы в `data/`

### Ссылки:

- Таблица: [вставьте ссылку на вашу таблицу]
- Последний экспорт: https://drive.google.com/drive/folders/12xAyhviznwrx86eb95UVCfLtKtIkC3X7

### Файлы данных:
```
data/
├── Категории_—_Список.csv (1006 строк, 291KB)
├── Категория_—_Новые_лорнеты.csv (57 строк)
├── Категория_—_лорнет.csv (57 строк)
├── Категория_—_Театральные.csv (157 строк)
├── История_позиций.csv (2 строки)
├── Ключевые_слова.csv (5 строк)
└── Каталог_для_подбора.csv (большой, 1.3MB)
```

### Claude Code может:

- ✅ Анализировать структуру данных
- ✅ Находить проблемы и дубликаты
- ✅ Предлагать оптимизации кода на основе данных
- ✅ Генерировать валидацию и отчёты