/**
 * ========================================
 * ЛОГИКА УМНОЙ СОРТИРОВКИ (Pure Logic)
 * ========================================
 * 
 * Этот файл содержит чистую логику сортировки, отвязанную от UI Google Sheets.
 * Используется как для сортировки на листе, так и для прямой сортировки через API.
 */

/**
 * Вычисляет порядок товаров для умной сортировки
 * 
 * @param {Array<Object>} products - Массив товаров. Каждый объект должен иметь:
 *   {
 *     id: number|string,
 *     title: string,
 *     inStock: boolean, // true если в наличии
 *     brand: string,    // Бренд (нормализованный)
 *     originalPosition: number // Исходная позиция (для сохранения относительного порядка)
 *   }
 * @param {number} itemsPerPage - Количество товаров на первой странице
 * @returns {Array<Object>} - Отсортированный массив товаров
 */
function calculateSmartSortOrder(products, itemsPerPage) {
    // 1. Разделяем на "В наличии" и "Нет в наличии"
    const inStockItems = products.filter(p => p.inStock);
    const outOfStockItems = products.filter(p => !p.inStock);

    // 2. Группируем товары в наличии по брендам
    const inStockByBrand = {};
    const inStockNoBrand = [];

    inStockItems.forEach(item => {
        if (item.brand && item.brand !== 'Other') {
            if (!inStockByBrand[item.brand]) inStockByBrand[item.brand] = [];
            inStockByBrand[item.brand].push(item);
        } else {
            inStockNoBrand.push(item);
        }
    });

    // 3. ВНЕДРЕНИЕ СЛУЧАЙНОСТИ
    // Перемешиваем товары внутри каждого бренда
    Object.keys(inStockByBrand).forEach(brand => {
        shuffleArray(inStockByBrand[brand]);
    });
    // Перемешиваем товары без бренда
    shuffleArray(inStockNoBrand);

    // 4. ФОРМИРОВАНИЕ ПЕРВОЙ СТРАНИЦЫ
    const page1Items = [];

    // Перемешиваем список брендов, чтобы порядок брендов на 1-й странице менялся
    const brands = Object.keys(inStockByBrand);
    shuffleArray(brands);

    // Шаг 4.1: Берем по одному представителю каждого бренда
    brands.forEach(brand => {
        if (inStockByBrand[brand].length > 0) {
            page1Items.push(inStockByBrand[brand].shift());
        }
    });

    // Шаг 4.2: Заполняем оставшееся место на 1-й странице (Round Robin)
    let hasMore = true;

    // Для Round Robin создадим временную структуру
    const roundRobinGroups = [];
    brands.forEach(brand => {
        if (inStockByBrand[brand].length > 0) roundRobinGroups.push(inStockByBrand[brand]);
    });
    if (inStockNoBrand.length > 0) roundRobinGroups.push(inStockNoBrand);

    while (page1Items.length < itemsPerPage && hasMore) {
        hasMore = false;

        for (let i = 0; i < roundRobinGroups.length; i++) {
            if (page1Items.length >= itemsPerPage) break;

            const group = roundRobinGroups[i];
            if (group.length > 0) {
                page1Items.push(group.shift());
                hasMore = true;
            }
        }
    }

    // 5. ФОРМИРОВАНИЕ ХВОСТА (Остальные товары)
    // Собираем все оставшиеся товары, которые не попали на 1-ю страницу.
    // Важно: мы хотим сохранить их "стабильный" порядок, насколько это возможно,
    // или просто собрать остатки.
    // В текущей реализации мы уже "потратили" часть товаров из inStockByBrand.
    // Просто соберем всё, что осталось.

    const tailItems = [];

    // Собираем остатки из брендов
    brands.forEach(brand => {
        if (inStockByBrand[brand].length > 0) {
            tailItems.push(...inStockByBrand[brand]);
        }
    });
    // Собираем остатки без бренда
    if (inStockNoBrand.length > 0) {
        tailItems.push(...inStockNoBrand);
    }

    // 6. Сборка финального списка
    // Сначала 1-я страница
    // Потом Хвост (в наличии)
    // Потом Нет в наличии
    return [...page1Items, ...tailItems, ...outOfStockItems];
}

/**
 * Вспомогательная функция перемешивания массива (Fisher-Yates shuffle)
 */
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}

/**
 * Вспомогательная функция для определения бренда (нормализация)
 */
function normalizeBrand(brandName, productTitle) {
    let key = 'Other';

    if (brandName && brandName.toString().trim() !== '') {
        key = brandName.toString().trim();
    } else if (productTitle) {
        // Fallback: парсинг из названия
        const IGNORED_WORDS = new Set([
            'бинокль', 'монокуляр', 'труба', 'телескоп', 'микроскоп',
            'лупа', 'прицел', 'дальномер', 'тепловизор', 'пнв',
            'аксессуар', 'штатив', 'сумка', 'чехол', 'цифровой', 'ночного', 'видения'
        ]);

        const parts = productTitle.split(/[\s,.\(\)\-]+/).filter(p => p.trim() !== '');
        for (const part of parts) {
            const cleanWord = part.replace(/[^a-zA-Zа-яА-Я0-9]/g, '');
            if (cleanWord.length > 1 && !IGNORED_WORDS.has(cleanWord.toLowerCase())) {
                key = cleanWord;
                break; // Берем первое подходящее слово
            }
        }
    }

    // Нормализация ключа (Первая буква заглавная, остальные строчные)
    if (key && key !== 'Other') {
        return key.charAt(0).toUpperCase() + key.slice(1).toLowerCase();
    }
    return 'Other';
}
