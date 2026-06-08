# FeedGen v3.0

Універсальний AI-генератор для товарних фідів Google Merchant Center на базі Claude API. Працює з будь-якою товарною категорією.

## Що нового у v3.0

**1. Мульти-формат імпорту/експорту**
Тепер сервіс приймає XLSX, CSV та XML (формат Google Shopping RSS). Експортувати результат можна в будь-якому з цих форматів незалежно від формату імпорту. CSV автоматично визначає роздільник (кома/крапка з комою/таб) і кодування (UTF-8, CP1251).

**2. Налаштування структури заголовка (drag-and-drop)**
На кроці налаштувань є список атрибутів, які можна перетягувати щоб задати порядок їх появи в заголовку. Кожен атрибут можна увімкнути/вимкнути перемикачем. Внизу показується live-превʼю структури. Порядок передається в AI як інструкція.

**3. Доповнення відсутніх атрибутів**
Якщо у фіді бракує стандартних GMC-атрибутів (color, material, gender, age_group, size, pattern, product_type, google_product_category тощо), сервіс пропонує їх згенерувати. AI визначає значення з даних товару і додає окремими колонками. Неповний фід стає повноцінним. Показуються лише ті атрибути, яких реально немає у фіді.

**4. Прибрано орієнтацію на нішу одягу**
Інтерфейс і движок працюють з повним набором з 20 GMC-атрибутів, а не з фіксованим clothing-набором. Для електроніки будуть свої поля, для книг свої.

**5. Opus прибрано**
Залишились Sonnet 4.6 і Haiku 4.5 (Opus 4.8 має інший API без temperature, несумісний з поточною логікою).

## Збережені фічі v2

Prompt caching (економія до 90%), retry logic з exponential backoff, dynamic column detection, валідація результатів, статистика генерації.

## Структура

```
feedgen-v3/
├── Dockerfile
├── requirements.txt
├── .gitignore
├── app.py
├── templates/
│   └── index.html
└── static/
    ├── style.css
    └── app.js
```

## Деплой на Railway

1. Завантаж вміст папки (не саму папку) в GitHub репозиторій
2. Railway → Deploy from GitHub
3. Settings → Build → Builder: **Dockerfile**
4. Variables → `ANTHROPIC_API_KEY=sk-ant-...`
5. Settings → Networking → Generate Domain (port 8000)

## API Endpoints

- `GET /` — головна
- `GET /api/schema` — список атрибутів (структуровані + генеровані)
- `POST /api/analyze` — аналіз фіду (xlsx/csv/xml), повертає present/missing атрибути
- `POST /api/generate` — запуск генерації з опціями (title_order, generate_attributes, output_format)
- `GET /api/status/{job_id}` — статус
- `GET /api/download/{job_id}` — завантаження в обраному форматі
- `GET /api/health` — health check

## Повний набір GMC-атрибутів

Розпізнаються і обробляються: id, title, description, link, image_link, product_type, google_product_category, brand, gtin, mpn, color, material, gender, age_group, size, pattern, condition, price, availability, product_highlight.

Генеровані (можуть доповнюватись): product_type, google_product_category, brand, color, material, gender, age_group, size, pattern, condition, product_highlight.

## Вартість

| Модель | Без кешу | З кешем |
|--------|----------|---------|
| Haiku 4.5 | ~$5.70 | ~$1.00 |
| Sonnet 4.6 | ~$17 | ~$3.00 |

(на фіді ~1000 рядків / ~260 унікальних товарів). Railway: $5/міс.

## TODO для подальшого розвитку

- Batch API для фідів 10k+ (50% знижка)
- Persistent storage (Redis/PostgreSQL) замість in-memory jobs
- Авторизація користувачів
- Preview результатів перед скачуванням
- Підтримка Opus 4.8 (адаптувати виклик під новий API без temperature)
- Google Sheets імпорт/експорт
- Збереження пресетів структури title для повторного використання
