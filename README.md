# FeedGen v4.0

AI-генератор заголовків, описів та атрибутів для товарних фідів Google Merchant Center на базі Claude API. Універсальний - працює з будь-якою категорією товарів. Формати: XLSX, CSV, XML (Google Shopping).

## Флоу користувача

1. Завантаж фід - сервіс сам розпізнає колонки (UA/EN aliases)
2. Налаштуй структуру title (drag-and-drop), вибери атрибути для доповнення
3. **Прев'ю на 5 товарах** - подивись якість і орієнтовну вартість повного запуску до того, як витрачати гроші
4. Підтверди повну генерацію
5. Скачай результат: всі оригінальні колонки + згенеровані

Генерацію можна скасувати в процесі - вже згенероване збережеться в частковий файл. Налаштування (модель, мова, порядок атрибутів) запам'ятовуються між сесіями.

## Основне під капотом

- **AsyncAnthropic + паралельна генерація** (semaphore, дефолт 5 потоків) з prompt caching і прогрівом кешу
- **Оцінка вартості** рахується з реальних токенів прев'ю, не з теорії
- **Токен-авторизація** з ізоляцією клієнтів: чужі файли і job'и недоступні
- **Валідація всього**: file_id, модель, формат, мова, розміри файлів, ліміт товарів
- **defusedxml** для парсингу, захист від path traversal, rate limit на логін
- Дедуплікація враховує атрибути варіантів (розмір/колір не злипаються при доповненні)
- Валідація результату: довжина title, banned words (з конфігу ніші), дублікати

## Структура

```
feedgen-v4/
├── app.py              # бекенд (FastAPI)
├── requirements.txt
├── templates/
│   ├── index.html
│   └── login.html
└── static/
    ├── style.css
    └── app.js
```

## Деплой на Railway

1. Залий вміст папки в GitHub репозиторій
2. Railway → New Project → Deploy from GitHub
3. Variables: `ANTHROPIC_API_KEY=sk-ant-...`
4. Опційно: `ACCESS_TOKENS=token1:client1,token2:client2` (без цієї змінної доступ відкритий)
5. Settings → Networking → Generate Domain (port 8000)

Інші env (опційно): `MAX_FEED_SIZE_MB=50`, `MAX_UNIQUE_PRODUCTS=3000`, `GENERATION_CONCURRENCY=5`, `COOKIE_INSECURE=1` (тільки для локального http).

Токени читаються при старті - зміна ACCESS_TOKENS вимагає redeploy.

## Локальний запуск

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...
export COOKIE_INSECURE=1
uvicorn app:app --host 0.0.0.0 --port 8000
```

## API

- `POST /api/analyze` - аналіз фіду, структура і колонки
- `POST /api/preview` - семпл на N товарів + оцінка вартості
- `POST /api/generate` - повна генерація (background job)
- `GET /api/status/{job_id}` - статус
- `POST /api/cancel/{job_id}` - скасувати, зберегти часткові результати
- `GET /api/download/{job_id}` - результат (частковий має суфікс _partial)
- `POST /api/login`, `/api/logout`, `GET /api/me` - авторизація
- `GET /api/schema`, `GET /api/health`

## Кастомізація під нішу

JSON-конфіг ніші при генерації стає system prompt: tone of voice, заборонені слова, правила структури. Banned words підтягуються з `banned_words`, `tone_of_voice.banned_words`, `tone_of_voice.word_taboos`, `brand_config.banned_words`.

## TODO

- Persistent storage (Redis) замість in-memory jobs
- Batch API для фідів 10k+ (50% знижка)
- Конфіг-білдер замість завантаження JSON
- Ручний маппінг колонок при помилці автодетекту
