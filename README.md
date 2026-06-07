# FeedGen v2.0

AI-генератор заголовків та описів для товарних фідів Google Merchant Center на базі Claude API. **Універсальний** — працює з будь-якою товарною категорією (одяг, електроніка, книги, меблі тощо).

## Що нового у v2.0

- **Універсальний движок** — жодної прив'язки до ніші. Атрибути товару передаються в промпт динамічно: для одягу це колір/матеріал/розмір, для електроніки — gtin/стан/ціна, і т.д. Сервіс бере тільки ті поля, що є у фіді.
- **Prompt caching** — system prompt кешується, економія до 90% на input токенах.
- **Retry logic** — exponential backoff (2/4/8 сек) при rate limit та серверних помилках.
- **Dynamic column detection** — автоматично розпізнає структуру фіду за назвами колонок (UA та EN aliases). Працює з будь-яким стандартним GMC фідом.
- **Result validation** — перевірка довжини title, заборонених слів (з конфігу), унікальності.
- **Статистика** — унікальність, fallback count, розподіл оцінок.

## Кастомізація під нішу/бренд

Сервіс універсальний за замовчуванням. Щоб заточити під конкретний бренд — завантаж JSON-конфіг ніші при генерації. Конфіг стає system prompt і може містити: tone of voice, заборонені слова, правила структури title, граматичні таблиці тощо.

Banned words для валідації беруться з конфігу автоматично — підтримуються поля `banned_words` (top-level), `tone_of_voice.banned_words`, `tone_of_voice.word_taboos`, `brand_config.banned_words`. Тобто будь-яка структура конфігу спрацює.

## Структура

```
feedgen-v2/
├── Dockerfile
├── requirements.txt
├── .gitignore
├── app.py              # Бекенд (FastAPI)
├── templates/
│   └── index.html
└── static/
    ├── style.css
    └── app.js
```

## Деплой на Railway

1. Завантаж вміст папки (не саму папку) в GitHub репозиторій
2. Railway → New Project → Deploy from GitHub
3. Settings → Build → Builder: **Dockerfile**
4. Variables → додай `ANTHROPIC_API_KEY=sk-ant-...`
5. Settings → Networking → Generate Domain (target port 8000)

## Локальний запуск

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...
uvicorn app:app --host 0.0.0.0 --port 8000
```

Відкрий http://localhost:8000

## Prompt caching — як працює

System prompt (JSON-конфіг ніші) позначається `cache_control: ephemeral`. Перший запит записує кеш (1.25x вартості), наступні читають з кешу (0.1x). Оскільки конфіг однаковий для всіх товарів у фіді, економія величезна.

Кеш живе 5 хвилин між запитами. При генерації 262 товарів підряд (з паузою 0.1 сек) кеш не встигає протухнути.

## Вартість (реальна, на фіді ~1000 рядків / ~260 унікальних товарів)

| Модель | Без кешу | З кешем |
|--------|----------|---------|
| Haiku 4.5 | ~$5.70 | ~$1.00 |
| Sonnet 4 | ~$17.30 | ~$3.00 |

Вартість залежить від розміру конфігу ніші (більший конфіг = дорожчий system prompt, але кешування нівелює різницю). Railway: $5/міс (Hobby план).

## API Endpoints

- `GET /` — головна сторінка
- `POST /api/analyze` — аналіз фіду, повертає структуру і розпізнані колонки
- `POST /api/generate` — запуск генерації (фоновий task)
- `GET /api/status/{job_id}` — статус генерації
- `GET /api/download/{job_id}` — завантаження результату
- `GET /api/health` — health check

## Що ще можна додати (TODO)

- Batch API для фідів 10k+ (50% знижка)
- Persistent storage (Redis/PostgreSQL) замість in-memory jobs
- Авторизація користувачів
- Preview результатів перед скачуванням
- UI для ручного маппінгу колонок (якщо автодетект помилився)
- Конфіг-білдер замість завантаження JSON
