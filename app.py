"""
FeedGen v2.0 - AI Feed Text Generator
======================================
Web service for generating optimized product titles and descriptions
for Google Merchant Center feeds using Claude AI.

Key features:
- Prompt caching (up to 90% cost reduction on system prompt)
- Retry logic with exponential backoff
- Dynamic column detection
- Result validation (length, banned words, uniqueness)
- Robust JSON parsing (4 strategies)

Usage:
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY=sk-ant-...
    uvicorn app:app --host 0.0.0.0 --port 8000
"""

import os
import json
import re
import uuid
import asyncio
import tempfile
from pathlib import Path
from datetime import datetime
from collections import Counter

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.requests import Request

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

app = FastAPI(title="FeedGen", version="2.0")

# Ensure required directories exist (prevents crash if not deployed correctly)
BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

jobs = {}
UPLOAD_DIR = Path(tempfile.gettempdir()) / "feedgen"
UPLOAD_DIR.mkdir(exist_ok=True)

# Models that support prompt caching
CACHE_SUPPORTED_MODELS = [
    "claude-sonnet-4-20250514",
    "claude-haiku-4-5-20251001",
]


# =============================================================================
# COLUMN DETECTION
# =============================================================================

# Standard Google Merchant Center attributes -> recognized header aliases.
# Neutral set - works for any product category (clothing, electronics, books, etc).
COLUMN_ALIASES = {
    "id": ["id", "item_id", "offer_id", "товар_id"],
    "title": ["title", "name", "назва", "найменування"],
    "description": ["description", "desc", "опис"],
    "product_type": ["product_type", "category", "google_product_category", "категорія", "тип_товару"],
    "brand": ["brand", "manufacturer", "бренд", "виробник"],
    "link": ["link", "url", "посилання"],
    # Optional attributes - passed to prompt if present, not required
    "gender": ["gender", "стать"],
    "color": ["color", "colour", "колір"],
    "material": ["material", "матеріал"],
    "size": ["size", "розмір"],
    "age_group": ["age_group", "вікова_група"],
    "pattern": ["pattern", "візерунок", "малюнок"],
    "condition": ["condition", "стан"],
    "gtin": ["gtin", "ean", "upc"],
    "mpn": ["mpn", "артикул"],
    "price": ["price", "ціна"],
    "product_highlight": ["product_highlight", "highlight"],
}

# Attributes that are core (always extracted). Everything else is optional
# and passed to the prompt dynamically only when present in the feed.
CORE_ATTRIBUTES = ["id", "title", "description", "product_type", "brand"]
OPTIONAL_ATTRIBUTES = [
    "gender", "color", "material", "size", "age_group",
    "pattern", "condition", "product_highlight",
]


def detect_columns(headers):
    """Maps feed headers to internal keys. Returns dict {key: column_index}."""
    mapping = {}
    headers_lower = [str(h).lower().strip() if h else "" for h in headers]

    for key, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in headers_lower:
                mapping[key] = headers_lower.index(alias)
                break

    return mapping


def detect_header_row(first_row):
    """Checks if first row is a header (contains known column names)."""
    if not first_row or not isinstance(first_row[0], str):
        return False
    first_lower = str(first_row[0]).lower().strip()
    return first_lower in ["id", "item_id", "товар_id"]


# =============================================================================
# ROUTES
# =============================================================================

@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")


@app.post("/api/analyze")
async def analyze_feed(file: UploadFile = File(...)):
    """Analyzes uploaded xlsx and returns feed structure with detected columns."""
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Підтримуються лише .xlsx файли")

    tmp_path = UPLOAD_DIR / f"{uuid.uuid4().hex}_{file.filename}"
    content = await file.read()
    tmp_path.write_bytes(content)

    try:
        wb = openpyxl.load_workbook(tmp_path, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))

        if not rows:
            raise HTTPException(400, "Файл порожній")

        has_header = detect_header_row(rows[0])

        if has_header:
            headers = [str(h) if h else "" for h in rows[0]]
            data_rows = rows[1:]
            col_map = detect_columns(headers)
        else:
            # No header - assume standard Google Merchant Center order
            headers = [f"col_{i}" for i in range(len(rows[0]))]
            data_rows = rows
            # Default GMC column positions
            col_map = {
                "id": 0, "title": 1, "description": 2,
                "product_type": 7, "brand": 16, "gender": 18,
                "color": 21, "material": 22,
            }

        # Unique products by title
        title_idx = col_map.get("title", 1)
        titles = Counter(
            row[title_idx] for row in data_rows
            if len(row) > title_idx and row[title_idx]
        )

        # Sample values for detected columns
        sample = {}
        for key, idx in col_map.items():
            if key in ["id", "description", "link"]:
                continue
            vals = set()
            for row in data_rows[:100]:
                if len(row) > idx and row[idx]:
                    vals.add(str(row[idx]))
            sample[key] = sorted(vals)[:10]

        return {
            "file_id": tmp_path.stem,
            "filename": file.filename,
            "total_rows": len(data_rows),
            "unique_titles": len(titles),
            "has_header": has_header,
            "columns": headers[:30],
            "detected_columns": col_map,
            "sample_values": sample,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, f"Помилка читання файлу: {str(e)}")


@app.post("/api/generate")
async def start_generation(
    background_tasks: BackgroundTasks,
    file_id: str = Form(...),
    config: UploadFile = File(None),
    config_text: str = Form(None),
    model: str = Form("claude-sonnet-4-20250514"),
    language: str = Form("uk"),
    column_map: str = Form(None),
):
    """Starts background generation."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(500, "ANTHROPIC_API_KEY не налаштований на сервері")

    matching = list(UPLOAD_DIR.glob(f"{file_id}*"))
    if not matching:
        raise HTTPException(404, "Файл не знайдено. Завантажте ще раз.")
    input_path = matching[0]

    niche_config = None
    if config:
        niche_config = json.loads(await config.read())
    elif config_text:
        niche_config = json.loads(config_text)

    # Parse column map override
    col_map = None
    if column_map:
        try:
            col_map = json.loads(column_map)
        except json.JSONDecodeError:
            pass

    job_id = uuid.uuid4().hex[:12]
    jobs[job_id] = {
        "status": "processing",
        "progress": 0,
        "total": 0,
        "message": "Підготовка...",
        "created": datetime.now().isoformat(),
        "output_file": None,
        "errors": [],
        "stats": {},
    }

    background_tasks.add_task(
        run_generation, job_id, str(input_path), api_key,
        model, language, niche_config, col_map
    )

    return {"job_id": job_id}


@app.get("/api/status/{job_id}")
async def get_status(job_id: str):
    if job_id not in jobs:
        raise HTTPException(404, "Завдання не знайдено")
    return jobs[job_id]


@app.get("/api/download/{job_id}")
async def download_result(job_id: str):
    from fastapi.responses import FileResponse
    if job_id not in jobs:
        raise HTTPException(404, "Завдання не знайдено")
    job = jobs[job_id]
    if job["status"] != "done" or not job["output_file"]:
        raise HTTPException(400, "Файл ще не готовий")
    return FileResponse(
        job["output_file"],
        filename=f"generated_feed_{job_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/health")
async def health():
    return {
        "status": "ok",
        "version": "2.0",
        "api_key_configured": bool(os.environ.get("ANTHROPIC_API_KEY")),
    }


# =============================================================================
# GENERATION ENGINE
# =============================================================================

def build_system_prompt(niche_config, language):
    """Builds system prompt from niche config or default."""
    if niche_config:
        return json.dumps(niche_config, ensure_ascii=False, indent=2)

    lang_name = "українська" if language == "uk" else "English"
    return f"""You are a specialized AI copywriter for Google Merchant Center product feeds.
Generation language: {lang_name}.

Your task: based on the provided product attributes, generate:
1. An optimized title (max 150 characters) — key attributes first, for maximum search relevance
2. An optimized description (50-80 words) — structured, factual, highlighting the product's attributes

Rules:
- Use ONLY the attributes provided. Never invent specifications that are not in the data.
- Title: most important information in the first 70 characters
- Natural language, no keyword stuffing
- Each title must be unique
- Maintain grammatical correctness in the target language (agreement, cases, etc.)
- Adapt tone and structure to the product category implied by the attributes

Respond ONLY with valid JSON: {{"title": "...", "description": "..."}}"""


def extract_product_data(row, col_map):
    """Extracts all available attributes using dynamic column mapping.
    Returns dict with core fields always present and optional fields only when in feed."""
    def get(key, default=""):
        idx = col_map.get(key)
        if idx is not None and len(row) > idx and row[idx]:
            return row[idx]
        return default

    data = {
        "id": get("id"),
        "title": get("title"),
        "description": str(get("description"))[:500],
        "product_type": get("product_type"),
        "brand": get("brand"),
        "attributes": {},
    }
    # Collect any optional attribute that exists in this feed
    for attr in OPTIONAL_ATTRIBUTES:
        val = get(attr)
        if val:
            data["attributes"][attr] = str(val)
    return data


def parse_claude_response(text):
    """4-strategy JSON extraction from Claude response."""
    text = text.strip()

    # Strategy 1: direct parse
    try:
        result = json.loads(text)
        if "title" in result:
            return result["title"], result.get("description", "")
    except json.JSONDecodeError:
        pass

    # Strategy 2: markdown code block
    code_block = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if code_block:
        try:
            result = json.loads(code_block.group(1))
            if "title" in result:
                return result["title"], result.get("description", "")
        except json.JSONDecodeError:
            pass

    # Strategy 3: simple brace match
    brace = re.search(r'\{[^{}]*"title"[^{}]*\}', text, re.DOTALL)
    if brace:
        try:
            result = json.loads(brace.group(0))
            if "title" in result:
                return result["title"], result.get("description", "")
        except json.JSONDecodeError:
            pass

    # Strategy 4: first { to last }
    first, last = text.find("{"), text.rfind("}")
    if first != -1 and last > first:
        try:
            result = json.loads(text[first:last + 1])
            if "title" in result:
                return result["title"], result.get("description", "")
        except json.JSONDecodeError:
            pass

    return None, None


async def call_claude_with_retry(client, model, system_blocks, product_data, max_retries=3):
    """Calls Claude API with prompt caching and exponential backoff retry."""
    import anthropic

    # Build attribute lines dynamically - only what's present in this feed
    attr_lines = [
        f"Current title: {product_data['title']}",
        f"Description: {product_data['description'][:300]}",
    ]
    if product_data.get("product_type"):
        attr_lines.append(f"Product type: {product_data['product_type']}")
    if product_data.get("brand"):
        attr_lines.append(f"Brand: {product_data['brand']}")
    for attr_name, attr_val in product_data.get("attributes", {}).items():
        attr_lines.append(f"{attr_name.replace('_', ' ').capitalize()}: {attr_val}")

    user_message = (
        "Generate an optimized title and description for this product.\n\n"
        + "\n".join(attr_lines)
        + '\n\nRespond ONLY with valid JSON: {"title": "...", "description": "..."}'
    )

    for attempt in range(max_retries):
        try:
            response = client.messages.create(
                model=model,
                max_tokens=1024,
                temperature=0.3,
                system=system_blocks,  # list with cache_control
                messages=[{"role": "user", "content": user_message}],
            )
            text = response.content[0].text.strip()
            title, desc = parse_claude_response(text)
            if title is not None:
                return title, desc, None
            return None, None, f"parse_failed: {text[:80]}"

        except anthropic.RateLimitError:
            wait = (2 ** attempt) * 2  # 2, 4, 8 seconds
            await asyncio.sleep(wait)
            continue
        except anthropic.APIStatusError as e:
            if e.status_code in (429, 500, 502, 503, 529):
                wait = (2 ** attempt) * 2
                await asyncio.sleep(wait)
                continue
            return None, None, f"api_error: {str(e)[:80]}"
        except Exception as e:
            return None, None, f"error: {str(e)[:80]}"

    return None, None, "max_retries_exceeded"


def validate_result(title, banned_words, seen_titles):
    """Validates generated title. Returns (score, comment)."""
    score = 5
    comments = []

    if len(title) > 150:
        score = min(score, 3)
        comments.append(f"title {len(title)} chars (>150)")

    title_lower = title.lower()
    for word in banned_words:
        if word.lower() in title_lower:
            score = min(score, 2)
            comments.append(f"banned word '{word}'")

    if title in seen_titles:
        score = min(score, 3)
        comments.append("duplicate title")

    return score, "; ".join(comments)


async def run_generation(job_id, input_path, api_key, model, language, niche_config, col_map_override):
    """Background generation task with prompt caching."""
    import anthropic

    try:
        client = anthropic.Anthropic(api_key=api_key)
        system_text = build_system_prompt(niche_config, language)

        # Build system blocks with prompt caching
        use_caching = model in CACHE_SUPPORTED_MODELS
        if use_caching:
            system_blocks = [
                {
                    "type": "text",
                    "text": system_text,
                    "cache_control": {"type": "ephemeral"},
                }
            ]
        else:
            system_blocks = system_text

        # Extract banned words from config for validation.
        # Checks multiple possible locations so any config structure works.
        banned_words = []
        if niche_config:
            candidates = []
            # Top-level neutral field
            candidates += niche_config.get("banned_words", []) or []
            # Nested under tone_of_voice (common in brand configs)
            tov = niche_config.get("tone_of_voice", {})
            if isinstance(tov, dict):
                candidates += tov.get("banned_words", []) or []
                candidates += tov.get("word_taboos", []) or []
            # Nested under brand_config
            bc = niche_config.get("brand_config", {})
            if isinstance(bc, dict):
                candidates += bc.get("banned_words", []) or []
            banned_words = [w for w in candidates if isinstance(w, str) and len(w) > 2]

        # Read feed
        wb = openpyxl.load_workbook(input_path, read_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))

        has_header = detect_header_row(all_rows[0])
        if has_header:
            headers = [str(h) if h else "" for h in all_rows[0]]
            data_rows = all_rows[1:]
            col_map = col_map_override or detect_columns(headers)
        else:
            data_rows = all_rows
            col_map = col_map_override or {
                "id": 0, "title": 1, "description": 2,
                "product_type": 7, "brand": 16, "gender": 18,
                "color": 21, "material": 22,
            }

        title_idx = col_map.get("title", 1)
        desc_idx = col_map.get("description", 2)
        id_idx = col_map.get("id", 0)

        # Deduplicate
        unique_products = {}
        row_keys = []
        for row in data_rows:
            key = (
                row[title_idx] if len(row) > title_idx else None,
                row[desc_idx] if len(row) > desc_idx else None,
            )
            if key not in unique_products:
                unique_products[key] = extract_product_data(row, col_map)
            row_keys.append(key)

        total_unique = len(unique_products)
        jobs[job_id]["total"] = total_unique
        jobs[job_id]["message"] = f"Генерація {total_unique} унікальних товарів..."

        # Generate
        generated = {}
        seen_titles = set()
        processed = 0
        fallback_count = 0
        cache_hits = 0

        for key, product in unique_products.items():
            gen_title, gen_desc, error = await call_claude_with_retry(
                client, model, system_blocks, product
            )

            if gen_title is None:
                gen_title = product["title"]
                gen_desc = str(product.get("description", ""))
                fallback_count += 1
                if len(jobs[job_id]["errors"]) < 50:
                    jobs[job_id]["errors"].append(f"ID {product['id']}: {error}")
                score, comment = 1, f"fallback: {error}"
            else:
                score, comment = validate_result(gen_title, banned_words, seen_titles)

            seen_titles.add(gen_title)
            generated[key] = (gen_title, gen_desc, score, comment)
            processed += 1
            jobs[job_id]["progress"] = processed
            jobs[job_id]["message"] = f"Оброблено {processed}/{total_unique}"

            # With caching, no need for aggressive rate limiting
            await asyncio.sleep(0.1)

        # Build output
        output_rows = []
        for i, row in enumerate(data_rows):
            key = row_keys[i]
            g = generated[key]
            rid = row[id_idx] if len(row) > id_idx else ""
            try:
                rid = int(rid) if rid else ""
            except (ValueError, TypeError):
                pass
            orig_title = row[title_idx] if len(row) > title_idx else ""
            orig_desc = row[desc_idx] if len(row) > desc_idx else ""
            output_rows.append((rid, g[0], orig_title, g[1], orig_desc, g[2], g[3]))

        # Write xlsx
        out_wb = Workbook()
        out_ws = out_wb.active
        out_ws.title = "Generated Feed"

        out_headers = ["id", "generated_title", "title", "generated_description",
                       "description", "quality_score", "quality_comment"]
        hfont = Font(bold=True, size=11, name="Arial")
        hfill = PatternFill("solid", start_color="D9E1F2")

        for ci, h in enumerate(out_headers, 1):
            c = out_ws.cell(row=1, column=ci, value=h)
            c.font, c.fill = hfont, hfill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for ri, rd in enumerate(output_rows, 2):
            for ci, val in enumerate(rd, 1):
                c = out_ws.cell(row=ri, column=ci, value=val)
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(vertical="top", wrap_text=True)

        widths = {"A": 10, "B": 60, "C": 45, "D": 75, "E": 50, "F": 12, "G": 45}
        for col, w in widths.items():
            out_ws.column_dimensions[col].width = w
        out_ws.auto_filter.ref = f"A1:G{len(output_rows)+1}"

        output_path = UPLOAD_DIR / f"generated_{job_id}.xlsx"
        out_wb.save(str(output_path))

        # Stats
        score_dist = Counter(g[2] for g in generated.values())
        unique_gen = len(set(g[0] for g in generated.values()))

        jobs[job_id]["status"] = "done"
        jobs[job_id]["output_file"] = str(output_path)
        jobs[job_id]["stats"] = {
            "total_rows": len(output_rows),
            "unique_products": total_unique,
            "unique_generated": unique_gen,
            "fallback_count": fallback_count,
            "caching_enabled": use_caching,
            "score_distribution": dict(score_dist),
        }
        jobs[job_id]["message"] = (
            f"Готово! {len(output_rows)} товарів. "
            f"Fallback: {fallback_count}. Кешування: {'так' if use_caching else 'ні'}"
        )

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["message"] = f"Помилка: {str(e)}"
