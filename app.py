"""
FeedGen v3.0 - AI Feed Text Generator
======================================
Universal web service for generating optimized product titles, descriptions,
and missing attributes for Google Merchant Center feeds using Claude AI.

v3 features:
- Multi-format import: XLSX, CSV, XML (Google Shopping feed format)
- Full Google Merchant Center attribute set (not clothing-specific)
- Configurable attribute ordering for title/description structure (drag-and-drop)
- Attribute completion: fill missing GMC fields (color, material, gender, etc.)
- Prompt caching, retry logic with backoff, dynamic column detection, validation

Usage:
    pip install -r requirements.txt
    export ANTHROPIC_API_KEY=sk-ant-...
    uvicorn app:app --host 0.0.0.0 --port 8000
"""

import os
import io
import csv
import json
import re
import uuid
import asyncio
import tempfile
import xml.etree.ElementTree as ET
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

app = FastAPI(title="FeedGen", version="3.0")

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

# Models that support prompt caching (Opus excluded - different API surface)
CACHE_SUPPORTED_MODELS = [
    "claude-sonnet-4-6",
    "claude-haiku-4-5",
    "claude-haiku-4-5-20251001",
]


# =============================================================================
# GOOGLE MERCHANT CENTER ATTRIBUTE SCHEMA
# =============================================================================
# Full neutral GMC attribute set. Each attribute can be: detected from feed,
# used in title/description structure, and generated/completed if missing.

GMC_ATTRIBUTES = {
    # Core identifiers
    "id": {"aliases": ["id", "item_id", "offer_id", "товар_id"], "generatable": False, "label": "ID"},
    "title": {"aliases": ["title", "name", "назва", "найменування"], "generatable": False, "label": "Title"},
    "description": {"aliases": ["description", "desc", "опис"], "generatable": False, "label": "Description"},
    "link": {"aliases": ["link", "url", "посилання"], "generatable": False, "label": "Link"},
    "image_link": {"aliases": ["image_link", "image", "зображення"], "generatable": False, "label": "Image"},
    # Category
    "product_type": {"aliases": ["product_type", "тип_товару"], "generatable": True, "label": "Product Type"},
    "google_product_category": {"aliases": ["google_product_category", "google_category", "gpc"], "generatable": True, "label": "Google Category"},
    # Brand & identifiers
    "brand": {"aliases": ["brand", "manufacturer", "бренд", "виробник"], "generatable": True, "label": "Brand"},
    "gtin": {"aliases": ["gtin", "ean", "upc", "barcode"], "generatable": False, "label": "GTIN"},
    "mpn": {"aliases": ["mpn", "артикул"], "generatable": False, "label": "MPN"},
    # Variant attributes (generatable)
    "color": {"aliases": ["color", "colour", "колір"], "generatable": True, "label": "Color"},
    "material": {"aliases": ["material", "матеріал"], "generatable": True, "label": "Material"},
    "gender": {"aliases": ["gender", "стать"], "generatable": True, "label": "Gender"},
    "age_group": {"aliases": ["age_group", "вікова_група"], "generatable": True, "label": "Age Group"},
    "size": {"aliases": ["size", "розмір"], "generatable": True, "label": "Size"},
    "pattern": {"aliases": ["pattern", "візерунок", "малюнок"], "generatable": True, "label": "Pattern"},
    # Other
    "condition": {"aliases": ["condition", "стан"], "generatable": True, "label": "Condition"},
    "price": {"aliases": ["price", "ціна"], "generatable": False, "label": "Price"},
    "availability": {"aliases": ["availability", "наявність"], "generatable": False, "label": "Availability"},
    "product_highlight": {"aliases": ["product_highlight", "highlight"], "generatable": True, "label": "Highlight"},
}

CORE_ATTRIBUTES = ["id", "title", "description", "product_type", "brand"]

# Attributes that can appear in title/description structure (orderable)
STRUCTURABLE_ATTRIBUTES = [
    "brand", "gender", "age_group", "product_type", "color",
    "material", "pattern", "size", "google_product_category",
]

# Attributes the AI can generate/complete when missing from feed
GENERATABLE_ATTRIBUTES = [k for k, v in GMC_ATTRIBUTES.items() if v["generatable"]]

# Default ordering for title structure
DEFAULT_TITLE_ORDER = ["brand", "gender", "product_type", "color", "material", "pattern"]


def detect_columns(headers):
    """Maps feed headers to GMC attribute keys. Returns {key: column_index}."""
    mapping = {}
    headers_lower = [str(h).lower().strip() if h else "" for h in headers]
    for key, meta in GMC_ATTRIBUTES.items():
        for alias in meta["aliases"]:
            if alias in headers_lower:
                mapping[key] = headers_lower.index(alias)
                break
    return mapping


def detect_header_row(first_row):
    """Checks if first row is a header."""
    if not first_row or not isinstance(first_row[0], str):
        return False
    return str(first_row[0]).lower().strip() in ["id", "item_id", "offer_id", "товар_id"]


# =============================================================================
# MULTI-FORMAT FEED PARSING
# =============================================================================

def parse_feed_file(file_path, filename):
    """Parses XLSX, CSV, or XML feed. Returns (headers, data_rows)."""
    ext = filename.lower().rsplit(".", 1)[-1]

    if ext in ("xlsx", "xls"):
        return _parse_xlsx(file_path)
    elif ext == "csv":
        return _parse_csv(file_path)
    elif ext == "xml":
        return _parse_xml(file_path)
    else:
        raise ValueError(f"Непідтримуваний формат: .{ext}")


def _parse_xlsx(file_path):
    wb = openpyxl.load_workbook(file_path, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    if detect_header_row(rows[0]):
        headers = [str(h) if h else "" for h in rows[0]]
        return headers, [list(r) for r in rows[1:]]
    headers = [f"col_{i}" for i in range(len(rows[0]))]
    return headers, [list(r) for r in rows]


def _parse_csv(file_path):
    # Try common delimiters and encodings
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            with open(file_path, "r", encoding=encoding, newline="") as f:
                sample = f.read(4096)
                f.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
                    delimiter = dialect.delimiter
                except csv.Error:
                    delimiter = ","
                reader = csv.reader(f, delimiter=delimiter)
                rows = [row for row in reader if any(cell.strip() for cell in row)]
            if not rows:
                return [], []
            headers = [h.strip() for h in rows[0]]
            return headers, rows[1:]
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError("Не вдалося прочитати CSV (проблема з кодуванням)")


def _parse_xml(file_path):
    """Parses Google Shopping XML feed (RSS 2.0 with g: namespace)."""
    tree = ET.parse(file_path)
    root = tree.getroot()

    ns = {"g": "http://base.google.com/ns/1.0"}
    items = root.findall(".//item")
    if not items:
        # Try without channel wrapper
        items = root.findall(".//{http://base.google.com/ns/1.0}item") or root.findall("item")

    if not items:
        raise ValueError("XML не містить <item> елементів (очікується формат Google Shopping)")

    # Collect all attribute names across items
    all_keys = []
    seen_keys = set()
    parsed_items = []

    for item in items:
        item_data = {}
        for child in item:
            tag = child.tag
            # Strip namespace
            if "}" in tag:
                tag = tag.split("}")[-1]
            tag = tag.lower()
            val = (child.text or "").strip()
            item_data[tag] = val
            if tag not in seen_keys:
                seen_keys.add(tag)
                all_keys.append(tag)
        parsed_items.append(item_data)

    headers = all_keys
    data_rows = [[item.get(k, "") for k in headers] for item in parsed_items]
    return headers, data_rows


def write_output(output_rows, out_headers, output_format, output_path):
    """Writes results in requested format (xlsx, csv, or xml)."""
    if output_format == "csv":
        with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(out_headers)
            for row in output_rows:
                writer.writerow(row)
    elif output_format == "xml":
        ET.register_namespace("g", "http://base.google.com/ns/1.0")
        rss = ET.Element("rss", {"version": "2.0"})
        channel = ET.SubElement(rss, "channel")
        for row in output_rows:
            item = ET.SubElement(channel, "item")
            for h, val in zip(out_headers, row):
                el = ET.SubElement(item, f"{{http://base.google.com/ns/1.0}}{h}")
                el.text = str(val) if val is not None else ""
        tree = ET.ElementTree(rss)
        tree.write(output_path, encoding="utf-8", xml_declaration=True)
    else:  # xlsx
        out_wb = Workbook()
        out_ws = out_wb.active
        out_ws.title = "Generated Feed"
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
        out_wb.save(str(output_path))


# =============================================================================
# ROUTES
# =============================================================================

@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse(request=request, name="index.html")


@app.get("/api/schema")
async def get_schema():
    """Returns GMC attribute schema for frontend (orderable + generatable lists)."""
    return {
        "structurable": [
            {"key": k, "label": GMC_ATTRIBUTES[k]["label"]}
            for k in STRUCTURABLE_ATTRIBUTES
        ],
        "generatable": [
            {"key": k, "label": GMC_ATTRIBUTES[k]["label"]}
            for k in GENERATABLE_ATTRIBUTES
        ],
        "default_title_order": DEFAULT_TITLE_ORDER,
    }


@app.post("/api/analyze")
async def analyze_feed(file: UploadFile = File(...)):
    """Analyzes uploaded feed (xlsx/csv/xml). Returns structure + detected columns."""
    ext = file.filename.lower().rsplit(".", 1)[-1]
    if ext not in ("xlsx", "xls", "csv", "xml"):
        raise HTTPException(400, "Підтримуються формати: XLSX, CSV, XML")

    tmp_path = UPLOAD_DIR / f"{uuid.uuid4().hex}_{file.filename}"
    content = await file.read()
    tmp_path.write_bytes(content)

    try:
        headers, data_rows = parse_feed_file(tmp_path, file.filename)
        if not data_rows:
            raise HTTPException(400, "Файл порожній або не містить даних")

        col_map = detect_columns(headers)

        title_idx = col_map.get("title", 0)
        titles = Counter(
            row[title_idx] for row in data_rows
            if len(row) > title_idx and row[title_idx]
        )

        # Which generatable attributes are MISSING (could be completed)
        missing_generatable = [
            k for k in GENERATABLE_ATTRIBUTES if k not in col_map
        ]
        present_attributes = [k for k in col_map.keys()]

        # Sample values
        sample = {}
        for key, idx in col_map.items():
            if key in ["id", "description", "link", "image_link"]:
                continue
            vals = set()
            for row in data_rows[:100]:
                if len(row) > idx and row[idx]:
                    vals.add(str(row[idx]))
            if vals:
                sample[key] = sorted(vals)[:10]

        return {
            "file_id": tmp_path.stem,
            "filename": file.filename,
            "format": ext,
            "total_rows": len(data_rows),
            "unique_titles": len(titles),
            "columns": headers[:40],
            "detected_columns": col_map,
            "present_attributes": present_attributes,
            "missing_generatable": missing_generatable,
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
    model: str = Form("claude-sonnet-4-6"),
    language: str = Form("uk"),
    column_map: str = Form(None),
    title_order: str = Form(None),
    generate_attributes: str = Form(None),
    output_format: str = Form("xlsx"),
):
    """Starts background generation with v3 options."""
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

    col_map = json.loads(column_map) if column_map else None
    order = json.loads(title_order) if title_order else DEFAULT_TITLE_ORDER
    gen_attrs = json.loads(generate_attributes) if generate_attributes else []

    job_id = uuid.uuid4().hex[:12]
    jobs[job_id] = {
        "status": "processing", "progress": 0, "total": 0,
        "message": "Підготовка...", "created": datetime.now().isoformat(),
        "output_file": None, "errors": [], "stats": {},
    }

    background_tasks.add_task(
        run_generation, job_id, str(input_path), matching[0].name, api_key,
        model, language, niche_config, col_map, order, gen_attrs, output_format,
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
    fmt = job.get("stats", {}).get("output_format", "xlsx")
    media_types = {
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "csv": "text/csv",
        "xml": "application/xml",
    }
    return FileResponse(
        job["output_file"],
        filename=f"generated_feed_{job_id}.{fmt}",
        media_type=media_types.get(fmt, "application/octet-stream"),
    )


@app.get("/api/health")
async def health():
    return {
        "status": "ok", "version": "3.0",
        "api_key_configured": bool(os.environ.get("ANTHROPIC_API_KEY")),
    }


# =============================================================================
# GENERATION ENGINE
# =============================================================================

def build_system_prompt(niche_config, language, title_order, gen_attrs):
    """Builds system prompt. Uses niche config if provided, else neutral default."""
    if niche_config:
        base = json.dumps(niche_config, ensure_ascii=False, indent=2)
        # Append structure/generation instructions
        extra = f"\n\n--- STRUCTURE & GENERATION SETTINGS ---\n"
        extra += f"Title attribute order: {' → '.join(title_order)}\n"
        if gen_attrs:
            extra += f"Generate these missing attributes if possible: {', '.join(gen_attrs)}\n"
        return base + extra

    lang_name = "українська" if language == "uk" else "English"
    order_str = " → ".join(title_order)

    prompt = f"""You are a specialized AI copywriter for Google Merchant Center product feeds.
Generation language: {lang_name}.

TASK — for each product, based ONLY on the provided attributes, generate:
1. An optimized title (max 150 chars). Follow this attribute order: {order_str}.
   Most important info in the first 70 chars. Skip attributes that are absent.
2. An optimized description (50-80 words) — factual, structured, highlighting key attributes.
"""

    if gen_attrs:
        attrs_list = ", ".join(gen_attrs)
        prompt += f"""3. Infer and fill these MISSING attributes when they can be reliably determined
   from the title/description/other data: {attrs_list}.
   Only fill an attribute if you are confident. Leave empty ("") if uncertain.
   Never invent values that contradict the data.
"""

    prompt += """
RULES:
- Use ONLY provided data. Never invent specifications.
- Each title must be unique.
- Maintain grammatical correctness in the target language (agreement, cases).
- Adapt tone to the product category implied by the attributes.

RESPONSE FORMAT — respond ONLY with valid JSON:
"""
    if gen_attrs:
        attrs_json = ", ".join(f'"{a}": "..."' for a in gen_attrs)
        prompt += f'{{"title": "...", "description": "...", "attributes": {{{attrs_json}}}}}'
    else:
        prompt += '{"title": "...", "description": "..."}'

    return prompt


def extract_product_data(row, col_map):
    """Extracts all available attributes dynamically."""
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
    for key in GMC_ATTRIBUTES:
        if key in CORE_ATTRIBUTES:
            continue
        val = get(key)
        if val:
            data["attributes"][key] = str(val)
    return data


def parse_claude_response(text, gen_attrs=None):
    """Extracts JSON from Claude response. Returns (title, description, attributes_dict)."""
    text = text.strip()
    candidates = []

    # Strategy 1: direct
    try:
        candidates.append(json.loads(text))
    except json.JSONDecodeError:
        pass
    # Strategy 2: markdown block
    cb = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if cb:
        try:
            candidates.append(json.loads(cb.group(1)))
        except json.JSONDecodeError:
            pass
    # Strategy 3: first { to last }
    f, l = text.find("{"), text.rfind("}")
    if f != -1 and l > f:
        try:
            candidates.append(json.loads(text[f:l + 1]))
        except json.JSONDecodeError:
            pass

    for result in candidates:
        if isinstance(result, dict) and "title" in result:
            return (
                result["title"],
                result.get("description", ""),
                result.get("attributes", {}) if gen_attrs else {},
            )
    return None, None, {}


async def call_claude_with_retry(client, model, system_blocks, product_data, gen_attrs, max_retries=3):
    """Calls Claude API with prompt caching and exponential backoff."""
    import anthropic

    attr_lines = [
        f"Current title: {product_data['title']}",
        f"Description: {product_data['description'][:300]}",
    ]
    if product_data.get("product_type"):
        attr_lines.append(f"Product type: {product_data['product_type']}")
    if product_data.get("brand"):
        attr_lines.append(f"Brand: {product_data['brand']}")
    for k, v in product_data.get("attributes", {}).items():
        attr_lines.append(f"{k.replace('_', ' ').capitalize()}: {v}")

    user_message = (
        "Generate optimized content for this product.\n\n"
        + "\n".join(attr_lines)
    )

    for attempt in range(max_retries):
        try:
            response = client.messages.create(
                model=model,
                max_tokens=1500,
                temperature=0.3,
                system=system_blocks,
                messages=[{"role": "user", "content": user_message}],
            )
            text = response.content[0].text.strip()
            title, desc, attrs = parse_claude_response(text, gen_attrs)
            if title is not None:
                return title, desc, attrs, None
            return None, None, {}, f"parse_failed: {text[:80]}"

        except anthropic.RateLimitError:
            await asyncio.sleep((2 ** attempt) * 2)
            continue
        except anthropic.APIStatusError as e:
            if e.status_code in (429, 500, 502, 503, 529):
                await asyncio.sleep((2 ** attempt) * 2)
                continue
            return None, None, {}, f"api_error: {str(e)[:120]}"
        except Exception as e:
            return None, None, {}, f"error: {str(e)[:120]}"

    return None, None, {}, "max_retries_exceeded"


def validate_result(title, banned_words, seen_titles):
    score = 5
    comments = []
    if len(title) > 150:
        score = min(score, 3)
        comments.append(f"title {len(title)} chars (>150)")
    tl = title.lower()
    for w in banned_words:
        if w.lower() in tl:
            score = min(score, 2)
            comments.append(f"banned word '{w}'")
    if title in seen_titles:
        score = min(score, 3)
        comments.append("duplicate title")
    return score, "; ".join(comments)


async def run_generation(job_id, input_path, filename, api_key, model, language,
                         niche_config, col_map_override, title_order, gen_attrs, output_format):
    """Background generation task."""
    import anthropic

    try:
        client = anthropic.Anthropic(api_key=api_key)
        system_text = build_system_prompt(niche_config, language, title_order, gen_attrs)

        use_caching = model in CACHE_SUPPORTED_MODELS
        if use_caching:
            system_blocks = [{"type": "text", "text": system_text, "cache_control": {"type": "ephemeral"}}]
        else:
            system_blocks = system_text

        # Banned words from config (flexible locations)
        banned_words = []
        if niche_config:
            cands = list(niche_config.get("banned_words", []) or [])
            tov = niche_config.get("tone_of_voice", {})
            if isinstance(tov, dict):
                cands += tov.get("banned_words", []) or []
                cands += tov.get("word_taboos", []) or []
            bc = niche_config.get("brand_config", {})
            if isinstance(bc, dict):
                cands += bc.get("banned_words", []) or []
            banned_words = [w for w in cands if isinstance(w, str) and len(w) > 2]

        # Parse feed (any format)
        headers, data_rows = parse_feed_file(Path(input_path), filename)
        col_map = col_map_override or detect_columns(headers)

        title_idx = col_map.get("title", 0)
        desc_idx = col_map.get("description", 1)
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

        for key, product in unique_products.items():
            gt, gd, gattrs, error = await call_claude_with_retry(
                client, model, system_blocks, product, gen_attrs
            )
            if gt is None:
                gt = product["title"]
                gd = str(product.get("description", ""))
                gattrs = {}
                fallback_count += 1
                if len(jobs[job_id]["errors"]) < 50:
                    jobs[job_id]["errors"].append(f"ID {product['id']}: {error}")
                score, comment = 1, f"fallback: {error}"
            else:
                score, comment = validate_result(gt, banned_words, seen_titles)

            seen_titles.add(gt)
            generated[key] = (gt, gd, gattrs, score, comment)
            processed += 1
            jobs[job_id]["progress"] = processed
            jobs[job_id]["message"] = f"Оброблено {processed}/{total_unique}"
            await asyncio.sleep(0.1)

        # Build output headers: original + generated + filled attributes
        out_headers = ["id", "generated_title", "title", "generated_description", "description"]
        for attr in gen_attrs:
            out_headers.append(f"generated_{attr}")
        out_headers += ["quality_score", "quality_comment"]

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
            out_row = [rid, g[0], orig_title, g[1], orig_desc]
            for attr in gen_attrs:
                out_row.append(g[2].get(attr, ""))
            out_row += [g[3], g[4]]
            output_rows.append(out_row)

        # Write in requested format
        output_path = UPLOAD_DIR / f"generated_{job_id}.{output_format}"
        write_output(output_rows, out_headers, output_format, str(output_path))

        score_dist = Counter(g[3] for g in generated.values())
        unique_gen = len(set(g[0] for g in generated.values()))

        jobs[job_id]["status"] = "done"
        jobs[job_id]["output_file"] = str(output_path)
        jobs[job_id]["stats"] = {
            "total_rows": len(output_rows),
            "unique_products": total_unique,
            "unique_generated": unique_gen,
            "fallback_count": fallback_count,
            "caching_enabled": use_caching,
            "generated_attributes": gen_attrs,
            "output_format": output_format,
            "score_distribution": dict(score_dist),
        }
        jobs[job_id]["message"] = (
            f"Готово! {len(output_rows)} товарів. Fallback: {fallback_count}. "
            f"Формат: {output_format.upper()}"
        )

    except Exception as e:
        jobs[job_id]["status"] = "error"
        jobs[job_id]["message"] = f"Помилка: {str(e)}"
