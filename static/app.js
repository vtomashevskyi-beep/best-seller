// FeedGen v4.0 Frontend

const $ = (s) => document.querySelector(s);
const show = (el) => el.classList.remove('hidden');
const hide = (el) => el.classList.add('hidden');
// Escape user-controlled data before inserting into innerHTML (XSS protection)
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
));

const SETTINGS_KEY = 'feedgen_settings_v4';

let currentFileId = null;
let currentJobId = null;
let configFile = null;
let detectedColumns = {};
let missingGeneratable = [];
let schema = { structurable: [], generatable: [], default_title_order: [] };
let attrOrder = [];
let selectedGenAttrs = new Set();
let pollFailures = 0;
let previewDone = false;

// =============================================================================
// Settings persistence (localStorage)
// =============================================================================

function saveSettings() {
    try {
        localStorage.setItem(SETTINGS_KEY, JSON.stringify({
            model: $('#model-select').value,
            language: $('#lang-select').value,
            output_format: $('#format-select').value,
            attr_order: attrOrder.map(a => ({ key: a.key, enabled: a.enabled })),
            gen_attrs: [...selectedGenAttrs],
        }));
    } catch (e) { /* localStorage may be unavailable - not critical */ }
}

function loadSettings() {
    try {
        return JSON.parse(localStorage.getItem(SETTINGS_KEY)) || null;
    } catch (e) { return null; }
}

function applySavedSelects(saved) {
    if (!saved) return;
    if (saved.model && [...$('#model-select').options].some(o => o.value === saved.model)) {
        $('#model-select').value = saved.model;
    }
    if (saved.language) $('#lang-select').value = saved.language;
    if (saved.output_format) $('#format-select').value = saved.output_format;
}

// =============================================================================
// Init: load schema and user info
// =============================================================================

(async () => {
    try {
        const [schemaRes, meRes] = await Promise.all([
            fetch('/api/schema'),
            fetch('/api/me'),
        ]);
        schema = await schemaRes.json();
        const me = await meRes.json();
        if (me.auth_enabled && me.client) {
            const bar = document.getElementById('user-bar');
            if (bar) {
                bar.style.display = 'block';
                document.getElementById('user-label').textContent = `Клієнт: ${me.client}`;
            }
        }
        applySavedSelects(loadSettings());
    } catch (e) { console.error('Init failed', e); }
})();

// =============================================================================
// STEP 1: Upload
// =============================================================================

const dropzone = $('#dropzone');
const fileInput = $('#file-input');
const DROPZONE_DEFAULT = `
    <div class="upload-content">
        <div class="upload-icon">📄</div>
        <p>Перетягни файл сюди</p>
        <p class="upload-hint">XLSX, CSV або XML · або натисни для вибору</p>
    </div>`;

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', (e) => {
    e.preventDefault(); dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => { if (e.target.files.length) handleFile(e.target.files[0]); });

async function handleFile(file) {
    if (!file.name.match(/\.(xlsx|csv|xml)$/i)) {
        alert('Підтримуються XLSX, CSV, XML. Старий формат .xls треба пересохранити як .xlsx');
        return;
    }
    dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">⏳</div><p>Аналізую...</p></div>`;

    const fd = new FormData();
    fd.append('file', file);
    try {
        const res = await fetch('/api/analyze', { method: 'POST', body: fd });
        if (!res.ok) {
            let detail = `HTTP ${res.status}`;
            try { detail = (await res.json()).detail || detail; } catch (_) {}
            throw new Error(detail);
        }
        const d = await res.json();

        currentFileId = d.file_id;
        detectedColumns = d.detected_columns || {};
        missingGeneratable = d.missing_generatable || [];
        previewDone = false;
        hide($('#preview-panel'));

        dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">✅</div><p>${esc(d.filename)}</p><p class="upload-hint">${esc(d.format.toUpperCase())}</p></div>`;

        show($('#file-info'));
        $('#file-info').querySelector('.file-stats').innerHTML = `
            <div class="stat"><span class="stat-label">Рядків</span><span class="stat-value">${esc(d.total_rows)}</span></div>
            <div class="stat"><span class="stat-label">Унікальних товарів</span><span class="stat-value">${esc(d.unique_titles)}</span></div>
        `;

        const present = d.present_attributes || [];
        const tagsHtml = present.map(c => `<span class="col-tag found">${esc(c)} ✓</span>`).join('');
        let colsHtml = `<span style="font-size:12px;color:var(--text-dim);width:100%;margin-bottom:4px;">Знайдено у фіді:</span>${tagsHtml}`;
        if (d.over_limit) {
            colsHtml += `<span class="col-tag missing" style="width:100%">⚠ Унікальних товарів більше за ліміт (${esc(d.max_unique_products)}) - генерація не запуститься, розбий фід на частини</span>`;
        }
        $('#detected-cols').innerHTML = colsHtml;

        buildAttrOrder();
        buildGenAttrs();
        show($('#step-config'));
    } catch (err) {
        dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">❌</div><p>${esc(err.message)}</p><p class="upload-hint">Спробуй ще раз</p></div>`;
    }
}

// =============================================================================
// STEP 2: Attribute ordering (drag-and-drop)
// =============================================================================

function buildAttrOrder() {
    const saved = loadSettings();
    const labelMap = {};
    schema.structurable.forEach(s => labelMap[s.key] = s.label);

    if (saved && Array.isArray(saved.attr_order) && saved.attr_order.length) {
        // Restore saved order/toggles, keep only keys that still exist in schema
        attrOrder = saved.attr_order
            .filter(a => labelMap[a.key])
            .map(a => ({ key: a.key, label: labelMap[a.key], enabled: !!a.enabled }));
    } else {
        const order = schema.default_title_order.length ? schema.default_title_order : ['brand', 'product_type', 'color'];
        attrOrder = order.map(key => ({
            key,
            label: labelMap[key] || key,
            enabled: detectedColumns[key] !== undefined || key === 'brand' || key === 'product_type',
        }));
    }

    schema.structurable.forEach(s => {
        if (!attrOrder.find(a => a.key === s.key)) {
            attrOrder.push({ key: s.key, label: s.label, enabled: false });
        }
    });

    renderAttrOrder();
}

function renderAttrOrder() {
    const list = $('#attr-order-list');
    list.innerHTML = '';
    attrOrder.forEach((attr, idx) => {
        const li = document.createElement('li');
        li.className = 'attr-order-item' + (attr.enabled ? '' : ' disabled');
        li.draggable = true;
        li.dataset.key = attr.key;
        li.innerHTML = `
            <span class="attr-drag-handle">⠿</span>
            <span class="attr-order-num">${idx + 1}</span>
            <span class="attr-name">${esc(attr.label)}</span>
            <span class="attr-toggle ${attr.enabled ? 'on' : ''}" data-key="${esc(attr.key)}"></span>
        `;
        list.appendChild(li);
    });

    list.querySelectorAll('.attr-toggle').forEach(t => {
        t.addEventListener('click', (e) => {
            e.stopPropagation();
            const key = t.dataset.key;
            const attr = attrOrder.find(a => a.key === key);
            attr.enabled = !attr.enabled;
            invalidatePreview();
            renderAttrOrder();
            updateTitlePreview();
        });
    });

    let dragEl = null;
    list.querySelectorAll('.attr-order-item').forEach(item => {
        item.addEventListener('dragstart', () => { dragEl = item; item.classList.add('dragging'); });
        item.addEventListener('dragend', () => {
            item.classList.remove('dragging');
            const newOrder = [...list.querySelectorAll('.attr-order-item')].map(el => {
                const key = el.dataset.key;
                return attrOrder.find(a => a.key === key);
            });
            attrOrder = newOrder;
            invalidatePreview();
            renderAttrOrder();
            updateTitlePreview();
        });
        item.addEventListener('dragover', (e) => {
            e.preventDefault();
            const after = getDragAfter(list, e.clientY);
            if (after == null) list.appendChild(dragEl);
            else list.insertBefore(dragEl, after);
        });
    });

    updateTitlePreview();
}

function getDragAfter(list, y) {
    const els = [...list.querySelectorAll('.attr-order-item:not(.dragging)')];
    return els.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        if (offset < 0 && offset > closest.offset) return { offset, element: child };
        return closest;
    }, { offset: -Infinity }).element;
}

function updateTitlePreview() {
    const enabled = attrOrder.filter(a => a.enabled).map(a => a.label);
    const preview = enabled.length
        ? enabled.map(l => `<strong>{${esc(l)}}</strong>`).join(' ')
        : '<em>немає увімкнених атрибутів</em>';
    $('#title-preview').innerHTML = `Приклад: ${preview}`;
}

// =============================================================================
// STEP 2: Generatable attributes
// =============================================================================

function buildGenAttrs() {
    const container = $('#gen-attrs-list');
    selectedGenAttrs.clear();

    const labelMap = {};
    schema.generatable.forEach(g => labelMap[g.key] = g.label);
    const available = missingGeneratable.filter(k => labelMap[k]);

    if (available.length === 0) {
        container.innerHTML = '<span class="gen-attrs-empty">Усі атрибути вже присутні у фіді - доповнювати нічого.</span>';
        return;
    }

    const saved = loadSettings();
    const savedGen = new Set((saved && saved.gen_attrs) || []);

    container.innerHTML = '';
    available.forEach(key => {
        const chip = document.createElement('div');
        chip.className = 'gen-attr-chip';
        chip.dataset.key = key;
        chip.innerHTML = `<span class="check">＋</span> ${esc(labelMap[key])}`;
        const select = () => {
            selectedGenAttrs.add(key);
            chip.classList.add('selected');
            chip.querySelector('.check').textContent = '✓';
        };
        if (savedGen.has(key)) select();
        chip.addEventListener('click', () => {
            if (selectedGenAttrs.has(key)) {
                selectedGenAttrs.delete(key);
                chip.classList.remove('selected');
                chip.querySelector('.check').textContent = '＋';
            } else {
                select();
            }
            invalidatePreview();
        });
        container.appendChild(chip);
    });
}

// =============================================================================
// Config inputs
// =============================================================================

$('#config-input').addEventListener('change', (e) => {
    if (e.target.files.length) {
        configFile = e.target.files[0];
        $('#config-filename').textContent = configFile.name;
        invalidatePreview();
    }
});
['#model-select', '#lang-select', '#format-select'].forEach(sel => {
    $(sel).addEventListener('change', invalidatePreview);
});

function invalidatePreview() {
    // Settings changed after a preview - the sample no longer reflects them
    if (previewDone) {
        previewDone = false;
        const note = $('#preview-stale-note');
        if (note) show(note);
    }
}

function buildFormData(extra = {}) {
    const titleOrder = attrOrder.filter(a => a.enabled).map(a => a.key);
    const fd = new FormData();
    fd.append('file_id', currentFileId);
    fd.append('model', $('#model-select').value);
    fd.append('language', $('#lang-select').value);
    fd.append('column_map', JSON.stringify(detectedColumns));
    fd.append('title_order', JSON.stringify(titleOrder));
    fd.append('generate_attributes', JSON.stringify([...selectedGenAttrs]));
    for (const [k, v] of Object.entries(extra)) fd.append(k, v);
    if (configFile) fd.append('config', configFile);
    return fd;
}

// =============================================================================
// STEP 2.5: Preview (new in v4)
// =============================================================================

$('#btn-preview').addEventListener('click', async () => {
    const btn = $('#btn-preview');
    btn.disabled = true;
    btn.innerHTML = '<span class="btn-icon">⏳</span> Генерую прев\'ю...';
    saveSettings();

    try {
        const res = await fetch('/api/preview', { method: 'POST', body: buildFormData() });
        if (!res.ok) {
            let detail = `HTTP ${res.status}`;
            try { detail = (await res.json()).detail || detail; } catch (_) {}
            throw new Error(detail);
        }
        const d = await res.json();
        renderPreview(d);
        previewDone = true;
        hide($('#preview-stale-note'));
    } catch (err) {
        alert(`Помилка прев'ю: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<span class="btn-icon">🔍</span> Прев\'ю на 5 товарах';
    }
});

function renderPreview(d) {
    const panel = $('#preview-panel');
    show(panel);

    const rows = d.results.map(r => {
        const scoreClass = r.score >= 5 ? 'score-good' : (r.score >= 3 ? 'score-mid' : 'score-bad');
        const attrs = Object.entries(r.attributes || {})
            .map(([k, v]) => `${esc(k)}: ${esc(v)}`).join(', ');
        return `
            <div class="preview-item">
                <div class="preview-row"><span class="preview-label">ID</span><span>${esc(r.id)}</span><span class="preview-score ${scoreClass}">${esc(r.score)}/5</span></div>
                <div class="preview-row"><span class="preview-label">Було</span><span class="preview-old">${esc(r.original_title)}</span></div>
                <div class="preview-row"><span class="preview-label">Title</span><span class="preview-new">${esc(r.generated_title ?? '(помилка)')}</span></div>
                <div class="preview-row"><span class="preview-label">Опис</span><span>${esc(r.generated_description ?? '')}</span></div>
                ${attrs ? `<div class="preview-row"><span class="preview-label">Атрибути</span><span>${attrs}</span></div>` : ''}
                ${r.comment ? `<div class="preview-row preview-comment"><span class="preview-label">⚠</span><span>${esc(r.comment)}</span></div>` : ''}
            </div>`;
    }).join('');
    $('#preview-results').innerHTML = rows;

    let confirmHtml = `<div class="confirm-line">Повна генерація: <strong>${esc(d.total_unique)}</strong> унікальних товарів`;
    if (d.estimate) {
        confirmHtml += `, орієнтовно до <strong>$${esc(d.estimate.est_cost_usd_max)}</strong>`;
        confirmHtml += `<span class="confirm-note">${esc(d.estimate.note)}</span>`;
    }
    confirmHtml += `</div>`;
    if (d.over_limit) {
        confirmHtml += `<div class="confirm-line" style="color:var(--error)">⚠ Перевищено ліміт ${esc(d.max_unique_products)} товарів - повний запуск заблоковано</div>`;
    }
    $('#preview-confirm-info').innerHTML = confirmHtml;
    $('#btn-generate').disabled = !!d.over_limit;

    panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// =============================================================================
// Full generation (confirmation step)
// =============================================================================

$('#btn-generate').addEventListener('click', async () => {
    if (!previewDone) {
        const ok = confirm('Прев\'ю не запускалось (або налаштування змінились після нього). Запустити повну генерацію без перевірки на семплі?');
        if (!ok) return;
    }
    const btn = $('#btn-generate');
    btn.disabled = true;
    btn.innerHTML = '<span class="btn-icon">⏳</span> Запускаю...';
    saveSettings();

    try {
        const fd = buildFormData({ output_format: $('#format-select').value });
        const res = await fetch('/api/generate', { method: 'POST', body: fd });
        if (!res.ok) {
            let detail = `HTTP ${res.status}`;
            try { detail = (await res.json()).detail || detail; } catch (_) {}
            throw new Error(detail);
        }
        const d = await res.json();
        currentJobId = d.job_id;
        pollFailures = 0;
        hide($('#step-upload')); hide($('#step-config')); show($('#step-progress'));
        pollStatus();
    } catch (err) {
        alert(`Помилка: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<span class="btn-icon">⚡</span> Запустити повну генерацію';
    }
});

// =============================================================================
// Progress + Cancel + Result
// =============================================================================

$('#btn-cancel').addEventListener('click', async () => {
    const btn = $('#btn-cancel');
    btn.disabled = true;
    btn.textContent = 'Скасовую...';
    try {
        await fetch(`/api/cancel/${currentJobId}`, { method: 'POST' });
        // Status polling will pick up the "cancelled" state and show partial result
    } catch (err) {
        btn.disabled = false;
        btn.textContent = 'Скасувати';
    }
});

async function pollStatus() {
    try {
        const res = await fetch(`/api/status/${currentJobId}`);
        if (res.status === 404) {
            showError('Завдання не знайдено. Схоже, сервер перезапустився - завантаж фід і запусти генерацію ще раз.');
            return;
        }
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const d = await res.json();
        pollFailures = 0;
        const pct = d.total > 0 ? Math.round((d.progress / d.total) * 100) : 0;
        $('#progress-fill').style.width = `${pct}%`;
        $('#progress-pct').textContent = `${pct}%`;
        $('#progress-text').textContent = d.message;
        if (d.status === 'done') { showResult(d, false); return; }
        if (d.status === 'cancelled') { showResult(d, true); return; }
        if (d.status === 'error') { showError(d.message); return; }
        setTimeout(pollStatus, 1500);
    } catch (err) {
        pollFailures++;
        if (pollFailures >= 10) {
            showError('Втрачено зв\'язок із сервером. Онови сторінку і спробуй ще раз.');
            return;
        }
        setTimeout(pollStatus, 3000);
    }
}

function showResult(d, cancelled) {
    hide($('#step-progress')); show($('#step-result'));
    $('.result-icon').textContent = cancelled ? '⏹' : '✅';
    $('#result-message').textContent = d.message;
    show($('#btn-download'));
    if (cancelled) {
        $('#btn-download').innerHTML = '<span class="btn-icon">📥</span> Завантажити частковий файл';
    }
    if (d.stats) {
        const s = d.stats;
        const sd = s.score_distribution || {};
        const genAttrs = (s.generated_attributes || []).length;
        $('#result-stats').innerHTML = `
            <div class="result-stat"><span class="result-stat-label">Всього рядків</span><span class="result-stat-value">${esc(s.total_rows)}</span></div>
            <div class="result-stat"><span class="result-stat-label">Оброблено</span><span class="result-stat-value">${esc(s.processed_unique)}/${esc(s.unique_products)}</span></div>
            <div class="result-stat"><span class="result-stat-label">Оцінка 5/5</span><span class="result-stat-value">${esc(sd['5'] || 0)}</span></div>
            <div class="result-stat"><span class="result-stat-label">Доповнено атрибутів</span><span class="result-stat-value">${esc(genAttrs)}</span></div>
        `;
    }
    if (d.errors && d.errors.length) {
        show($('#result-errors'));
        $('#result-errors').innerHTML = `<strong>Попередження (${d.errors.length}):</strong><br>` +
            d.errors.slice(0, 20).map(e => `• ${esc(e)}`).join('<br>');
    }
}

function showError(msg) {
    hide($('#step-progress')); show($('#step-result'));
    $('.result-icon').textContent = '❌';
    $('#result-message').textContent = msg;
    hide($('#btn-download'));
}

$('#btn-download').addEventListener('click', () => { window.location.href = `/api/download/${currentJobId}`; });

// "New generation" WITHOUT page reload: settings survive, only file state resets
$('#btn-restart').addEventListener('click', () => {
    currentFileId = null;
    currentJobId = null;
    configFile = null;
    detectedColumns = {};
    missingGeneratable = [];
    previewDone = false;
    pollFailures = 0;
    fileInput.value = '';
    $('#config-filename').textContent = 'Не вибрано';
    dropzone.innerHTML = DROPZONE_DEFAULT;
    hide($('#file-info'));
    hide($('#step-config'));
    hide($('#step-progress'));
    hide($('#step-result'));
    hide($('#preview-panel'));
    hide($('#result-errors'));
    $('#result-errors').innerHTML = '';
    $('#progress-fill').style.width = '0%';
    $('.result-icon').textContent = '✅';
    show($('#step-upload'));
    window.scrollTo({ top: 0, behavior: 'smooth' });
});
