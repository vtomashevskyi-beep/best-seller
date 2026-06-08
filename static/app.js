// FeedGen v3 Frontend

const $ = (s) => document.querySelector(s);
const show = (el) => el.classList.remove('hidden');
const hide = (el) => el.classList.add('hidden');

let currentFileId = null;
let currentJobId = null;
let configFile = null;
let detectedColumns = {};
let missingGeneratable = [];
let schema = { structurable: [], generatable: [], default_title_order: [] };
let attrOrder = [];        // [{key, label, enabled}]
let selectedGenAttrs = new Set();

// Load schema on startup
(async () => {
    try {
        const res = await fetch('/api/schema');
        schema = await res.json();
    } catch (e) { console.error('Schema load failed', e); }
})();

// =============================================================================
// STEP 1: Upload
// =============================================================================

const dropzone = $('#dropzone');
const fileInput = $('#file-input');

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', (e) => {
    e.preventDefault(); dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => { if (e.target.files.length) handleFile(e.target.files[0]); });

async function handleFile(file) {
    if (!file.name.match(/\.(xlsx?|csv|xml)$/i)) {
        alert('Підтримуються XLSX, CSV, XML');
        return;
    }
    dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">⏳</div><p>Аналізую...</p></div>`;

    const fd = new FormData();
    fd.append('file', file);
    try {
        const res = await fetch('/api/analyze', { method: 'POST', body: fd });
        if (!res.ok) throw new Error((await res.json()).detail);
        const d = await res.json();

        currentFileId = d.file_id;
        detectedColumns = d.detected_columns || {};
        missingGeneratable = d.missing_generatable || [];

        dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">✅</div><p>${d.filename}</p><p class="upload-hint">${d.format.toUpperCase()}</p></div>`;

        show($('#file-info'));
        $('#file-info').querySelector('.file-stats').innerHTML = `
            <div class="stat"><span class="stat-label">Рядків</span><span class="stat-value">${d.total_rows}</span></div>
            <div class="stat"><span class="stat-label">Унікальних товарів</span><span class="stat-value">${d.unique_titles}</span></div>
        `;

        // Detected columns
        const present = d.present_attributes || [];
        const tagsHtml = present.map(c => `<span class="col-tag found">${c} ✓</span>`).join('');
        $('#detected-cols').innerHTML = `<span style="font-size:12px;color:var(--text-dim);width:100%;margin-bottom:4px;">Знайдено у фіді:</span>${tagsHtml}`;

        buildAttrOrder();
        buildGenAttrs();
        show($('#step-config'));
    } catch (err) {
        dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">❌</div><p>${err.message}</p><p class="upload-hint">Спробуй ще раз</p></div>`;
    }
}

// =============================================================================
// STEP 2: Attribute ordering (drag-and-drop)
// =============================================================================

function buildAttrOrder() {
    // Start from default order, include only attrs that make sense
    const order = schema.default_title_order.length ? schema.default_title_order : ['brand', 'product_type', 'color'];
    const labelMap = {};
    schema.structurable.forEach(s => labelMap[s.key] = s.label);

    attrOrder = order.map(key => ({
        key,
        label: labelMap[key] || key,
        enabled: detectedColumns[key] !== undefined || key === 'brand' || key === 'product_type',
    }));

    // Add any structurable attrs not in default order (disabled by default)
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
            <span class="attr-name">${attr.label}</span>
            <span class="attr-toggle ${attr.enabled ? 'on' : ''}" data-key="${attr.key}"></span>
        `;
        list.appendChild(li);
    });

    // Toggle handlers
    list.querySelectorAll('.attr-toggle').forEach(t => {
        t.addEventListener('click', (e) => {
            e.stopPropagation();
            const key = t.dataset.key;
            const attr = attrOrder.find(a => a.key === key);
            attr.enabled = !attr.enabled;
            renderAttrOrder();
            updateTitlePreview();
        });
    });

    // Drag handlers
    let dragEl = null;
    list.querySelectorAll('.attr-order-item').forEach(item => {
        item.addEventListener('dragstart', () => { dragEl = item; item.classList.add('dragging'); });
        item.addEventListener('dragend', () => {
            item.classList.remove('dragging');
            // Rebuild attrOrder from DOM
            const newOrder = [...list.querySelectorAll('.attr-order-item')].map(el => {
                const key = el.dataset.key;
                return attrOrder.find(a => a.key === key);
            });
            attrOrder = newOrder;
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
        ? enabled.map(l => `<strong>{${l}}</strong>`).join(' ')
        : '<em>немає увімкнених атрибутів</em>';
    $('#title-preview').innerHTML = `Приклад: ${preview}`;
}

// =============================================================================
// STEP 2: Generatable attributes
// =============================================================================

function buildGenAttrs() {
    const container = $('#gen-attrs-list');
    selectedGenAttrs.clear();

    // Only show generatable attrs that are MISSING from the feed
    const labelMap = {};
    schema.generatable.forEach(g => labelMap[g.key] = g.label);
    const available = missingGeneratable.filter(k => labelMap[k]);

    if (available.length === 0) {
        container.innerHTML = '<span class="gen-attrs-empty">Усі атрибути вже присутні у фіді — доповнювати нічого.</span>';
        return;
    }

    container.innerHTML = '';
    available.forEach(key => {
        const chip = document.createElement('div');
        chip.className = 'gen-attr-chip';
        chip.dataset.key = key;
        chip.innerHTML = `<span class="check">＋</span> ${labelMap[key]}`;
        chip.addEventListener('click', () => {
            if (selectedGenAttrs.has(key)) {
                selectedGenAttrs.delete(key);
                chip.classList.remove('selected');
                chip.querySelector('.check').textContent = '＋';
            } else {
                selectedGenAttrs.add(key);
                chip.classList.add('selected');
                chip.querySelector('.check').textContent = '✓';
            }
        });
        container.appendChild(chip);
    });
}

// =============================================================================
// Config file
// =============================================================================

$('#config-input').addEventListener('change', (e) => {
    if (e.target.files.length) {
        configFile = e.target.files[0];
        $('#config-filename').textContent = configFile.name;
    }
});

// =============================================================================
// Generate
// =============================================================================

$('#btn-generate').addEventListener('click', async () => {
    const btn = $('#btn-generate');
    btn.disabled = true;
    btn.innerHTML = '<span class="btn-icon">⏳</span> Запускаю...';

    const titleOrder = attrOrder.filter(a => a.enabled).map(a => a.key);

    const fd = new FormData();
    fd.append('file_id', currentFileId);
    fd.append('model', $('#model-select').value);
    fd.append('language', $('#lang-select').value);
    fd.append('column_map', JSON.stringify(detectedColumns));
    fd.append('title_order', JSON.stringify(titleOrder));
    fd.append('generate_attributes', JSON.stringify([...selectedGenAttrs]));
    fd.append('output_format', $('#format-select').value);
    if (configFile) fd.append('config', configFile);

    try {
        const res = await fetch('/api/generate', { method: 'POST', body: fd });
        if (!res.ok) throw new Error((await res.json()).detail);
        const d = await res.json();
        currentJobId = d.job_id;
        hide($('#step-upload')); hide($('#step-config')); show($('#step-progress'));
        pollStatus();
    } catch (err) {
        alert(`Помилка: ${err.message}`);
        btn.disabled = false;
        btn.innerHTML = '<span class="btn-icon">⚡</span> Запустити генерацію';
    }
});

// =============================================================================
// Progress + Result
// =============================================================================

async function pollStatus() {
    try {
        const res = await fetch(`/api/status/${currentJobId}`);
        const d = await res.json();
        const pct = d.total > 0 ? Math.round((d.progress / d.total) * 100) : 0;
        $('#progress-fill').style.width = `${pct}%`;
        $('#progress-pct').textContent = `${pct}%`;
        $('#progress-text').textContent = d.message;
        if (d.status === 'done') { showResult(d); return; }
        if (d.status === 'error') { showError(d.message); return; }
        setTimeout(pollStatus, 1500);
    } catch (err) { setTimeout(pollStatus, 3000); }
}

function showResult(d) {
    hide($('#step-progress')); show($('#step-result'));
    $('#result-message').textContent = d.message;
    if (d.stats) {
        const s = d.stats;
        const sd = s.score_distribution || {};
        const genAttrs = (s.generated_attributes || []).length;
        $('#result-stats').innerHTML = `
            <div class="result-stat"><span class="result-stat-label">Всього рядків</span><span class="result-stat-value">${s.total_rows}</span></div>
            <div class="result-stat"><span class="result-stat-label">Унікальних</span><span class="result-stat-value">${s.unique_generated}/${s.unique_products}</span></div>
            <div class="result-stat"><span class="result-stat-label">Оцінка 5/5</span><span class="result-stat-value">${sd['5'] || 0}</span></div>
            <div class="result-stat"><span class="result-stat-label">Доповнено атрибутів</span><span class="result-stat-value">${genAttrs}</span></div>
        `;
    }
    if (d.errors && d.errors.length) {
        show($('#result-errors'));
        $('#result-errors').innerHTML = `<strong>Попередження (${d.errors.length}):</strong><br>` +
            d.errors.slice(0, 20).map(e => `• ${e}`).join('<br>');
    }
}

function showError(msg) {
    hide($('#step-progress')); show($('#step-result'));
    $('.result-icon').textContent = '❌';
    $('#result-message').textContent = msg;
    hide($('#btn-download'));
}

$('#btn-download').addEventListener('click', () => { window.location.href = `/api/download/${currentJobId}`; });
$('#btn-restart').addEventListener('click', () => location.reload());
