// FeedGen v2 Frontend Logic

const $ = (sel) => document.querySelector(sel);
const show = (el) => el.classList.remove('hidden');
const hide = (el) => el.classList.add('hidden');

let currentFileId = null;
let currentJobId = null;
let configFile = null;
let detectedColumns = {};

// =============================================================================
// STEP 1: Upload
// =============================================================================

const dropzone = $('#dropzone');
const fileInput = $('#file-input');

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});

async function handleFile(file) {
    if (!file.name.match(/\.xlsx?$/i)) {
        alert('Підтримуються лише .xlsx файли');
        return;
    }

    dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">⏳</div><p>Аналізую файл...</p></div>`;

    const formData = new FormData();
    formData.append('file', file);

    try {
        const res = await fetch('/api/analyze', { method: 'POST', body: formData });
        if (!res.ok) throw new Error((await res.json()).detail);
        const data = await res.json();

        currentFileId = data.file_id;
        detectedColumns = data.detected_columns || {};

        dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">✅</div><p>${data.filename}</p></div>`;

        const infoEl = $('#file-info');
        show(infoEl);
        infoEl.querySelector('.file-stats').innerHTML = `
            <div class="stat"><span class="stat-label">Рядків</span><span class="stat-value">${data.total_rows}</span></div>
            <div class="stat"><span class="stat-label">Унікальних товарів</span><span class="stat-value">${data.unique_titles}</span></div>
        `;

        // Show detected columns
        const requiredCols = ['title', 'description', 'product_type', 'color', 'material', 'gender'];
        const colsHtml = requiredCols.map(col => {
            const found = detectedColumns[col] !== undefined;
            return `<span class="col-tag ${found ? 'found' : 'missing'}">${col}${found ? ' ✓' : ' ✗'}</span>`;
        }).join('');
        $('#detected-cols').innerHTML = `<span style="font-size:12px;color:var(--text-dim);width:100%;margin-bottom:4px;">Розпізнані колонки:</span>${colsHtml}`;

        show($('#step-config'));

    } catch (err) {
        dropzone.innerHTML = `<div class="upload-content"><div class="upload-icon">❌</div><p>${err.message}</p><p class="upload-hint">Спробуй ще раз</p></div>`;
    }
}

// =============================================================================
// STEP 2: Config
// =============================================================================

$('#config-input').addEventListener('change', (e) => {
    if (e.target.files.length) {
        configFile = e.target.files[0];
        $('#config-filename').textContent = configFile.name;
    }
});

// Update cache hint based on model
$('#model-select').addEventListener('change', (e) => {
    $('#cache-hint').textContent = '✓ Prompt caching увімкнено — економія до 90%';
});

$('#btn-generate').addEventListener('click', async () => {
    const btn = $('#btn-generate');
    btn.disabled = true;
    btn.innerHTML = '<span class="btn-icon">⏳</span> Запускаю...';

    const formData = new FormData();
    formData.append('file_id', currentFileId);
    formData.append('model', $('#model-select').value);
    formData.append('language', $('#lang-select').value);
    formData.append('column_map', JSON.stringify(detectedColumns));
    if (configFile) formData.append('config', configFile);

    try {
        const res = await fetch('/api/generate', { method: 'POST', body: formData });
        if (!res.ok) throw new Error((await res.json()).detail);
        const data = await res.json();

        currentJobId = data.job_id;
        hide($('#step-upload'));
        hide($('#step-config'));
        show($('#step-progress'));
        pollStatus();

    } catch (err) {
        alert(`Помилка: ${err.message}`);
        btn.disabled = false;
        btn.innerHTML = '<span class="btn-icon">⚡</span> Запустити генерацію';
    }
});

// =============================================================================
// STEP 3: Progress
// =============================================================================

async function pollStatus() {
    try {
        const res = await fetch(`/api/status/${currentJobId}`);
        const data = await res.json();

        const pct = data.total > 0 ? Math.round((data.progress / data.total) * 100) : 0;
        $('#progress-fill').style.width = `${pct}%`;
        $('#progress-pct').textContent = `${pct}%`;
        $('#progress-text').textContent = data.message;

        if (data.status === 'done') { showResult(data); return; }
        if (data.status === 'error') { showError(data.message); return; }
        setTimeout(pollStatus, 1500);
    } catch (err) {
        setTimeout(pollStatus, 3000);
    }
}

// =============================================================================
// STEP 4: Result
// =============================================================================

function showResult(data) {
    hide($('#step-progress'));
    show($('#step-result'));
    $('#result-message').textContent = data.message;

    // Show stats
    if (data.stats) {
        const s = data.stats;
        const scoreDist = s.score_distribution || {};
        const score5 = scoreDist['5'] || 0;
        $('#result-stats').innerHTML = `
            <div class="result-stat"><span class="result-stat-label">Всього рядків</span><span class="result-stat-value">${s.total_rows}</span></div>
            <div class="result-stat"><span class="result-stat-label">Унікальних</span><span class="result-stat-value">${s.unique_generated}/${s.unique_products}</span></div>
            <div class="result-stat"><span class="result-stat-label">Оцінка 5/5</span><span class="result-stat-value">${score5}</span></div>
            <div class="result-stat"><span class="result-stat-label">Fallback</span><span class="result-stat-value">${s.fallback_count}</span></div>
        `;
    }

    if (data.errors && data.errors.length > 0) {
        const errEl = $('#result-errors');
        show(errEl);
        errEl.innerHTML = `<strong>Попередження (${data.errors.length}):</strong><br>` +
            data.errors.slice(0, 20).map(e => `• ${e}`).join('<br>');
    }
}

function showError(message) {
    hide($('#step-progress'));
    show($('#step-result'));
    $('.result-icon').textContent = '❌';
    $('#result-message').textContent = message;
    hide($('#btn-download'));
}

$('#btn-download').addEventListener('click', () => {
    window.location.href = `/api/download/${currentJobId}`;
});

$('#btn-restart').addEventListener('click', () => location.reload());
