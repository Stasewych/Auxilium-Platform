// ========================================
// AUXILIUM — ОБ'ЄДНАННЯ EXCEL-ФАЙЛІВ
// ========================================

// Довший збіг на початку двох файлів вважається дублікатом файлу, а не шапкою
const MAX_HEADER_ROWS = 5;
const MAX_MANUAL_ROWS = 3; // найбільше значення в перемикачі «Рядків шапки»
const MAX_PAIR_NOTES = 5;
const PREVIEW_DATA_ROWS = 3;
const PREVIEW_MAX_COLS = 30;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

let files = [];            // файли в порядку об'єднання
let headerMode = 'auto';   // 'auto' або кількість рядків шапки
let manualOrder = false;
let nameEdited = false;
let busy = false;
let nextId = 1;
const openPreviews = new Set();

const ICONS = {
    up:     '<polyline points="18 15 12 9 6 15"/>',
    down:   '<polyline points="6 9 12 15 18 9"/>',
    remove: '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>'
};

// ========================================
// HELPERS
// ========================================
function plural(n, one, few, many) {
    const d = n % 10, h = n % 100;
    if (d === 1 && h !== 11) return one;
    if (d >= 2 && d <= 4 && (h < 12 || h > 14)) return few;
    return many;
}

function num(n) {
    return n.toLocaleString('uk-UA');
}

function rowsText(n) {
    return num(n) + ' ' + plural(n, 'рядок', 'рядки', 'рядків');
}

function el(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
}

function setStatus(text, isError) {
    const status = document.getElementById('status');
    status.textContent = text || '';
    status.classList.toggle('error', !!isError);
}

function validFiles() {
    return files.filter(f => !f.error);
}

// ========================================
// READ
// ========================================
function cellKey(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date) return 'd' + v.getTime();
    if (typeof v === 'object') return 'o' + JSON.stringify(v);
    return (typeof v)[0] + v;
}

function cellText(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toLocaleDateString('uk-UA', { timeZone: 'UTC' });
    if (typeof v === 'number') return v.toLocaleString('uk-UA', { maximumFractionDigits: 10 });
    if (typeof v === 'object') {
        if (v.richText) return v.richText.map(part => part.text).join('');
        if (v.text !== undefined) return String(v.text);
        if (v.error) return String(v.error);
        return '';
    }
    return String(v);
}

async function readWorkbook(file) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    const ws = wb.worksheets[0];
    if (!ws) throw new Error('no worksheet');

    const rows = [];
    let colCount = 0;
    let formulas = 0;
    let blankFormulas = 0;

    for (let n = 1; n <= ws.rowCount; n++) {
        const row = ws.findRow(n);
        const cells = [];
        const keys = [];

        if (row) {
            row.eachCell({ includeEmpty: true }, (cell, col) => {
                let v = cell.value;
                if (cell.isMerged && cell.master !== cell) {
                    v = null;
                } else if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
                    // Після зсуву рядків посилання у формулах стали б хибними, тому береться значення
                    formulas++;
                    if (v.result === undefined) blankFormulas++;
                    v = v.result;
                }
                if (v === undefined) v = null;

                const style = cell.style && Object.keys(cell.style).length ? cell.style : null;
                cells[col - 1] = (v === null && !style) ? null : { v: v, s: style };
                keys[col - 1] = cellKey(v);
            });
        }

        while (keys.length && keys[keys.length - 1] === '') keys.pop();
        colCount = Math.max(colCount, keys.length);
        rows.push({ height: row ? row.height : undefined, cells: cells, key: keys.join('\u0001') });
    }

    // Порожній хвіст аркуша в результат не йде, інакше між файлами з'являться розриви
    while (rows.length && rows[rows.length - 1].key === '') rows.pop();

    const columns = [];
    for (let c = 1; c <= ws.columnCount; c++) {
        const col = ws.getColumn(c);
        columns.push({ width: col.width, hidden: col.hidden });
    }

    const merges = [];
    if (ws.hasMerges) {
        (ws.model.merges || []).forEach(ref => {
            const m = /^[A-Z]+\d+:[A-Z]+(\d+)$/.exec(ref);
            if (m && Number(m[1]) <= MAX_HEADER_ROWS) merges.push({ ref: ref, bottom: Number(m[1]) });
        });
    }

    return {
        sheetName: ws.name,
        sheetCount: wb.worksheets.length,
        rows: rows,
        colCount: colCount,
        formulas: formulas,
        blankFormulas: blankFormulas,
        sheet: {
            columns: columns,
            merges: merges,
            properties: ws.properties,
            pageSetup: ws.pageSetup,
            headerFooter: ws.headerFooter,
            views: ws.views
        }
    };
}

async function addFiles(fileList) {
    const incoming = Array.from(fileList);
    if (!incoming.length || busy) return;

    busy = true;
    render();
    const skipped = [];

    for (let i = 0; i < incoming.length; i++) {
        const file = incoming[i];
        if (files.some(f => f.name === file.name && f.size === file.size && f.lastModified === file.lastModified)) {
            skipped.push(file.name);
            continue;
        }

        setStatus('Читаю файл ' + (i + 1) + ' із ' + incoming.length + ': ' + file.name);
        const entry = { id: nextId++, name: file.name, size: file.size, lastModified: file.lastModified, rows: [], h: 0 };

        if (!/\.xlsx$/i.test(file.name)) {
            entry.error = 'Підтримується лише формат .xlsx. Файли .xls і .csv спершу збережіть в Excel як .xlsx.';
        } else {
            try {
                Object.assign(entry, await readWorkbook(file));
                if (!entry.rows.length) entry.error = 'Перший аркуш файлу порожній.';
            } catch (err) {
                console.error(err);
                entry.error = 'Файл не вдалося прочитати. Можливо, він пошкоджений або захищений паролем.';
            }
        }
        files.push(entry);
    }

    if (!manualOrder) sortFiles();
    busy = false;
    setStatus(skipped.length ? 'Уже в списку: ' + skipped.join(', ') : '');
    render();
}

// ========================================
// PLAN
// ========================================
function commonPrefix(a, b) {
    const limit = Math.min(a.rows.length, b.rows.length);
    let n = 0;
    while (n < limit && a.rows[n].key === b.rows[n].key) n++;
    return n;
}

function computePlan() {
    const list = validFiles();
    const notes = [];
    let headerSrc = null;
    let pairs = 0;

    if (headerMode === 'auto') {
        list.forEach(f => { f.h = 0; });
        for (let i = 0; i < list.length; i++) {
            for (let j = i + 1; j < list.length; j++) {
                const a = list[i], b = list[j];
                const p = commonPrefix(a, b);
                if (p > MAX_HEADER_ROWS) {
                    if (++pairs > MAX_PAIR_NOTES) continue;
                    const pair = 'Файли «' + a.name + '» і «' + b.name + '» ';
                    notes.push({
                        warn: true,
                        text: p === a.rows.length && p === b.rows.length
                            ? pair + 'збігаються дослівно. Схоже, той самий файл додано двічі.'
                            : pair + 'починаються з ' + num(p) + ' однакових рядків. Для шапки це забагато, тому ці рядки лишаються в результаті. Перевірте, чи це різні вивантаження.'
                    });
                } else {
                    a.h = Math.max(a.h, p);
                    b.h = Math.max(b.h, p);
                }
            }
        }
        if (pairs > MAX_PAIR_NOTES) {
            notes.push({ warn: true, text: 'Таких пар файлів ще ' + num(pairs - MAX_PAIR_NOTES) + '.' });
        }
        headerSrc = list.find(f => f.h > 0) || null;
        if (headerSrc && headerSrc !== list[0]) {
            notes.push({ text: 'У файлі «' + list[0].name + '» шапки нема, тому її взято з файлу «' + headerSrc.name + '».' });
        }
        // Шапка з різним заголовком над назвами колонок не дає спільного початку, тому про неї лише підказка
        if (!headerSrc && list.length > 1) {
            for (let k = MAX_MANUAL_ROWS - 1; k > 0; k--) {
                const key = list[0].rows[k] && list[0].rows[k].key;
                if (key && list.every(f => f.rows[k] && f.rows[k].key === key)) {
                    notes.push({
                        warn: true,
                        text: 'Рядок ' + (k + 1) + ' однаковий в усіх файлах, а рядки над ним різняться. Якщо це шапка з різним заголовком у кожному файлі, задайте кількість рядків шапки вручну: ' + (k + 1) + '.'
                    });
                    break;
                }
            }
        }
    } else {
        const n = Number(headerMode);
        list.forEach(f => { f.h = Math.min(n, f.rows.length); });
        headerSrc = n > 0 && list.length ? list[0] : null;
    }

    const other = list.find(f => f.colCount !== list[0].colCount);
    if (other) {
        notes.push({
            warn: true,
            text: 'Файли мають різну кількість колонок: ' + list[0].colCount + ' у «' + list[0].name + '» і ' + other.colCount + ' у «' + other.name + '». Перевірте, чи це вивантаження одного формату.'
        });
    }

    let dataRows = 0;
    let repeats = 0;
    let formulas = 0;
    let blankFormulas = 0;
    const seen = new Set();
    list.forEach(f => {
        formulas += f.formulas;
        blankFormulas += f.blankFormulas;
        for (let i = f.h; i < f.rows.length; i++) {
            const key = f.rows[i].key;
            dataRows++;
            if (key === '') continue;
            if (seen.has(key)) repeats++;
            else seen.add(key);
        }
    });

    if (repeats) {
        notes.push({
            text: num(repeats) + ' ' + plural(repeats, 'рядок даних дослівно повторює інший', 'рядки даних дослівно повторюють інші', 'рядків даних дослівно повторюють інші') + '. У результаті вони лишаються.'
        });
    }
    if (formulas) {
        notes.push({
            text: 'Формули записуються в результат як значення (' + num(formulas) + ' ' + plural(formulas, 'клітинка', 'клітинки', 'клітинок') + ').'
        });
    }
    if (blankFormulas) {
        notes.push({
            warn: true,
            text: num(blankFormulas) + ' ' + plural(blankFormulas, 'формула не має', 'формули не мають', 'формул не мають') + ' збереженого значення, тому в результаті ці клітинки будуть порожні. Щоб значення з\'явилися, відкрийте вихідний файл в Excel і збережіть його ще раз.'
        });
    }

    return {
        files: list,
        headerSrc: headerSrc,
        headerRows: headerSrc ? headerSrc.h : 0,
        dataRows: dataRows,
        notes: notes
    };
}

// ========================================
// BUILD
// ========================================
async function buildFile(plan) {
    const base = plan.files[0];
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(base.sheetName, {
        properties: base.sheet.properties,
        pageSetup: base.sheet.pageSetup,
        headerFooter: base.sheet.headerFooter,
        views: base.sheet.views
    });

    base.sheet.columns.forEach((c, i) => {
        if (c.width === undefined && !c.hidden) return;
        const col = ws.getColumn(i + 1);
        if (c.width !== undefined) col.width = c.width;
        if (c.hidden) col.hidden = true;
    });

    let n = 0;
    const put = (src) => {
        n++;
        if (src.height === undefined && !src.cells.length) return;
        const row = ws.getRow(n);
        if (src.height !== undefined) row.height = src.height;
        for (let i = 0; i < src.cells.length; i++) {
            const c = src.cells[i];
            if (!c) continue;
            const cell = row.getCell(i + 1);
            if (c.v !== null) cell.value = c.v;
            if (c.s) cell.style = c.s;
        }
    };

    if (plan.headerSrc) {
        plan.headerSrc.rows.slice(0, plan.headerRows).forEach(put);
        plan.headerSrc.sheet.merges
            .filter(m => m.bottom <= plan.headerRows)
            .forEach(m => ws.mergeCells(m.ref));
    }
    plan.files.forEach(f => {
        for (let i = f.h; i < f.rows.length; i++) put(f.rows[i]);
    });

    return wb.xlsx.writeBuffer();
}

function outputName() {
    const typed = document.getElementById('outName').value
        .replace(/\.xlsx$/i, '')
        .replace(/[\\/:*?"<>|]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    return (typed || 'Об\'єднаний файл') + '.xlsx';
}

function defaultName() {
    const names = validFiles().map(f => f.name.replace(/\.xlsx$/i, ''));
    if (!names.length) return '';

    let prefix = names[0];
    names.forEach(name => {
        let i = 0;
        while (i < prefix.length && i < name.length && prefix[i] === name[i]) i++;
        prefix = prefix.slice(0, i);
    });
    // Спільний початок не має обриватися посеред слова чи числа
    if (/\S$/.test(prefix) && names.some(name => name.length > prefix.length && /\S/.test(name[prefix.length]))) {
        prefix = prefix.replace(/\S+$/, '');
    }
    prefix = prefix.replace(/[\s_.,(-]+$/, '');

    return prefix.length >= 3 ? prefix + ' (об\'єднано)' : 'Об\'єднаний файл';
}

async function mergeAndDownload() {
    const plan = computePlan();
    if (busy || plan.files.length < 2) return;

    busy = true;
    render();
    setStatus('Збираю файл…');
    // Пауза дає браузеру показати статус до важкої синхронної роботи
    await new Promise(resolve => setTimeout(resolve, 30));

    try {
        const buffer = await buildFile(plan);
        const name = outputName();
        const link = document.createElement('a');
        link.href = URL.createObjectURL(new Blob([buffer], { type: XLSX_MIME }));
        link.download = name;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(link.href), 1000);
        setStatus('Файл «' + name + '» збережено: ' + rowsText(plan.headerRows + plan.dataRows) + '.');
    } catch (err) {
        console.error(err);
        setStatus('Об\'єднати файли не вдалося: ' + (err && err.message ? err.message : 'невідома помилка') + '.', true);
    }

    busy = false;
    render();
}

// ========================================
// LIST ACTIONS
// ========================================
function sortFiles() {
    files.sort((a, b) => a.name.localeCompare(b.name, 'uk', { numeric: true }));
}

function sortByName() {
    manualOrder = false;
    sortFiles();
    setStatus('');
    render();
}

function moveFile(id, step) {
    const from = files.findIndex(f => f.id === id);
    const to = from + step;
    if (from < 0 || to < 0 || to >= files.length) return;
    files.splice(to, 0, files.splice(from, 1)[0]);
    manualOrder = true;
    setStatus('');
    render();
}

function removeFile(id) {
    files = files.filter(f => f.id !== id);
    openPreviews.delete(id);
    if (!files.length) resetList();
    setStatus('');
    render();
}

function resetList() {
    files = [];
    manualOrder = false;
    nameEdited = false;
    openPreviews.clear();
}

function clearFiles() {
    resetList();
    setStatus('');
    render();
}

function setHeaderMode(btn) {
    headerMode = btn.dataset.mode;
    document.querySelectorAll('.seg').forEach(seg => seg.classList.toggle('active', seg === btn));
    setStatus('');
    render();
}

function pickFiles() {
    document.getElementById('fileInput').click();
}

// ========================================
// RENDER
// ========================================
function iconButton(icon, title, disabled, onClick) {
    const btn = el('button', 'icon-btn');
    btn.type = 'button';
    btn.title = title;
    btn.setAttribute('aria-label', title);
    btn.disabled = disabled;
    btn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + ICONS[icon] + '</svg>';
    btn.onclick = onClick;
    return btn;
}

function fileMeta(file, plan) {
    const parts = [rowsText(file.rows.length)];
    if (file.h > 0) {
        parts.push('шапка ' + rowsText(file.h) + (file === plan.headerSrc ? ', лишається' : ', прибирається'));
    } else if (headerMode === 'auto') {
        parts.push('шапки не знайдено');
    }
    if (file.sheetCount > 1) parts.push('взято перший аркуш із ' + file.sheetCount);
    return parts.join(' · ');
}

function fillPreview(box, file, plan) {
    const table = el('table', 'preview-table');
    const count = Math.min(file.rows.length, file.h + PREVIEW_DATA_ROWS);
    const cols = Math.min(file.colCount, PREVIEW_MAX_COLS);

    for (let r = 0; r < count; r++) {
        const tr = el('tr');
        let tag = 'дані';
        if (r < file.h) {
            const kept = file === plan.headerSrc;
            tag = kept ? 'шапка, лишається' : 'шапка, прибирається';
            tr.className = kept ? 'is-header' : 'is-header is-dropped';
        }
        tr.appendChild(el('td', 'row-tag', tag));
        for (let c = 0; c < cols; c++) {
            const cell = file.rows[r].cells[c];
            const text = cell ? cellText(cell.v) : '';
            const td = el('td', '', text);
            td.title = text;
            tr.appendChild(td);
        }
        table.appendChild(tr);
    }

    box.replaceChildren(table);
}

function renderFile(file, index, plan) {
    const li = el('li', file.error ? 'file failed' : 'file');
    const row = el('div', 'file-row');
    const main = el('div', 'file-main');

    row.appendChild(el('span', 'file-num', String(index + 1)));
    main.appendChild(el('span', 'file-name', file.name));
    main.appendChild(el('span', 'file-meta', file.error || fileMeta(file, plan)));
    row.appendChild(main);

    const actions = el('div', 'file-actions');
    actions.appendChild(iconButton('up', 'Вище', busy || index === 0, () => moveFile(file.id, -1)));
    actions.appendChild(iconButton('down', 'Нижче', busy || index === files.length - 1, () => moveFile(file.id, 1)));
    actions.appendChild(iconButton('remove', 'Прибрати зі списку', busy, () => removeFile(file.id)));
    row.appendChild(actions);
    li.appendChild(row);

    if (!file.error) {
        const details = el('details', 'file-preview');
        const box = el('div', 'preview-scroll');
        details.appendChild(el('summary', '', 'Початок файлу'));
        details.appendChild(box);
        details.open = openPreviews.has(file.id);
        if (details.open) fillPreview(box, file, plan);
        details.addEventListener('toggle', () => {
            if (details.open) {
                openPreviews.add(file.id);
                fillPreview(box, file, plan);
            } else {
                openPreviews.delete(file.id);
            }
        });
        li.appendChild(details);
    }

    return li;
}

function headerHint() {
    if (headerMode === 'auto') {
        return 'Авто порівнює початок файлів: рядки, що повторюються на тому самому місці в кількох файлах, вважаються шапкою.';
    }
    if (headerMode === '0') return 'Рядки з усіх файлів ідуть у результат без змін.';
    return 'У першому файлі шапка лишається, а з кожного наступного прибирається перших рядків: ' + headerMode + '.';
}

function summaryText(plan) {
    const count = plan.files.length;
    if (count < 2) return 'Для об\'єднання потрібно щонайменше два файли.';

    const from = ' з ' + count + ' ' + plural(count, 'файлу', 'файлів', 'файлів') + '.';
    if (plan.headerRows > 0) {
        return 'У результаті буде ' + rowsText(plan.headerRows + plan.dataRows) + ': ' + rowsText(plan.headerRows) + ' шапки і ' + rowsText(plan.dataRows) + ' даних' + from;
    }
    return 'У результаті буде ' + rowsText(plan.dataRows) + ' даних' + from;
}

function render() {
    const plan = computePlan();

    document.getElementById('panel').hidden = files.length === 0;
    document.getElementById('pickBtn').disabled = busy;
    document.getElementById('sortBtn').disabled = busy || !manualOrder;
    document.getElementById('mergeBtn').disabled = busy || plan.files.length < 2;
    document.getElementById('headerHint').textContent = headerHint();
    document.getElementById('summary').textContent = summaryText(plan);

    document.getElementById('fileList').replaceChildren(
        ...files.map((file, index) => renderFile(file, index, plan))
    );
    document.getElementById('notes').replaceChildren(
        ...plan.notes.map(note => el('li', note.warn ? 'note warn' : 'note', note.text))
    );

    const outName = document.getElementById('outName');
    if (!nameEdited) outName.value = defaultName();
}

// ========================================
// EVENTS
// ========================================
(function init() {
    const dropzone = document.getElementById('dropzone');
    const input = document.getElementById('fileInput');

    input.addEventListener('change', () => {
        addFiles(input.files);
        input.value = '';
    });

    ['dragenter', 'dragover'].forEach(type => {
        dropzone.addEventListener(type, event => {
            event.preventDefault();
            dropzone.classList.add('over');
        });
    });
    ['dragleave', 'drop'].forEach(type => {
        dropzone.addEventListener(type, () => dropzone.classList.remove('over'));
    });
    dropzone.addEventListener('drop', event => {
        event.preventDefault();
        addFiles(event.dataTransfer.files);
    });

    // Файл, кинутий повз зону, не має відкритися замість інструмента
    ['dragover', 'drop'].forEach(type => {
        window.addEventListener(type, event => event.preventDefault());
    });

    render();
})();
