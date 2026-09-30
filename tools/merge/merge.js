// ========================================
// AUXILIUM — ЗВЕДЕННЯ РЕЄСТРІВ
// ========================================

// Довший збіг на початку двох файлів вважається дублікатом файлу, а не шапкою
const MAX_HEADER_ROWS = 5;
const MAX_MANUAL_ROWS = 3; // найбільше значення в перемикачі «Рядків шапки»
const MAX_PAIR_NOTES = 5;
const PREVIEW_HEAD_ROWS = 8;
const PREVIEW_MORE_ROWS = 20;
const PREVIEW_TAIL_ROWS = 3;
const PREVIEW_MAX_COLS = 30;
const RESULT_TOP_ROWS = 3;
const RESULT_MAX_JOINS = 6;
const SOURCE_COL_TITLE = 'Файл';
const SOURCE_COL_WIDTH = 36;
const BUILD_CHUNK = 2000;
const RENDER_EVERY_MS = 150;
const ENTER_MS = 320; // тривалість появи рядка в списку, та сама, що в merge.css
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

let files = [];            // файли в порядку зведення
let headerMode = 'auto';   // 'auto' або кількість рядків шапки
let dedupe = false;
let sourceCol = false;
let manualOrder = false;
let nameEdited = false;
let busy = false;          // іде читання файлів або запис результату
let building = false;      // іде саме запис результату
let nextId = 1;
let lastPlan = null;
const openPreviews = new Set();

const ICONS = {
    up:     '<polyline points="18 15 12 9 6 15"/>',
    down:   '<polyline points="6 9 12 15 18 9"/>',
    remove: '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>',
    done:   '<circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.8 2.8L16 9.5"/>'
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

function svg(icon) {
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + ICONS[icon] + '</svg>';
}

// Пауза дає браузеру перемалювати сторінку між важкими синхронними кроками
function tick() {
    return new Promise(resolve => setTimeout(resolve, 0));
}

function setStatus(text, kind) {
    const status = document.getElementById('status');
    status.className = kind ? 'status ' + kind : 'status';
    status.replaceChildren();
    if (!text) return;
    if (kind === 'done') {
        const icon = el('span', 'status-icon');
        icon.innerHTML = svg('done');
        status.appendChild(icon);
    }
    status.appendChild(el('span', '', text));
}

// total = 0 означає крок без відомої тривалості
function setProgress(label, done, total) {
    const box = document.getElementById('progress');
    box.hidden = label === null;
    if (label === null) return;
    box.classList.toggle('indeterminate', !total);
    document.getElementById('progressLabel').textContent = label;
    document.getElementById('progressCount').textContent = total ? num(done) + ' / ' + num(total) : '';
    document.getElementById('progressBar').style.transform = 'scaleX(' + (total ? Math.max(0.02, done / total) : 1) + ')';
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

function parseSheet(ws) {
    const rows = [];
    let colCount = 0;
    let width = 0;
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

        while (cells.length && !cells[cells.length - 1]) cells.pop();
        while (keys.length && keys[keys.length - 1] === '') keys.pop();
        colCount = Math.max(colCount, keys.length);
        width = Math.max(width, cells.length);
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
            const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
            if (m && Number(m[4]) <= MAX_HEADER_ROWS) {
                merges.push({ left: m[1], top: Number(m[2]), right: m[3], bottom: Number(m[4]) });
            }
        });
    }

    return {
        sheetName: ws.name,
        hidden: ws.state === 'hidden' || ws.state === 'veryHidden',
        rows: rows,
        colCount: colCount,
        width: width,
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

async function readWorkbook(file) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    if (!wb.worksheets.length) throw new Error('no worksheet');
    return wb.worksheets.map(parseSheet);
}

function selectSheet(entry, index) {
    Object.assign(entry, entry.sheets[index]);
    entry.sheetIdx = index;
    entry.overrides = new Map();      // номер рядка → true, якщо рядок прибрано вручну, і false, якщо лишено
    entry.headRows = PREVIEW_HEAD_ROWS;
}

// ========================================
// SOURCES: файли, папки, архіви
// ========================================
function isXlsx(name) {
    return /\.xlsx$/i.test(name);
}

// Тимчасові файли Excel (~$назва.xlsx) і службові файли macOS
function isJunk(name) {
    return /^(~\$|\.)/.test(name);
}

// Назва без позначки UTF-8 приходить з архіву побайтно: macOS пише такі назви в UTF-8, Windows у кодуванні OEM
function zipName(raw) {
    if (!/[\u0080-ÿ]/.test(raw) || /[Ā-￿]/.test(raw)) return raw;
    const bytes = Uint8Array.from(raw, ch => ch.charCodeAt(0));
    try {
        return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (err) {
        return new TextDecoder('ibm866').decode(bytes);
    }
}

async function unzip(item) {
    let other = 0;
    const found = fflate.unzipSync(new Uint8Array(await item.file.arrayBuffer()), {
        filter: info => {
            const path = info.name.replace(/\\/g, '/');
            const name = path.split('/').pop();
            if (!name || isJunk(name) || path.includes('__MACOSX/')) return false;
            if (isXlsx(name)) return true;
            other++;
            return false;
        }
    });

    const items = Object.keys(found).map(raw => {
        const parts = zipName(raw).normalize('NFC').replace(/\\/g, '/').split('/');
        const name = parts.pop();
        return {
            file: new File([found[raw]], name, { lastModified: item.file.lastModified }),
            name: name,
            origin: [item.name].concat(parts).join(' / '),
            bulk: true
        };
    });
    return { items: items, other: other };
}

function makeEntry(item) {
    return {
        id: nextId++,
        name: item.name,
        origin: item.origin,
        key: [item.origin, item.name, item.file.size, item.file.lastModified].join('|'),
        sheets: [],
        rows: [],
        h: 0,
        overrides: new Map(),
        repeats: new Set(),
        trimmed: new Set()
    };
}

function addFailed(item, message) {
    const entry = makeEntry(item);
    entry.error = message;
    files.push(entry);
}

// sources: [{ file, dir }], де dir задано для файлів із папки
async function addSources(sources) {
    if (!sources.length || busy) return;

    busy = true;
    setStatus('');
    setProgress('Готую файли…', 0, 0);
    render();
    await tick();

    const queue = [];
    const listed = [];
    let other = 0;
    let added = 0;

    for (const source of sources) {
        const name = source.file.name.normalize('NFC');
        const item = { file: source.file, name: name, origin: (source.dir || '').normalize('NFC'), bulk: !!source.dir };

        if (isJunk(name)) continue;
        if (/\.zip$/i.test(name)) {
            try {
                const unpacked = await unzip(item);
                if (unpacked.items.length) {
                    other += unpacked.other;
                    queue.push(...unpacked.items);
                } else {
                    addFailed(item, 'В архіві нема файлів .xlsx.');
                }
            } catch (err) {
                console.error(err);
                addFailed(item, 'Архів не вдалося розпакувати. Можливо, він пошкоджений або захищений паролем.');
            }
        } else if (isXlsx(name)) {
            queue.push(item);
        } else if (item.bulk) {
            other++;
        } else if (/\.(rar|7z)$/i.test(name)) {
            addFailed(item, 'З архівів підтримується лише ZIP. Архів RAR або 7z спершу розпакуйте й додайте папку.');
        } else {
            addFailed(item, 'Підтримується лише формат .xlsx. Файли .xls і .csv спершу збережіть в Excel як .xlsx.');
        }
    }

    let painted = 0;
    for (let i = 0; i < queue.length; i++) {
        const item = queue[i];
        const entry = makeEntry(item);
        if (files.some(f => f.key === entry.key)) {
            listed.push(item.name);
            continue;
        }

        setProgress('Читаю: ' + item.name, i, queue.length);
        await tick();

        try {
            entry.sheets = await readWorkbook(item.file);
            const first = entry.sheets.findIndex(sheet => sheet.rows.length);
            if (first >= 0) selectSheet(entry, first);
            else entry.error = entry.sheets.length > 1 ? 'Усі аркуші файлу порожні.' : 'Аркуш файлу порожній.';
        } catch (err) {
            console.error(err);
            entry.error = 'Файл не вдалося прочитати. Можливо, він пошкоджений або захищений паролем.';
        }

        files.push(entry);
        if (!entry.error) added++;
        if (!manualOrder) sortFiles();

        if (performance.now() - painted > RENDER_EVERY_MS) {
            render();
            painted = performance.now();
        }
    }

    if (!manualOrder) sortFiles();
    busy = false;
    setProgress(null);

    const report = [];
    if (added) report.push('Додано ' + num(added) + ' ' + plural(added, 'файл', 'файли', 'файлів') + '.');
    if (other) report.push('Пропущено файлів іншого формату: ' + num(other) + '.');
    if (listed.length) {
        report.push('Уже в списку: ' + listed.slice(0, 3).join(', ') + (listed.length > 3 ? ' і ще ' + num(listed.length - 3) : '') + '.');
    }
    if (!files.length) report.unshift('Файлів .xlsx не знайдено.');
    setStatus(report.join(' '), added || !report.length ? '' : 'error');
    render();
}

// Обхід папки, яку перетягнули в зону
async function walkEntry(entry, dir, out) {
    if (entry.isFile) {
        const file = await new Promise((resolve, reject) => entry.file(resolve, reject));
        out.push({ file: file, dir: dir });
    } else if (entry.isDirectory) {
        const reader = entry.createReader();
        const path = dir ? dir + ' / ' + entry.name : entry.name;
        let batch;
        // readEntries віддає вміст порціями, тому читається до порожньої відповіді
        do {
            batch = await new Promise((resolve, reject) => reader.readEntries(resolve, reject));
            for (const child of batch) await walkEntry(child, path, out);
        } while (batch.length);
    }
}

async function addDropped(transfer) {
    const entries = [];
    const sources = [];

    // Елементи перетягування доступні лише під час самої події, тому збираються до першого await
    Array.from(transfer.items || []).forEach(item => {
        if (item.kind !== 'file') return;
        const entry = item.webkitGetAsEntry ? item.webkitGetAsEntry() : null;
        if (entry) {
            entries.push(entry);
        } else {
            const file = item.getAsFile();
            if (file) sources.push({ file: file, dir: '' });
        }
    });
    if (!entries.length && !sources.length) {
        Array.from(transfer.files || []).forEach(file => sources.push({ file: file, dir: '' }));
    }

    try {
        for (const entry of entries) await walkEntry(entry, '', sources);
    } catch (err) {
        console.error(err);
        setStatus('Папку не вдалося прочитати. Спробуйте кнопку «Вибрати папку».', 'error');
        return;
    }
    addSources(sources);
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

    // out — рядки результату по порядку: спершу шапка, далі дані кожного файлу
    const dropped = (f, i, byDefault) => f.overrides.has(i) ? f.overrides.get(i) : byDefault;
    const out = [];
    if (headerSrc) {
        for (let i = 0; i < headerSrc.h; i++) {
            if (!dropped(headerSrc, i, false)) out.push({ f: headerSrc, i: i });
        }
    }
    const headerRows = out.length;

    let repeats = 0;
    let formulas = 0;
    let blankFormulas = 0;
    let manual = 0;
    let width = 0;
    const seen = new Set();
    list.forEach(f => {
        formulas += f.formulas;
        blankFormulas += f.blankFormulas;
        manual += f.overrides.size;
        width = Math.max(width, f.width);
        f.repeats = new Set();
        f.trimmed = new Set();
        for (let i = 0; i < f.rows.length; i++) {
            const header = i < f.h;
            if (header && f === headerSrc) continue;
            if (dropped(f, i, header)) continue;
            const key = f.rows[i].key;
            if (key !== '') {
                if (seen.has(key)) {
                    repeats++;
                    f.repeats.add(i);
                    if (dedupe) continue;
                } else {
                    seen.add(key);
                }
            }
            out.push({ f: f, i: i });
        }
    });
    // Порожні рядки в самому кінці результату не потрібні, як і в кінці кожного файлу
    while (out.length > headerRows && out[out.length - 1].f.rows[out[out.length - 1].i].key === '') {
        const last = out.pop();
        last.f.trimmed.add(last.i);
    }

    if (repeats && !dedupe) {
        notes.push({
            text: num(repeats) + ' ' + plural(repeats, 'рядок даних дослівно повторює інший', 'рядки даних дослівно повторюють інші', 'рядків даних дослівно повторюють інші') + '. У результаті вони лишаються, їх прибирає перемикач «Прибирати повтори рядків».'
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
        headerRows: headerRows,
        dataRows: out.length - headerRows,
        out: out,
        width: width,
        repeats: repeats,
        manual: manual,
        notes: notes
    };
}

// Стан рядка у вихідному файлі: куди він потрапляє за поточних налаштувань
function rowState(file, i, plan) {
    const header = i < file.h;
    const manual = file.overrides.has(i);
    const dropped = manual ? file.overrides.get(i) : header && file !== plan.headerSrc;

    if (dedupe && file.repeats.has(i)) return { header: header, dropped: true, locked: true, tag: 'повтор, прибирається' };
    if (file.trimmed.has(i)) return { header: header, dropped: true, locked: true, tag: 'порожній у кінці, прибирається' };
    if (header) {
        if (manual) return { header: true, dropped: dropped, tag: dropped ? 'шапка, прибрано вручну' : 'шапка, лишено вручну' };
        return { header: true, dropped: dropped, tag: dropped ? 'шапка, прибирається' : 'шапка, лишається' };
    }
    if (dropped) return { dropped: true, tag: 'прибрано вручну' };
    return { dropped: false, tag: file.repeats.has(i) ? 'повтор' : 'дані' };
}

// ========================================
// BUILD
// ========================================
// Стиль для клітинки з назвою файлу береться з останньої оформленої клітинки рядка
function tailStyle(src) {
    for (let i = src.cells.length - 1; i >= 0; i--) {
        const s = src.cells[i] && src.cells[i].s;
        if (!s) continue;
        const style = {};
        if (s.font) style.font = s.font;
        if (s.border) style.border = s.border;
        if (s.fill) style.fill = s.fill;
        if (s.alignment) style.alignment = Object.assign({}, s.alignment, { horizontal: 'left' });
        return style;
    }
    return null;
}

// Закріплені рядки не мають заходити на дані, коли частину шапки прибрано
function fitViews(views, top, plan) {
    return (views || []).map(view => {
        if (view.state !== 'frozen' || !view.ySplit) return view;
        let split = 0;
        while (split < plan.out.length && plan.out[split].f === top && plan.out[split].i < view.ySplit) split++;
        if (split === view.ySplit) return view;

        const copy = Object.assign({}, view, { ySplit: split });
        delete copy.topLeftCell;
        if (!split && !copy.xSplit) copy.state = 'normal';
        return copy;
    });
}

async function buildFile(plan, onProgress) {
    const base = plan.files[0];
    const top = plan.headerSrc || base;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(base.sheetName, {
        properties: base.sheet.properties,
        pageSetup: base.sheet.pageSetup,
        headerFooter: base.sheet.headerFooter,
        views: fitViews(top.sheet.views, top, plan)
    });

    base.sheet.columns.forEach((c, i) => {
        if (c.width === undefined && !c.hidden) return;
        const col = ws.getColumn(i + 1);
        if (c.width !== undefined) col.width = c.width;
        if (c.hidden) col.hidden = true;
    });
    if (sourceCol) ws.getColumn(plan.width + 1).width = SOURCE_COL_WIDTH;

    const put = (ref, n) => {
        const src = ref.f.rows[ref.i];
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

        // У шапці назва колонки стоїть лише в останньому рядку, поряд із назвами решти колонок
        const label = sourceLabel(ref, n, plan);
        if (label) {
            const cell = row.getCell(plan.width + 1);
            const style = tailStyle(src);
            cell.value = label;
            if (style) cell.style = style;
        }
    };

    for (let n = 0; n < plan.out.length; n++) {
        put(plan.out[n], n + 1);
        if (n % BUILD_CHUNK === BUILD_CHUNK - 1) {
            onProgress(n + 1);
            await tick();
        }
    }

    if (plan.headerSrc) {
        const at = new Map(); // рядок шапки у вихідному файлі → рядок у результаті
        for (let n = 0; n < plan.headerRows; n++) at.set(plan.out[n].i + 1, n + 1);
        plan.headerSrc.sheet.merges.forEach(m => {
            for (let r = m.top; r <= m.bottom; r++) {
                if (!at.has(r)) return;
            }
            ws.mergeCells(m.left + at.get(m.top) + ':' + m.right + at.get(m.bottom));
        });
    }

    onProgress(null);
    await tick();
    return wb.xlsx.writeBuffer();
}

// Текст для колонки з назвою файлу: n — номер рядка в результаті
function sourceLabel(ref, n, plan) {
    if (!sourceCol) return '';
    if (n <= plan.headerRows) return n === plan.headerRows ? SOURCE_COL_TITLE : '';
    return ref.f.rows[ref.i].key === '' ? '' : ref.f.name;
}

function outputName() {
    const typed = document.getElementById('outName').value
        .replace(/\.xlsx$/i, '')
        .replace(/[\\/:*?"<>|]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    return (typed || 'Зведений реєстр') + '.xlsx';
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

    return prefix.length >= 3 ? prefix + ' (зведено)' : 'Зведений реєстр';
}

async function mergeAndDownload() {
    const plan = computePlan();
    if (busy || plan.files.length < 2 || !plan.out.length) return;

    busy = building = true;
    setStatus('');
    setProgress('Збираю рядки', 0, plan.out.length);
    render();
    await tick();

    try {
        const buffer = await buildFile(plan, done => {
            if (done === null) setProgress('Записую файл…', 0, 0);
            else setProgress('Збираю рядки', done, plan.out.length);
        });
        const name = outputName();
        const link = document.createElement('a');
        link.href = URL.createObjectURL(new Blob([buffer], { type: XLSX_MIME }));
        link.download = name;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(link.href), 1000);
        setStatus('Файл «' + name + '» збережено: ' + rowsText(plan.out.length) + '.', 'done');
    } catch (err) {
        console.error(err);
        setStatus('Звести файли не вдалося: ' + (err && err.message ? err.message : 'невідома помилка') + '.', 'error');
    }

    busy = building = false;
    setProgress(null);
    render();
}

// ========================================
// LIST ACTIONS
// ========================================
function sortFiles() {
    files.sort((a, b) => a.name.localeCompare(b.name, 'uk', { numeric: true }) || a.origin.localeCompare(b.origin, 'uk', { numeric: true }));
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

function setDedupe(on) {
    dedupe = on;
    setStatus('');
    render();
}

function setSourceCol(on) {
    sourceCol = on;
    setStatus('');
    render();
}

function setSheet(id, index) {
    const file = files.find(f => f.id === id);
    if (!file || busy) return;
    selectSheet(file, index);
    setStatus('');
    render();
}

function toggleRow(id, i) {
    const file = files.find(f => f.id === id);
    if (!file || busy || !lastPlan) return;
    const byDefault = i < file.h && file !== lastPlan.headerSrc;
    const now = file.overrides.has(i) ? file.overrides.get(i) : byDefault;
    if (!now === byDefault) file.overrides.delete(i);
    else file.overrides.set(i, !now);
    setStatus('');
    render();
}

function resetRows() {
    files.forEach(f => f.overrides.clear());
    setStatus('');
    render();
}

function showMoreRows(id) {
    const file = files.find(f => f.id === id);
    if (!file) return;
    file.headRows += PREVIEW_MORE_ROWS;
    render();
}

function pickFiles() {
    document.getElementById('fileInput').click();
}

function pickFolder() {
    document.getElementById('dirInput').click();
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
    btn.innerHTML = svg(icon);
    btn.onclick = onClick;
    return btn;
}

function fileMeta(file, plan) {
    const parts = [rowsText(file.rows.length)];
    if (file.h > 0) {
        let kept = 0;
        for (let i = 0; i < file.h; i++) {
            if (!rowState(file, i, plan).dropped) kept++;
        }
        parts.push('шапка ' + rowsText(file.h) + (kept === file.h ? ', лишається' : kept === 0 ? ', прибирається' : ', лишається ' + kept + ' з ' + file.h));
    } else if (headerMode === 'auto') {
        parts.push('шапки не знайдено');
    }
    if (file.overrides.size) parts.push('вручну змінено: ' + num(file.overrides.size));
    if (dedupe && file.repeats.size) parts.push('повторів прибрано: ' + num(file.repeats.size));
    return parts.join(' · ');
}

function dataCells(tr, src, cols) {
    for (let c = 0; c < cols; c++) {
        const cell = src.cells[c];
        const text = cell ? cellText(cell.v) : '';
        const td = el('td', '', text);
        td.title = text;
        tr.appendChild(td);
    }
}

function gapRow(text, cols, action) {
    const tr = el('tr', 'gap');
    const td = el('td', '', text);
    td.colSpan = cols + 1;
    if (action) td.appendChild(action);
    tr.appendChild(td);
    return tr;
}

function previewRow(file, i, plan, cols) {
    const state = rowState(file, i, plan);
    const tr = el('tr', (state.header ? 'is-header' : '') + (state.dropped ? ' is-dropped' : '') + (state.locked ? ' is-locked' : ''));
    const tag = el('td', 'row-tag');
    const pick = el('label', 'row-pick');
    const box = el('input');

    box.type = 'checkbox';
    box.checked = !state.dropped;
    box.disabled = busy || !!state.locked;
    box.dataset.focus = file.id + ':' + i;
    box.setAttribute('aria-label', 'Рядок ' + (i + 1) + ': лишити в результаті');
    box.onchange = () => toggleRow(file.id, i);

    pick.appendChild(box);
    pick.appendChild(el('span', 'row-num', String(i + 1)));
    pick.appendChild(el('span', 'row-state', state.tag));
    tag.appendChild(pick);
    tr.appendChild(tag);
    dataCells(tr, file.rows[i], cols);

    if (!state.locked) {
        tr.onclick = event => {
            if (!event.target.closest('.row-pick')) toggleRow(file.id, i);
        };
    }
    return tr;
}

function fillPreview(box, file, plan) {
    const table = el('table', 'preview-table pickable');
    const total = file.rows.length;
    const cols = Math.min(file.width, PREVIEW_MAX_COLS);
    const head = Math.min(total, Math.max(file.headRows, file.h + 3));
    const tail = Math.max(head, total - PREVIEW_TAIL_ROWS);

    for (let i = 0; i < head; i++) table.appendChild(previewRow(file, i, plan, cols));
    if (tail > head) {
        const more = el('button', 'link-btn', 'Показати ще ' + Math.min(PREVIEW_MORE_ROWS, tail - head));
        more.type = 'button';
        more.onclick = () => showMoreRows(file.id);
        table.appendChild(gapRow('Ще ' + rowsText(tail - head) + ' між початком і кінцем файлу. ', cols, more));
    }
    for (let i = tail; i < total; i++) table.appendChild(previewRow(file, i, plan, cols));

    box.replaceChildren(table);
}

function renderFile(file, index, plan) {
    const li = el('li', file.error ? 'file failed' : 'file');
    const row = el('div', 'file-row');
    const main = el('div', 'file-main');

    // Список перемальовується цілком, тому поява нового рядка триває з того місця, де її перервало перемальовування
    if (!file.shownAt) file.shownAt = performance.now();
    const elapsed = performance.now() - file.shownAt;
    if (elapsed < ENTER_MS) {
        li.classList.add('enter');
        li.style.animationDelay = -Math.round(elapsed) + 'ms';
    }

    row.appendChild(el('span', 'file-num', String(index + 1)));
    main.appendChild(el('span', 'file-name', file.name));
    if (file.origin) main.appendChild(el('span', 'file-origin', file.origin));
    main.appendChild(el('span', 'file-meta', file.error || fileMeta(file, plan)));

    if (!file.error && file.sheets.length > 1) {
        const pick = el('label', 'sheet-pick', 'Аркуш');
        const select = el('select');
        file.sheets.forEach((sheet, i) => {
            const size = sheet.rows.length ? rowsText(sheet.rows.length) : 'порожній';
            const option = el('option', '', sheet.sheetName + ' · ' + size + (sheet.hidden ? ' · прихований' : ''));
            option.value = i;
            option.disabled = !sheet.rows.length;
            option.selected = i === file.sheetIdx;
            select.appendChild(option);
        });
        select.disabled = busy;
        select.dataset.focus = file.id + ':sheet';
        select.onchange = () => setSheet(file.id, Number(select.value));
        pick.appendChild(select);
        main.appendChild(pick);
    }
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
        box.dataset.scroll = file.id;
        details.appendChild(el('summary', '', 'Переглянути рядки й вибрати, які прибрати'));
        details.appendChild(box);
        details.appendChild(el('p', 'preview-hint', 'Клік по рядку прибирає його з результату або повертає назад.'));
        details.open = openPreviews.has(file.id);
        if (details.open) fillPreview(box, file, plan);
        details.addEventListener('toggle', () => {
            // Відкрите превю після перемальовування теж дає подію toggle, її треба пропустити
            if (details.open === openPreviews.has(file.id)) return;
            if (details.open) {
                openPreviews.add(file.id);
                box.classList.add('reveal');
                fillPreview(box, file, lastPlan);
            } else {
                openPreviews.delete(file.id);
            }
        });
        li.appendChild(details);
    }

    return li;
}

// Превю результату: початок і кінець, а також місця, де один файл змінюється іншим
function fillResult(plan) {
    const box = document.getElementById('resultBox');
    const note = document.getElementById('resultNote');
    const total = plan.out.length;
    if (plan.files.length < 2 || !total) {
        box.replaceChildren();
        note.textContent = '';
        return;
    }

    const pick = new Set();
    const joins = [];
    for (let n = 0; n < Math.min(total, plan.headerRows + RESULT_TOP_ROWS); n++) pick.add(n);
    for (let n = plan.headerRows + 1; n < total; n++) {
        if (plan.out[n].f !== plan.out[n - 1].f) joins.push(n);
    }
    joins.slice(0, RESULT_MAX_JOINS).forEach(n => { pick.add(n - 1); pick.add(n); });
    for (let n = Math.max(0, total - 2); n < total; n++) pick.add(n);

    const cols = Math.min(plan.width, PREVIEW_MAX_COLS);
    const span = cols + (sourceCol ? 2 : 1);
    const table = el('table', 'preview-table result-table');
    let prev = -1;
    Array.from(pick).sort((a, b) => a - b).forEach(n => {
        if (n - prev > 1) table.appendChild(gapRow('Ще ' + rowsText(n - prev - 1), span));
        const ref = plan.out[n];
        const header = n < plan.headerRows;
        const tr = el('tr', header ? 'is-header' : '');
        const tag = el('td', 'row-tag');
        tag.appendChild(el('span', 'row-num', num(n + 1)));
        tag.appendChild(el('span', 'row-state', (header ? 'шапка, ' : '') + 'файл ' + (files.indexOf(ref.f) + 1) + ', рядок ' + num(ref.i + 1)));
        tag.title = ref.f.name;
        tr.appendChild(tag);
        dataCells(tr, ref.f.rows[ref.i], cols);
        if (sourceCol) tr.appendChild(el('td', 'source-cell', sourceLabel(ref, n + 1, plan)));
        table.appendChild(tr);
        prev = n;
    });

    box.replaceChildren(table);
    note.textContent = 'Показано початок і кінець результату, а також місця, де один файл змінюється іншим. Біля номера рядка вказано, звідки його взято.'
        + (joins.length > RESULT_MAX_JOINS ? ' Таких місць ' + num(joins.length) + ', показано перші ' + RESULT_MAX_JOINS + '.' : '');
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
    if (count < 2) return 'Для зведення потрібно щонайменше два файли.';
    if (!plan.out.length) return 'У результаті не лишилося жодного рядка.';

    const from = ' з ' + count + ' ' + plural(count, 'файлу', 'файлів', 'файлів') + '.';
    const removed = dedupe && plan.repeats ? ' Повторів прибрано: ' + num(plan.repeats) + '.' : '';
    if (plan.headerRows > 0) {
        return 'У результаті буде ' + rowsText(plan.out.length) + ': ' + rowsText(plan.headerRows) + ' шапки і ' + rowsText(plan.dataRows) + ' даних' + from + removed;
    }
    return 'У результаті буде ' + rowsText(plan.dataRows) + ' даних' + from + removed;
}

function render() {
    const plan = lastPlan = computePlan();

    // Перемальовування не має збивати прокрутку превю і фокус
    const scrolls = new Map();
    document.querySelectorAll('[data-scroll]').forEach(box => scrolls.set(box.dataset.scroll, box.scrollLeft));
    const focused = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.focus : '';

    document.getElementById('panel').hidden = files.length === 0;
    document.getElementById('dropzone').classList.toggle('busy', busy);
    ['pickBtn', 'dirBtn', 'dedupe', 'sourceCol'].forEach(id => { document.getElementById(id).disabled = busy; });
    document.querySelectorAll('.seg').forEach(seg => { seg.disabled = busy; });
    document.getElementById('sortBtn').disabled = busy || !manualOrder;
    document.getElementById('clearBtn').disabled = busy;
    document.getElementById('mergeBtn').disabled = busy || plan.files.length < 2 || !plan.out.length;
    document.getElementById('mergeBtn').classList.toggle('loading', building);
    document.getElementById('headerHint').textContent = headerHint();
    document.getElementById('summary').textContent = summaryText(plan);
    document.getElementById('manual').hidden = !plan.manual;
    document.getElementById('manualText').textContent = 'Вручну змінено рядків: ' + num(plan.manual) + '.';
    document.getElementById('resetRowsBtn').disabled = busy;
    document.getElementById('resultPreview').hidden = plan.files.length < 2;

    document.getElementById('fileList').replaceChildren(
        ...files.map((file, index) => renderFile(file, index, plan))
    );
    document.getElementById('notes').replaceChildren(
        ...plan.notes.map(note => el('li', note.warn ? 'note warn' : 'note', note.text))
    );
    fillResult(plan);

    const outName = document.getElementById('outName');
    if (!nameEdited) outName.value = defaultName();

    document.querySelectorAll('[data-scroll]').forEach(box => {
        if (scrolls.has(box.dataset.scroll)) box.scrollLeft = scrolls.get(box.dataset.scroll);
    });
    if (focused) {
        const target = document.querySelector('[data-focus="' + focused + '"]');
        if (target && !target.disabled) target.focus({ preventScroll: true });
    }
}

// ========================================
// EVENTS
// ========================================
(function init() {
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const dirInput = document.getElementById('dirInput');

    fileInput.addEventListener('change', () => {
        addSources(Array.from(fileInput.files).map(file => ({ file: file, dir: '' })));
        fileInput.value = '';
    });
    dirInput.addEventListener('change', () => {
        addSources(Array.from(dirInput.files).map(file => ({
            file: file,
            dir: (file.webkitRelativePath || '').split('/').slice(0, -1).join(' / ') || 'папка'
        })));
        dirInput.value = '';
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
        if (!busy) addDropped(event.dataTransfer);
    });

    // Файл, кинутий повз зону, не має відкритися замість інструмента
    ['dragover', 'drop'].forEach(type => {
        window.addEventListener(type, event => event.preventDefault());
    });

    render();
})();
