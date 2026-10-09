// ========================================
// AUXILIUM — ІДЕНТИФІКАТОРИ АСВП
// ========================================

const PDFJS_URL = '../vendor/pdfjs/pdf.min.mjs';
const PDFJS_WORKER_URL = '../vendor/pdfjs/pdf.worker.min.mjs';
const PAGE_ROWS = 100;
const MAX_NAMES = 3;
const MAX_CONFLICT_NOTES = 3;
const RENDER_EVERY_MS = 150;
const ENTER_MS = 320; // тривалість появи рядка в таблиці, та сама, що в asvp.css
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const DEFAULT_NAME = 'Ідентифікатори АСВП';
const SHEET_NAME = 'Ідентифікатори АСВП';
const COL_VP = 'Номер ВП';
const COL_ID = 'Ідентифікатор доступу';
const COL_FILE = 'Файл';

// Колонтитул кожної сторінки постанови: «Реєстраційний номер виконавчого провадження: …» і «Ідентифікатор доступу: …, адреса для ознайомлення…».
// Якщо колонтитула нема, номер береться із заголовка «ВП № …», а ідентифікатор з пункту «Ідентифікатор для доступу …».
// Ідентифікатор містить кириличні літери, тому переноситься символ у символ, без заміни на латиницю.
const ID_PART = '([0-9A-ZА-ЯІЇЄҐ]{8,20})(?![0-9A-Za-zА-Яа-яІЇЄҐіїєґ])';
const RE_VP = /Реєстраційний\s*номер\s*виконавчого\s*провадження\s*:?\s*(\d+)/;
const RE_VP_TITLE = /(?<![A-ZА-ЯІЇЄҐ])ВП\s*№\s*(\d+)/;
const RE_ID = new RegExp('Ідентифікатор\\s*доступу\\s*:?\\s*' + ID_PART);
const RE_ID_BODY = new RegExp('Ідентифікатор\\s*для\\s*доступу\\s*:?\\s*' + ID_PART);

let files = [];            // прочитані файли в порядку додавання
let includeNoId = false;
let sourceCol = false;
let filter = 'all';        // 'all' або 'attention'
let shownRows = PAGE_ROWS;
let nameEdited = false;
let busy = false;          // іде читання файлів або запис результату
let building = false;      // іде саме запис результату
let nextId = 1;
let pdfjs = null;          // бібліотека читання PDF вантажиться разом із першим файлом
let pdfWorker = null;
const pdfUrls = new Map(); // id файлу → адреса для перегляду PDF у новій вкладці

const ICONS = {
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

function namesText(list) {
    return list.slice(0, MAX_NAMES).join(', ') + (list.length > MAX_NAMES ? ' і ще ' + num(list.length - MAX_NAMES) : '');
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

// ========================================
// READ PDF
// ========================================
async function loadPdfjs() {
    if (pdfjs) return;
    const lib = await import(new URL(PDFJS_URL, document.baseURI).href);
    lib.GlobalWorkerOptions.workerSrc = new URL(PDFJS_WORKER_URL, document.baseURI).href;
    // Один фоновий потік на всі файли, щоб не запускати новий для кожного PDF
    pdfWorker = new lib.PDFWorker({ name: 'asvp', verbosity: 0 });
    await pdfWorker.promise;
    pdfjs = lib;
}

function pageText(content) {
    return content.items.map(item => (item.str || '') + (item.hasEOL ? '\n' : '')).join('');
}

function parsePage(text) {
    const vp = RE_VP.exec(text) || RE_VP_TITLE.exec(text);
    const id = RE_ID.exec(text) || RE_ID_BODY.exec(text);
    return { vp: vp ? vp[1] : '', id: id ? id[1] : '' };
}

async function readPdf(file) {
    const task = pdfjs.getDocument({
        data: new Uint8Array(await file.arrayBuffer()),
        worker: pdfWorker,
        isEvalSupported: false,
        verbosity: 0
    });
    try {
        const doc = await task.promise;
        const pairs = [];
        let hasText = false;
        // Сторінки читаються всі: в одному PDF можуть бути зшиті постанови різних проваджень
        for (let p = 1; p <= doc.numPages; p++) {
            const page = await doc.getPage(p);
            const text = pageText(await page.getTextContent());
            page.cleanup();
            if (/\S/.test(text)) hasText = true;
            const pair = parsePage(text);
            if (pair.vp && !pairs.some(x => x.vp === pair.vp && x.id === pair.id)) pairs.push(pair);
        }
        // Сторінка без ідентифікатора не рахується, якщо інша сторінка того самого ВП його має
        return { hasText: hasText, pairs: pairs.filter(pair => pair.id || !pairs.some(x => x.vp === pair.vp && x.id)) };
    } finally {
        task.destroy();
    }
}

// ========================================
// SOURCES: файли, папки, архіви
// ========================================
function isPdf(name) {
    return /\.pdf$/i.test(name);
}

// Службові файли macOS і тимчасові файли Office
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
            if (isPdf(name)) return true;
            other++;
            return false;
        }
    });

    const items = Object.keys(found).map(raw => {
        const parts = zipName(raw).normalize('NFC').replace(/\\/g, '/').split('/');
        const name = parts.pop();
        return {
            file: new File([found[raw]], name, { type: 'application/pdf', lastModified: item.file.lastModified }),
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
        file: item.file,
        pairs: [],
        kind: '',     // причина, з якої файл не дав жодного ВП
        error: ''
    };
}

function readError(err) {
    if (err && err.name === 'PasswordException') return { kind: 'locked', text: 'PDF захищений паролем, тож прочитати його не вдалося.' };
    return { kind: 'broken', text: 'Файл не вдалося прочитати. Можливо, він пошкоджений або це не PDF.' };
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

    const addFailed = (item, kind, message) => {
        const entry = makeEntry(item);
        if (files.some(f => f.key === entry.key)) {
            listed.push(item.name);
            return;
        }
        entry.kind = kind;
        entry.error = message;
        files.push(entry);
    };

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
                    addFailed(item, 'format', 'В архіві нема файлів PDF.');
                }
            } catch (err) {
                console.error(err);
                addFailed(item, 'broken', 'Архів не вдалося розпакувати. Можливо, він пошкоджений або захищений паролем.');
            }
        } else if (isPdf(name)) {
            queue.push(item);
        } else if (item.bulk) {
            other++;
        } else if (/\.(rar|7z)$/i.test(name)) {
            addFailed(item, 'format', 'З архівів підтримується лише ZIP. Архів RAR або 7z спершу розпакуйте й додайте папку.');
        } else {
            addFailed(item, 'format', 'Підтримується лише формат PDF.');
        }
    }

    if (queue.length && !pdfjs) {
        setProgress('Готую модуль читання PDF…', 0, 0);
        try {
            await loadPdfjs();
        } catch (err) {
            console.error(err);
            busy = false;
            setProgress(null);
            setStatus('Модуль читання PDF не завантажився. Оновіть сторінку, а якщо це не допоможе, оновіть браузер.', 'error');
            render();
            return;
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

        try {
            const parsed = await readPdf(item.file);
            entry.pairs = parsed.pairs;
            if (!parsed.pairs.length && !parsed.hasText) {
                entry.kind = 'scan';
                entry.error = 'У файлі нема текстового шару, схоже, це скан.';
            } else if (!parsed.pairs.length) {
                entry.kind = 'foreign';
                entry.error = 'Номер ВП не знайдено, схоже, це не постанова АСВП.';
            }
        } catch (err) {
            // Пароль і пошкоджений файл очікувані, тож у консоль іде лише невідома помилка
            if (!/^(Password|InvalidPDF)Exception$/.test(err && err.name)) console.error(err);
            const failure = readError(err);
            entry.kind = failure.kind;
            entry.error = failure.text;
        }

        files.push(entry);
        if (!entry.error) added++;

        if (performance.now() - painted > RENDER_EVERY_MS) {
            render();
            painted = performance.now();
        }
    }

    busy = false;
    setProgress(null);

    const report = [];
    if (added) report.push('Додано ' + num(added) + ' ' + plural(added, 'файл', 'файли', 'файлів') + '.');
    if (other) report.push('Пропущено файлів іншого формату: ' + num(other) + '.');
    if (listed.length) report.push('Уже в списку: ' + namesText(listed) + '.');
    if (!files.length) report.unshift('Файлів PDF не знайдено.');
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
// Номери ВП порівнюються як числа будь-якої довжини
function vpOrder(a, b) {
    return a.length - b.length || (a < b ? -1 : a > b ? 1 : 0);
}

function isAttention(row) {
    return row.state === 'error' || row.state === 'noId' || row.state === 'conflict';
}

function computePlan() {
    const rows = [];
    files.forEach(file => {
        if (file.error) rows.push({ file: file, vp: '', id: '', state: 'error' });
        else file.pairs.forEach(pair => rows.push({ file: file, vp: pair.vp, id: pair.id, state: 'ok' }));
    });
    // У межах одного ВП рядки йдуть у порядку додавання, тож повтором позначається пізніший файл
    rows.sort((a, b) => (a.state === 'error') - (b.state === 'error') || vpOrder(a.vp, b.vp) || a.file.id - b.file.id || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

    const groups = new Map();
    rows.forEach(row => {
        if (row.state === 'error') return;
        if (!groups.has(row.vp)) groups.set(row.vp, []);
        groups.get(row.vp).push(row);
    });

    const out = [];
    const conflicts = [];
    const missing = [];
    let repeats = 0;

    groups.forEach((group, vp) => {
        const ids = [];
        group.forEach(row => { if (row.id && !ids.includes(row.id)) ids.push(row.id); });
        if (ids.length > 1) conflicts.push({ vp: vp, rows: group.filter(row => row.id) });
        if (!ids.length) missing.push(vp);

        const taken = new Map(); // ідентифікатор → рядок результату
        group.forEach(row => {
            if (!row.id && ids.length) {
                row.state = 'covered';
                return;
            }
            if (!row.id) row.state = 'noId';
            else if (ids.length > 1) row.state = 'conflict';

            if (taken.has(row.id)) {
                row.repeat = true;
                repeats++;
                const line = taken.get(row.id);
                if (!line.names.includes(row.file.name)) line.names.push(row.file.name);
                return;
            }
            if (row.state === 'noId' && !includeNoId) return;
            const line = { vp: vp, id: row.id, names: [row.file.name] };
            taken.set(row.id, line);
            out.push(line);
        });
    });

    // Файл, усі провадження якого вже є в таблиці з інших файлів, у нотатку про кілька проваджень не йде
    const fresh = new Set(rows.filter(row => row.state !== 'error' && !row.repeat).map(row => row.file));

    return {
        rows: rows,
        out: out,
        conflicts: conflicts,
        missing: missing,
        repeats: repeats,
        attention: rows.filter(isAttention).length,
        failed: files.filter(f => f.error),
        multi: files.filter(f => f.pairs.length > 1 && fresh.has(f)),
        notes: []
    };
}

function addNotes(plan) {
    const notes = plan.notes;
    const failed = kinds => plan.failed.filter(f => kinds.includes(f.kind)).map(f => f.name);

    const scans = failed(['scan']);
    if (scans.length) {
        notes.push({ key: 'scan', warn: true, text: 'Текстового шару нема в ' + num(scans.length) + ' ' + plural(scans.length, 'файлі', 'файлах', 'файлах') + ': ' + namesText(scans)
            + (scans.length === 1 ? '. Схоже, це скан' : '. Схоже, це скани') + ', а номер ВП читається лише з PDF, згенерованих в АСВП.' });
    }
    const foreign = failed(['foreign']);
    if (foreign.length) {
        notes.push({ key: 'foreign', warn: true, text: 'Номера ВП не знайдено в ' + num(foreign.length) + ' ' + plural(foreign.length, 'файлі', 'файлах', 'файлах') + ': ' + namesText(foreign)
            + (foreign.length === 1 ? '. Схоже, це не постанова АСВП.' : '. Схоже, це не постанови АСВП.') });
    }
    const broken = failed(['broken', 'locked', 'format']);
    if (broken.length) {
        notes.push({ key: 'broken', warn: true, text: 'Не вдалося прочитати ' + num(broken.length) + ' ' + plural(broken.length, 'файл', 'файли', 'файлів') + ': ' + namesText(broken)
            + '. Причину видно в таблиці біля кожного з них.' });
    }

    if (plan.missing.length) {
        const count = plan.missing.length;
        notes.push({ key: 'missing', warn: true, text: 'Ідентифікатора доступу нема для ' + num(count) + ' ' + plural(count, 'провадження', 'проваджень', 'проваджень')
            + ': ВП № ' + namesText(plan.missing) + '. Його зазначають лише в постановах, які надсилають стягувачу й боржнику.'
            + (includeNoId ? ' Ці номери йдуть у таблицю з порожньою коміркою.' : ' Щоб ці номери потрапили в таблицю з порожньою коміркою, увімкніть перемикач під таблицею.') });
    }

    plan.conflicts.slice(0, MAX_CONFLICT_NOTES).forEach(conflict => {
        const variants = [];
        conflict.rows.forEach(row => {
            const known = variants.find(v => v.id === row.id);
            if (known) known.names.push(row.file.name);
            else variants.push({ id: row.id, names: [row.file.name] });
        });
        notes.push({ key: 'conflict ' + conflict.vp, warn: true, text: 'ВП № ' + conflict.vp + ' має різні ідентифікатори: '
            + variants.map(v => v.id + ' у файлі ' + namesText(v.names)).join(', ')
            + '. У таблицю йдуть усі, тож варто перевірити, який із них чинний.' });
    });
    if (plan.conflicts.length > MAX_CONFLICT_NOTES) {
        const rest = plan.conflicts.length - MAX_CONFLICT_NOTES;
        notes.push({ key: 'conflicts', warn: true, text: 'Ще ' + num(rest) + ' ' + plural(rest, 'провадження має', 'провадження мають', 'проваджень мають')
            + ' різні ідентифікатори в різних файлах. Їх видно у фільтрі «Потребують уваги».' });
    }

    if (plan.repeats) {
        notes.push({ key: 'repeats', text: 'Повторів: ' + num(plan.repeats) + '. Те саме провадження з тим самим ідентифікатором трапилось у кількох файлах, і в таблицю воно йде один раз.' });
    }
    if (plan.multi.length) {
        const names = plan.multi.map(f => f.name + ' (' + f.pairs.length + ')');
        notes.push({ key: 'multi', text: 'Кілька проваджень знайдено в ' + num(names.length) + ' ' + plural(names.length, 'файлі', 'файлах', 'файлах') + ': ' + namesText(names)
            + '. Кожне з них іде в таблицю окремим рядком.' });
    }
}

// ========================================
// OUTPUT
// ========================================
// Номер ВП записується числом, щоб Excel сортував і шукав його так само, як інші числа
function vpValue(vp) {
    return /^[1-9]\d{0,14}$/.test(vp) ? Number(vp) : vp;
}

function outputName() {
    const typed = document.getElementById('outName').value
        .replace(/\.xlsx$/i, '')
        .replace(/[\\/:*?"<>|]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    return (typed || DEFAULT_NAME) + '.xlsx';
}

// Коли всі файли взято з однієї папки чи архіву, її назва потрапляє в назву результату
function defaultName() {
    const roots = new Set(files.map(f => f.origin.split(' / ')[0]));
    const root = roots.size === 1 ? Array.from(roots)[0].replace(/\.zip$/i, '').trim() : '';
    return root ? DEFAULT_NAME + ' (' + root + ')' : DEFAULT_NAME;
}

async function buildFile(plan) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(SHEET_NAME, { views: [{ state: 'frozen', ySplit: 1 }] });
    const columns = [
        { header: COL_VP, key: 'vp', width: 14 },
        { header: COL_ID, key: 'id', width: 24 }
    ];
    if (sourceCol) columns.push({ header: COL_FILE, key: 'file', width: 50 });
    ws.columns = columns;

    plan.out.forEach(line => {
        ws.addRow({ vp: vpValue(line.vp), id: line.id || null, file: line.names.join(', ') });
    });
    ws.getRow(1).font = { bold: true };
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };

    return wb.xlsx.writeBuffer();
}

async function saveExcel() {
    const plan = computePlan();
    if (busy || !plan.out.length) return;

    busy = building = true;
    setStatus('');
    setProgress('Записую файл…', 0, 0);
    render();
    await tick();

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
        setStatus('Файл «' + name + '» збережено: ' + rowsText(plan.out.length) + ' під шапкою.', 'done');
    } catch (err) {
        console.error(err);
        setStatus('Записати файл не вдалося: ' + (err && err.message ? err.message : 'невідома помилка') + '.', 'error');
    }

    busy = building = false;
    setProgress(null);
    render();
}

async function copyTable() {
    const plan = computePlan();
    if (busy || !plan.out.length) return;

    const text = plan.out.map(line => [line.vp, line.id].concat(sourceCol ? [line.names.join(', ')] : []).join('\t')).join('\r\n');
    try {
        await navigator.clipboard.writeText(text);
    } catch (err) {
        // Якщо браузер не дав доступу до буфера обміну, спрацьовує старий спосіб через виділений текст
        const area = el('textarea');
        area.value = text;
        area.style.position = 'fixed';
        area.style.opacity = '0';
        document.body.appendChild(area);
        area.select();
        const copied = document.execCommand('copy');
        area.remove();
        if (!copied) {
            setStatus('Скопіювати не вдалося, бо браузер не дав доступу до буфера обміну.', 'error');
            return;
        }
    }
    setStatus('Скопійовано ' + rowsText(plan.out.length) + ' без шапки. Їх можна вставити в Excel або Google Таблиці.', 'done');
}

// ========================================
// LIST ACTIONS
// ========================================
function openPdf(entry) {
    if (!pdfUrls.has(entry.id)) {
        pdfUrls.set(entry.id, URL.createObjectURL(new Blob([entry.file], { type: 'application/pdf' })));
    }
    if (!window.open(pdfUrls.get(entry.id), '_blank')) {
        setStatus('Браузер не відкрив нову вкладку. Дозвольте спливні вікна для цієї сторінки.', 'error');
    }
}

function forgetUrl(id) {
    if (!pdfUrls.has(id)) return;
    URL.revokeObjectURL(pdfUrls.get(id));
    pdfUrls.delete(id);
}

function removeFile(id) {
    if (busy) return;
    files = files.filter(f => f.id !== id);
    forgetUrl(id);
    if (!files.length) resetList();
    setStatus('');
    render();
}

function resetList() {
    files = [];
    filter = 'all';
    shownRows = PAGE_ROWS;
    nameEdited = false;
    Array.from(pdfUrls.keys()).forEach(forgetUrl);
}

function clearFiles() {
    resetList();
    setStatus('');
    render();
}

function setFilter(btn) {
    filter = btn.dataset.filter;
    shownRows = PAGE_ROWS;
    render();
}

function showMore() {
    shownRows += PAGE_ROWS;
    render();
}

function setIncludeNoId(on) {
    includeNoId = on;
    setStatus('');
    render();
}

function setSourceCol(on) {
    sourceCol = on;
    setStatus('');
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

function rowTag(row) {
    if (row.state === 'covered') return { text: 'ідентифікатора тут нема, але він є в іншому файлі', dropped: true };
    if (row.state === 'noId') {
        if (row.repeat) return { text: 'ідентифікатора нема, номер уже є в таблиці', dropped: true, warn: true };
        return includeNoId
            ? { text: 'ідентифікатора нема, у таблицю йде з порожньою коміркою', warn: true }
            : { text: 'ідентифікатора нема, у таблицю не йде', dropped: true, warn: true };
    }
    if (row.state === 'conflict') {
        return { text: 'для цього ВП в іншому файлі інший ідентифікатор' + (row.repeat ? ', цей уже є в таблиці' : ''), dropped: !!row.repeat, warn: true };
    }
    if (row.repeat) return { text: 'повтор, у таблицю не йде', dropped: true };
    return { text: '' };
}

function fileCell(file) {
    const td = el('td', 'c-file');
    if (isPdf(file.name) && file.kind !== 'format') {
        const link = el('button', 'file-link', file.name);
        link.type = 'button';
        link.title = 'Відкрити PDF у новій вкладці';
        link.onclick = () => openPdf(file);
        td.appendChild(link);
    } else {
        td.appendChild(el('span', '', file.name));
    }
    if (file.origin) td.appendChild(el('span', 'file-origin', file.origin));
    return td;
}

function rowView(row, index) {
    const file = row.file;
    const tag = row.state === 'error' ? { text: '' } : rowTag(row);
    const tr = el('tr', (tag.dropped ? 'is-dropped' : '') + (tag.warn ? ' is-warn' : ''));

    // Таблиця перемальовується цілком, тому поява нового рядка триває з того місця, де її перервало перемальовування
    if (!file.shownAt) file.shownAt = performance.now();
    const elapsed = performance.now() - file.shownAt;
    if (elapsed < ENTER_MS) {
        tr.classList.add('enter');
        tr.style.animationDelay = -Math.round(elapsed) + 'ms';
    }

    tr.appendChild(el('td', 'c-num', num(index + 1)));
    if (row.state === 'error') {
        const td = el('td', 'row-error', file.error);
        td.colSpan = 2;
        tr.appendChild(td);
    } else {
        tr.appendChild(el('td', 'c-vp', row.vp));
        tr.appendChild(el('td', row.id ? 'c-id' : 'c-id empty', row.id || 'нема'));
    }

    const name = fileCell(file);
    if (tag.text) name.appendChild(el('span', 'row-tag', tag.text));
    tr.appendChild(name);

    const actions = el('td', 'c-act');
    actions.appendChild(iconButton('remove', 'Прибрати файл зі списку', busy, () => removeFile(file.id)));
    tr.appendChild(actions);
    return tr;
}

function fillRows(plan) {
    const list = filter === 'attention' ? plan.rows.filter(isAttention) : plan.rows;
    const shown = list.slice(0, shownRows);
    const trs = shown.map(rowView);

    if (list.length > shown.length) {
        const tr = el('tr', 'gap');
        const td = el('td', '', 'Показано ' + num(shown.length) + ' з ' + num(list.length) + '.');
        const more = el('button', 'link-btn', 'Показати ще ' + num(Math.min(PAGE_ROWS, list.length - shown.length)));
        more.type = 'button';
        more.onclick = showMore;
        td.colSpan = 5;
        td.appendChild(more);
        tr.appendChild(td);
        trs.push(tr);
    }
    if (!list.length) {
        const tr = el('tr', 'blank');
        const td = el('td', '', 'Рядків, що потребують уваги, нема.');
        td.colSpan = 5;
        tr.appendChild(td);
        trs.push(tr);
    }
    document.getElementById('rows').replaceChildren(...trs);
}

// Нотатка, що вже на екрані, лише оновлює текст: нова поява на кожному перемальовуванні виглядала б як блимання
function fillNotes(notes) {
    const list = document.getElementById('notes');
    const shown = new Map(Array.from(list.children).map(li => [li.dataset.key, li]));
    Array.from(list.children).forEach(li => { if (!notes.some(note => note.key === li.dataset.key)) li.remove(); });
    notes.forEach((note, i) => {
        let li = shown.get(note.key);
        if (!li) {
            li = el('li');
            li.dataset.key = note.key;
        }
        li.className = note.warn ? 'note warn' : 'note';
        if (li.textContent !== note.text) li.textContent = note.text;
        // Вузол переставляється лише тоді, коли стоїть не на своєму місці, бо вставка заново перезапускає анімацію
        if (list.children[i] !== li) list.insertBefore(li, list.children[i] || null);
    });
}

function summaryText(plan) {
    const count = files.length;
    const from = ' з ' + num(count) + ' ' + plural(count, 'файлу', 'файлів', 'файлів');
    const attention = plan.attention ? ' Потребують уваги: ' + num(plan.attention) + '.' : '';
    if (!plan.out.length) return 'Таблиця поки порожня: у ' + (count === 1 ? 'файлі' : 'файлах') + ' не знайдено ВП з ідентифікатором доступу.' + attention;
    return 'У таблицю піде ' + rowsText(plan.out.length) + from + '.' + attention;
}

function render() {
    const plan = computePlan();
    addNotes(plan);

    if (!plan.attention) filter = 'all';

    document.getElementById('panel').hidden = files.length === 0;
    document.getElementById('dropzone').classList.toggle('busy', busy);
    ['pickBtn', 'dirBtn', 'includeNoId', 'sourceCol', 'clearBtn'].forEach(id => { document.getElementById(id).disabled = busy; });
    document.getElementById('copyBtn').disabled = busy || !plan.out.length;
    document.getElementById('saveBtn').disabled = busy || !plan.out.length;
    document.getElementById('saveBtn').classList.toggle('loading', building);
    document.getElementById('summary').textContent = summaryText(plan);
    document.getElementById('noIdToggle').hidden = !plan.missing.length;

    document.getElementById('filter').hidden = !plan.attention;
    document.getElementById('countAll').textContent = num(plan.rows.length);
    document.getElementById('countAttention').textContent = num(plan.attention);
    document.querySelectorAll('.seg').forEach(seg => {
        const active = seg.dataset.filter === filter;
        seg.classList.toggle('active', active);
        seg.setAttribute('aria-pressed', active);
        seg.disabled = busy;
    });

    fillNotes(plan.notes);
    fillRows(plan);

    const outName = document.getElementById('outName');
    if (!nameEdited) outName.value = defaultName();
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
