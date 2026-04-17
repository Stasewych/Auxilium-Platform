// ========================================
// AUXILIUM PLATFORM
// ========================================

const toolNames = {
    iban:     'Фільтрація IBAN',
    calc:     'Калькулятор судових зборів',
    ocr:     'Сканер документів',
    zvirka:  'Звірка надходжень',
    debtors: 'Перевірка боржників',
    clicker: 'Клікер'
};

const downloadInfo = {
    clicker: {
        title: 'Клікер',
        desc: 'Десктопна програма-клікер. Посилання на завантаження буде додано найближчим часом.'
    }
};

function openTool(btn, type, url) {
    const tool = btn.dataset.tool;

    // Update active nav
    document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
    btn.classList.add('active');

    // Hide all views
    document.getElementById('welcome').style.display = 'none';
    document.getElementById('iframeWrap').classList.remove('active');
    document.getElementById('downloadPage').classList.remove('active');

    if (type === 'iframe') {
        const frame = document.getElementById('toolFrame');
        const wrap = document.getElementById('iframeWrap');
        document.getElementById('iframeTitle').textContent = toolNames[tool] || '';
        frame.src = url;
        wrap.classList.add('active');
    } else if (type === 'download') {
        const info = downloadInfo[tool] || {};
        document.getElementById('downloadTitle').textContent = info.title || 'Завантажити програму';
        document.getElementById('downloadDesc').textContent = info.desc || '';

        const dlBtn = document.getElementById('downloadBtn');
        if (url && url !== '#') {
            dlBtn.href = url;
            dlBtn.classList.remove('disabled');
            dlBtn.textContent = 'Завантажити';
        } else {
            dlBtn.href = '#';
            dlBtn.classList.add('disabled');
            dlBtn.textContent = 'Незабаром';
        }

        document.getElementById('downloadPage').classList.add('active');
    }

    // Close mobile sidebar
    closeSidebar();
}

function closeTool() {
    document.getElementById('iframeWrap').classList.remove('active');
    document.getElementById('downloadPage').classList.remove('active');
    document.getElementById('welcome').style.display = 'flex';
    document.getElementById('toolFrame').src = '';

    document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
}

function toggleSidebar() {
    document.getElementById('sidebar').classList.toggle('open');
    document.getElementById('overlay').classList.toggle('open');
}

function closeSidebar() {
    document.getElementById('sidebar').classList.remove('open');
    document.getElementById('overlay').classList.remove('open');
}
