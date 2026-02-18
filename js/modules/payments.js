function generateMonthOptions(showAll = false) {
    const months = [
        'Січень', 'Лютий', 'Березень', 'Квітень', 'Травень', 'Червень',
        'Липень', 'Серпень', 'Вересень', 'Жовтень', 'Листопад', 'Грудень'
    ];
    const select = document.getElementById('monthSelectPayments');
    const currentValue = select.value;

    select.innerHTML = '<option value="">Оберіть місяць</option>';

    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();

    const oneMonthAgo = new Date(currentYear, currentMonth - 1, 1);

    for (let year = currentYear - 2; year <= currentYear + 1; year++) {
        months.forEach((month, index) => {
            const monthDate = new Date(year, index, 1);

            if (!showAll) {
                if (monthDate < oneMonthAgo) {
                    return;
                }
            }

            const option = document.createElement('option');
            option.value = `Оплата ${month} ${year}`;
            option.textContent = `${month} ${year}`;

            if (year === currentYear && index === currentMonth) {
                option.textContent += ' (поточний)';
            }

            select.appendChild(option);
        });
    }

    if (currentValue) {
        const options = Array.from(select.options);
        const matchingOption = options.find(opt => opt.value === currentValue);
        if (matchingOption) {
            select.value = currentValue;
        }
    }
}

generateMonthOptions(false);

document.getElementById('showOlderMonthsPayments').addEventListener('change', function (e) {
    generateMonthOptions(e.target.checked);
});

// Payment tab variables
let file1DataPayments = null;
let file2DataPayments = null;
let processedWorkbookPayments = null;
let validationErrorsPayments = [];

// ===========================================
// PAYMENTS TAB LOGIC
// ===========================================

// Конфігурації банків
const BANK_CONFIGS = {
    'original': {
        name: 'Оригінальний формат (Аксіліум)',
        sheetName: 'загальний',
        pipColumn: 5,      // E - ПІП Позичальника
        ipnColumn: 6,      // F - ІПН Позичальника
        contractColumn: 7  // G - Номер кредитного договору
    },
    'idea': {
        name: 'Ідея Банк',
        sheetName: 'реєстр',
        pipColumn: 13,     // M - ПІБ Боржника
        ipnColumn: 14,     // N - ІПН
        contractColumn: 7  // G - Номер Первинного договору
    },
    'task': {
        name: 'Таскомбанк',
        sheetName: 'sql results',
        pipColumn: 8,      // H - Прізвище Ім'я по-батькові Позичальника
        ipnColumn: 7,      // G - Ідентифікаційний номер Позичальника
        contractColumn: 12 // L - Номер Кредитного договору
    }
};

let selectedBankType = null; // Зберігає обраний тип банку

// Обробник вибору типу банку
document.getElementById('bankTypeSelect').addEventListener('change', function (e) {
    selectedBankType = e.target.value;
    checkReadyToProcessPayments();

    // Якщо файл вже завантажений, але тип банку змінився - треба перевірити файл заново
    // Але для простоти, можна просто скинути файл 1, якщо він є
    if (file1DataPayments) {
        // Можна додати логіку перевалідації, але простіше попросити завантажити наново або просто ігнорувати колізії поки не натиснуть обробку
        // Краще скинути, щоб не було плутанини
        file1DataPayments = null;
        document.getElementById('file1Payments').value = '';
        const label = document.getElementById('label1Payments');
        label.classList.remove('has-file', 'error');
        label.innerHTML = `
                        <div class="file-icon">
                            <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round"
                                    d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                            </svg>
                        </div>
                        <div class="file-details">
                            <div class="file-name">Оберіть основний реєстр</div>
                            <div class="file-info">Excel файл з інформацією про боржників (.xlsx або .xls)</div>
                        </div>
                    `;
        checkReadyToProcessPayments();
    }
});

async function validateFile1Payments(arrayBuffer) {
    const errors = [];

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        if (!workbook.worksheets || workbook.worksheets.length === 0) {
            errors.push('Файл не містить жодного аркуша');
            return errors;
        }

        // ВАЖЛИВО: Валідація тепер залежить від обраного банку
        if (!selectedBankType) {
            errors.push('Спочатку оберіть тип банку');
            return errors;
        }

        const bankConfig = BANK_CONFIGS[selectedBankType];
        const requiredSheet = workbook.worksheets.find(sheet =>
            sheet.name.toLowerCase().trim() === bankConfig.sheetName.toLowerCase()
        );

        if (!requiredSheet) {
            const sheetNames = workbook.worksheets.map(s => s.name).join(', ');
            errors.push(`Відсутній обов'язковий аркуш "${bankConfig.sheetName}". Знайдені аркуші: ${sheetNames}`);
            return errors;
        }

        if (requiredSheet.rowCount === 0) {
            errors.push(`Аркуш "${bankConfig.sheetName}" пустий`);
            return errors;
        }

        // EXTRACT COLUMNS & POPULATE UI
        const columns = extractColumns(requiredSheet, 1);
        populateColumnSelect('colMapFile1Name', columns, ['ПІП', 'ПІБ', 'Позичальник', 'Name', 'Прізвище']);
        populateColumnSelect('colMapFile1Contract', columns, ['Договір', 'Contract', 'Credit', 'Agreement', 'Кредитний']);
        populateColumnSelect('colMapFile1Ipn', columns, ['ІПН', 'IPN', 'Code', 'Код', 'РНОКПП']);

        // Show mapping section
        toggleColumnMapping(true);

        if (requiredSheet.rowCount < 2) {
            errors.push('Відсутні рядки з даними (лише заголовки)');
            return errors;
        }

        // Перевіряємо чи є дані у колонці ПІБ
        let hasValidData = false;
        // Check first 10 rows (skip header)
        for (let i = 2; i <= Math.min(requiredSheet.rowCount, 11); i++) {
            const cell = requiredSheet.getRow(i).getCell(bankConfig.pipColumn);
            if (cell.value && cell.value.toString().trim()) {
                hasValidData = true;
                break;
            }
        }

        if (!hasValidData) {
            errors.push('Не знайдено валідних ПІБ у перших 10 рядках (перевірте правильність формату банку)');
        }

        return errors;
    } catch (error) {
        console.error(error);
        errors.push(`Помилка читання файлу: ${error.message}`);
        return errors;
    }
}

async function validateFile2Payments(arrayBuffer) {
    const errors = [];

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        if (!workbook.worksheets || workbook.worksheets.length === 0) {
            errors.push('Файл не містить жодного аркуша');
            return errors;
        }

        const sheet = workbook.worksheets[0];

        if (sheet.rowCount === 0) {
            errors.push('Файл пустий');
            return errors;
        }

        // EXTRACT COLUMNS & POPULATE UI
        // Try to find a header row or just use preview
        // We'll use row 10 (or 1 if < 10) for preview
        const previewRow = Math.min(sheet.rowCount, 10);
        const columns = extractColumnsWithPreview(sheet, previewRow);

        populateColumnSelect('colMapFile2Name', columns, ['ПІП', 'ПІБ', 'Платник', 'Name', 'Призначення']);
        populateColumnSelect('colMapFile2Amount', columns, ['Сума', 'Amount', 'Credit', 'Кредит']);

        if (sheet.rowCount < 2) {
            errors.push('Недостатньо даних (менше 2 рядків)');
        }

        return errors;
    } catch (error) {
        console.error(error);
        errors.push(`Помилка читання файлу: ${error.message}`);
        return errors;
    }
}
document.getElementById('file1Payments').addEventListener('change', async function (e) {
    const file = e.target.files[0];
    if (file) {
        const label = document.getElementById('label1Payments');

        if (!selectedBankType) {
            alert('Спочатку оберіть тип банку');
            e.target.value = ''; // Очищуємо input
            return;
        }

        label.classList.remove('has-file', 'error');
        label.innerHTML = `
                    <div class="file-icon"><div class="spinner"></div></div>
                    <div class="file-details">
                        <div class="file-name">Обробка...</div>
                        <div class="file-info">Перевірка структури файлу</div>
                    </div>
                `;

        const reader = new FileReader();
        reader.onload = async function (e) {
            let arrayBuffer = e.target.result;
            try {
                // Конвертуємо .xls в .xlsx якщо потрібно
                arrayBuffer = await convertXlsToXlsx(arrayBuffer, file.name);

                const errors = await validateFile1Payments(arrayBuffer);
                validationErrorsPayments = validationErrorsPayments.filter(e => !e.startsWith('Файл 1'));

                if (errors.length > 0) {
                    validationErrorsPayments.push(...errors.map(e => `Файл 1: ${e}`));
                    label.classList.add('error');
                    label.innerHTML = `
                                <div class="file-icon">
                                    <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12" />
                                    </svg>
                                </div>
                                <div class="file-details">
                                    <div class="file-name">${file.name}</div>
                                    <div class="file-info">Валідація не пройдена</div>
                                </div>
                            `;
                    showValidationErrorsPayments(validationErrorsPayments);
                    file1DataPayments = null;
                } else {
                    file1DataPayments = arrayBuffer;

                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);

                    const bankConfig = BANK_CONFIGS[selectedBankType];
                    const zagalniySheet = workbook.worksheets.find(sheet =>
                        sheet.name.toLowerCase().trim() === bankConfig.sheetName.toLowerCase()
                    );

                    const rowCount = zagalniySheet.rowCount - 1;

                    let sheetsInfo = '';
                    if (workbook.worksheets.length > 1) {
                        sheetsInfo = ` • ${workbook.worksheets.length} аркушів`;
                    }

                    label.classList.add('has-file');
                    label.innerHTML = `
                                <div class="file-icon">
                                    <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7" />
                                    </svg>
                                </div>
                                <div class="file-details">
                                    <div class="file-name">${file.name}</div>
                                    <div class="file-info">${rowCount.toLocaleString('uk-UA')} боржників у реєстрі${sheetsInfo}</div>
                                </div>
                            `;

                    if (validationErrorsPayments.length === 0) {
                        hideValidationErrorsPayments();
                    }
                }
            } catch (error) {
                validationErrorsPayments.push(`Файл 1: Помилка читання - ${error.message}`);
                label.classList.add('error');
                label.innerHTML = `
                            <div class="file-icon">
                                <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12" />
                                </svg>
                            </div>
                            <div class="file-details">
                                <div class="file-name">${file.name}</div>
                                <div class="file-info">Неможливо прочитати файл</div>
                            </div>
                        `;
                showValidationErrorsPayments(validationErrorsPayments);
                file1DataPayments = null;
            }

            checkReadyToProcessPayments();
        };
        reader.readAsArrayBuffer(file);
    }
});

document.getElementById('file2Payments').addEventListener('change', async function (e) {
    const file = e.target.files[0];
    if (file) {
        const label = document.getElementById('label2Payments');
        label.classList.remove('has-file', 'error');
        label.innerHTML = `
                    <div class="file-icon"><div class="spinner"></div></div>
                    <div class="file-details">
                        <div class="file-name">Обробка...</div>
                        <div class="file-info">Перевірка структури файлу</div>
                    </div>
                `;

        const reader = new FileReader();
        reader.onload = async function (e) {
            let arrayBuffer = e.target.result;
            try {
                // Конвертуємо .xls в .xlsx якщо потрібно
                arrayBuffer = await convertXlsToXlsx(arrayBuffer, file.name);

                const errors = await validateFile2Payments(arrayBuffer);
                validationErrorsPayments = validationErrorsPayments.filter(e => !e.startsWith('Файл 2'));

                if (errors.length > 0) {
                    validationErrorsPayments.push(...errors.map(e => `Файл 2: ${e}`));
                    label.classList.add('error');
                    label.innerHTML = `
                                <div class="file-icon">
                                    <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12" />
                                    </svg>
                                </div>
                                <div class="file-details">
                                    <div class="file-name">${file.name}</div>
                                    <div class="file-info">Валідація не пройдена</div>
                                </div>
                            `;
                    showValidationErrorsPayments(validationErrorsPayments);
                    file2DataPayments = null;
                } else {
                    file2DataPayments = arrayBuffer;

                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(arrayBuffer);
                    const sheet = workbook.worksheets[0];
                    const paymentCount = sheet.rowCount - 4;

                    label.classList.add('has-file');
                    label.innerHTML = `
                                <div class="file-icon">
                                    <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7" />
                                    </svg>
                                </div>
                                <div class="file-details">
                                    <div class="file-name">${file.name}</div>
                                    <div class="file-info">${paymentCount.toLocaleString('uk-UA')} платіжних транзакцій</div>
                                </div>
                            `;

                    if (validationErrorsPayments.length === 0) {
                        hideValidationErrorsPayments();
                    }
                }
            } catch (error) {
                validationErrorsPayments.push(`Файл 2: Помилка читання - ${error.message}`);
                label.classList.add('error');
                label.innerHTML = `
                            <div class="file-icon">
                                <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12" />
                                </svg>
                            </div>
                            <div class="file-details">
                                <div class="file-name">${file.name}</div>
                                <div class="file-info">Неможливо прочитати файл</div>
                            </div>
                        `;
                showValidationErrorsPayments(validationErrorsPayments);
                file2DataPayments = null;
            }

            checkReadyToProcessPayments();
        };
        reader.readAsArrayBuffer(file);
    }
});

document.getElementById('colMapFile1Name').addEventListener('change', checkReadyToProcessPayments);
document.getElementById('colMapFile1Contract').addEventListener('change', checkReadyToProcessPayments);
document.getElementById('colMapFile1Ipn').addEventListener('change', checkReadyToProcessPayments);
document.getElementById('colMapFile2Name').addEventListener('change', checkReadyToProcessPayments);
document.getElementById('colMapFile2Amount').addEventListener('change', checkReadyToProcessPayments);

document.getElementById('monthSelectPayments').addEventListener('change', checkReadyToProcessPayments);

function checkReadyToProcessPayments() {
    const monthSelected = document.getElementById('monthSelectPayments').value !== '';

    const file1NameCol = document.getElementById('colMapFile1Name').value;
    const file1ContractCol = document.getElementById('colMapFile1Contract').value;
    const file1IpnCol = document.getElementById('colMapFile1Ipn').value;
    const file2NameCol = document.getElementById('colMapFile2Name').value;
    const file2AmountCol = document.getElementById('colMapFile2Amount').value;

    const colsSelected = file1NameCol && file1ContractCol && file1IpnCol && file2NameCol && file2AmountCol;

    const btn = document.getElementById('processBtnPayments');
    btn.disabled = !(file1DataPayments && file2DataPayments && monthSelected && colsSelected && validationErrorsPayments.length === 0);
}
document.getElementById('processBtnPayments').addEventListener('click', processFilesPayments);

async function processFilesPayments() {
    showProgress('Обробка файлів', 'Звірка платежів боржників', [
        { title: 'Підготовка файлів', desc: 'Читання структур даних Excel' },
        { title: 'Аналіз структури', desc: 'Пошук необхідних колонок' },
        { title: 'Збір платежів', desc: 'Обробка даних бухгалтерії' },
        { title: 'Створення індексу', desc: 'Зіставлення з боржниками' },
        { title: 'Підготовка колонки', desc: 'Вставка нової колонки' },
        { title: 'Обробка платежів', desc: 'Зіставлення даних' },
        { title: 'Створення файлу', desc: 'Генерація результату' }
    ]);

    document.getElementById('processBtnPayments').disabled = true;
    hideValidationErrorsPayments();

    try {
        updateProgress(10, 1);
        await delay(500);

        const monthColumn = document.getElementById('monthSelectPayments').value;

        // GET MAPPED COLUMNS
        const f1NameIdx = parseInt(document.getElementById('colMapFile1Name').value);
        const f1ContractIdx = parseInt(document.getElementById('colMapFile1Contract').value);
        // IPN unused for processing logic itself, just for validation/location implicitly

        const f2NameIdx = parseInt(document.getElementById('colMapFile2Name').value);
        const f2AmountIdx = parseInt(document.getElementById('colMapFile2Amount').value);

        const workbook1 = new ExcelJS.Workbook();
        await workbook1.xlsx.load(file1DataPayments);

        const workbook2 = new ExcelJS.Workbook();
        await workbook2.xlsx.load(file2DataPayments);

        updateProgress(20, 2);
        await delay(400);

        const bankConfig = BANK_CONFIGS[selectedBankType];
        const zagalniySheet = workbook1.worksheets.find(sheet =>
            sheet.name.toLowerCase().trim() === bankConfig.sheetName.toLowerCase()
        );

        if (!zagalniySheet) {
            throw new Error(`Аркуш "${bankConfig.sheetName}" не знайдено у файлі боржників`);
        }

        // Header check for naming collisions
        const headerRow = zagalniySheet.getRow(1);
        const headers = [];
        headerRow.eachCell({ includeEmpty: true }, (cell) => {
            headers.push(cell.value);
        });

        // Override file 1 column indices with hardcoded bank config
        // (Since we are using 1-based indices from config directly in logic below)
        // Note: The UI mapping dropdowns are still there but effectively we are enforcing 
        // the known structure for these specific bank formats as requested.
        // However, to be safe and consistent with previous code using `f1NameIdx`, 
        // let's update those variables or usage.

        // We will use the indices from BANK_CONFIGS for File 1 data extraction
        const pipColumnIndex = bankConfig.pipColumn;
        const contractColumnIndex = bankConfig.contractColumn;

        updateProgress(30, 2);
        await delay(300);

        let newColumnName = monthColumn;
        let counter = 2;
        while (headers.includes(newColumnName)) {
            newColumnName = `${monthColumn} (${counter})`;
            counter++;
        }

        // Insert AFTER Name column (or default to E+1 if we want strict position)
        // For consistency with original logic: insert after Name column
        const newColumnIndex = pipColumnIndex + 1;

        updateProgress(40, 3);
        await delay(500);

        // Збираємо платежі з Файлу 2
        // Ключ: ПІБ + Номер договору
        const sheet2 = workbook2.worksheets[0];
        const payments = {}; // {identifier: сума}
        const paymentDetails = {}; // Зберігаємо деталі для звіту

        // Start from row 5 as per "Accounting Report" usually having header clutter
        // Or maybe start closer to 1 if user chose dynamic columns? 
        // Providing a safe buffer (e.g. Row 2) might be better if dynamic. 
        // But let's stick to 5 to avoid junk if top rows are empty/titles.
        for (let i = 5; i <= sheet2.rowCount; i++) {
            const row = sheet2.getRow(i);
            const pipCell = row.getCell(f2NameIdx);
            const amountCell = row.getCell(f2AmountIdx);

            if (pipCell.value && amountCell.value) {
                const cellText = pipCell.value.toString();
                const lines = cellText.split('\n');

                if (lines.length >= 2) {
                    const pipName = lines[0].trim();
                    const contractNumber = extractContractNumber(cellText);

                    if (pipName && contractNumber) {
                        const identifier = createIdentifier(pipName, contractNumber);
                        const amount = parseFloat(amountCell.value);

                        if (!isNaN(amount)) {
                            // Якщо вже є платіж з таким ключем - додаємо до суми
                            payments[identifier] = (payments[identifier] || 0) + amount;

                            // Зберігаємо деталі для звіту
                            if (!paymentDetails[identifier]) {
                                paymentDetails[identifier] = {
                                    name: pipName,
                                    contract: contractNumber,
                                    amount: 0
                                };
                            }
                            paymentDetails[identifier].amount += amount;
                        }
                    }
                }
            }
        }

        const paymentCount = Object.keys(payments).length;
        updateProgress(50, 4);
        await delay(400);

        // Створюємо мапу боржників: {identifier: [rowIndex1, rowIndex2, ...]}
        // Один ідентифікатор може мати кілька рядків (дублікати)
        const debtorsMap = {};

        for (let i = 2; i <= zagalniySheet.rowCount; i++) {
            const row = zagalniySheet.getRow(i);
            const pipCell = row.getCell(pipColumnIndex); // Use config index
            const contractCell = row.getCell(contractColumnIndex); // Use config index

            if (pipCell.value && contractCell.value) {
                const identifier = createIdentifier(pipCell.value, contractCell.value);

                if (!debtorsMap[identifier]) {
                    debtorsMap[identifier] = [];
                }
                debtorsMap[identifier].push(i);
            }
        }

        updateProgress(60, 5);
        await delay(300);

        zagalniySheet.spliceColumns(newColumnIndex, 0, []);

        const newHeaderCell = zagalniySheet.getRow(1).getCell(newColumnIndex);
        newHeaderCell.value = newColumnName;

        const referenceCell = zagalniySheet.getRow(1).getCell(pipColumnIndex);
        if (referenceCell.style) {
            newHeaderCell.style = Object.assign({}, referenceCell.style);
        }

        updateProgress(70, 6);
        await delay(600);

        let matchedCount = 0;
        let totalAmount = 0;
        let processedCount = 0;
        const notFoundPeople = []; // Для тих, кого не знайдено

        for (const [identifier, amount] of Object.entries(payments)) {
            processedCount++;

            if (processedCount % 5 === 0) {
                const progress = 70 + (processedCount / paymentCount) * 20;
                updateProgress(progress, 6);
                await delay(50);
            }

            const rowIndexes = debtorsMap[identifier];

            if (rowIndexes && rowIndexes.length > 0) {
                // Вставляємо суму в усі знайдені рядки (якщо є дублікати)
                for (const rowIndex of rowIndexes) {
                    const row = zagalniySheet.getRow(rowIndex);
                    const cell = row.getCell(newColumnIndex);
                    cell.value = amount;

                    const referenceCell = row.getCell(pipColumnIndex);
                    if (referenceCell.style) {
                        cell.style = Object.assign({}, referenceCell.style);
                    }

                    cell.numFmt = '#,##0.00';

                    matchedCount++;
                    totalAmount += amount;
                }
            } else {
                // Не знайдено - додаємо до списку незіставлених
                const details = paymentDetails[identifier];
                if (details) {
                    notFoundPeople.push({
                        name: details.name,
                        contract: details.contract,
                        amount: details.amount
                    });
                }
            }
        }

        updateProgress(90, 7);
        await delay(500);

        // Unhide all rows/columns in the output
        // zagalniySheet is our target sheet in workbook1
        zagalniySheet.eachRow((row) => {
            row.hidden = false;
        });

        // Also unhide columns if any are hidden
        for (let i = 1; i <= zagalniySheet.columnCount; i++) {
            const col = zagalniySheet.getColumn(i);
            if (col.hidden) {
                col.hidden = false;
            }
        }

        // DEEP CLEAN STRATEGY: Create a brand new workbook and copy ONLY values/styles
        // This completely eliminates any corrupted formula metadata or shared formula ghosts
        const cleanWorkbook = new ExcelJS.Workbook();
        const cleanSheet = cleanWorkbook.addWorksheet(zagalniySheet.name || 'Sheet1');

        // Copy rows and cells manually
        zagalniySheet.eachRow({ includeEmpty: true }, (srcRow, rowIndex) => {
            const destRow = cleanSheet.getRow(rowIndex);

            srcRow.eachCell({ includeEmpty: true }, (srcCell, colNumber) => {
                const destCell = destRow.getCell(colNumber);

                // 1. Copy Value (Resolve formulas to flat values)
                if (srcCell.type === ExcelJS.ValueType.Formula) {
                    // If it has a result, use it. usage of 'result' property covers most cases
                    if (srcCell.result !== undefined && srcCell.result !== null) {
                        destCell.value = srcCell.result;
                    } else if (srcCell.value && typeof srcCell.value === 'object' && srcCell.value.result) {
                        destCell.value = srcCell.value.result;
                    } else {
                        // Last resort: if we can't get a result, we leave it empty to avoid crashing
                        // or try to coerce value if possible
                        destCell.value = null;
                    }
                } else if (srcCell.type === ExcelJS.ValueType.Object) {
                    // Sometimes rich text or hyperlinks. 
                    // For safety in this specific "fix it" mode, let's try to extract text if it's rich text
                    // or just copy value if it looks safe.
                    // For now, simple value copy usually works for non-formulas
                    destCell.value = srcCell.value;
                } else {
                    destCell.value = srcCell.value;
                }

                // 2. Copy Style (Deep copy to detach references)
                if (srcCell.style) {
                    // JSON parse/stringify is slow but ensures no reference leaks
                    destCell.style = JSON.parse(JSON.stringify(srcCell.style));
                }

                // 3. Copy Format
                if (srcCell.numFmt) {
                    destCell.numFmt = srcCell.numFmt;
                }
            });
            destRow.commit();
        });

        // Copy column widths (visuals)
        for (let i = 1; i <= zagalniySheet.columnCount; i++) {
            const srcCol = zagalniySheet.getColumn(i);
            const destCol = cleanSheet.getColumn(i);
            if (srcCol.width) {
                destCol.width = srcCol.width;
            }
        }

        processedWorkbookPayments = cleanWorkbook;

        updateProgress(100, 7);
        await delay(500);

        hideProgress();
        displayResultsPayments(matchedCount, totalAmount, notFoundPeople, newColumnName, paymentCount);

    } catch (error) {
        hideProgress();
        alert('Помилка обробки: ' + error.message);
        console.error(error);
    } finally {
        document.getElementById('processBtnPayments').disabled = false;
    }
}

function displayResultsPayments(matchedCount, totalAmount, notFoundPeople, columnName, totalPayments) {
    const resultsDiv = document.getElementById('resultsPayments');

    let html = `
                <div class="result-item success">
                    <div class="result-title">Обробку завершено</div>
                    <div class="result-details">
                        Створено колонку: <strong>${columnName}</strong><br>
                        Всього платежів у звіті: <strong>${totalPayments}</strong><br>
                        Знайдено збігів: <strong>${matchedCount}</strong><br>
                        Загальна сума знайдених платежів: <strong>${formatNumber(totalAmount)} грн</strong>
                    </div>
                </div>
            `;

    if (notFoundPeople.length > 0) {
        const notFoundAmount = notFoundPeople.reduce((sum, p) => sum + p.amount, 0);
        html += `
                    <div class="result-item warning">
                        <div class="result-title">Незіставлені записи</div>
                        <div class="result-details">
                            Кількість: <strong>${notFoundPeople.length}</strong><br>
                            Загальна сума: <strong>${formatNumber(notFoundAmount)} грн</strong>
                            <div class="not-found-list">
                                ${notFoundPeople.map(p =>
            `<div>
                                        <div>
                                            <strong>${p.name}</strong><br>
                                            <span style="font-size: 12px; color: var(--gray-500);">Договір: ${p.contract}</span>
                                        </div>
                                        <span>${formatNumber(p.amount)} грн</span>
                                    </div>`
        ).join('')}
                            </div>
                        </div>
                    </div>
                `;
    }

    html += `
                <button class="download-btn" onclick="downloadFilePayments()">
                    <span>Завантажити оброблений файл</span>
                </button>
            `;

    resultsDiv.innerHTML = html;
    resultsDiv.classList.add('show');
}

async function downloadFilePayments() {
    try {
        if (!processedWorkbookPayments) {
            alert('Помилка: Немає даних для завантаження. Спробуйте обробити файл ще раз.');
            console.error('processedWorkbookPayments is null');
            return;
        }

        const monthSelect = document.getElementById('monthSelectPayments');
        if (!monthSelect || !monthSelect.value) {
            alert('Помилка: Не обрано місяць.');
            return;
        }

        const monthName = monthSelect.value
            .replace('Оплата ', '')
            .replace(/ /g, '_');
        const fileName = `Реєстр_боржників_${monthName}.xlsx`;

        // Видаляємо зайві дані, якщо є
        // Для Таскомбанку "sql results" може містити щось зайве, але writeBuffer мав би працювати
        const buffer = await processedWorkbookPayments.xlsx.writeBuffer();

        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });

        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    } catch (error) {
        console.error('Download error:', error);
        alert(`Помилка при завантаженні файлу: ${error.message}`);
    }
}

// Ensure global access
window.downloadFilePayments = downloadFilePayments;
