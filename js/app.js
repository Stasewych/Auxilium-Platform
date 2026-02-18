        // AUTHENTICATION
        const CREDENTIALS = {
            username: 'AuxiliumUser',
            password: 'Auxilium2026!'
        };

        document.getElementById('loginForm').addEventListener('submit', function (e) {
            e.preventDefault();

            const username = document.getElementById('username').value;
            const password = document.getElementById('password').value;
            const errorMessage = document.getElementById('loginError');

            if (username === CREDENTIALS.username && password === CREDENTIALS.password) {
                errorMessage.classList.remove('show');
                document.getElementById('loginContainer').style.display = 'none';
                document.getElementById('mainContainer').classList.add('show');
                sessionStorage.setItem('authenticated', 'true');
            } else {
                errorMessage.classList.add('show');
                document.getElementById('username').classList.add('error');
                document.getElementById('password').classList.add('error');

                setTimeout(() => {
                    document.getElementById('username').classList.remove('error');
                    document.getElementById('password').classList.remove('error');
                }, 2000);
            }
        });

        function logout() {
            sessionStorage.removeItem('authenticated');
            document.getElementById('mainContainer').classList.remove('show');
            document.getElementById('loginContainer').style.display = 'block';
            document.getElementById('username').value = '';
            document.getElementById('password').value = '';
            document.getElementById('loginError').classList.remove('show');

            // Reset tabs
            switchTab('payments');
        }

        window.addEventListener('load', function () {
            if (sessionStorage.getItem('authenticated') === 'true') {
                document.getElementById('loginContainer').style.display = 'none';
                document.getElementById('mainContainer').classList.add('show');
            }
        });

        // TAB SWITCHING
        function switchTab(tabName) {
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.remove('active');
            });

            // Remove active class from all tab buttons
            document.querySelectorAll('.tab').forEach(btn => {
                btn.classList.remove('active');
            });

            // Show selected tab
            if (tabName === 'payments') {
                document.getElementById('paymentsTab').classList.add('active');
                document.querySelector('.tab:nth-child(1)').classList.add('active');
            } else if (tabName === 'iban') {
                document.getElementById('ibanTab').classList.add('active');
                document.querySelector('.tab:nth-child(2)').classList.add('active');
            }
        }

        // PROGRESS FUNCTIONS (shared)
        function showProgress(title, subtitle, steps) {
            document.getElementById('progressTitle').textContent = title;
            document.getElementById('progressSubtitle').textContent = subtitle;

            const stepsContainer = document.getElementById('progressSteps');
            stepsContainer.innerHTML = steps.map((step, index) => `
                <div class="progress-step" id="progressStep${index + 1}">
                    <div class="progress-step-number">${index + 1}</div>
                    <div class="progress-step-text">
                        <div class="progress-step-title">${step.title}</div>
                        <div class="progress-step-desc">${step.desc}</div>
                    </div>
                </div>
            `).join('');

            document.getElementById('progressOverlay').classList.add('show');
            updateProgress(0, 1);
        }

        function hideProgress() {
            document.getElementById('progressOverlay').classList.remove('show');
            document.getElementById('progressPercentage').textContent = '0%';
            document.getElementById('progressBarFill').style.width = '0%';
        }

        function updateProgress(percentage, currentStep) {
            document.getElementById('progressPercentage').textContent = Math.round(percentage) + '%';
            document.getElementById('progressBarFill').style.width = percentage + '%';

            const steps = document.querySelectorAll('.progress-step');
            steps.forEach((step, index) => {
                const stepNum = index + 1;
                if (stepNum < currentStep) {
                    step.classList.remove('active');
                    step.classList.add('completed');
                } else if (stepNum === currentStep) {
                    step.classList.remove('completed');
                    step.classList.add('active');
                } else {
                    step.classList.remove('active', 'completed');
                }
            });
        }

        function delay(ms) {
            return new Promise(resolve => setTimeout(resolve, ms));
        }

        // ===========================================
        // PAYMENTS TAB LOGIC
        // ===========================================

        // Функція конвертації .xls в .xlsx для роботи з ExcelJS
        async function convertXlsToXlsx(arrayBuffer, fileName) {
            // Перевіряємо чи це .xls файл
            const isXls = fileName.toLowerCase().endsWith('.xls');

            if (!isXls) {
                // Якщо це вже .xlsx - повертаємо як є
                return arrayBuffer;
            }

            try {
                // Get selected encoding
                const encoding = parseInt(document.getElementById('encodingSelect').value) || 1251;

                // Використовуємо SheetJS для читання .xls
                const data = new Uint8Array(arrayBuffer);
                // Use selected encoding
                const workbook = XLSX.read(data, { type: 'array', codepage: encoding });

                // Конвертуємо в .xlsx формат
                const xlsxBuffer = XLSX.write(workbook, {
                    bookType: 'xlsx',
                    type: 'array'
                });

                return xlsxBuffer;
            } catch (error) {
                throw new Error(`Помилка конвертації .xls файлу: ${error.message}`);
            }
        }

        // Initialize months for payments
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

        // Helper: Extract columns from worksheet
        function extractColumns(worksheet, headerRowIndex = 1) {
            const columns = [];
            const row = worksheet.getRow(headerRowIndex);

            // If row is empty, generate generic columns based on worksheet.columnCount
            if (row.cellCount === 0) {
                const colCount = Math.max(worksheet.columnCount, 26);
                for (let i = 1; i <= colCount; i++) {
                    const letter = worksheet.getColumn(i).letter || getColumnLetter(i);
                    columns.push({
                        index: i,
                        letter: letter,
                        value: `Column ${letter}`
                    });
                }
            } else {
                row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                    columns.push({
                        index: colNumber,
                        letter: worksheet.getColumn(colNumber).letter,
                        value: cell.value ? cell.value.toString() : `(Empty)`
                    });
                });
            }
            return columns;
        }

        // Helper: Extract columns with preview (for File 2)
        function extractColumnsWithPreview(worksheet, previewRowIndex = 10) {
            const columns = [];
            // Scan first 20 columns max (or more if needed)
            const maxCols = Math.max(worksheet.columnCount, 30);
            const previewRow = worksheet.getRow(previewRowIndex);

            for (let i = 1; i <= maxCols; i++) {
                const letter = worksheet.getColumn(i).letter || getColumnLetter(i);

                // Try to find a header-like value in first 10 rows? 
                // No, just show preview.
                const previewVal = previewRow.getCell(i).value;
                const previewStr = previewVal ? String(previewVal).substring(0, 15) : '';

                columns.push({
                    index: i,
                    letter: letter,
                    value: `Колонка ${letter} ${previewStr ? `(${previewStr}...)` : ''}`
                });
            }
            return columns;
        }

        // Helper: Get column letter from index
        function getColumnLetter(colIndex) {
            let temp, letter = '';
            while (colIndex > 0) {
                temp = (colIndex - 1) % 26;
                letter = String.fromCharCode(temp + 65) + letter;
                colIndex = (colIndex - temp - 1) / 2;
            }
            return letter;
        }

        // Helper: Populate Select Key
        function populateColumnSelect(selectId, columns, keywords = []) {
            const select = document.getElementById(selectId);
            const currentVal = select.value;
            select.innerHTML = '<option value="">Оберіть колонку...</option>';

            let autoMatch = null;

            columns.forEach(col => {
                const option = document.createElement('option');
                option.value = col.index;
                option.textContent = `${col.letter}: ${col.value}`;
                select.appendChild(option);

                if (!autoMatch && keywords.length > 0) {
                    const text = col.value.toLowerCase();
                    if (keywords.some(k => text.includes(k.toLowerCase()))) {
                        autoMatch = col.index;
                    }
                }
            });

            if (autoMatch) {
                select.value = autoMatch;
            }
        }

        function toggleColumnMapping(show) {
            const div = document.getElementById('columnMappingSection');
            if (show) div.classList.remove('hidden');
            else div.classList.add('hidden');
        }

        // Payment validation functions
        async function validateFile1Payments(arrayBuffer) {
            const errors = [];

            try {
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(arrayBuffer);

                if (!workbook.worksheets || workbook.worksheets.length === 0) {
                    errors.push('Файл не містить жодного аркуша');
                    return errors;
                }

                const zagalniySheet = workbook.worksheets.find(sheet =>
                    sheet.name.toLowerCase().trim() === 'загальний'
                );

                if (!zagalniySheet) {
                    const sheetNames = workbook.worksheets.map(s => s.name).join(', ');
                    errors.push(`Відсутній обов'язковий аркуш "Загальний". Знайдені аркуші: ${sheetNames}`);
                    return errors;
                }

                if (zagalniySheet.rowCount === 0) {
                    errors.push('Аркуш "Загальний" пустий');
                    return errors;
                }

                // EXTRACT COLUMNS & POPULATE UI
                const columns = extractColumns(zagalniySheet, 1);
                populateColumnSelect('colMapFile1Name', columns, ['ПІП', 'ПІБ', 'Позичальник', 'Name', 'Прізвище']);
                populateColumnSelect('colMapFile1Contract', columns, ['Договір', 'Contract', 'Credit', 'Agreement', 'Кредитний']);
                populateColumnSelect('colMapFile1Ipn', columns, ['ІПН', 'IPN', 'Code', 'Код', 'РНОКПП']);

                // Show mapping section
                toggleColumnMapping(true);

                if (zagalniySheet.rowCount < 2) {
                    errors.push('Відсутні рядки з даними (лише заголовки)');
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

        function showValidationErrorsPayments(errors) {
            const errorDiv = document.getElementById('validationErrorsPayments');
            const errorList = document.getElementById('errorListPayments');

            errorList.innerHTML = '';
            errors.forEach(error => {
                const li = document.createElement('li');
                li.textContent = error;
                errorList.appendChild(li);
            });

            errorDiv.classList.add('show');
        }

        function hideValidationErrorsPayments() {
            document.getElementById('validationErrorsPayments').classList.remove('show');
        }

        // File upload handlers for payments
        document.getElementById('file1Payments').addEventListener('change', async function (e) {
            const file = e.target.files[0];
            if (file) {
                const label = document.getElementById('label1Payments');
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

                            const zagalniySheet = workbook.worksheets.find(sheet =>
                                sheet.name.toLowerCase().trim() === 'загальний'
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

        // PAYMENTS PROCESSING
        function normalizeName(name) {
            if (!name || typeof name !== 'string') return '';
            name = name.split('\n')[0].trim();
            name = name.replace(/\s+/g, ' ');
            return name.toLowerCase().trim();
        }

        function normalizeContract(contract) {
            if (!contract || typeof contract !== 'string') return '';
            // Видаляємо зайві пробіли та переводимо в lower case
            return contract.trim().toLowerCase();
        }

        function extractContractNumber(text) {
            // Витягуємо номер договору до слова "від"
            // Приклад: "CL-214019 від 16.08.2019" -> "CL-214019"
            if (!text || typeof text !== 'string') return '';

            const lines = text.split('\n');
            if (lines.length < 2) return '';

            const contractLine = lines[1].trim();
            const parts = contractLine.split(' від ');
            if (parts.length > 0) {
                return parts[0].trim();
            }

            return contractLine;
        }

        function createIdentifier(name, contract) {
            // Створюємо унікальний ключ: normalize(ПІБ) + "|" + normalize(договір)
            const normName = normalizeName(name);
            const normContract = normalizeContract(contract);
            return normName + '|' + normContract;
        }

        function formatNumber(num) {
            if (num === null || num === undefined || num === '') return '';
            const number = parseFloat(num);
            if (isNaN(number)) return '';
            return number.toLocaleString('uk-UA', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
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

                const zagalniySheet = workbook1.worksheets.find(sheet =>
                    sheet.name.toLowerCase().trim() === 'загальний'
                );

                if (!zagalniySheet) {
                    throw new Error('Аркуш "Загальний" не знайдено у файлі боржників');
                }

                // Header check for naming collisions
                const headerRow = zagalniySheet.getRow(1);
                const headers = [];
                headerRow.eachCell({ includeEmpty: true }, (cell) => {
                    headers.push(cell.value);
                });

                updateProgress(30, 2);
                await delay(300);

                let newColumnName = monthColumn;
                let counter = 2;
                while (headers.includes(newColumnName)) {
                    newColumnName = `${monthColumn} (${counter})`;
                    counter++;
                }

                // Insert AFTER Name column
                const newColumnIndex = f1NameIdx + 1;

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
                    const pipCell = row.getCell(f1NameIdx);
                    const contractCell = row.getCell(f1ContractIdx);

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

                const referenceCell = zagalniySheet.getRow(1).getCell(f1NameIdx);
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

                            const referenceCell = row.getCell(f1NameIdx);
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

                processedWorkbookPayments = workbook1;

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
            if (processedWorkbookPayments) {
                const monthName = document.getElementById('monthSelectPayments').value
                    .replace('Оплата ', '')
                    .replace(/ /g, '_');
                const fileName = `Реєстр_боржників_${monthName}.xlsx`;

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
            }
        }

        // ===========================================
        // IBAN TAB LOGIC
        // ===========================================
        let file1DataIban = null;
        let file2DataIban = null;
        let file1SelectedColumnIban = null;
        let file2SelectedColumnIban = null;
        let resultWorkbookIban = null;

        document.getElementById('file1Iban').addEventListener('change', function (e) {
            handleFileUploadIban(e.target.files[0], 1);
        });

        document.getElementById('file2Iban').addEventListener('change', function (e) {
            handleFileUploadIban(e.target.files[0], 2);
        });

        document.getElementById('file1ColumnIban').addEventListener('change', function (e) {
            handleColumnSelectionIban(1, e.target.value);
        });

        document.getElementById('file2ColumnIban').addEventListener('change', function (e) {
            handleColumnSelectionIban(2, e.target.value);
        });

        document.getElementById('generateBtnIban').addEventListener('click', generateResultsIban);
        document.getElementById('downloadBtnIban').addEventListener('click', downloadResultsIban);

        function handleFileUploadIban(file, fileNumber) {
            if (!file) return;

            const label = document.getElementById(`label${fileNumber}Iban`);
            const select = document.getElementById(`file${fileNumber}ColumnIban`);

            const reader = new FileReader();

            reader.onload = function (e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                    const fileData = {
                        name: file.name,
                        workbook: workbook,
                        sheetName: firstSheetName,
                        allRows: jsonData
                    };

                    if (fileNumber === 1) {
                        file1DataIban = fileData;
                        file1SelectedColumnIban = null;
                    } else {
                        file2DataIban = fileData;
                        file2SelectedColumnIban = null;
                    }

                    label.classList.add('has-file');
                    const rowCount = jsonData.length - 1;
                    label.innerHTML = `
                        <div class="file-icon">
                            <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7" />
                            </svg>
                        </div>
                        <div class="file-details">
                            <div class="file-name">${file.name}</div>
                            <div class="file-info">${rowCount.toLocaleString('uk-UA')} записів завантажено</div>
                        </div>
                    `;

                    const maxCols = Math.max(...jsonData.map(row => row.length));
                    select.innerHTML = '<option value="">Оберіть колонку з IBAN</option>';
                    for (let i = 0; i < maxCols; i++) {
                        const colLetter = String.fromCharCode(65 + i);
                        const firstValue = jsonData.length > 1 && jsonData[1][i] ? jsonData[1][i] : '';
                        select.innerHTML += `<option value="${i}">${colLetter} - ${firstValue}</option>`;
                    }
                    select.classList.remove('hidden');

                    checkIfReadyToGenerateIban();

                } catch (error) {
                    alert(`Помилка читання файлу: ${error.message}`);
                }
            };

            reader.readAsArrayBuffer(file);
        }

        function handleColumnSelectionIban(fileNumber, columnIndex) {
            if (columnIndex === '') {
                if (fileNumber === 1) {
                    file1SelectedColumnIban = null;
                } else {
                    file2SelectedColumnIban = null;
                }
                checkIfReadyToGenerateIban();
                return;
            }

            const colIndex = parseInt(columnIndex);

            if (fileNumber === 1) {
                file1SelectedColumnIban = colIndex;
            } else {
                file2SelectedColumnIban = colIndex;
            }

            checkIfReadyToGenerateIban();
        }

        function checkIfReadyToGenerateIban() {
            const canGenerate = file1DataIban && file2DataIban &&
                file1SelectedColumnIban !== null &&
                file2SelectedColumnIban !== null;

            document.getElementById('generateBtnIban').disabled = !canGenerate;

            if (canGenerate) {
                updateStatsIban();
            }
        }

        function updateStatsIban() {
            if (!file1DataIban || !file2DataIban ||
                file1SelectedColumnIban === null || file2SelectedColumnIban === null) return;

            const values1 = [];
            file1DataIban.allRows.forEach((row, idx) => {
                if (idx > 0 && row[file1SelectedColumnIban]) {
                    values1.push(String(row[file1SelectedColumnIban]).trim());
                }
            });

            const values2 = [];
            file2DataIban.allRows.forEach((row, idx) => {
                if (idx > 0 && row[file2SelectedColumnIban]) {
                    values2.push(String(row[file2SelectedColumnIban]).trim());
                }
            });

            const set2 = new Set(values2);
            const uniqueIBANs1 = [...new Set(values1)];
            const excludedCount = uniqueIBANs1.filter(x => set2.has(x)).length;
            const resultCount = uniqueIBANs1.filter(x => !set2.has(x)).length;

            document.getElementById('statFile1Iban').textContent = values1.length.toLocaleString('uk-UA');
            document.getElementById('statFile2Iban').textContent = values2.length.toLocaleString('uk-UA');
            document.getElementById('statExcludedIban').textContent = excludedCount.toLocaleString('uk-UA');
            document.getElementById('statResultIban').textContent = resultCount.toLocaleString('uk-UA');
        }

        async function generateResultsIban() {
            showProgress('Обробка файлів', 'Фільтрація IBAN номерів', [
                { title: 'Читання файлів', desc: 'Завантаження даних з Excel' },
                { title: 'Аналіз IBAN', desc: 'Збір зарплатних рахунків' },
                { title: 'Фільтрація', desc: 'Виключення рахунків' },
                { title: 'Створення файлу', desc: 'Генерація результату' }
            ]);

            try {
                updateProgress(10, 1);
                await delay(300);

                updateProgress(25, 1);
                await delay(300);

                updateProgress(30, 2);
                await delay(300);

                const values2Set = new Set();
                file2DataIban.allRows.forEach((row, idx) => {
                    if (idx > 0 && row[file2SelectedColumnIban]) {
                        values2Set.add(String(row[file2SelectedColumnIban]).trim());
                    }
                });

                updateProgress(50, 2);
                await delay(400);

                updateProgress(55, 3);
                await delay(300);

                const resultRows = [file1DataIban.allRows[0]];
                file1DataIban.allRows.forEach((row, idx) => {
                    if (idx > 0) {
                        const value = row[file1SelectedColumnIban] ? String(row[file1SelectedColumnIban]).trim() : '';
                        if (value && !values2Set.has(value)) {
                            resultRows.push(row);
                        }
                    }
                });

                updateProgress(75, 3);
                await delay(400);

                updateProgress(80, 4);
                await delay(300);

                const wb = XLSX.utils.book_new();

                const ws1 = XLSX.utils.aoa_to_sheet(file1DataIban.allRows);
                formatIBANColumn(ws1, file1SelectedColumnIban, file1DataIban.allRows.length);
                XLSX.utils.book_append_sheet(wb, ws1, 'Всі рахунки');

                updateProgress(85, 4);
                await delay(200);

                const ws2 = XLSX.utils.aoa_to_sheet(file2DataIban.allRows);
                formatIBANColumn(ws2, file2SelectedColumnIban, file2DataIban.allRows.length);
                XLSX.utils.book_append_sheet(wb, ws2, 'Зарплатні');

                updateProgress(90, 4);
                await delay(200);

                const resultRowsRenamed = resultRows.map((row, idx) => {
                    if (idx === 0) {
                        const newRow = [...row];
                        if (newRow.length > 0) newRow[0] = 'asvp_id';
                        if (newRow.length > 1) newRow[1] = 'iban';
                        return newRow;
                    }
                    return row;
                });

                const ws3 = XLSX.utils.aoa_to_sheet(resultRowsRenamed);
                formatIBANColumn(ws3, file1SelectedColumnIban, resultRowsRenamed.length);
                XLSX.utils.book_append_sheet(wb, ws3, 'Результат');

                resultWorkbookIban = wb;

                updateProgress(100, 4);
                await delay(500);

                hideProgress();

                document.getElementById('successMessageIban').classList.add('show');
                document.getElementById('downloadBtnIban').classList.remove('hidden');

            } catch (error) {
                hideProgress();
                alert(`Помилка: ${error.message}`);
            }
        }

        function formatIBANColumn(worksheet, columnIndex, rowCount) {
            const colLetter = XLSX.utils.encode_col(columnIndex);

            for (let row = 0; row < rowCount; row++) {
                const cellAddress = colLetter + (row + 1);
                const cell = worksheet[cellAddress];

                if (cell && cell.v) {
                    const cellValue = String(cell.v).trim();
                    cell.t = 's';
                    cell.v = cellValue;
                    cell.z = '@';
                    delete cell.w;
                }
            }
        }

        function downloadResultsIban() {
            if (!resultWorkbookIban) return;

            try {
                const wbout = XLSX.write(resultWorkbookIban, {
                    bookType: 'xlsx',
                    type: 'array',
                    cellDates: false,
                    bookSST: false
                });

                const blob = new Blob([wbout], { type: 'application/octet-stream' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'IBAN_результат.xlsx';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);

            } catch (error) {
                alert(`Помилка завантаження: ${error.message}`);
            }
        }
