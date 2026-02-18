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
