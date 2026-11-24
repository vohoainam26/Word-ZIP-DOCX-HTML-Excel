/**
 * Product Label Checker - Excel Reader Module
 * Parses uploaded Excel workbooks into structured product records with 16 language columns.
 */
(function (global) {
    'use strict';

    const LANGUAGES = [
        'EN', 'FR', 'DE', 'CZ', 'PL', 'HU', 'ES', 'PT',
        'IT', 'RO', 'SK', 'NL', 'DK', 'FI', 'SE', 'EE'
    ];

    const CODE_CANDIDATES = [
        'internal reference', 'default_code', 'product code', 'reference',
        'mã sp', 'mã sản phẩm', 'code', 'sku', 'ref'
    ];

    const NAME_CANDIDATES = [
        'name', 'product name', 'tên sp', 'tên sản phẩm', 'title', 'tên'
    ];

    function findKey(row, candidates) {
        const keys = Object.keys(row);
        for (const candidate of candidates) {
            const match = keys.find(k => k.trim().toLowerCase() === candidate);
            if (match !== undefined) return row[match];
        }
        return '';
    }

    function findLanguageValue(row, lang) {
        const target = lang.toLowerCase();
        const keys = Object.keys(row);
        const match = keys.find(k => k.trim().toLowerCase() === target);
        return match !== undefined ? String(row[match] ?? '') : '';
    }

    function chooseSheet(workbook) {
        const names = workbook.SheetNames || [];
        if (!names.length) {
            throw new Error('Workbook không chứa bất kỳ sheet nào.');
        }

        // Prioritize standard sheet names
        const priorityOrder = ['done', 'products', 'sản phẩm', 'patch', 'sheet1'];
        for (const target of priorityOrder) {
            const found = names.find(n => n.trim().toLowerCase() === target);
            if (found) return found;
        }
        return names[0];
    }

    function parseWorkbookData(arrayBuffer, options = {}) {
        const XLSX = global.XLSX;
        if (!XLSX) {
            throw new Error('Thư viện XLSX chưa sẵn sàng. Vui lòng kiểm tra lại tài nguyên.');
        }

        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = options.sheetName || chooseSheet(workbook);
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
            throw new Error(`Không tìm thấy sheet "${sheetName}".`);
        }

        const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
        if (!rawRows.length) {
            throw new Error(`Sheet "${sheetName}" không có dữ liệu dòng nào.`);
        }

        const products = [];
        let skippedEmpty = 0;

        rawRows.forEach((rawRow, index) => {
            const code = String(findKey(rawRow, CODE_CANDIDATES) || '').trim();
            const name = String(findKey(rawRow, NAME_CANDIDATES) || '').trim();

            if (!code && !name) {
                skippedEmpty++;
                return;
            }

            const languages = {};
            LANGUAGES.forEach(lang => {
                languages[lang] = findLanguageValue(rawRow, lang);
            });

            products.push({
                rowIndex: index + 2, // 1-based header + 1
                code: code || `ROW-${index + 2}`,
                name: name || '(Không có tên sản phẩm)',
                languages
            });
        });

        if (!products.length) {
            throw new Error(`Không tìm thấy cột thông tin sản phẩm (Internal Reference/Name) trong sheet "${sheetName}".`);
        }

        return {
            sheetName,
            availableSheets: workbook.SheetNames,
            totalRows: rawRows.length,
            products,
            skippedEmpty
        };
    }

    function parseFile(file, options = {}) {
        return new Promise((resolve, reject) => {
            if (!file) {
                return reject(new Error('Chưa chọn file Excel.'));
            }

            const reader = new FileReader();
            reader.onload = function (event) {
                try {
                    const result = parseWorkbookData(event.target.result, options);
                    result.fileName = file.name;
                    result.fileSize = file.size;
                    resolve(result);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = function () {
                reject(new Error('Lỗi khi đọc file từ ổ đĩa trình duyệt.'));
            };
            reader.readAsArrayBuffer(file);
        });
    }

    global.ProductLabelExcelReader = {
        LANGUAGES,
        parseWorkbookData,
        parseFile
    };

})(typeof window !== 'undefined' ? window : this);
