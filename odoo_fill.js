(function () {
    'use strict';

    const REQUIRED_LANGUAGES = ['EN', 'FR', 'DE', 'CZ', 'PL', 'HU', 'ES', 'PT', 'IT', 'RO', 'SK', 'NL', 'DK', 'FI', 'SE'];
    const OPTIONAL_LANGUAGES = ['EE'];
    const ALL_OUTPUT_LANGUAGES = [...REQUIRED_LANGUAGES, ...OPTIONAL_LANGUAGES];
    const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    const HANDLE_DATABASE = 'word-zip-odoo-working-file';
    const HANDLE_STORE = 'handles';
    const HANDLE_KEY = 'odoo-workbook';

    const state = {
        workbookHandle: null,
        workbookFileName: '',
        workbookFingerprint: '',
        workbookBuffer: null,
        workbookSheetName: '',
        headerMap: new Map(),
        odooRows: [],
        groups: [],
        errors: [],
        ignoredCount: 0,
        processed: false,
        processing: false,
        saving: false,
        previewGroupId: null,
        previewLanguage: '',
        previewMode: 'visual',
        previewRenderToken: 0
    };

    const el = {};

    function clean(value) {
        return String(value == null ? '' : value).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    }

    function escapeHtml(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function normalizeName(value) {
        return clean(value)
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toLowerCase()
            .replace(/&/g, ' and ')
            .replace(/\b(?:specification|label|labels)\b/g, ' ')
            .replace(/[^a-z0-9]+/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function normalizeReference(value) {
        return clean(value).toLowerCase().replace(/[^a-z0-9-]/g, '');
    }

    function fileStem(value) {
        return clean(String(value || '').split('/').pop().split('\\').pop()).replace(/\.docx$/i, '');
    }

    function companyMode() {
        return document.querySelector('input[name="odoo-company-mode"]:checked')?.value || 'importer';
    }

    function languageOrder(language) {
        const index = ALL_OUTPUT_LANGUAGES.indexOf(language);
        return index < 0 ? 999 : index;
    }

    function openHandleDatabase() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(HANDLE_DATABASE, 1);
            request.onupgradeneeded = () => {
                if (!request.result.objectStoreNames.contains(HANDLE_STORE)) request.result.createObjectStore(HANDLE_STORE);
            };
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error || new Error('Không mở được bộ nhớ file của trình duyệt.'));
        });
    }

    async function saveWorkingHandle(handle) {
        try {
            const database = await openHandleDatabase();
            await new Promise((resolve, reject) => {
                const transaction = database.transaction(HANDLE_STORE, 'readwrite');
                transaction.objectStore(HANDLE_STORE).put(handle, HANDLE_KEY);
                transaction.oncomplete = resolve;
                transaction.onerror = () => reject(transaction.error);
            });
            database.close();
        } catch (_error) {
            // Vẫn dùng được file trong phiên hiện tại nếu trình duyệt chặn IndexedDB.
        }
    }

    async function getSavedWorkingHandle() {
        try {
            const database = await openHandleDatabase();
            const handle = await new Promise((resolve, reject) => {
                const transaction = database.transaction(HANDLE_STORE, 'readonly');
                const request = transaction.objectStore(HANDLE_STORE).get(HANDLE_KEY);
                request.onsuccess = () => resolve(request.result || null);
                request.onerror = () => reject(request.error);
            });
            database.close();
            return handle;
        } catch (_error) {
            return null;
        }
    }

    async function hasReadWritePermission(handle, requestAccess) {
        if (!handle) return false;
        const options = { mode: 'readwrite' };
        if (await handle.queryPermission(options) === 'granted') return true;
        if (!requestAccess) return false;
        return (await handle.requestPermission(options)) === 'granted';
    }

    function fileFingerprint(file) {
        return `${file.name}|${file.size}|${file.lastModified}`;
    }

    function findWorkbookSheet(workbook) {
        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: true, range: 0 });
            if (!rows.length) continue;
            const headings = rows[0].map(value => clean(value).toUpperCase());
            if (headings.includes('NAME') && headings.includes('INTERNAL REFERENCE')) {
                return { sheetName, worksheet, rows };
            }
        }
        throw new Error('Không tìm thấy sheet có đủ cột Name và Internal Reference.');
    }

    function readWorkbookFile(file, buffer, preserveResults) {
        const workbook = XLSX.read(buffer, { type: 'array', cellStyles: true, cellDates: true });
        const selected = findWorkbookSheet(workbook);
        const headers = selected.rows[0].map(value => clean(value));
        const headerMap = new Map();
        headers.forEach((heading, index) => headerMap.set(heading.toUpperCase(), index));
        const missingColumns = REQUIRED_LANGUAGES.filter(language => !headerMap.has(language));
        if (missingColumns.length) throw new Error(`Excel thiếu các cột ngôn ngữ: ${missingColumns.join(', ')}.`);

        const nameColumn = headerMap.get('NAME');
        const referenceColumn = headerMap.get('INTERNAL REFERENCE');
        const categoryColumn = headerMap.get('PRODUCT CATEGORY');
        const rows = selected.rows.slice(1).map((row, index) => {
            const name = clean(row[nameColumn]);
            const reference = clean(row[referenceColumn]);
            const category = categoryColumn == null ? '' : clean(row[categoryColumn]);
            const existingHtml = Object.create(null);
            for (const language of ALL_OUTPUT_LANGUAGES) {
                const column = headerMap.get(language);
                if (column != null && clean(row[column])) existingHtml[language] = String(row[column]);
            }
            return {
                rowNumber: index + 2,
                name,
                reference,
                category,
                normalizedName: normalizeName(name),
                normalizedReference: normalizeReference(reference),
                existingHtml
            };
        }).filter(row => row.name || row.reference);

        state.workbookFileName = file.name;
        state.workbookFingerprint = fileFingerprint(file);
        state.workbookBuffer = buffer.slice(0);
        state.workbookSheetName = selected.sheetName;
        state.headerMap = headerMap;
        state.odooRows = rows;
        const filledCells = rows.reduce((total, row) => total + REQUIRED_LANGUAGES.filter(language => row.existingHtml[language]).length, 0);
        el.workbookMeta.textContent = `${file.name} — ${rows.length} sản phẩm, ${filledCells} ô ngôn ngữ đã có. Mọi lượt duyệt sẽ ghi trực tiếp vào file này.`;
        el.workbookButton.textContent = 'Đổi file Odoo đang làm việc';
        if (!preserveResults) clearResults();
    }

    async function loadWorkbookFromHandle(handle, requestAccess, preserveResults) {
        try {
            if (!await hasReadWritePermission(handle, requestAccess)) {
                state.workbookHandle = handle;
                el.workbookMeta.textContent = `Đã nhớ file ${handle.name || 'Odoo'}, nhưng cần bạn bấm nút để cấp lại quyền đọc/ghi.`;
                el.workbookButton.textContent = 'Mở lại file Odoo đã chọn';
                return false;
            }
            const file = await handle.getFile();
            const buffer = await file.arrayBuffer();
            state.workbookHandle = handle;
            readWorkbookFile(file, buffer, preserveResults);
            await saveWorkingHandle(handle);
            return true;
        } catch (error) {
            el.workbookMeta.textContent = `Lỗi: ${error.message || error}`;
            return false;
        }
    }

    async function chooseWorkingWorkbook() {
        if (!window.showOpenFilePicker) {
            alert('Trình duyệt này không hỗ trợ ghi trực tiếp vào file. Hãy mở tool bằng Chrome hoặc Edge trên HTTPS/Vercel.');
            return;
        }
        el.workbookButton.disabled = true;
        try {
            if (state.workbookHandle && !state.workbookBuffer && await loadWorkbookFromHandle(state.workbookHandle, true, false)) return;
            const handles = await window.showOpenFilePicker({
                multiple: false,
                types: [{
                    description: 'Excel Odoo (.xlsx)',
                    accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
                }]
            });
            if (handles[0]) await loadWorkbookFromHandle(handles[0], true, false);
        } catch (error) {
            if (error?.name !== 'AbortError') el.workbookMeta.textContent = `Lỗi: ${error.message || error}`;
        } finally {
            el.workbookButton.disabled = false;
        }
    }

    async function restoreWorkingWorkbook() {
        if (!window.showOpenFilePicker) {
            el.workbookMeta.textContent = 'Hãy dùng Chrome hoặc Edge để chọn và ghi trực tiếp vào cùng file Odoo.';
            return;
        }
        const handle = await getSavedWorkingHandle();
        if (!handle) return;
        state.workbookHandle = handle;
        await loadWorkbookFromHandle(handle, false, false);
    }

    async function collectDescriptors(files) {
        const descriptors = [];
        const errors = [];
        for (const sourceFile of files) {
            if (/\.docx$/i.test(sourceFile.name)) {
                descriptors.push({
                    file: sourceFile,
                    fileName: sourceFile.name,
                    sourceName: sourceFile.name,
                    sourcePath: '',
                    folderInfo: null,
                    sourceTimestamp: Number(sourceFile.lastModified || 0)
                });
                continue;
            }
            if (!/\.zip$/i.test(sourceFile.name)) continue;
            try {
                const zip = await JSZip.loadAsync(await sourceFile.arrayBuffer());
                const entries = Object.values(zip.files).filter(entry => {
                    return !entry.dir && /\.docx$/i.test(entry.name) && !/(?:^|\/)__MACOSX(?:\/|$)/i.test(entry.name) && !/(?:^|\/)~\$/i.test(entry.name);
                });
                for (const entry of entries) {
                    descriptors.push({
                        entry,
                        file: null,
                        fileName: entry.name.split('/').pop(),
                        sourceName: sourceFile.name,
                        sourcePath: entry.name,
                        folderInfo: window.batchWordTool.detectFolderLanguage(entry.name),
                        sourceTimestamp: entry.date instanceof Date ? entry.date.getTime() : 0
                    });
                }
                if (!entries.length) errors.push({
                    fileName: sourceFile.name,
                    sourceName: sourceFile.name,
                    inputPath: sourceFile.name,
                    detail: 'ZIP không có file DOCX hợp lệ.'
                });
            } catch (error) {
                errors.push({
                    fileName: sourceFile.name,
                    sourceName: sourceFile.name,
                    inputPath: sourceFile.name,
                    detail: `Không đọc được ZIP: ${error.message || error}`
                });
            }
        }
        return { descriptors, errors };
    }

    async function descriptorFile(descriptor) {
        if (descriptor.file) return descriptor.file;
        const blob = await descriptor.entry.async('blob');
        return new File([blob], descriptor.fileName, {
            type: DOCX_MIME,
            lastModified: Number(descriptor.sourceTimestamp || Date.now())
        });
    }

    function folderMonthTimestamp(sourcePath) {
        const parts = String(sourcePath || '').replace(/\\/g, '/').split('/');
        if (parts.length) parts.pop();
        const monthNames = {
            january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
            july: 7, august: 8, september: 9, october: 10, november: 11, december: 12
        };
        let newest = 0;
        for (let index = 0; index < parts.length; index += 1) {
            const part = clean(parts[index]);
            let match = part.match(/(?:^|\D)(0?[1-9]|1[0-2])\s*[_\-.]\s*(20\d{2})(?:\D|$)/);
            if (match) newest = Math.max(newest, Date.UTC(Number(match[2]), Number(match[1]) - 1, 1));
            match = part.match(/(?:^|\D)(20\d{2})\s*[_\-.]\s*(0?[1-9]|1[0-2])(?:\D|$)/);
            if (match) newest = Math.max(newest, Date.UTC(Number(match[1]), Number(match[2]) - 1, 1));
            if (/^(0?[1-9]|1[0-2])$/.test(part) && /^(20\d{2})(?:\D|$)/.test(clean(parts[index + 1]))) {
                newest = Math.max(newest, Date.UTC(Number(clean(parts[index + 1]).match(/^(20\d{2})/)[1]), Number(part) - 1, 1));
            }
            const monthNameMatch = part.match(/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(20\d{2})\b/i);
            if (monthNameMatch) {
                newest = Math.max(newest, Date.UTC(Number(monthNameMatch[2]), monthNames[monthNameMatch[1].toLowerCase()] - 1, 1));
            }
        }
        return newest;
    }

    function descriptorVersionTimestamp(descriptor) {
        return folderMonthTimestamp(descriptor.sourcePath) || Number(descriptor.sourceTimestamp || 0);
    }

    function preferCzSevenByFive(descriptors) {
        const meta = descriptors.map(descriptor => {
            const identity = summaryIdentity(descriptor.fileName);
            if (!identity) return { descriptor, key: '', rank: 0 };
            const language = descriptor.folderInfo?.language || (identity.languages.length === 1 ? identity.languages[0] : '');
            if (language !== 'CZ') return { descriptor, key: '', rank: 0 };
            let rank = Number(descriptor.folderInfo?.czRank || 0);
            if (!rank && /7\s*[x×]\s*5/i.test(descriptor.fileName)) rank = 2;
            if (!rank && /5\s*[x×]\s*3/i.test(descriptor.fileName)) rank = 1;
            return {
                descriptor,
                key: `${identity.productCode}::${normalizeName(identity.productName)}`,
                rank
            };
        });
        const maxRank = new Map();
        meta.forEach(item => {
            if (item.key) maxRank.set(item.key, Math.max(maxRank.get(item.key) || 0, item.rank));
        });
        return meta
            .filter(item => !item.key || item.rank >= (maxRank.get(item.key) || 0))
            .map(item => item.descriptor);
    }

    function summaryIdentity(fileName) {
        try {
            return window.batchWordTool.parseFileIdentity(fileName);
        } catch (_error) {
            return null;
        }
    }

    function groupKey(name, code) {
        const normalized = normalizeName(name);
        const reference = normalizeReference(code);
        if (normalized && reference) return `${reference}::${normalized}`;
        return normalized || `code:${reference}`;
    }

    function ensureGroup(map, name, code, matchedRow, matchMode, matchDetail) {
        const key = matchedRow ? `odoo:${matchedRow.rowNumber}` : groupKey(name, code);
        if (!map.has(key)) {
            map.set(key, {
                id: `odoo-group-${map.size + 1}`,
                key,
                name: clean(name),
                code: clean(code),
                cells: Object.create(null),
                cellMeta: Object.create(null),
                sourceFiles: new Set(),
                warnings: [],
                selectedRowNumber: matchedRow ? String(matchedRow.rowNumber) : '',
                selectedRow: matchedRow || null,
                matchCandidates: matchedRow ? [matchedRow] : [],
                matchMode: matchMode || 'none',
                matchDetail: matchDetail || 'Chưa ghép được sản phẩm Odoo.',
                approved: false,
                saved: false
            });
        }
        const group = map.get(key);
        if (!group.name && name) group.name = clean(name);
        if (!group.code && code) group.code = clean(code);
        return group;
    }

    function addLanguage(group, language, html, meta) {
        const normalizedLanguage = clean(language).toUpperCase();
        if (!ALL_OUTPUT_LANGUAGES.includes(normalizedLanguage) || !clean(html)) return false;
        const existing = group.cellMeta[normalizedLanguage];
        const incomingRank = Number(meta?.rank || 0);
        const existingRank = Number(existing?.rank || 0);
        const incomingVersion = Number(meta?.versionTimestamp || 0);
        const existingVersion = Number(existing?.versionTimestamp || 0);
        if (existing && existingRank > incomingRank) return false;
        if (existing && existingRank === incomingRank && existingVersion > incomingVersion) return false;
        if (existing && existingRank === incomingRank && existingVersion === incomingVersion) {
            group.warnings.push(`${normalizedLanguage} bị lặp cùng ngày; giữ kết quả đầu tiên từ ${existing.fileName}.`);
            return false;
        }
        group.cells[normalizedLanguage] = html;
        group.cellMeta[normalizedLanguage] = { ...meta, rank: incomingRank, versionTimestamp: incomingVersion };
        return true;
    }

    async function convertSummary(descriptor, identity) {
        const file = await descriptorFile(descriptor);
        const folderLanguage = descriptor.folderInfo?.language;
        const identityLanguage = identity.languages.length > 1 ? 'ES > PT' : identity.languages[0];
        const outputLanguage = folderLanguage || identityLanguage;
        if (outputLanguage === 'ES > PT') return null;
        const item = {
            file,
            entry: null,
            fileName: descriptor.fileName,
            outputLanguage,
            languages: [],
            productCode: identity.productCode,
            productName: identity.productName
        };
        const html = await window.batchWordTool.convertSummaryItem(item);
        return {
            kind: 'summary',
            name: item.productName,
            code: item.productCode,
            languages: { [outputLanguage]: html },
            rank: outputLanguage === 'CZ' ? Number(descriptor.folderInfo?.czRank || 0) : 0
        };
    }

    async function convertImport(descriptor, index, inspection = null, preparedFile = null, fallbackCandidates = []) {
        const file = preparedFile || await descriptorFile(descriptor);
        let result;
        try {
            result = await window.importLabelsTool.convertFile(file, index, companyMode(), inspection);
        } catch (error) {
            if (!/Không tìm thấy NTF dùng chung/i.test(error.message || '') || !fallbackCandidates.length) throw error;
            const nutritionFallbacks = [];
            let lastError = error;
            for (const candidate of fallbackCandidates) {
                if (candidate.preflight.identity) continue;
                try {
                    const fallbackFile = candidate.preflight.preparedFile || await descriptorFile(candidate.descriptor);
                    const fallbackInspection = candidate.preflight.inspection || await window.importLabelsTool.inspectFile(fallbackFile);
                    nutritionFallbacks.push({
                        inspection: fallbackInspection,
                        fileName: candidate.descriptor.fileName,
                        sourcePath: candidate.descriptor.sourcePath
                    });
                    result = await window.importLabelsTool.convertFile(
                        file,
                        index,
                        companyMode(),
                        inspection,
                        nutritionFallbacks
                    );
                    break;
                } catch (fallbackError) {
                    lastError = fallbackError;
                }
            }
            if (!result) throw lastError;
        }
        return {
            kind: 'import',
            name: result.name || fileStem(descriptor.fileName),
            code: result.defaultCode || '',
            languages: result.htmlByLanguage,
            warnings: result.warnings || [],
            nutritionFallbackPath: result.nutritionFallbackPath || '',
            rank: 0
        };
    }

    function missingLanguages(group) {
        return REQUIRED_LANGUAGES.filter(language => !effectiveHtml(group, language));
    }

    function effectiveHtml(group, language) {
        const existing = group.selectedRow?.existingHtml?.[language];
        return existing || group.cells[language] || '';
    }

    function writableLanguages(group) {
        if (!group.selectedRow) return [];
        return ALL_OUTPUT_LANGUAGES.filter(language => group.cells[language] && !group.selectedRow.existingHtml[language]);
    }

    function conflictLanguages(group) {
        if (!group.selectedRow) return [];
        return ALL_OUTPUT_LANGUAGES.filter(language => {
            const existing = group.selectedRow.existingHtml[language];
            const incoming = group.cells[language];
            return existing && incoming && clean(existing) !== clean(incoming);
        });
    }

    function levenshtein(a, b) {
        if (!a) return b.length;
        if (!b) return a.length;
        const previous = Array.from({ length: b.length + 1 }, (_, index) => index);
        for (let row = 1; row <= a.length; row++) {
            let diagonal = previous[0];
            previous[0] = row;
            for (let column = 1; column <= b.length; column++) {
                const saved = previous[column];
                previous[column] = Math.min(
                    previous[column] + 1,
                    previous[column - 1] + 1,
                    diagonal + (a[row - 1] === b[column - 1] ? 0 : 1)
                );
                diagonal = saved;
            }
        }
        return previous[b.length];
    }

    function similarity(a, b) {
        if (!a || !b) return 0;
        return 1 - levenshtein(a, b) / Math.max(a.length, b.length, 1);
    }

    function sourceIncludesExactReference(source, reference) {
        const value = clean(source).toLowerCase();
        const token = clean(reference).toLowerCase();
        if (!value || !token) return false;
        const escaped = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        return new RegExp(`(?:^|[^a-z0-9])${escaped}(?=$|[^a-z0-9])`, 'i').test(value);
    }

    function matchOdooProduct(name, code, sourcePath) {
        const normalizedCode = normalizeReference(code);
        if (normalizedCode) {
            const referenceMatches = state.odooRows.filter(row => row.normalizedReference && row.normalizedReference === normalizedCode);
            if (referenceMatches.length === 1) {
                return { row: referenceMatches[0], mode: 'reference', detail: `Khớp chính xác Internal Reference: ${referenceMatches[0].reference}.` };
            }
        }

        const normalizedName = normalizeName(name);
        if (!normalizedName) return null;
        const exactMatches = state.odooRows.filter(row => row.normalizedName === normalizedName);
        if (exactMatches.length === 1) {
            return { row: exactMatches[0], mode: 'name', detail: 'Đã tìm thấy chính xác tên sản phẩm trong Odoo.' };
        }

        const suggestions = state.odooRows.map(row => ({ row, score: similarity(normalizedName, row.normalizedName) }))
            .sort((a, b) => b.score - a.score);
        const best = suggestions[0];
        const second = suggestions[1];
        if (best?.score >= 0.90 && (!second || best.score - second.score >= 0.01)) {
            return {
                row: best.row,
                mode: 'suggestion',
                detail: `Tên gần khớp Odoo ${Math.round(best.score * 100)}% — hãy kiểm tra trước khi duyệt.`
            };
        }

        if (sourcePath) {
            const sourceReferenceMatches = state.odooRows.filter(row =>
                row.normalizedReference && sourceIncludesExactReference(sourcePath, row.reference)
            );
            if (sourceReferenceMatches.length === 1) {
                return { row: sourceReferenceMatches[0], mode: 'reference', detail: `Tìm thấy Internal Reference ${sourceReferenceMatches[0].reference} trong đường dẫn file.` };
            }
        }
        return null;
    }

    function importNameFromFile(fileName) {
        const stem = fileStem(fileName);
        return clean(stem
            .replace(/^specification\s+\d+(?:-\d+)?[\s_-]+/i, '')
            .replace(/^\d+(?:-\d+)?[\s_-]+/i, ''));
    }

    function isAuxiliaryInstructionDescriptor(descriptor) {
        const parts = String(descriptor.sourcePath || descriptor.fileName || '')
            .split(/[>\\/]+/)
            .map(part => clean(part).replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').toLowerCase());
        return parts.some(part => /\bs\s+navod\s+labels?\b/.test(part));
    }

    function preflightLanguageKey(descriptor, preflight) {
        const languages = [];
        const folderLanguage = clean(descriptor.folderInfo?.language).toUpperCase();
        if (folderLanguage) languages.push(folderLanguage);
        if (!languages.length && preflight.identity?.languages?.length) languages.push(...preflight.identity.languages);
        return [...new Set(languages)].sort((left, right) => languageOrder(left) - languageOrder(right)).join('+') || 'MULTI';
    }

    function preferLatestPrepared(prepared) {
        const byProductLanguage = new Map();
        prepared.forEach(item => {
            const key = `${item.preflight.row.rowNumber}::${preflightLanguageKey(item.descriptor, item.preflight)}`;
            if (!byProductLanguage.has(key)) byProductLanguage.set(key, []);
            byProductLanguage.get(key).push({
                ...item,
                versionTimestamp: descriptorVersionTimestamp(item.descriptor),
                sourceTimestamp: Number(item.descriptor.sourceTimestamp || 0)
            });
        });
        const selected = [];
        byProductLanguage.forEach(items => {
            items.sort((left, right) =>
                right.versionTimestamp - left.versionTimestamp ||
                right.sourceTimestamp - left.sourceTimestamp ||
                right.index - left.index
            );
            selected.push({ ...items[0], fallbackCandidates: items.slice(1) });
        });
        return selected.sort((left, right) => left.index - right.index);
    }

    async function preflightDescriptor(descriptor) {
        // Supplier-instruction notices such as "PALDO_S NAVOD labels" only say
        // to use the existing warehouse instructions. They are not product
        // labels and intentionally contain no language, Ingredients or NTF.
        if (isAuxiliaryInstructionDescriptor(descriptor)) return null;
        const identity = summaryIdentity(descriptor.fileName);
        const name = identity ? identity.productName : importNameFromFile(descriptor.fileName);
        const code = identity ? identity.productCode : (fileStem(descriptor.fileName).match(/(?:^|[\s_-])(\d+(?:-\d+)?)(?=[\s_-]|$)/)?.[1] || '');
        const match = matchOdooProduct(name, code, descriptor.sourcePath || descriptor.fileName);
        if (match) return { identity, name, code, ...match };
        if (identity) return null;
        try {
            const preparedFile = await descriptorFile(descriptor);
            const inspection = await window.importLabelsTool.inspectFile(preparedFile);
            const titleMatch = matchOdooProduct(inspection.name, inspection.defaultCode, descriptor.sourcePath || descriptor.fileName);
            return titleMatch ? {
                identity: null,
                name: inspection.name,
                code: inspection.defaultCode,
                inspection,
                preparedFile,
                ...titleMatch
            } : null;
        } catch (_error) {
            return null;
        }
    }

    function sourceContainsReference(group, row) {
        if (!row.normalizedReference) return false;
        if (normalizeReference(group.code) === row.normalizedReference) return true;
        return [...group.sourceFiles].some(source => sourceIncludesExactReference(source, row.reference));
    }

    function autoMatchGroup(group) {
        const exactReference = state.odooRows.filter(row => sourceContainsReference(group, row));
        if (exactReference.length === 1) {
            group.matchCandidates = exactReference;
            setGroupMatch(group, exactReference[0], 'reference', `Khớp chính xác Internal Reference: ${exactReference[0].reference}.`);
            return;
        }
        const normalizedName = normalizeName(group.name);
        const exactName = state.odooRows.filter(row => row.normalizedName && row.normalizedName === normalizedName);
        if (exactName.length === 1) {
            group.matchCandidates = exactName;
            setGroupMatch(group, exactName[0], 'name', 'Khớp chính xác theo tên sản phẩm. Hãy kiểm tra rồi tick duyệt.');
            return;
        }
        const suggestions = state.odooRows.map(row => ({ row, score: similarity(normalizedName, row.normalizedName) }))
            .sort((a, b) => b.score - a.score);
        group.matchCandidates = suggestions.slice(0, 12).map(item => item.row);
        if (suggestions[0]?.score >= 0.72) {
            const percent = Math.round(suggestions[0].score * 100);
            setGroupMatch(group, suggestions[0].row, 'suggestion', `Gợi ý gần đúng ${percent}% — cần kiểm tra và tick duyệt.`);
            return;
        }
        setGroupMatch(group, null, 'none', 'Không tìm thấy sản phẩm phù hợp trong toàn bộ file Odoo.');
    }

    function setGroupMatch(group, row, mode, detail) {
        group.selectedRow = row || null;
        group.selectedRowNumber = row ? String(row.rowNumber) : '';
        group.matchMode = mode;
        group.matchDetail = detail;
        group.approved = false;
        group.saved = false;
    }

    async function processSources() {
        if (state.processing) return;
        if (!state.workbookBuffer) {
            alert('Hãy chọn file Odoo đang làm việc trước.');
            return;
        }
        const files = [...el.sourceFiles.files];
        if (!files.length) {
            alert('Hãy chọn ít nhất một file ZIP hoặc DOCX.');
            return;
        }

        state.processing = true;
        state.processed = false;
        state.groups = [];
        state.errors = [];
        state.ignoredCount = 0;
        el.processButton.disabled = true;
        el.processButton.textContent = 'Đang đọc dữ liệu…';
        renderResults();

        try {
            const collected = await collectDescriptors(files);
            state.errors.push(...collected.errors);
            const descriptors = preferCzSevenByFive(collected.descriptors);
            const prepared = [];
            for (let index = 0; index < descriptors.length; index += 1) {
                const descriptor = descriptors[index];
                el.processButton.textContent = `Đang kiểm tra ${index + 1}/${descriptors.length}…`;
                await new Promise(resolve => setTimeout(resolve, 0));
                const preflight = await preflightDescriptor(descriptor);
                if (!preflight) {
                    state.ignoredCount += 1;
                    continue;
                }
                prepared.push({ descriptor, preflight, index });
            }
            const selected = preferLatestPrepared(prepared);
            state.ignoredCount += prepared.length - selected.length;
            const groups = new Map();
            for (let selectedIndex = 0; selectedIndex < selected.length; selectedIndex += 1) {
                const { descriptor, preflight, index, fallbackCandidates } = selected[selectedIndex];
                el.processButton.textContent = `Đang xử lý ${selectedIndex + 1}/${selected.length}…`;
                await new Promise(resolve => setTimeout(resolve, 0));
                try {
                    const converted = preflight.identity
                        ? await convertSummary(descriptor, preflight.identity)
                        : await convertImport(
                            descriptor,
                            index + 1,
                            preflight.inspection,
                            preflight.preparedFile,
                            fallbackCandidates
                        );
                    if (!converted) continue;
                    const group = ensureGroup(
                        groups,
                        converted.name,
                        converted.code,
                        preflight.row,
                        preflight.mode,
                        preflight.detail
                    );
                    let acceptedLanguage = false;
                    for (const [language, html] of Object.entries(converted.languages || {})) {
                        acceptedLanguage = addLanguage(group, language, html, {
                            fileName: descriptor.fileName,
                            sourcePath: descriptor.sourcePath,
                            kind: converted.kind,
                            rank: converted.rank,
                            versionTimestamp: descriptorVersionTimestamp(descriptor)
                        }) || acceptedLanguage;
                    }
                    if (acceptedLanguage) group.sourceFiles.add(descriptor.sourcePath || descriptor.fileName);
                    if (acceptedLanguage && converted.nutritionFallbackPath) {
                        group.sourceFiles.add(`NTF từ bản cũ: ${converted.nutritionFallbackPath}`);
                    }
                    if (converted.warnings?.length) group.warnings.push(...converted.warnings);
                } catch (error) {
                    state.errors.push({
                        fileName: descriptor.fileName,
                        sourceName: descriptor.sourceName,
                        sourcePath: descriptor.sourcePath,
                        inputPath: descriptor.sourcePath || descriptor.fileName,
                        productName: preflight.row.name,
                        productCode: preflight.row.reference,
                        detail: error.message || String(error)
                    });
                }
            }
            state.groups = [...groups.values()].sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
            state.processed = true;
        } finally {
            state.processing = false;
            el.processButton.disabled = false;
            el.processButton.textContent = 'Kiểm tra Odoo & chuyển HTML';
            renderResults();
        }
    }

    function matchOptions(group) {
        const rows = [...group.matchCandidates];
        if (group.selectedRow && !rows.some(row => row.rowNumber === group.selectedRow.rowNumber)) rows.unshift(group.selectedRow);
        const options = ['<option value="">— Chưa ghép —</option>'];
        for (const row of rows) {
            const selected = String(row.rowNumber) === group.selectedRowNumber ? ' selected' : '';
            const reference = row.reference ? ` — ${row.reference}` : '';
            options.push(`<option value="${row.rowNumber}"${selected}>${escapeHtml(row.name)}${escapeHtml(reference)}</option>`);
        }
        return options.join('');
    }

    function groupStatus(group) {
        const missing = missingLanguages(group);
        if (!group.selectedRow) return { className: 'error', text: 'Chưa ghép Odoo', detail: group.matchDetail };
        if (group.saved) return { className: 'success', text: 'Đã lưu', detail: `Đã ghi trực tiếp vào ${state.workbookFileName}.` };
        const writable = writableLanguages(group);
        const conflicts = conflictLanguages(group);
        if (!writable.length && conflicts.length) return { className: 'error', text: 'Chỉ có xung đột', detail: `Giữ dữ liệu cũ tại: ${conflicts.join(', ')}.` };
        if (!writable.length) return { className: 'success', text: 'Không có ô mới', detail: 'Các ngôn ngữ trong ZIP đã có sẵn trong Odoo.' };
        if (conflicts.length) return { className: 'processing', text: 'Có xung đột', detail: `Sẽ thêm ${writable.length} ô trống; giữ dữ liệu cũ tại ${conflicts.join(', ')}.` };
        return {
            className: missing.length ? 'processing' : 'success',
            text: missing.length ? `Sẵn sàng lưu ${writable.length} ngôn ngữ` : 'Sẵn sàng — đủ 15',
            detail: missing.length ? `Sau khi lưu còn thiếu: ${missing.join(', ')}.` : 'Sau khi gộp sẽ đủ 15 ngôn ngữ.'
        };
    }

    function renderResults() {
        el.resultsBody.innerHTML = '';
        if (!state.groups.length && !state.errors.length) {
            const emptyMessage = state.processing
                ? 'Đang kiểm tra danh sách Odoo…'
                : state.processed
                    ? 'Không có sản phẩm nào trong ZIP khớp với file Odoo.'
                    : 'Chưa có dữ liệu để kiểm tra.';
            el.resultsBody.innerHTML = `<tr><td colspan="7" class="empty-state">${emptyMessage}</td></tr>`;
            updateSummary();
            return;
        }

        for (const group of state.groups) {
            const missing = missingLanguages(group);
            const available = REQUIRED_LANGUAGES.length - missing.length;
            const writable = writableLanguages(group);
            const conflicts = conflictLanguages(group);
            const status = groupStatus(group);
            const row = document.createElement('tr');
            row.innerHTML = `
                <td class="odoo-approve"><input type="checkbox" ${group.approved ? 'checked' : ''} ${(!group.selectedRow || !writable.length || group.saved || state.saving) ? 'disabled' : ''} aria-label="Duyệt ${escapeHtml(group.name)}"></td>
                <td><div class="file-name">${escapeHtml(group.name || 'Không có tên')}</div><div class="source-name">${escapeHtml([...group.sourceFiles].join(' • '))}</div></td>
                <td>${escapeHtml(group.code || '—')}</td>
                <td><select class="odoo-match-select">${matchOptions(group)}</select><div class="odoo-match-note">${escapeHtml(group.matchDetail)}</div></td>
                <td><span class="coverage-badge ${missing.length ? 'incomplete' : ''}">${available}/15</span><div class="odoo-match-note">${writable.length ? `+${writable.length} ô mới` : 'Không có ô mới'}</div>${missing.length ? `<div class="odoo-missing">${escapeHtml(missing.join(', '))}</div>` : ''}${conflicts.length ? `<div class="odoo-missing">Xung đột: ${escapeHtml(conflicts.join(', '))}</div>` : ''}</td>
                <td><span class="status-badge status-${status.className}">${escapeHtml(status.text)}</span><div class="odoo-match-note">${escapeHtml(status.detail)}</div></td>
                <td class="action-cell"><button class="small-btn" type="button">Xem kết quả</button></td>`;

            const checkbox = row.querySelector('input[type="checkbox"]');
            checkbox.addEventListener('change', () => {
                group.approved = checkbox.checked;
                updateSummary();
            });
            const select = row.querySelector('select');
            select.addEventListener('change', () => {
                const selectedRow = state.odooRows.find(candidate => String(candidate.rowNumber) === select.value) || null;
                setGroupMatch(
                    group,
                    selectedRow,
                    selectedRow ? 'manual' : 'none',
                    selectedRow ? 'Bạn đã chọn thủ công — hãy kiểm tra rồi tick duyệt.' : 'Chưa ghép được sản phẩm Odoo.'
                );
                renderResults();
            });
            row.querySelector('button').addEventListener('click', () => openPreview(group.id));
            el.resultsBody.appendChild(row);
        }

        for (const error of state.errors) {
            const row = document.createElement('tr');
            row.className = 'row-error';
            const inputLocation = error.sourcePath
                ? `${error.sourceName || 'ZIP'} > ${error.sourcePath}`
                : (error.inputPath || error.fileName);
            row.innerHTML = `<td>—</td><td><div class="file-name">${escapeHtml(error.productName || error.fileName)}</div><div class="source-name">Vị trí input: ${escapeHtml(inputLocation)}</div></td><td>${escapeHtml(error.productCode || '—')}</td><td>${error.productName ? 'Đã tìm thấy trong Odoo' : '—'}</td><td>—</td><td><span class="status-badge status-error">! Lỗi chuyển đổi</span><div class="odoo-match-note">${escapeHtml(error.detail)}</div></td><td>—</td>`;
            el.resultsBody.appendChild(row);
        }
        updateSummary();
    }

    function updateSummary() {
        const complete = state.groups.filter(group => missingLanguages(group).length === 0).length;
        const eligibleGroups = state.groups.filter(group => group.selectedRow && writableLanguages(group).length > 0 && !group.saved);
        const approvedGroups = state.groups.filter(group => group.approved && group.selectedRow && writableLanguages(group).length > 0);
        const approvedCells = approvedGroups.reduce((total, group) => total + writableLanguages(group).length, 0);
        const review = state.groups.filter(group => !group.saved && (!group.selectedRow || writableLanguages(group).length === 0 || !group.approved)).length;
        el.countTotal.textContent = state.groups.length;
        el.countReady.textContent = complete;
        el.countReview.textContent = review;
        el.countError.textContent = state.errors.length;
        el.reviewPanel.classList.toggle('visible', state.processed);
        el.reviewMessage.textContent = state.processed
            ? `${complete}/${state.groups.length} sản phẩm đủ 15 sau khi gộp; đã duyệt ${approvedGroups.length} sản phẩm với ${approvedCells} ô HTML mới.`
            : 'Có thể duyệt và lưu từng phần; dữ liệu đã lưu sẽ được dùng cho lượt tiếp theo.';
        el.fillButton.disabled = !state.processed || approvedGroups.length === 0 || state.saving;
        el.fillButton.textContent = state.saving ? 'Đang lưu trực tiếp…' : `Duyệt & Lưu vào file Odoo (${approvedGroups.length})`;
        const allApproved = eligibleGroups.length > 0 && eligibleGroups.every(group => group.approved);
        el.approveAllButton.disabled = !state.processed || eligibleGroups.length === 0 || state.saving;
        el.approveAllButton.textContent = allApproved ? `Bỏ duyệt tất cả (${eligibleGroups.length})` : `Duyệt tất cả (${eligibleGroups.length})`;
        el.errorLogButton.disabled = state.errors.length === 0;
        el.errorLogButton.textContent = `Tải log lỗi (${state.errors.length})`;
    }

    function approveAll() {
        const eligibleGroups = state.groups.filter(group => group.selectedRow && writableLanguages(group).length > 0 && !group.saved);
        if (!eligibleGroups.length) return;
        const shouldApprove = !eligibleGroups.every(group => group.approved);
        eligibleGroups.forEach(group => { group.approved = shouldApprove; });
        renderResults();
    }

    function exportErrorLog() {
        if (!state.errors.length) return;
        const data = [[
            '#',
            'Odoo Product',
            'Internal Reference',
            'Input Source',
            'DOCX File',
            'Location in Input',
            'Error Reason'
        ]];
        state.errors.forEach((error, index) => {
            const inputLocation = error.sourcePath
                ? `${error.sourceName || 'ZIP'} > ${error.sourcePath}`
                : (error.inputPath || error.fileName);
            data.push([
                index + 1,
                error.productName || '',
                error.productCode || '',
                error.sourceName || error.fileName || '',
                error.fileName || '',
                inputLocation,
                error.detail || ''
            ]);
        });
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        worksheet['!cols'] = [
            { wch: 6 }, { wch: 48 }, { wch: 22 }, { wch: 34 },
            { wch: 48 }, { wch: 95 }, { wch: 75 }
        ];
        worksheet['!autofilter'] = { ref: `A1:G${data.length}` };
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Errors');
        const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, '').replace('T', '_');
        XLSX.writeFile(workbook, `odoo_error_log_${stamp}.xlsx`, { compression: true });
    }

    function sanitizePreviewHtml(html) {
        const template = document.createElement('template');
        template.innerHTML = html || '';
        template.content.querySelectorAll('script,style,iframe,object,embed,link,meta').forEach(node => node.remove());
        template.content.querySelectorAll('*').forEach(node => {
            for (const attribute of [...node.attributes]) {
                const name = attribute.name.toLowerCase();
                if (name.startsWith('on') || name === 'src' || name === 'href' || name === 'srcdoc') node.removeAttribute(attribute.name);
                if (name === 'style' && !/^\s*font-weight\s*:\s*400\s*;?\s*$/i.test(attribute.value)) node.removeAttribute(attribute.name);
            }
        });
        return template.innerHTML;
    }

    function previewGroup() {
        return state.groups.find(group => group.id === state.previewGroupId) || null;
    }

    function openPreview(groupId) {
        const group = state.groups.find(candidate => candidate.id === groupId);
        if (!group) return;
        const languages = ALL_OUTPUT_LANGUAGES.filter(language => effectiveHtml(group, language));
        if (!languages.length) return;
        state.previewGroupId = groupId;
        state.previewLanguage = languages[0];
        el.previewTitle.textContent = `ODOO FILL — ${group.name}`;
        el.previewLanguage.innerHTML = languages.map(language => `<option value="${language}">${language}</option>`).join('');
        el.previewLanguage.value = state.previewLanguage;
        el.previewBackdrop.classList.add('visible');
        setPreviewMode('visual');
        updatePreview();
    }

    function closePreview() {
        state.previewRenderToken += 1;
        state.previewGroupId = null;
        el.previewBackdrop.classList.remove('visible');
    }

    function setPreviewMode(mode) {
        state.previewMode = mode;
        const visual = mode === 'visual';
        el.previewVisualPanel.hidden = !visual;
        el.previewCode.hidden = visual;
        el.previewVisualTab.classList.toggle('active', visual);
        el.previewCodeTab.classList.toggle('active', !visual);
        el.previewVisualTab.setAttribute('aria-selected', String(visual));
        el.previewCodeTab.setAttribute('aria-selected', String(!visual));
        if (visual) renderPreviewImage();
    }

    function updatePreview() {
        const group = previewGroup();
        if (!group) return;
        state.previewLanguage = el.previewLanguage.value;
        const html = effectiveHtml(group, state.previewLanguage);
        el.previewCode.textContent = html;
        el.previewRendered.innerHTML = sanitizePreviewHtml(html);
        if (state.previewMode === 'visual') renderPreviewImage();
    }

    async function renderPreviewImage() {
        const group = previewGroup();
        if (!group || !el.previewBackdrop.classList.contains('visible')) return;
        const token = ++state.previewRenderToken;
        el.previewImage.hidden = true;
        el.previewRendered.hidden = false;
        el.previewStatus.hidden = false;
        el.previewStatus.textContent = 'Đang tạo ảnh từ HTML…';
        try {
            await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            const canvas = await html2canvas(el.previewRendered, { backgroundColor: '#ffffff', scale: 1.5, logging: false, useCORS: true });
            if (token !== state.previewRenderToken) return;
            el.previewImage.src = canvas.toDataURL('image/png');
            el.previewImage.hidden = false;
            el.previewRendered.hidden = true;
            el.previewStatus.hidden = true;
        } catch (_error) {
            if (token !== state.previewRenderToken) return;
            el.previewStatus.textContent = 'Không tạo được ảnh; đang hiển thị HTML trực tiếp.';
            el.previewRendered.hidden = false;
        }
    }

    async function copyPreview() {
        const text = el.previewCode.textContent || '';
        if (!text) return;
        try {
            await navigator.clipboard.writeText(text);
        } catch (_error) {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            textarea.remove();
        }
        el.copyPreview.textContent = 'Đã copy';
        setTimeout(() => { el.copyPreview.textContent = 'Copy HTML'; }, 1200);
    }

    function rebindMatchedRows() {
        for (const group of state.groups) {
            group.selectedRow = state.odooRows.find(row => String(row.rowNumber) === group.selectedRowNumber) || null;
        }
    }

    async function fillWorkbook() {
        if (state.saving) return;
        const approved = state.groups.filter(group => group.approved && group.selectedRow && writableLanguages(group).length > 0);
        if (!approved.length || !state.workbookBuffer || !state.workbookHandle) return;
        state.saving = true;
        renderResults();
        try {
            if (!await hasReadWritePermission(state.workbookHandle, true)) throw new Error('Bạn chưa cấp quyền ghi vào file Odoo.');
            const currentFile = await state.workbookHandle.getFile();
            if (fileFingerprint(currentFile) !== state.workbookFingerprint) {
                await loadWorkbookFromHandle(state.workbookHandle, false, true);
                rebindMatchedRows();
                state.groups.forEach(group => { group.approved = false; });
                alert('File Odoo đã thay đổi bên ngoài tool. Tool đã đọc lại bản mới nhất; hãy kiểm tra và tick duyệt lại để tránh ghi đè nhầm.');
                return;
            }

            const workbook = XLSX.read(state.workbookBuffer.slice(0), { type: 'array', cellStyles: true, cellDates: true });
            const worksheet = workbook.Sheets[state.workbookSheetName];
            if (!worksheet) throw new Error('Không tìm thấy sheet Odoo ban đầu. Hãy mở lại file Excel.');
            let writtenCells = 0;
            for (const group of approved) {
                for (const language of writableLanguages(group)) {
                    const html = group.cells[language];
                    const column = state.headerMap.get(language);
                    if (!html || column == null) continue;
                    const address = XLSX.utils.encode_cell({ r: group.selectedRow.rowNumber - 1, c: column });
                    if (clean(worksheet[address]?.v)) continue;
                    worksheet[address] = { t: 's', v: html };
                    writtenCells += 1;
                }
            }
            if (!writtenCells) throw new Error('Không còn ô trống nào để lưu. Dữ liệu cũ trong Odoo được giữ nguyên.');

            const output = XLSX.write(workbook, { type: 'array', bookType: 'xlsx', compression: true, cellStyles: true });
            const writable = await state.workbookHandle.createWritable();
            await writable.write(output);
            await writable.close();

            const updatedFile = await state.workbookHandle.getFile();
            const updatedBuffer = await updatedFile.arrayBuffer();
            readWorkbookFile(updatedFile, updatedBuffer, true);
            rebindMatchedRows();
            for (const group of approved) {
                group.approved = false;
                group.saved = true;
            }
            el.workbookMeta.textContent = `${updatedFile.name} — vừa lưu trực tiếp ${writtenCells} ô HTML cho ${approved.length} sản phẩm. File này sẵn sàng cho lượt Daily/Import tiếp theo.`;
        } catch (error) {
            alert(`Không lưu được vào file Odoo: ${error.message || error}`);
        } finally {
            state.saving = false;
            renderResults();
        }
    }

    function clearResults() {
        state.groups = [];
        state.errors = [];
        state.ignoredCount = 0;
        state.processed = false;
        state.processing = false;
        renderResults();
    }

    function resetAll() {
        el.sourceFiles.value = '';
        clearResults();
    }

    function init() {
        Object.assign(el, {
            workbookButton: document.getElementById('odoo-workbook-button'),
            workbookMeta: document.getElementById('odoo-workbook-meta'),
            sourceFiles: document.getElementById('odoo-source-files'),
            processButton: document.getElementById('odoo-process-button'),
            resetButton: document.getElementById('odoo-reset-button'),
            resultsBody: document.getElementById('odoo-results-body'),
            countTotal: document.getElementById('odoo-count-total'),
            countReady: document.getElementById('odoo-count-ready'),
            countReview: document.getElementById('odoo-count-review'),
            countError: document.getElementById('odoo-count-error'),
            reviewPanel: document.getElementById('odoo-review-panel'),
            reviewMessage: document.getElementById('odoo-review-message'),
            errorLogButton: document.getElementById('odoo-error-log-button'),
            approveAllButton: document.getElementById('odoo-approve-all-button'),
            fillButton: document.getElementById('odoo-fill-button'),
            previewBackdrop: document.getElementById('odoo-preview-backdrop'),
            previewTitle: document.getElementById('odoo-preview-title'),
            previewLanguage: document.getElementById('odoo-preview-language'),
            previewCode: document.getElementById('odoo-preview-code'),
            previewVisualPanel: document.getElementById('odoo-preview-visual-panel'),
            previewRendered: document.getElementById('odoo-preview-rendered'),
            previewImage: document.getElementById('odoo-preview-image'),
            previewStatus: document.getElementById('odoo-preview-status'),
            previewVisualTab: document.getElementById('odoo-preview-visual-tab'),
            previewCodeTab: document.getElementById('odoo-preview-code-tab'),
            closePreview: document.getElementById('odoo-close-preview'),
            copyPreview: document.getElementById('odoo-copy-preview')
        });

        el.workbookButton.addEventListener('click', chooseWorkingWorkbook);
        el.processButton.addEventListener('click', processSources);
        el.resetButton.addEventListener('click', resetAll);
        el.errorLogButton.addEventListener('click', exportErrorLog);
        el.approveAllButton.addEventListener('click', approveAll);
        el.fillButton.addEventListener('click', fillWorkbook);
        el.previewLanguage.addEventListener('change', updatePreview);
        el.previewVisualTab.addEventListener('click', () => setPreviewMode('visual'));
        el.previewCodeTab.addEventListener('click', () => setPreviewMode('code'));
        el.closePreview.addEventListener('click', closePreview);
        el.previewBackdrop.addEventListener('click', event => {
            if (event.target === el.previewBackdrop) closePreview();
        });
        el.copyPreview.addEventListener('click', copyPreview);
        document.addEventListener('keydown', event => {
            if (event.key === 'Escape' && el.previewBackdrop.classList.contains('visible')) closePreview();
        });
        clearResults();
        restoreWorkingWorkbook();
    }

    window.odooFillTool = { REQUIRED_LANGUAGES, OPTIONAL_LANGUAGES, normalizeName };
    document.addEventListener('DOMContentLoaded', init);
})();
