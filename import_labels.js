(function () {
    'use strict';

    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const LANGUAGES = ['EN', 'FR', 'DE', 'CZ', 'PL', 'HU', 'ES', 'PT', 'IT', 'RO', 'SK', 'NL', 'DK', 'FI', 'SE', 'EE'];
    const EXPORT_COLUMNS = [...LANGUAGES, 'ES > PT'];

    const INGREDIENT_HEADINGS = {
        EN: 'Ingredients', FR: 'Ingrédients', DE: 'Zutaten', CZ: 'Složení',
        PL: 'Składniki', HU: 'Összetevők', ES: 'Ingredientes', PT: 'Ingredientes',
        IT: 'Ingredienti', RO: 'Ingrediente', SK: 'Zloženie', NL: 'Ingrediënten',
        DK: 'Ingredienser', FI: 'Ainesosat', SE: 'Ingredienser', EE: 'Koostisosad'
    };

    const NUTRITION_HEADINGS = {
        EN: 'Nutrition Facts per 100g', FR: 'Valeurs Nutritionnelles pour 100g',
        DE: 'Nährwertangaben pro 100g', CZ: 'Výživové Údaje na 100g',
        PL: 'Wartości Odżywcze na 100g', HU: 'Tápanyagtartalom 100g-ban',
        ES: 'Información Nutricional por 100g', PT: 'Declaração Nutricional por 100g',
        IT: 'Valori Nutrizionali per 100g', RO: 'Informații Nutriționale per 100g',
        SK: 'Výživové Údaje na 100g', NL: 'Voedingswaarden per 100g',
        DK: 'Næringsindhold pr 100g', FI: 'Ravintoarvot per 100g',
        SE: 'Näringsvärde per 100g', EE: 'Toitumisalane teave 100g kohta'
    };

    const NUTRITION_HEADING_ALIASES = {
        EN: ['Nutrition Facts'], FR: ['Valeurs Nutritionnelles'], DE: ['Nährwertangaben'],
        CZ: ['Výživové Údaje'], PL: ['Wartości Odżywcze'], HU: ['Tápanyagtartalom'],
        ES: ['Información Nutricional'], PT: ['Declaração Nutricional'], IT: ['Valori Nutrizionali'],
        RO: ['Informații Nutriționale'], SK: ['Výživové Údaje'], NL: ['Voedingswaarde', 'Voedingswaarden'],
        DK: ['Næringsindhold'], FI: ['Ravintoarvot', 'Ravintosisältö'], SE: ['Näringsdeklaration', 'Näringsvärde'],
        EE: ['Toitumisalane Teave']
    };

    // These are recognition aliases, not a fixed nutrition schema. Every row
    // found in Word is kept. Unknown rows are accepted when the source table
    // provides one unambiguous label per detected language.
    const NUTRIENT_ALIASES = {
        EN: ['Energy', 'Fat', 'Of Which Saturated', 'Carbohydrates', 'Of Which Sugars', 'Dietary Fiber', 'Fiber', 'Protein', 'Salt', 'Sodium'],
        FR: ['Énergie', 'Grasse', 'Matières Grasses', 'Dont Acides Gras Saturés', 'Glucides', 'Dont Sucres', 'Fibres Alimentaires', 'Fibres', 'Protéines', 'Sel', 'Sodium'],
        DE: ['Energie', 'Fett', 'Davon Gesättigte Fettsäuren', 'Kohlenhydrate', 'Davon Zucker', 'Ballaststoffe', 'Eiweiss', 'Eiweiß', 'Salz', 'Natrium'],
        CZ: ['Energie', 'Tuky', 'Z Toho Nasycené Mastné Kyseliny', 'Sacharidy', 'Z Toho Cukry', 'Vláknina', 'Bílkoviny', 'Sůl', 'Sodík'],
        PL: ['Energia', 'Tłuszcz', 'W Tym Kwasy Tłuszczowe Nasycone', 'Węglowodany', 'W Tym Cukry', 'Błonnik', 'Białko', 'Sól', 'Sód'],
        HU: ['Energia', 'Zsír', 'Amelyből Telített Zsírsavak', 'Szénhidrá', 'Szénhidrát', 'Amelyből Cukrok', 'Élelmi Rost', 'Fehérje', 'Só', 'Nátrium'],
        ES: ['Energía', 'Grasas', 'De Las Cuales Saturadas', 'Hidratos De Carbono', 'De Los Cuales Azúcares', 'Fibra Alimentaria', 'Proteínas', 'Sal', 'Sodio'],
        PT: ['Energia', 'Lípidos', 'Dos Quais Saturados', 'Hidratos De Carbono', 'Dos Quais Açúcares', 'Fibra Alimentar', 'Proteínas', 'Sal', 'Sódio'],
        IT: ['Energia', 'Grassi', 'Di Cui Acidi Grassi Saturi', 'Carboidrati', 'Di Cui Zuccheri', 'Fibre', 'Proteine', 'Sale', 'Sodio'],
        RO: ['Energie', 'Grăsimi', 'Din Care Acizi Grași Saturați', 'Glucide', 'Din Care Zaharuri', 'Fibre', 'Proteine', 'Sare', 'Sodiu'],
        SK: ['Energia', 'Tuky', 'Z Toho Nasýtené Mastné Kyseliny', 'Sacharidy', 'Z Toho Cukry', 'Vláknina', 'Bielkoviny', 'Soľ', 'Sodík'],
        NL: ['Energie', 'Vetten', 'Waarvan Verzadigd', 'Koolhydraten', 'Waarvan Suikers', 'Voedingsvezels', 'Eiwitten', 'Zout', 'Natrium'],
        DK: ['Energi', 'Fedt', 'Heraf Mættede', 'Kulhydrat', 'Heraf Sukker', 'Kostfibre', 'Kostfibe', 'Protein', 'Salt', 'Natrium'],
        FI: ['Energia', 'Rasva', 'Josta Tyydyttynyttä', 'Hiilihydraatit', 'Josta Sokereita', 'Ravintokuitu', 'Proteiini', 'Suola', 'Natrium'],
        SE: ['Energi', 'Fett', 'Varav Mättat', 'Varav Mättat Fett', 'Kolhydrater', 'Varav Sockerarter', 'Kostfiber', 'Protein', 'Salt', 'Natrium'],
        EE: ['Energia', 'Rasvad', 'Millest Küllastunud', 'Süsivesikud', 'Millest Suhkrud', 'Kiudained', 'Valgud', 'Sool', 'Naatrium']
    };

    const STORAGE = {
        EN: ['Storage.', 'Store at room temperature.'],
        CZ: ['Skladování.', 'Skladujte při pokojové teplotě.']
    };

    const WARNING_TERMS = {
        EN: [/^contains\b/, /^may contain\b/, /^warning\b/],
        FR: [/^contient\b/, /^peut contenir\b/, /^avertissement\b/],
        DE: [/^enthalt\b/, /^kann\b.*\benthalten\b/, /^warnung\b/],
        CZ: [/^obsahuje\b/, /^muze obsahovat\b/, /^varovani\b/],
        PL: [/^zawiera\b/, /^moze zawierac\b/, /^ostrzezenie\b/],
        HU: [/\btartalmaz\b/, /^figyelmeztetes\b/],
        ES: [/^contiene\b/, /^puede contener\b/, /^advertencia\b/],
        PT: [/^contem\b/, /^pode conter\b/, /^aviso\b/],
        IT: [/^contiene\b/, /^puo contenere\b/, /^avvertenza\b/],
        RO: [/^contine\b/, /^poate contine\b/, /^avertisment\b/],
        SK: [/^obsahuje\b/, /^moze obsahovat\b/, /^upozornenie\b/],
        NL: [/^bevat\b/, /^kan bevatten\b/, /^waarschuwing\b/],
        DK: [/^indeholder\b/, /^kan indeholde\b/, /^advarsel\b/],
        FI: [/^sisaltaa\b/, /^saattaa sisaltaa\b/, /^varoitus\b/],
        SE: [/^innehaller\b/, /^kan innehalla\b/, /^varning\b/],
        EE: [/^sisaldab\b/, /^voib sisaldada\b/, /^hoiatus\b/]
    };

    const COMPANY_INFO = {
        EN: {
            importer: 'Importer: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prague 4 Czech Republic.',
            seller: 'Seller: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prague 4 Czech Republic.',
            contact: 'Web: www.pepis.company. Tel.: +420 608 999 488.'
        },
        FR: {
            importer: 'Importateur: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prague 4, République tchèque.',
            seller: 'Vendeur: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prague 4, République tchèque.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        DE: {
            importer: 'Importeur: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tschechische Republik.',
            seller: 'Vertrieb: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tschechische Republik.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        CZ: {
            importer: 'Dovozce: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Česká republika.',
            seller: 'Prodejce: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Česká republika.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        IT: {
            importer: 'Importatore: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Repubblica Ceca.',
            seller: 'Venditore: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Repubblica Ceca.',
            contact: 'Web: www.pepis.company, Telefono: +420 608 999 488.'
        },
        PL: {
            importer: 'Importer: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Republika Czeska.',
            seller: 'Sprzedawca: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Republika Czeska.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        HU: {
            importer: 'Importőr: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prága 4, Csehország.',
            seller: 'Forgalmazó: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prága 4, Csehország.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        ES: {
            importer: 'Importador: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, República Checa.',
            seller: 'Vendedor: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, República Checa.',
            contact: 'Web: www.pepis.company, Teléfono: +420 608 999 488.'
        },
        PT: {
            importer: 'Importador: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, República Checa.',
            seller: 'Vendedor: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, República Checa.',
            contact: 'Web: www.pepis.company, Telefone: +420 608 999 488.'
        },
        RO: {
            importer: 'Importator: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Republica Cehă.',
            seller: 'Vânzător: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Republica Cehă.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        SK: {
            importer: 'Importér: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Česká republika.',
            seller: 'Predajca: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Česká republika.',
            contact: 'Web: www.pepis.company, Telefón: +420 608 999 488.'
        },
        NL: {
            importer: 'Importeur: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praag 4, Tsjechië.',
            seller: 'Verkoper: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praag 4, Tsjechië.',
            contact: 'Web: www.pepis.company, Telefoon: +420 608 999 488.'
        },
        DK: {
            importer: 'Importør: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tjekkiet.',
            seller: 'Sælger: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tjekkiet.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        FI: {
            importer: 'Maahantuoja: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Tšekki.',
            seller: 'Myyjä: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Tšekki.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        SE: {
            importer: 'Importör: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tjeckien.',
            seller: 'Säljare: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tjeckien.',
            contact: 'Web: www.pepis.company, Telefon: +420 608 999 488.'
        },
        EE: {
            importer: 'Turustaja: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Tšehhi Vabariik.',
            seller: 'Müüja: PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Tšehhi Vabariik.',
            contact: 'Web: www.pepis.company, Telefoon: +420 608 999 488.'
        }
    };

    const state = {
        items: [], processing: false, processed: false, exported: false,
        previewItemId: null, previewLanguage: '', previewMode: 'visual',
        previewImageUrl: '', previewRenderToken: 0, companyMode: 'importer'
    };

    const el = {};

    function clean(value) {
        return String(value || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    }

    function normalizedLabel(value) {
        return clean(value).normalize('NFD').replace(/\p{M}/gu, '').toLowerCase()
            .replace(/[^\p{L}\p{N}]+/gu, ' ').replace(/\s+/g, ' ').trim();
    }

    function escapeHtml(value) {
        return String(value || '')
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }

    function escapeRegExp(value) {
        return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    function runPropertyEnabled(run, localName) {
        const property = run.getElementsByTagNameNS(W_NS, localName)[0];
        if (!property) return false;
        const value = property.getAttributeNS(W_NS, 'val') || property.getAttribute('w:val');
        return !value || !/^(0|false|off)$/i.test(value);
    }

    function paragraphSegments(paragraph) {
        const segments = [];
        for (const run of paragraph.getElementsByTagNameNS(W_NS, 'r')) {
            let text = '';
            for (const child of run.childNodes) {
                if (child.namespaceURI !== W_NS) continue;
                if (child.localName === 't' || child.localName === 'delText') text += child.textContent || '';
                if (child.localName === 'tab') text += '\t';
                if (child.localName === 'br' || child.localName === 'cr') text += '\n';
            }
            if (!text) continue;
            segments.push({
                text,
                bold: runPropertyEnabled(run, 'b') || runPropertyEnabled(run, 'bCs'),
                italic: runPropertyEnabled(run, 'i') || runPropertyEnabled(run, 'iCs'),
                underline: runPropertyEnabled(run, 'u')
            });
        }
        return segments;
    }

    function paragraphData(paragraph) {
        const segments = paragraphSegments(paragraph);
        return { text: clean(segments.map(segment => segment.text).join('')), segments };
    }

    function directChildren(element, localName) {
        return [...element.childNodes].filter(node => node.namespaceURI === W_NS && node.localName === localName);
    }

    function tableData(table) {
        const rows = [];
        for (const row of table.getElementsByTagNameNS(W_NS, 'tr')) {
            const cells = [];
            for (const cell of directChildren(row, 'tc')) {
                const paragraphs = [...cell.getElementsByTagNameNS(W_NS, 'p')].map(paragraphData);
                cells.push({
                    text: clean(paragraphs.map(paragraph => paragraph.text).filter(Boolean).join(' ')),
                    paragraphs
                });
            }
            if (cells.length) rows.push(cells);
        }
        return rows;
    }

    async function readDocx(file) {
        let zip;
        try {
            zip = await JSZip.loadAsync(await file.arrayBuffer());
        } catch (_error) {
            throw new Error('DOCX bị hỏng hoặc không đúng định dạng Word.');
        }
        const part = zip.file('word/document.xml');
        if (!part) throw new Error('DOCX thiếu word/document.xml.');
        const xml = new DOMParser().parseFromString(await part.async('string'), 'application/xml');
        if (xml.querySelector('parsererror')) throw new Error('Không đọc được XML bên trong DOCX.');
        const body = xml.getElementsByTagNameNS(W_NS, 'body')[0];
        const paragraphs = [];
        const tables = [];
        for (const child of body.childNodes) {
            if (child.namespaceURI !== W_NS) continue;
            if (child.localName === 'p') {
                const paragraph = paragraphData(child);
                if (paragraph.text) paragraphs.push(paragraph);
            }
            if (child.localName === 'tbl') tables.push(tableData(child));
        }
        return { paragraphs, tables };
    }

    function sliceSegments(segments, startOffset, endOffset = Infinity) {
        const result = [];
        let cursor = 0;
        for (const segment of segments) {
            const start = cursor;
            const end = cursor + segment.text.length;
            cursor = end;
            const overlapStart = Math.max(start, startOffset);
            const overlapEnd = Math.min(end, endOffset);
            if (overlapStart >= overlapEnd) continue;
            const text = segment.text.slice(overlapStart - start, overlapEnd - start);
            if (!text) continue;
            const previous = result[result.length - 1];
            if (previous && previous.bold === segment.bold && previous.italic === segment.italic && previous.underline === segment.underline) {
                previous.text += text;
            } else {
                result.push({ ...segment, text });
            }
        }
        if (result.length) {
            result[0].text = result[0].text.replace(/^\s+/, '');
            result[result.length - 1].text = result[result.length - 1].text.replace(/\s+$/, '');
        }
        return result.filter(segment => segment.text);
    }

    function formatSegments(segments) {
        return segments.map(segment => {
            let content = escapeHtml(segment.text.replace(/\u00a0/g, ' ')).replace(/\r?\n/g, '<br>');
            if (!content) return '';
            if (!segment.bold && !segment.italic && !segment.underline) return `<span style="font-weight: 400;">${content}</span>`;
            if (segment.underline) content = `<u>${content}</u>`;
            if (segment.italic) content = `<i>${content}</i>`;
            if (segment.bold) content = `<b>${content}</b>`;
            return content;
        }).join('');
    }

    function languageBlocks(paragraphs) {
        const blocks = [];
        const markerPattern = /(?:^|[\r\n])[\t ]*►?[\t ]*\(([A-Z]{2})\)[\t ]*/gim;
        paragraphs.forEach((paragraph, paragraphIndex) => {
            const raw = paragraph.segments.map(segment => segment.text).join('');
            const markers = [];
            markerPattern.lastIndex = 0;
            let match;
            while ((match = markerPattern.exec(raw)) !== null) {
                const language = match[1].toUpperCase();
                if (!LANGUAGES.includes(language)) continue;
                const leadingBreakLength = /^[\r\n]/.test(match[0]) ? 1 : 0;
                markers.push({ language, start: match.index + leadingBreakLength });
            }
            markers.forEach((marker, markerIndex) => {
                const end = markers[markerIndex + 1]?.start ?? raw.length;
                const segments = sliceSegments(paragraph.segments, marker.start, end);
                if (!segments.length) return;
                blocks.push({
                    language: marker.language,
                    paragraphIndex,
                    paragraph: { text: clean(segments.map(segment => segment.text).join('')), segments }
                });
            });
        });
        return blocks;
    }

    function trailingWarningRange(value, language) {
        const terms = WARNING_TERMS[language] || [];
        if (!terms.length) return null;
        const sentences = [];
        const sentencePattern = /[^.!?\r\n]+(?:[.!?]+|$)/g;
        let match;
        while ((match = sentencePattern.exec(value)) !== null) {
            const leading = match[0].match(/^\s*/)?.[0].length || 0;
            const text = clean(match[0]);
            if (text) sentences.push({ start: match.index + leading, end: match.index + match[0].length, text });
        }
        let warningStart = value.length;
        for (let index = sentences.length - 1; index >= 0; index -= 1) {
            const normalized = normalizedLabel(sentences[index].text);
            if (!terms.some(pattern => pattern.test(normalized))) break;
            warningStart = sentences[index].start;
        }
        return warningStart < value.length ? { start: warningStart, end: value.length } : null;
    }

    function parseLanguageParagraph(paragraph, language) {
        const raw = paragraph.segments.map(segment => segment.text).join('');
        const marker = raw.match(/^\s*►?\s*\(([A-Z]{2})\)\s*/i);
        if (!marker || marker[1].toUpperCase() !== language) throw new Error(`Khối ${language} không hợp lệ.`);
        const ingredientHeading = INGREDIENT_HEADINGS[language];
        const headingPattern = new RegExp(`${escapeRegExp(ingredientHeading)}\\s*:`, 'i');
        const afterMarker = raw.slice(marker[0].length);
        const headingMatch = headingPattern.exec(afterMarker);
        if (!headingMatch) throw new Error(`Không tìm thấy tiêu đề thành phần trong khối ${language}.`);
        const headingStart = marker[0].length + headingMatch.index;
        const ingredientStart = headingStart + headingMatch[0].length;
        const description = clean(raw.slice(marker[0].length, headingStart));
        const warningRange = trailingWarningRange(raw.slice(ingredientStart), language);
        const ingredientEnd = warningRange ? ingredientStart + warningRange.start : raw.length;
        const ingredientSegments = sliceSegments(paragraph.segments, ingredientStart, ingredientEnd);
        const warningSegments = warningRange
            ? sliceSegments(paragraph.segments, ingredientStart + warningRange.start, ingredientStart + warningRange.end)
            : [];
        const ingredientText = clean(ingredientSegments.map(segment => segment.text).join(''));
        if (!description || !ingredientText) throw new Error(`Khối ${language} thiếu mô tả hoặc thành phần.`);
        return {
            language,
            description,
            ingredientHeading: `${ingredientHeading}:`,
            ingredientSegments,
            warningText: clean(warningSegments.map(segment => segment.text).join(''))
        };
    }

    function findNutritionTable(tables) {
        return tables.find(rows => rows.length >= 2 && rows[0] && rows[0].length >= 2 && /100\s*g/i.test(rows[0][1].text));
    }

    function sourceVariants(value) {
        return clean(value).split(/\s*\/\s*/).map(clean).filter(Boolean);
    }

    function resolveSourceLabel(source, language, aliasesByLanguage, languages, languageIndex, context) {
        const variants = sourceVariants(source);
        const aliases = aliasesByLanguage[language] || [];
        const normalizedVariants = variants.map(variant => normalizedLabel(variant));
        for (const alias of aliases) {
            const matchIndex = normalizedVariants.indexOf(normalizedLabel(alias));
            if (matchIndex >= 0) return variants[matchIndex];
        }
        if (variants.length === languages.length && variants[languageIndex]) {
            return variants[languageIndex];
        }
        throw new Error(`Không xác định được nhãn ${context} cho ${language} từ bảng Word.`);
    }

    function parseNutrition(table, languages) {
        if (!table) throw new Error('Không tìm thấy bảng NTF dùng chung có mốc 100g.');
        const dataRows = table.slice(1).filter(row => row.length >= 2 && clean(row[1].text));
        if (!dataRows.length) throw new Error('Bảng NTF không có chỉ tiêu nào có dữ liệu.');
        const headingSource = clean(table[0][0].text);
        const headingLabels = Object.create(null);
        languages.forEach((language, index) => {
            headingLabels[language] = resolveSourceLabel(
                headingSource, language, NUTRITION_HEADING_ALIASES, languages, index, 'tiêu đề NTF'
            );
        });
        const metrics = dataRows.map((row, rowIndex) => {
            const value = clean(row[1].text).replace(/\s*\/\s*/g, ' / ');
            if (!/\d/.test(value)) throw new Error(`Giá trị NTF dòng ${rowIndex + 1} không hợp lệ.`);
            const sourceLabel = clean(row[0].text);
            if (!sourceLabel) throw new Error(`Nhãn NTF dòng ${rowIndex + 1} đang trống.`);
            const labels = Object.create(null);
            languages.forEach((language, languageIndex) => {
                labels[language] = resolveSourceLabel(
                    sourceLabel, language, NUTRIENT_ALIASES, languages, languageIndex, `NTF dòng ${rowIndex + 1}`
                );
            });
            return { value, sourceLabel, labels };
        });
        return { headingLabels, metrics };
    }

    function isSubItem(label) {
        return /^(?:of\s+which|dont\b|davon\b|z\s+toho\b|w\s+tym\b|amelyből\b|de\s+las\b|de\s+los\b|dos\s+quais\b|di\s+cui\b|din\s+care\b|waarvan\b|heraf\b|josta\b|varav\b|millest\b)/i
            .test(clean(label).replace(/^-\s*/, ''));
    }

    function formatCompanySentence(value) {
        const match = clean(value).match(/^([^:]+):\s*(.+)$/);
        if (!match) return `<p><span style="font-weight: 400;">${escapeHtml(clean(value))}</span></p>`;
        return `<p><b>${escapeHtml(match[1])}:</b> <span style="font-weight: 400;">${escapeHtml(match[2])}</span></p>`;
    }

    function formatContactSentence(value) {
        const match = clean(value).match(/^([^:]+):\s*(.*?)(?:,\s+|\.\s+)([^:]+):\s*(.+)$/);
        if (!match) return formatCompanySentence(value);
        const separator = String(value).includes('. Tel.') ? '. ' : ', ';
        return `<p><b>${escapeHtml(match[1])}:</b> <span style="font-weight: 400;">${escapeHtml(match[2])}${escapeHtml(separator)}</span><b>${escapeHtml(match[3])}:</b> <span style="font-weight: 400;">${escapeHtml(match[4])}</span></p>`;
    }

    function companyInfoHtml(language, companyMode) {
        const info = COMPANY_INFO[language];
        if (!info) throw new Error(`Thiếu thông tin công ty cho ${language}.`);
        const selected = companyMode === 'seller' ? info.seller : info.importer;
        return [formatCompanySentence(selected), formatContactSentence(info.contact)];
    }

    function buildLanguageHtml(block, nutrition, companyMode = 'importer') {
        const metrics = nutrition.metrics.map(metric => ({ label: metric.labels[block.language], value: metric.value }));
        const lines = [
            `<p><i>(${block.language}) ${escapeHtml(block.description)}</i></p>`,
            `<p><b>${escapeHtml(block.ingredientHeading)}</b> ${formatSegments(block.ingredientSegments)}</p>`
        ];
        if (block.warningText) lines.push(`<p><b>${escapeHtml(block.warningText)}</b></p>`);
        const heading = NUTRITION_HEADINGS[block.language];
        if (block.language === 'EN' || block.language === 'CZ') {
            lines.push(`<p><b>${escapeHtml(heading)}:</b></p>`);
            metrics.forEach(metric => {
                const text = `${metric.label}: ${metric.value}`;
                if (isSubItem(metric.label)) lines.push(`<p><i>- ${escapeHtml(text)}</i></p>`);
                else lines.push(`<p><b>${escapeHtml(metric.label)}:</b> ${escapeHtml(metric.value)}</p>`);
            });
            const storage = STORAGE[block.language];
            lines.push(`<p><strong>${escapeHtml(storage[0])}</strong></p>`);
            lines.push(`<p><span>${escapeHtml(storage[1])}</span></p>`);
        } else {
            const inlineMetrics = metrics.map(metric => `${metric.label}: ${metric.value}`).join(', ');
            lines.push(`<p><b>${escapeHtml(heading)}:</b> <span style="font-weight: 400;">${escapeHtml(inlineMetrics)}.</span></p>`);
        }
        lines.push(...companyInfoHtml(block.language, companyMode));
        return lines.join('\n');
    }

    function optionalCode(fileName) {
        const stem = fileName.replace(/\.docx$/i, '');
        const match = stem.match(/(?:^|[\s_-])(\d+(?:-\d+)?)(?=[\s_-]|$)/);
        return match ? match[1] : '';
    }

    async function convertFile(file, id, companyMode = 'importer') {
        const doc = await readDocx(file);
        const blocks = languageBlocks(doc.paragraphs);
        if (!blocks.length) throw new Error('Không tìm thấy khối ngôn ngữ dạng (EN), (FR), (DE)...');
        const firstMarkerIndex = blocks[0].paragraphIndex;
        const titleParagraph = doc.paragraphs.slice(0, firstMarkerIndex).find(paragraph => paragraph.text);
        const name = clean(titleParagraph ? titleParagraph.text : file.name.replace(/\.docx$/i, ''));
        if (!name) throw new Error('Không tìm thấy tên sản phẩm.');
        const detectedLanguages = [...new Set(blocks.map(block => block.language))];
        const nutrition = parseNutrition(findNutritionTable(doc.tables), detectedLanguages);
        const htmlByLanguage = Object.create(null);
        const warnings = [];
        for (const blockData of blocks) {
            if (htmlByLanguage[blockData.language]) {
                warnings.push(`Bỏ qua khối ${blockData.language} bị lặp.`);
                continue;
            }
            const block = parseLanguageParagraph(blockData.paragraph, blockData.language);
            htmlByLanguage[blockData.language] = buildLanguageHtml(block, nutrition, companyMode);
        }
        return {
            id, file, fileName: file.name, name, defaultCode: optionalCode(file.name),
            languages: Object.keys(htmlByLanguage), htmlByLanguage, warnings,
            companyMode,
            status: 'success', detail: `Đã tách ${Object.keys(htmlByLanguage).length} ngôn ngữ — ${companyMode === 'seller' ? 'Seller' : 'Importer'}`
        };
    }

    function statusLabel(status) {
        return status === 'success' ? '✓ Hoàn thành' : status === 'error' ? '! Lỗi' : status === 'processing' ? '◌ Đang xử lý' : '• Đang chờ';
    }

    function renderRows() {
        el.body.innerHTML = '';
        if (!state.items.length) {
            el.body.innerHTML = '<tr><td colspan="8" class="empty-state">Chưa có file DOCX nào được quét.</td></tr>';
            return;
        }
        state.items.forEach((item, index) => {
            const row = document.createElement('tr');
            row.classList.toggle('row-error', item.status === 'error');
            const codeValue = escapeHtml(item.defaultCode || '');
            const sourceText = item.sourceType === 'zip'
                ? `ZIP: ${item.sourceName}${item.sourcePath ? ` — ${item.sourcePath}` : ''}`
                : 'File DOCX tải trực tiếp';
            row.innerHTML = `
                <td>${index + 1}</td>
                <td><div class="file-name">${escapeHtml(item.fileName)}</div><div class="source-name">${escapeHtml(sourceText)}</div></td>
                <td><input class="import-code-input" type="text" value="${codeValue}" placeholder="Không bắt buộc"></td>
                <td>${escapeHtml(item.name || '—')}</td>
                <td>${escapeHtml((item.languages || []).join(', ') || '—')}</td>
                <td><span class="status-badge status-${item.status}">${statusLabel(item.status)}</span></td>
                <td class="detail-cell">${escapeHtml(item.detail || '—')}</td>
                <td class="action-cell"></td>`;
            const input = row.querySelector('.import-code-input');
            input.addEventListener('input', () => { item.defaultCode = input.value.trim(); });
            if (item.status === 'success') {
                const button = document.createElement('button');
                button.className = 'small-btn';
                button.type = 'button';
                button.textContent = 'Xem kết quả';
                button.addEventListener('click', () => openPreview(item.id));
                row.querySelector('.action-cell').appendChild(button);
            }
            el.body.appendChild(row);
        });
    }

    function updateSummary() {
        const total = state.items.length;
        const success = state.items.filter(item => item.status === 'success').length;
        const error = state.items.filter(item => item.status === 'error').length;
        el.total.textContent = total;
        el.success.textContent = success;
        el.error.textContent = error;
        el.review.classList.toggle('visible', state.processed);
        el.reviewMessage.textContent = state.processed ? `Hoàn thành ${success} file, lỗi ${error} file. default_code có thể để trống.` : '';
        el.exportButton.disabled = !state.processed || !success;
        el.exportButton.textContent = `Duyệt & Xuất Excel (${success} dòng)`;
    }

    function basename(path) {
        return String(path || '').split(/[\\/]/).pop() || '';
    }

    async function scanImportSources(sourceFiles) {
        const descriptors = [];
        const docxMime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        for (const sourceFile of sourceFiles) {
            if (/\.docx$/i.test(sourceFile.name)) {
                if (/^~\$/i.test(sourceFile.name)) continue;
                descriptors.push({
                    file: sourceFile,
                    fileName: sourceFile.name,
                    sourceType: 'docx',
                    sourceName: sourceFile.name,
                    sourcePath: ''
                });
                continue;
            }
            if (!/\.zip$/i.test(sourceFile.name)) continue;

            let archive;
            try {
                archive = await JSZip.loadAsync(sourceFile);
            } catch (error) {
                descriptors.push({
                    file: null,
                    fileName: sourceFile.name,
                    sourceType: 'zip',
                    sourceName: sourceFile.name,
                    sourcePath: '',
                    preError: `Không mở được ZIP: ${error.message}`
                });
                continue;
            }

            const entries = Object.values(archive.files)
                .filter(entry => !entry.dir)
                .filter(entry => /\.docx$/i.test(entry.name))
                .filter(entry => !/(^|\/)__MACOSX\//i.test(entry.name))
                .filter(entry => !/^~\$/i.test(basename(entry.name)))
                .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));

            if (!entries.length) {
                descriptors.push({
                    file: null,
                    fileName: sourceFile.name,
                    sourceType: 'zip',
                    sourceName: sourceFile.name,
                    sourcePath: '',
                    preError: 'ZIP không chứa file DOCX.'
                });
                continue;
            }

            entries.forEach(entry => descriptors.push({
                file: null,
                entry,
                fileName: basename(entry.name),
                sourceType: 'zip',
                sourceName: sourceFile.name,
                sourcePath: entry.name,
                docxMime
            }));
        }
        return descriptors;
    }

    async function processFiles() {
        const sourceFiles = [...el.files.files].filter(file => /\.(zip|docx)$/i.test(file.name));
        if (!sourceFiles.length) {
            alert('Hãy chọn ít nhất một file ZIP hoặc DOCX.');
            return;
        }
        state.processing = true;
        state.processed = false;
        state.exported = false;
        state.companyMode = document.querySelector('input[name="import-company-mode"]:checked')?.value || 'importer';
        el.processButton.disabled = true;
        const descriptors = await scanImportSources(sourceFiles);
        if (!descriptors.length) {
            state.processing = false;
            el.processButton.disabled = false;
            alert('Không tìm thấy file DOCX hợp lệ trong dữ liệu đã chọn.');
            return;
        }
        state.items = descriptors.map((descriptor, index) => ({
            ...descriptor,
            id: index + 1,
            name: '',
            defaultCode: optionalCode(descriptor.fileName),
            languages: [],
            htmlByLanguage: Object.create(null),
            status: descriptor.preError ? 'error' : 'waiting',
            detail: descriptor.preError || 'Đang chờ xử lý'
        }));
        renderRows();
        updateSummary();
        for (let index = 0; index < state.items.length; index++) {
            const item = state.items[index];
            if (item.status === 'error') continue;
            item.status = 'processing';
            item.detail = item.sourceType === 'zip' ? 'Đang giải nén và đọc DOCX' : 'Đang đọc DOCX';
            renderRows();
            try {
                const file = item.file || new File(
                    [await item.entry.async('blob')],
                    item.fileName,
                    { type: item.docxMime, lastModified: Date.now() }
                );
                const converted = await convertFile(file, item.id, state.companyMode);
                converted.sourceType = item.sourceType;
                converted.sourceName = item.sourceName;
                converted.sourcePath = item.sourcePath;
                state.items[index] = converted;
            } catch (error) {
                item.status = 'error';
                item.detail = error.message || 'Lỗi không xác định';
            }
            renderRows();
            updateSummary();
            await new Promise(resolve => setTimeout(resolve, 0));
        }
        state.processing = false;
        state.processed = true;
        el.processButton.disabled = false;
        updateSummary();
    }

    function reset() {
        state.items = [];
        state.processing = false;
        state.processed = false;
        state.exported = false;
        el.files.value = '';
        renderRows();
        updateSummary();
    }

    function sanitizePreviewHtml(html) {
        const template = document.createElement('template');
        template.innerHTML = html || '';
        template.content.querySelectorAll('script, style, iframe, object, embed, link, meta').forEach(node => node.remove());
        template.content.querySelectorAll('*').forEach(node => {
            [...node.attributes].forEach(attribute => {
                const name = attribute.name.toLowerCase();
                const value = attribute.value.trim();
                if (name.startsWith('on') || ['src', 'href', 'srcdoc'].includes(name)) {
                    node.removeAttribute(attribute.name);
                } else if (name === 'style' && !/^font-weight\s*:\s*(?:400|normal)\s*;?$/i.test(value)) {
                    node.removeAttribute(attribute.name);
                }
            });
        });
        return template.innerHTML;
    }

    function setPreviewMode(mode) {
        state.previewMode = mode === 'code' ? 'code' : 'visual';
        const showVisual = state.previewMode === 'visual';
        el.previewVisualPanel.hidden = !showVisual;
        el.previewCode.hidden = showVisual;
        el.previewVisualTab.classList.toggle('active', showVisual);
        el.previewCodeTab.classList.toggle('active', !showVisual);
        el.previewVisualTab.setAttribute('aria-selected', String(showVisual));
        el.previewCodeTab.setAttribute('aria-selected', String(!showVisual));
    }

    async function renderPreviewImage() {
        const item = state.items.find(candidate => candidate.id === state.previewItemId);
        const language = state.previewLanguage;
        const html = item?.htmlByLanguage?.[language] || '';
        if (!item || !html) return;

        const renderToken = ++state.previewRenderToken;
        state.previewImageUrl = '';
        el.previewImage.hidden = true;
        el.previewImage.removeAttribute('src');
        el.previewRendered.hidden = false;
        el.previewRendered.innerHTML = sanitizePreviewHtml(html);
        el.previewStatus.textContent = `Đang tạo ảnh nhãn ${language}…`;

        await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        if (renderToken !== state.previewRenderToken) return;

        if (typeof html2canvas !== 'function') {
            el.previewStatus.textContent = 'Không tải được bộ tạo ảnh; đang hiển thị bản render trực tiếp.';
            return;
        }

        try {
            const canvas = await html2canvas(el.previewRendered, {
                backgroundColor: '#ffffff',
                scale: 2,
                useCORS: true,
                logging: false
            });
            if (renderToken !== state.previewRenderToken) return;
            state.previewImageUrl = canvas.toDataURL('image/png');
            el.previewImage.src = state.previewImageUrl;
            el.previewImage.alt = `Ảnh nhãn ${language} — ${item.name}`;
            el.previewImage.hidden = false;
            el.previewRendered.hidden = true;
            el.previewStatus.textContent = `Ảnh nhãn ${language} được tạo trực tiếp từ mã HTML.`;
        } catch (error) {
            console.error('Không thể tạo ảnh IMPORT LABELS:', error);
            el.previewStatus.textContent = 'Không thể tạo ảnh; đang hiển thị bản render trực tiếp để kiểm tra.';
        }
    }

    function openPreview(itemId) {
        const item = state.items.find(candidate => candidate.id === itemId);
        if (!item || item.status !== 'success') return;
        state.previewItemId = itemId;
        el.previewSelect.innerHTML = item.languages.map(language => `<option value="${language}">${language}</option>`).join('');
        state.previewLanguage = item.languages[0];
        el.previewSelect.value = state.previewLanguage;
        el.previewTitle.textContent = `${item.name} — ${state.previewLanguage}`;
        el.previewCode.textContent = item.htmlByLanguage[state.previewLanguage];
        setPreviewMode('visual');
        el.previewBackdrop.classList.add('visible');
        renderPreviewImage();
    }

    function updatePreviewLanguage() {
        const item = state.items.find(candidate => candidate.id === state.previewItemId);
        if (!item) return;
        state.previewLanguage = el.previewSelect.value;
        el.previewTitle.textContent = `${item.name} — ${state.previewLanguage}`;
        el.previewCode.textContent = item.htmlByLanguage[state.previewLanguage] || '';
        renderPreviewImage();
    }

    function closePreview() {
        state.previewRenderToken += 1;
        state.previewItemId = null;
        state.previewLanguage = '';
        state.previewImageUrl = '';
        el.previewImage.removeAttribute('src');
        el.previewBackdrop.classList.remove('visible');
    }

    function exportExcel() {
        const successful = state.items.filter(item => item.status === 'success');
        if (!successful.length || state.exported) return;
        const data = [
            ['default_code', 'name', ...EXPORT_COLUMNS],
            ...successful.map(item => [
                item.defaultCode || '', item.name,
                ...EXPORT_COLUMNS.map(language => item.htmlByLanguage[language] || 'X')
            ])
        ];
        const worksheet = XLSX.utils.aoa_to_sheet(data);
        worksheet['!cols'] = [{ wch: 18 }, { wch: 45 }, ...EXPORT_COLUMNS.map(() => ({ wch: 80 }))];
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Results');
        const stamp = new Date().toISOString().replace(/[-:]/g, '').slice(0, 15);
        XLSX.writeFile(workbook, `import_labels_${stamp}.xlsx`);
        state.exported = true;
        el.exportButton.disabled = true;
        el.exportButton.textContent = 'Đã duyệt & xuất Excel';
    }

    function initTabs() {
        document.querySelectorAll('.tool-tab').forEach(button => {
            button.addEventListener('click', () => {
                const target = button.dataset.toolTab;
                document.querySelectorAll('.tool-tab').forEach(tab => tab.classList.toggle('active', tab === button));
                document.querySelectorAll('.tool-view').forEach(view => view.hidden = view.dataset.toolView !== target);
            });
        });
    }

    function init() {
        Object.assign(el, {
            files: document.getElementById('import-source-files'),
            processButton: document.getElementById('import-process-button'),
            resetButton: document.getElementById('import-reset-button'),
            body: document.getElementById('import-results-body'),
            total: document.getElementById('import-count-total'),
            success: document.getElementById('import-count-success'),
            error: document.getElementById('import-count-error'),
            review: document.getElementById('import-review-panel'),
            reviewMessage: document.getElementById('import-review-message'),
            exportButton: document.getElementById('import-export-button'),
            previewBackdrop: document.getElementById('import-preview-backdrop'),
            previewTitle: document.getElementById('import-preview-title'),
            previewSelect: document.getElementById('import-preview-language'),
            previewCode: document.getElementById('import-preview-code'),
            previewVisualPanel: document.getElementById('import-preview-visual-panel'),
            previewRendered: document.getElementById('import-preview-rendered'),
            previewImage: document.getElementById('import-preview-image'),
            previewStatus: document.getElementById('import-preview-status'),
            previewVisualTab: document.getElementById('import-preview-visual-tab'),
            previewCodeTab: document.getElementById('import-preview-code-tab'),
            closePreview: document.getElementById('import-close-preview'),
            copyPreview: document.getElementById('import-copy-preview')
        });
        initTabs();
        el.processButton.addEventListener('click', processFiles);
        el.resetButton.addEventListener('click', reset);
        el.exportButton.addEventListener('click', exportExcel);
        el.previewSelect.addEventListener('change', updatePreviewLanguage);
        el.previewVisualTab.addEventListener('click', () => setPreviewMode('visual'));
        el.previewCodeTab.addEventListener('click', () => setPreviewMode('code'));
        el.closePreview.addEventListener('click', closePreview);
        el.previewBackdrop.addEventListener('click', event => {
            if (event.target === el.previewBackdrop) closePreview();
        });
        document.addEventListener('keydown', event => {
            if (event.key === 'Escape' && el.previewBackdrop.classList.contains('visible')) closePreview();
        });
        el.copyPreview.addEventListener('click', async () => {
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
        });
        reset();
    }

    window.importLabelsTool = { convertFile, parseLanguageParagraph, parseNutrition, buildLanguageHtml };
    document.addEventListener('DOMContentLoaded', init);
})();
