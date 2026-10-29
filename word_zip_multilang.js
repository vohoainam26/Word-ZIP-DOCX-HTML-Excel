(function (global) {
    'use strict';

    const SUPPORTED_LANGUAGES = new Set([
        'EN', 'FR', 'DE', 'CZ', 'PL', 'HU', 'ES', 'PT',
        'IT', 'RO', 'SK', 'NL', 'DK', 'FI', 'SE', 'EE'
    ]);

    const STORAGE_TRANSLATIONS = {
        EN: { heading: 'Storage.', text: 'Store at room temperature.' },
        FR: { heading: 'Conservation.', text: 'Conserver à température ambiante.' },
        DE: { heading: 'Lagerung.', text: 'Bei Raumtemperatur lagern.' },
        CZ: { heading: 'Skladování.', text: 'Skladujte při pokojové teplotě.' },
        PL: { heading: 'Przechowywanie.', text: 'Przechowywać w temperaturze pokojowej.' },
        HU: { heading: 'Tárolás.', text: 'Szobahőmérsékleten tárolandó.' },
        ES: { heading: 'Conservación.', text: 'Conservar a temperatura ambiente.' },
        PT: { heading: 'Conservação.', text: 'Conservar à temperatura ambiente.' },
        IT: { heading: 'Conservazione.', text: 'Conservare a temperatura ambiente.' },
        RO: { heading: 'Păstrare.', text: 'A se păstra la temperatura camerei.' },
        SK: { heading: 'Skladovanie.', text: 'Skladujte pri izbovej teplote.' },
        NL: { heading: 'Bewaring.', text: 'Bewaren bij kamertemperatuur.' },
        DK: { heading: 'Opbevaring.', text: 'Opbevares ved stuetemperatur.' },
        FI: { heading: 'Säilytys.', text: 'Säilytä huoneenlämmössä.' },
        SE: { heading: 'Förvaring.', text: 'Förvaras i rumstemperatur.' },
        EE: { heading: 'Säilitamine.', text: 'Hoida toatemperatuuril.' }
    };

    const SUBITEM_PATTERNS = [
        /^of\s+which\b/i,
        /^dont\b/i,
        /^davon\b/i,
        /^z\s+toho\b/i,
        /^w\s+tym\b/i,
        /^amelyből\b/i,
        /^de\s+(?:las|los)\s+cuales\b/i,
        /^(?:dos|das)\s+quais\b/i,
        /^di\s+cui\b/i,
        /^din\s+care\b/i,
        /^waarvan\b/i,
        /^heraf\b/i,
        /^josta\b/i,
        /^varav\b/i,
        /^millest\b/i
    ];

    function cleanText(value) {
        return String(value || '')
            .replace(/\u00a0/g, ' ')
            .replace(/[ \t\r\f\v]+/g, ' ')
            .replace(/\s+([,.;:!?])/g, '$1')
            .trim();
    }

    function escapeHtml(value) {
        return String(value || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function getLanguageMarker(paragraph) {
        const match = String(paragraph && paragraph.text || '').match(/^\(([A-Z]{2})\)\s*/i);
        if (!match) return null;
        const language = match[1].toUpperCase();
        return SUPPORTED_LANGUAGES.has(language) ? language : null;
    }

    function splitLanguageBlocks(paragraphs, expectedLanguages) {
        const markers = [];
        paragraphs.forEach((paragraph, index) => {
            const language = getLanguageMarker(paragraph);
            if (language) markers.push({ language, index });
        });

        const blocks = [];
        let searchAfter = -1;
        for (const language of expectedLanguages) {
            const markerPosition = markers.findIndex(marker => marker.index > searchAfter && marker.language === language);
            if (markerPosition < 0) {
                throw new Error(`Không tìm thấy khối (${language}) trong Word.`);
            }
            const marker = markers[markerPosition];
            const nextMarker = markers.find(candidate => candidate.index > marker.index);
            const endIndex = nextMarker ? nextMarker.index : paragraphs.length;
            blocks.push({
                language,
                paragraphs: paragraphs.slice(marker.index, endIndex)
            });
            searchAfter = marker.index;
        }
        return blocks;
    }

    function buildTextStream(paragraphs) {
        let text = '';
        const segments = [];
        paragraphs.forEach((paragraph, paragraphIndex) => {
            if (paragraphIndex > 0) {
                const start = text.length;
                text += '\n';
                segments.push({ text: '\n', bold: false, italic: false, synthetic: true, start, end: text.length });
            }
            for (const segment of paragraph.segments || []) {
                const start = text.length;
                text += segment.text;
                segments.push({
                    text: segment.text,
                    bold: Boolean(segment.bold),
                    italic: Boolean(segment.italic),
                    synthetic: false,
                    start,
                    end: text.length
                });
            }
        });
        return { text, segments };
    }

    function findBoldGroups(stream) {
        const groups = [];
        for (const segment of stream.segments) {
            if (!segment.bold || !segment.text.trim()) continue;
            const previous = groups[groups.length - 1];
            if (previous && /^\s*$/.test(stream.text.slice(previous.end, segment.start))) {
                previous.end = segment.end;
            } else {
                groups.push({ start: segment.start, end: segment.end });
            }
        }
        return groups.map(group => ({
            ...group,
            text: cleanText(stream.text.slice(group.start, group.end))
        }));
    }

    function extendHeadingEnd(text, end) {
        const remaining = text.slice(end);
        const colon = remaining.match(/^\s*:\s*/);
        if (colon) return end + colon[0].length;
        const whitespace = remaining.match(/^\s*/);
        return end + (whitespace ? whitespace[0].length : 0);
    }

    function normalizeHeading(value) {
        const heading = cleanText(value).replace(/\s*:\s*$/, ':');
        return heading.endsWith(':') ? heading : `${heading}:`;
    }

    function sliceStreamSegments(segments, startOffset, endOffset) {
        const result = [];
        for (const segment of segments) {
            const overlapStart = Math.max(startOffset, segment.start);
            const overlapEnd = Math.min(endOffset, segment.end);
            if (overlapStart >= overlapEnd) continue;
            let text = segment.text.slice(overlapStart - segment.start, overlapEnd - segment.start);
            text = text.replace(/\s+/g, ' ');
            if (!text) continue;
            const previous = result[result.length - 1];
            if (previous && previous.bold === segment.bold && previous.italic === segment.italic) {
                previous.text += text;
            } else {
                result.push({ text, bold: segment.bold, italic: segment.italic });
            }
        }
        if (result.length) {
            result[0].text = result[0].text.replace(/^\s+/, '');
            result[result.length - 1].text = result[result.length - 1].text.replace(/\s+$/, '');
        }
        return result.filter(segment => segment.text);
    }

    function findNutritionSentence(value) {
        const body = String(value || '').replace(/^\s+/, '');
        const sentenceEnd = body.search(/\.(?=\s+\p{L})/u);
        if (sentenceEnd >= 0) return body.slice(0, sentenceEnd + 1);
        const lineEnd = body.indexOf('\n');
        return lineEnd >= 0 ? body.slice(0, lineEnd) : body;
    }

    function isSubItemLabel(label) {
        const normalized = cleanText(label).replace(/^-\s*/, '').toLowerCase();
        return SUBITEM_PATTERNS.some(pattern => pattern.test(normalized));
    }

    function parseMetrics(sentence, language) {
        const cleanSentence = cleanText(sentence).replace(/\.\s*$/, '');
        const pieces = cleanSentence.split(/[,;]\s+(?=(?:-\s*)?\p{L})/u);
        const metrics = [];

        for (const piece of pieces) {
            const text = cleanText(piece).replace(/^-\s*/, '');
            const digitIndex = text.search(/\d/);
            if (digitIndex <= 0) continue;
            const label = cleanText(text.slice(0, digitIndex)).replace(/[<>≤≥~]\s*$/, '').trim();
            let valueStart = digitIndex;
            while (valueStart > 0 && /[<>≤≥~]/.test(text[valueStart - 1])) valueStart--;
            const value = cleanText(text.slice(valueStart));
            if (!label || !value) continue;
            metrics.push({
                label,
                value,
                subItem: isSubItemLabel(label)
            });
        }

        if (metrics.length < 2) {
            throw new Error(`Không tách được các chỉ tiêu NTF trong khối ${language}.`);
        }
        return metrics;
    }

    function formatIngredientSegments(segments) {
        return segments.map(segment => {
            const text = escapeHtml(segment.text);
            return segment.bold ? `<b>${text}</b>` : `<span style="font-weight: 400;">${text}</span>`;
        }).join('');
    }

    function extractLanguageBlock(block) {
        const description = cleanText(block.paragraphs[0] && block.paragraphs[0].text);
        if (!description || !new RegExp(`^\\(${block.language}\\)`, 'i').test(description)) {
            throw new Error(`Không tìm thấy mô tả (${block.language}) trong Word.`);
        }

        const bodyParagraphs = block.paragraphs.slice(1);
        const stream = buildTextStream(bodyParagraphs);
        const boldGroups = findBoldGroups(stream);
        const nutritionHeading = boldGroups.find(group => /\b100\b/.test(group.text));
        if (!nutritionHeading) {
            throw new Error(`Không tìm thấy tiêu đề NTF có mốc 100 trong khối ${block.language}.`);
        }

        const ingredientsHeading = boldGroups.find(group => group.start < nutritionHeading.start);
        if (!ingredientsHeading) {
            throw new Error(`Không tìm thấy tiêu đề thành phần trong khối ${block.language}.`);
        }

        const ingredientsEndOfHeading = extendHeadingEnd(stream.text, ingredientsHeading.end);
        const nutritionEndOfHeading = extendHeadingEnd(stream.text, nutritionHeading.end);
        const ingredientSegments = sliceStreamSegments(
            stream.segments,
            ingredientsEndOfHeading,
            nutritionHeading.start
        );
        const ingredientText = cleanText(ingredientSegments.map(segment => segment.text).join(''));
        if (!ingredientText) {
            throw new Error(`Phần thành phần trong khối ${block.language} đang trống.`);
        }

        const nutritionSentence = findNutritionSentence(stream.text.slice(nutritionEndOfHeading));
        const metrics = parseMetrics(nutritionSentence, block.language);

        return {
            language: block.language,
            description,
            ingredientsHeading: normalizeHeading(stream.text.slice(ingredientsHeading.start, ingredientsEndOfHeading)),
            ingredientSegments,
            nutritionHeading: normalizeHeading(stream.text.slice(nutritionHeading.start, nutritionEndOfHeading)),
            metrics
        };
    }

    function buildLanguageHtml(data) {
        const lines = [
            `<p><i>${escapeHtml(data.description)}</i></p>`,
            `<p><b>${escapeHtml(data.ingredientsHeading)}</b> ${formatIngredientSegments(data.ingredientSegments)}</p>`,
            `<p><b>${escapeHtml(data.nutritionHeading)}</b></p>`
        ];

        for (const metric of data.metrics) {
            const content = `${metric.subItem ? '- ' : ''}${metric.label}: ${metric.value}`;
            if (metric.subItem) lines.push(`<p><i>${escapeHtml(content)}</i></p>`);
            else lines.push(`<p><b>${escapeHtml(metric.label)}:</b> ${escapeHtml(metric.value)}</p>`);
        }

        const storage = STORAGE_TRANSLATIONS[data.language];
        if (!storage) throw new Error(`Thiếu bản dịch Storage cho ngôn ngữ ${data.language}.`);
        lines.push(`<p><strong>${escapeHtml(storage.heading)}</strong></p>`);
        lines.push(`<p><span>${escapeHtml(storage.text)}</span></p>`);
        return lines.join('\n');
    }

    function convert(identity, paragraphs) {
        const blocks = splitLanguageBlocks(paragraphs, identity.languages);
        const extracted = blocks.map(extractLanguageBlock);
        const html = extracted.map(buildLanguageHtml).join('\n');
        return {
            html,
            languages: extracted.map(item => item.language),
            extracted
        };
    }

    global.multilingualWordParser = {
        convert,
        splitLanguageBlocks,
        extractLanguageBlock,
        parseMetrics,
        buildLanguageHtml
    };
})(window);


