/**
 * Product Label Checker - HTML Validator Module
 * Sanitizes and validates HTML labels for the 16 target European languages.
 */
(function (global) {
    'use strict';

    const ALL_LANGUAGES = [
        'EN', 'FR', 'DE', 'CZ', 'PL', 'HU', 'ES', 'PT',
        'IT', 'RO', 'SK', 'NL', 'DK', 'FI', 'SE', 'EE'
    ];

    const MARKER_REGEX = /(?:^|[>\s])(?:▶|►)?\s*\(\s*([A-Za-z]{2})\s*\)/g;

    function sanitizeHtml(rawHtml) {
        if (!rawHtml) return '';

        // Use DOMPurify if available
        if (global.DOMPurify && typeof global.DOMPurify.sanitize === 'function') {
            return global.DOMPurify.sanitize(rawHtml, {
                ALLOWED_TAGS: [
                    'p', 'b', 'i', 'strong', 'em', 'span', 'table', 'tbody', 'thead',
                    'tr', 'th', 'td', 'ul', 'ol', 'li', 'br', 'hr', 'div'
                ],
                ALLOWED_ATTR: ['style', 'class']
            });
        }

        // Safe browser DOMParser fallback
        try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(String(rawHtml || ''), 'text/html');
            const dangerousSelectors = 'script, style, iframe, object, embed, link, meta, form, input, button, audio, video';
            doc.querySelectorAll(dangerousSelectors).forEach(node => node.remove());

            doc.body.querySelectorAll('*').forEach(node => {
                [...node.attributes].forEach(attr => {
                    const name = attr.name.toLowerCase();
                    const val = attr.value.trim().toLowerCase();
                    if (name.startsWith('on') || val.startsWith('javascript:') || val.startsWith('data:text/html')) {
                        node.removeAttribute(attr.name);
                    }
                });
            });

            return doc.body.innerHTML;
        } catch {
            return String(rawHtml)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
        }
    }

    function extractMarkers(htmlString) {
        const markers = [];
        let match;
        const rx = new RegExp(MARKER_REGEX.source, 'gi');
        while ((match = rx.exec(htmlString)) !== null) {
            const code = match[1].toUpperCase();
            if (ALL_LANGUAGES.includes(code)) {
                markers.push(code);
            }
        }
        return markers;
    }

    function validateLabel(htmlContent, targetLanguage) {
        const raw = String(htmlContent || '').trim();
        const targetLang = String(targetLanguage || '').toUpperCase();

        if (!raw) {
            return {
                status: 'missing',
                statusText: 'Chưa có dữ liệu',
                cleanHtml: '',
                issues: ['Nội dung rỗng (chưa có nhãn cho ngôn ngữ này)'],
                markers: [],
                foreignMarkers: [],
                hasExpectedMarker: false
            };
        }

        const cleanHtml = sanitizeHtml(raw);
        const markers = extractMarkers(cleanHtml);
        const expectedMarker = targetLang;
        const hasExpectedMarker = markers.includes(expectedMarker);
        const foreignMarkers = markers.filter(m => m !== expectedMarker);

        const issues = [];

        // Check foreign marker contamination
        if (foreignMarkers.length > 0) {
            const uniqueForeign = [...new Set(foreignMarkers)];
            issues.push(`Phát hiện khối marker ngoại ngữ thừa: ${uniqueForeign.join(', ')}`);
        }

        // Check if marker matches
        if (!hasExpectedMarker) {
            issues.push(`Thiếu marker ngôn ngữ chuẩn (${expectedMarker})`);
        } else if (markers[0] !== expectedMarker) {
            issues.push(`Marker đầu tiên không khớp với cột (${markers[0]} thay vì ${expectedMarker})`);
        }

        // Content checks
        const textContent = cleanHtml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
        if (textContent.length < 30) {
            issues.push('Nội dung nhãn quá ngắn (dưới 30 ký tự)');
        }

        let status = 'ready';
        let statusText = 'Sẵn sàng';

        if (foreignMarkers.length > 0) {
            status = 'contamination';
            statusText = 'Nhiễm ngoại ngữ';
        } else if (issues.length > 0) {
            status = 'warning';
            statusText = 'Cảnh báo';
        }

        return {
            status,
            statusText,
            cleanHtml,
            textContent,
            issues,
            markers,
            foreignMarkers,
            hasExpectedMarker
        };
    }

    function formatHtmlCode(html) {
        if (!html) return '';
        // Pretty format paragraphs
        return html
            .replace(/><p/g, '>\n<p')
            .replace(/<\/p><p/g, '</p>\n<p')
            .replace(/><table/g, '>\n<table')
            .replace(/<\/table><p/g, '</table>\n<p')
            .trim();
    }

    global.ProductLabelHtmlValidator = {
        ALL_LANGUAGES,
        sanitizeHtml,
        validateLabel,
        formatHtmlCode
    };

})(typeof window !== 'undefined' ? window : this);
