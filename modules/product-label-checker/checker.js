/**
 * Product Label Checker - Main Controller Module
 * Coordinates UI, Excel reading, and HTML label rendering for 16 European languages.
 */
(function (global) {
    'use strict';

    const LANGUAGES = [
        'EN', 'FR', 'DE', 'CZ', 'PL', 'HU', 'ES', 'PT',
        'IT', 'RO', 'SK', 'NL', 'DK', 'FI', 'SE', 'EE'
    ];

    const LANG_NAMES = {
        EN: 'English', FR: 'Français', DE: 'Deutsch', CZ: 'Čeština',
        PL: 'Polski', HU: 'Magyar', ES: 'Español', PT: 'Português',
        IT: 'Italiano', RO: 'Română', SK: 'Slovenčina', NL: 'Nederlands',
        DK: 'Dansk', FI: 'Suomi', SE: 'Svenska', EE: 'Eesti'
    };

    let initialized = false;
    let products = [];
    let selectedIndex = -1;
    let currentModalLang = null;

    // DOM element cache
    let el = {};

    function escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = String(str ?? '');
        return div.innerHTML;
    }

    function showToast(message) {
        if (!el.toast) return;
        el.toast.textContent = message;
        el.toast.hidden = false;
        clearTimeout(showToast._timer);
        showToast._timer = setTimeout(() => {
            if (el.toast) el.toast.hidden = true;
        }, 2800);
    }

    function isProductComplete(product) {
        if (!product || !product.languages) return false;
        return LANGUAGES.every(lang => Boolean(String(product.languages[lang] || '').trim()));
    }

    function updateStats() {
        const total = products.length;
        const complete = products.filter(isProductComplete).length;
        const missing = total - complete;

        if (el.statProducts) el.statProducts.textContent = total;
        if (el.statComplete) el.statComplete.textContent = complete;
        if (el.statMissing) el.statMissing.textContent = missing;
    }

    function renderProductList(filter = '') {
        if (!el.productList) return;
        const query = filter.trim().toLowerCase();
        el.productList.innerHTML = '';

        if (!products.length) {
            el.productList.innerHTML = '<div style="padding:16px;text-align:center;color:var(--muted);font-size:0.85rem">Chưa có sản phẩm nào. Hãy tải file Excel hoặc chọn Dữ liệu mẫu.</div>';
            return;
        }

        let renderedCount = 0;
        products.forEach((product, idx) => {
            const searchHaystack = `${product.code} ${product.name}`.toLowerCase();
            if (query && !searchHaystack.includes(query)) return;

            renderedCount++;
            const complete = isProductComplete(product);
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = `checker-product-item${idx === selectedIndex ? ' active' : ''}`;
            btn.innerHTML = `
                <strong>${escapeHtml(product.name)}</strong>
                <span>${escapeHtml(product.code)} · ${complete ? 'Đủ 16 ngôn ngữ' : 'Thiếu ngôn ngữ'}</span>
            `;
            btn.addEventListener('click', () => selectProduct(idx));
            el.productList.appendChild(btn);
        });

        if (renderedCount === 0) {
            el.productList.innerHTML = '<div style="padding:16px;text-align:center;color:var(--muted);font-size:0.85rem">Không tìm thấy sản phẩm phù hợp với từ khóa.</div>';
        }
    }

    function renderLabels() {
        if (!el.labelsGrid) return;
        el.labelsGrid.innerHTML = '';

        const product = products[selectedIndex];
        if (!product) {
            el.labelsGrid.innerHTML = '<div style="grid-column:1/-1;padding:40px;text-align:center;color:var(--muted)">Chưa có sản phẩm nào được chọn.</div>';
            return;
        }

        const validator = global.ProductLabelHtmlValidator;

        LANGUAGES.forEach(lang => {
            const rawContent = product.languages[lang] || '';
            const valResult = validator
                ? validator.validateLabel(rawContent, lang)
                : {
                    status: rawContent ? 'ready' : 'missing',
                    statusText: rawContent ? 'Sẵn sàng' : 'Chưa có',
                    cleanHtml: rawContent,
                    issues: []
                };

            const card = document.createElement('div');
            card.className = `checker-label-card status-${valResult.status}`;

            const badgeClass = `badge-${valResult.status}`;
            const langFullName = LANG_NAMES[lang] || lang;

            card.innerHTML = `
                <div class="checker-card-head">
                    <div class="checker-lang-badge">
                        <span class="checker-lang-pill">${lang}</span>
                        <span>${langFullName}</span>
                    </div>
                    <span class="checker-status-badge ${badgeClass}" title="${escapeHtml(valResult.issues.join('; '))}">${valResult.statusText}</span>
                </div>
                <div class="checker-paper-wrap">
                    <div class="checker-paper ${!valResult.cleanHtml ? 'empty-paper' : ''}">
                        ${valResult.cleanHtml || '<p>Chưa có nội dung nhãn cho ngôn ngữ này.</p>'}
                    </div>
                </div>
                <div class="checker-card-footer">
                    <button type="button" class="checker-zoom-btn" data-lang="${lang}">
                        🔍 Xem chi tiết & HTML
                    </button>
                </div>
            `;

            // Click paper or zoom button to open modal
            const paperEl = card.querySelector('.checker-paper');
            const zoomBtn = card.querySelector('.checker-zoom-btn');

            const openDetail = () => openModal(product, lang, valResult);
            paperEl.addEventListener('click', openDetail);
            zoomBtn.addEventListener('click', openDetail);

            el.labelsGrid.appendChild(card);
        });
    }

    function selectProduct(index) {
        if (!products.length) return;
        selectedIndex = Math.max(0, Math.min(index, products.length - 1));
        const product = products[selectedIndex];

        if (el.productTitle) {
            el.productTitle.textContent = product.name || '(Không có tên sản phẩm)';
        }
        if (el.productCode) {
            el.productCode.textContent = `Mã sản phẩm: ${product.code} · Dòng ${product.rowIndex || index + 2}`;
        }
        if (el.counter) {
            el.counter.textContent = `${selectedIndex + 1} / ${products.length}`;
        }
        if (el.prevBtn) {
            el.prevBtn.disabled = selectedIndex <= 0;
        }
        if (el.nextBtn) {
            el.nextBtn.disabled = selectedIndex >= products.length - 1;
        }

        renderProductList(el.searchInput ? el.searchInput.value : '');
        renderLabels();
    }

    function openModal(product, lang, valResult) {
        currentModalLang = lang;
        if (!el.modal) return;

        const langName = LANG_NAMES[lang] || lang;
        el.modalTitle.textContent = `[${lang}] ${langName} · ${product.name} (${product.code})`;

        // Render tab
        el.modalRender.innerHTML = valResult.cleanHtml || '<p style="color:var(--muted);font-style:italic">Chưa có nội dung nhãn.</p>';

        // Code tab
        const validator = global.ProductLabelHtmlValidator;
        const formattedCode = validator
            ? validator.formatHtmlCode(product.languages[lang] || '')
            : (product.languages[lang] || '');
        el.modalCode.textContent = formattedCode || '(Nội dung rỗng)';

        switchModalTab('render');
        el.modal.hidden = false;
        document.body.style.overflow = 'hidden';
    }

    function closeModal() {
        if (!el.modal) return;
        el.modal.hidden = true;
        document.body.style.overflow = '';
    }

    function switchModalTab(tab) {
        const isRender = tab === 'render';
        if (el.modalRenderTab) el.modalRenderTab.classList.toggle('active', isRender);
        if (el.modalCodeTab) el.modalCodeTab.classList.toggle('active', !isRender);
        if (el.modalRenderWrap) el.modalRenderWrap.hidden = !isRender;
        if (el.modalCodeWrap) el.modalCodeWrap.hidden = isRender;
    }

    function copyModalHtml() {
        const product = products[selectedIndex];
        if (!product || !currentModalLang) return;
        const raw = product.languages[currentModalLang] || '';
        if (!raw) {
            showToast('Không có nội dung để sao chép.');
            return;
        }

        navigator.clipboard.writeText(raw).then(() => {
            showToast(`Đã sao chép HTML nhãn [${currentModalLang}] vào clipboard.`);
        }).catch(() => {
            showToast('Không thể sao chép tự động. Hãy chọn text để copy.');
        });
    }

    function loadProducts(newProducts, statusLabel) {
        products = newProducts || [];
        selectedIndex = products.length ? 0 : -1;
        if (el.fileStatus) {
            el.fileStatus.textContent = statusLabel || `Đã tải ${products.length} sản phẩm`;
        }
        if (el.searchInput) {
            el.searchInput.value = '';
        }
        updateStats();
        selectProduct(0);
    }

    function handleFile(file) {
        if (!file) return;
        const reader = global.ProductLabelExcelReader;
        if (!reader) {
            showToast('Lỗi: Chưa tải module đọc Excel.');
            return;
        }

        showToast(`Đang đọc file: ${file.name}...`);
        if (el.fileStatus) el.fileStatus.textContent = `Đang đọc ${file.name}...`;

        reader.parseFile(file).then(result => {
            loadProducts(result.products, `${result.fileName} (${result.products.length} sản phẩm · Sheet: ${result.sheetName})`);
            showToast(`Đã đọc thành công ${result.products.length} sản phẩm từ sheet "${result.sheetName}".`);
        }).catch(err => {
            showToast(`Lỗi đọc file: ${err.message}`);
            if (el.fileStatus) el.fileStatus.textContent = 'Lỗi đọc file Excel';
        });
    }

    // High quality demo data for immediate UI inspection
    function getDemoProducts() {
        return [
            {
                rowIndex: 2,
                code: 'CANT0007',
                name: 'Cantabile Watermelon Ade 230ml KOR',
                languages: {
                    EN: '<p><i>(EN) Watermelon-flavored drink with sugar and sweetener.</i></p>\n<p><b>Ingredients:</b> <span style="font-weight: 400;">Water, sugar, watermelon concentrate 0.5%, flavouring (watermelon), acidity regulators (E330, E331(iii)), thickener (E415), sweetener (E955), colour (E129).</span></p>\n<p><b>Warning: E129 may have an adverse effect on activity and attention in children.</b></p>\n<p><b>Nutrition facts per 100 ml:</b></p>\n<p><b>Energy:</b> 236,48 kJ / 56,52 kcal</p>\n<p><b>Fat:</b> 0 g</p>\n<p><i>- Of which saturated: 0 g</i></p>\n<p><b>Carbohydrates:</b> 14,35 g</p>\n<p><i>- Of which sugars: 13,04 g</i></p>\n<p><b>Protein:</b> 0 g</p>\n<p><b>Salt:</b> 0,02 g</p>\n<p><strong>Storage.</strong></p>\n<p><span>Store at room temperature.</span></p>',
                    FR: '<p><i>(FR) Boisson aromatisée à la pastèque avec sucre et édulcorant.</i></p>\n<p><b>Ingrédients:</b> <span style="font-weight: 400;">Eau, sucre, concentré de pastèque 0,5%, arôme (pastèque), correcteurs d\'acidité (E330, E331(iii)), épaississant (E415), édulcorant (E955), colorant (E129).</span></p>\n<p><b>Avertissement: E129 peut avoir des effets indésirables sur l\'activité et l\'attention chez les enfants.</b></p>\n<p><b>Déclaration nutritionnelle pour 100 ml:</b> <span style="font-weight: 400;">Énergie: 236,48 kJ / 56,52 kcal, Matières grasses: 0 g, Dont saturés: 0 g, Glucides: 14,35 g, Dont sucres: 13,04 g, Protéines: 0 g, Sel: 0,02 g.</span></p>\n<p><b>Vendeur:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prague 4, République tchèque.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Téléphone:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    DE: '<p><i>(DE) Getränk mit Wassermelonengeschmack mit Zucker und Süßungsmittel.</i></p>\n<p><b>Zutaten:</b> <span style="font-weight: 400;">Wasser, Zucker, Wassermelonenkonzentrat 0,5%, Aroma (Wassermelone), Säureregulatoren (E330, E331(iii)), Verdickungsmittel (E415), Süßungsmittel (E955), Farbstoff (E129).</span></p>\n<p><b>Warnung: E129 kann die Aktivität und Aufmerksamkeit von Kindern beeinträchtigen.</b></p>\n<p><b>Nährwertangaben pro 100 ml:</b> <span style="font-weight: 400;">Energie: 236,48 kJ / 56,52 kcal, Fett: 0 g, Davon gesättigte Fettsäuren: 0 g, Kohlenhydrate: 14,35 g, Davon Zucker: 13,04 g, Eiweiß: 0 g, Salz: 0,02 g.</span></p>\n<p><b>Vertrieb:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tschechische Republik.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    CZ: '<p><i>(CZ) Nápoj s příchutí vodního melounu s cukrem a sladidlem.</i></p>\n<p><b>Složení:</b> <span style="font-weight: 400;">Voda, cukr, melounový koncentrát 0,5%, aroma (meloun), regulátory kyselosti (E330, E331(iii)), zahušťovadlo (E415), sladidlo (E955), barvivo (E129).</span></p>\n<p><b>Varování: E129 může nepříznivě ovlivňovat činnost a pozornost dětí.</b></p>\n<p><b>Výživové údaje na 100 ml:</b></p>\n<p><b>Energie:</b> 236,48 kJ / 56,52 kcal</p>\n<p><b>Tuky:</b> 0 g</p>\n<p><i>- Z toho nasycené mastné kyseliny: 0 g</i></p>\n<p><b>Sacharidy:</b> 14,35 g</p>\n<p><i>- Z toho cukry: 13,04 g</i></p>\n<p><b>Bílkoviny:</b> 0 g</p>\n<p><b>Sůl:</b> 0,02 g</p>\n<p><strong>Skladování.</strong></p>\n<p><span>Skladujte při pokojové teplotě.</span></p>',
                    PL: '<p><i>(PL) Napój o smaku arbuza z cukrem i substancją słodzącą.</i></p>\n<p><b>Składniki:</b> <span style="font-weight: 400;">Woda, cukier, koncentrat arbuzowy 0,5%, aromat (arbuz), regulatory kwasowości (E330, E331(iii)), substancja zagęszczająca (E415), substancja słodząca (E955), barwnik (E129).</span></p>\n<p><b>Ostrzeżenie: E129 może mieć szkodliwy wpływ na aktywność i skupienie uwagi u dzieci.</b></p>\n<p><b>Wartość odżywcza w 100 ml:</b> <span style="font-weight: 400;">Energia: 236,48 kJ / 56,52 kcal, Tłuszcz: 0 g, W tym kwasy tłuszczowe nasycone: 0 g, Węglowodany: 14,35 g, W tym cukry: 13,04 g, Białko: 0 g, Sól: 0,02 g.</span></p>\n<p><b>Sprzedawca:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Republika Czeska.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    HU: '<p><i>(HU) Görögdinnye ízű ital cukorral és édesítőszerrel.</i></p>\n<p><b>Összetevők:</b> <span style="font-weight: 400;">Víz, cukor, görögdinnye koncentrátum 0,5%, aroma (görögdinnye), savanyúságot szabályozó anyagok (E330, E331(iii)), sűrítőanyag (E415), édesítőszer (E955), színezék (E129).</span></p>\n<p><b>Figyelmeztetés: Az E129 a gyermekek tevékenységére és figyelmére káros hatást gyakorolhat.</b></p>\n<p><b>Tápérték 100 ml-ben:</b> <span style="font-weight: 400;">Energia: 236,48 kJ / 56,52 kcal, Zsír: 0 g, Amelyből telített zsírsavak: 0 g, Szénhidrát: 14,35 g, Amelyből cukrok: 13,04 g, Fehérje: 0 g, Só: 0,02 g.</span></p>\n<p><b>Forgalmazó:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prága 4, Cseh Köztársaság.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    ES: '<p><i>(ES) Bebida con sabor a sandía con azúcar y edulcorante.</i></p>\n<p><b>Ingredientes:</b> <span style="font-weight: 400;">Agua, azúcar, concentrado de sandía 0,5%, aroma (sandía), correctores de acidez (E330, E331(iii)), espesante (E415), edulcorante (E955), colorante (E129).</span></p>\n<p><b>Advertencia: E129 puede tener efectos negativos sobre la actividad y la atención de los niños.</b></p>\n<p><b>Información nutricional por 100 ml:</b> <span style="font-weight: 400;">Energía: 236,48 kJ / 56,52 kcal, Grasas: 0 g, De las cuales saturadas: 0 g, Hidratos de carbono: 14,35 g, De los cuales azúcares: 13,04 g, Proteínas: 0 g, Sal: 0,02 g.</span></p>\n<p><b>Vendedor:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, República Checa.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Teléfono:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    PT: '<p><i>(PT) Bebida com sabor a melancia com açúcar e adoçante.</i></p>\n<p><b>Ingredientes:</b> <span style="font-weight: 400;">Água, açúcar, concentrado de melancia 0,5%, aroma (melancia), reguladores de acidez (E330, E331(iii)), espessante (E415), adoçante (E955), corante (E129).</span></p>\n<p><b>Aviso: E129 pode ter um efeito adverso na atividade e na atenção das crianças.</b></p>\n<p><b>Declaração nutricional por 100 ml:</b> <span style="font-weight: 400;">Energia: 236,48 kJ / 56,52 kcal, Lípidos: 0 g, Dos quais saturados: 0 g, Hidratos de carbono: 14,35 g, Dos quais açúcares: 13,04 g, Proteínas: 0 g, Sal: 0,02 g.</span></p>\n<p><b>Vendedor:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, República Checa.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefone:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    IT: '<p><i>(IT) Bevanda aromatizzata all\'anguria con zucchero ed edulcorante.</i></p>\n<p><b>Ingredienti:</b> <span style="font-weight: 400;">Acqua, zucchero, concentrato di anguria 0,5%, aroma (anguria), regolatori di acidità (E330, E331(iii)), addensante (E415), edulcorante (E955), colorante (E129).</span></p>\n<p><b>Avvertenza: E129 può influire negativamente sull\'attività e l\'attenzione dei bambini.</b></p>\n<p><b>Valori nutrizionali per 100 ml:</b> <span style="font-weight: 400;">Energia: 236,48 kJ / 56,52 kcal, Grassi: 0 g, Di cui acidi grassi saturi: 0 g, Carboidrati: 14,35 g, Di cui zuccheri: 13,04 g, Proteine: 0 g, Sale: 0,02 g.</span></p>\n<p><b>Venditore:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Repubblica Ceca.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefono:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    RO: '<p><i>(RO) Băutură cu aromă de pepene roșu cu zahăr și îndulcitor.</i></p>\n<p><b>Ingrediente:</b> <span style="font-weight: 400;">Apă, zahăr, concentrat de pepene roșu 0,5%, aromă (pepene roșu), regulatori de aciditate (E330, E331(iii)), agent de îngroșare (E415), îndulcitor (E955), colorant (E129).</span></p>\n<p><b>Avertisment: E129 poate afecta negativ activitatea și atenția copiilor.</b></p>\n<p><b>Informații nutriționale per 100 ml:</b> <span style="font-weight: 400;">Energie: 236,48 kJ / 56,52 kcal, Grăsimi: 0 g, Din care acizi grași saturați: 0 g, Glucide: 14,35 g, Din care zaharuri: 13,04 g, Proteine: 0 g, Sare: 0,02 g.</span></p>\n<p><b>Vânzător:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praga 4, Republica Cehă.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    SK: '<p><i>(SK) Nápoj s príchuťou vodného melónu s cukrom a sladidlom.</i></p>\n<p><b>Zloženie:</b> <span style="font-weight: 400;">Voda, cukor, melónový koncentrát 0,5%, aróma (melón), regulátory kyslosti (E330, E331(iii)), zahusťovadlo (E415), sladidlo (E955), farbivo (E129).</span></p>\n<p><b>Upozornenie: E129 môže mať nepriaznivý vplyv na činnosť a pozornosť detí.</b></p>\n<p><b>Výživové údaje na 100 ml:</b> <span style="font-weight: 400;">Energia: 236,48 kJ / 56,52 kcal, Tuky: 0 g, Z toho nasýtené mastné kyseliny: 0 g, Sacharidy: 14,35 g, Z toho cukry: 13,04 g, Bielkoviny: 0 g, Soľ: 0,02 g.</span></p>\n<p><b>Predajca:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Česká republika.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefón:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    NL: '<p><i>(NL) Drank met watermeloensmaak met suiker en zoetstof.</i></p>\n<p><b>Ingrediënten:</b> <span style="font-weight: 400;">Water, suiker, watermeloenconcentraat 0,5%, aroma (watermeloen), zuurteregelaars (E330, E331(iii)), verdikkingsmiddel (E415), zoetstof (E955), kleurstof (E129).</span></p>\n<p><b>Waarschuwing: E129 kan de activiteit of oplettendheid van kinderen nadelig beïnvloeden.</b></p>\n<p><b>Voedingswaarden per 100 ml:</b> <span style="font-weight: 400;">Energie: 236,48 kJ / 56,52 kcal, Vetten: 0 g, Waarvan verzadigd: 0 g, Koolhydraten: 14,35 g, Waarvan suikers: 13,04 g, Eiwitten: 0 g, Zout: 0,02 g.</span></p>\n<p><b>Verkoper:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praag 4, Tsjechië.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefoon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    DK: '<p><i>(DK) Drik med vandmelonsmag med sukker og sødestof.</i></p>\n<p><b>Ingredienser:</b> <span style="font-weight: 400;">Vand, sukker, vandmelonkoncentrat 0,5%, aroma (vandmelon), surhedsregulerende midler (E330, E331(iii)), fortykningsmiddel (E415), sødestof (E955), farve (E129).</span></p>\n<p><b>Advarsel: E129 kan have en negativ virkning på børns aktivitet og koncentrationsevne.</b></p>\n<p><b>Næringsindhold pr 100 ml:</b> <span style="font-weight: 400;">Energi: 236,48 kJ / 56,52 kcal, Fedt: 0 g, Heraf mættede: 0 g, Kulhydrater: 14,35 g, Heraf sukker: 13,04 g, Protein: 0 g, Salt: 0,02 g.</span></p>\n<p><b>Sælger:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tjekkiet.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    FI: '<p><i>(FI) Vesimeloninmakuinen juoma sokerilla ja makeutusaineella.</i></p>\n<p><b>Ainesosat:</b> <span style="font-weight: 400;">Vesi, sokeri, vesimelonitiiviste 0,5%, aromi (vesimeloni), happamuudensäätöaineet (E330, E331(iii)), sakeuttamisaine (E415), makeutusaine (E955), väri (E129).</span></p>\n<p><b>Varoitus: E129 voi vaikuttaa haitallisesti lasten aktiivisuustasoon ja heikentää tarkkaavaisuutta.</b></p>\n<p><b>Ravintoarvot per 100 ml:</b> <span style="font-weight: 400;">Energia: 236,48 kJ / 56,52 kcal, Rasva: 0 g, Josta tyydyttynyttä: 0 g, Hiilihydraatit: 14,35 g, Josta sokereita: 13,04 g, Proteiini: 0 g, Suola: 0,02 g.</span></p>\n<p><b>Myyjä:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Tšekki.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    SE: '<p><i>(SE) Dryck med vattenmelonsmak med socker och sötningsmedel.</i></p>\n<p><b>Ingredienser:</b> <span style="font-weight: 400;">Vatten, socker, vattenmelonkoncentrat 0,5%, arom (vattenmelon), surhetsreglerande medel (E330, E331(iii)), förtjockningsmedel (E415), sötningsmedel (E955), färgämne (E129).</span></p>\n<p><b>Varning: E129 kan ha en negativ effekt på barns beteende och koncentration.</b></p>\n<p><b>Näringsvärde per 100 ml:</b> <span style="font-weight: 400;">Energi: 236,48 kJ / 56,52 kcal, Fett: 0 g, Varav mättat fett: 0 g, Kolhydrater: 14,35 g, Varav sockerarter: 13,04 g, Protein: 0 g, Salt: 0,02 g.</span></p>\n<p><b>Säljare:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tjeckien.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>',
                    EE: '<p><i>(EE) Arbuusimaitseline jook suhkru ja magusainega.</i></p>\n<p><b>Koostisosad:</b> <span style="font-weight: 400;">Vesi, suhkur, arbuusikontsentraat 0,5%, lõhna- ja maitseaine (arbuus), happesuse regulaatorid (E330, E331(iii)), paksendaja (E415), magusaine (E955), toiduvärv (E129).</span></p>\n<p><b>Hoiatus: E129 võib avaldada kahjulikku mõju laste aktiivsusele ja tähelepanuvõimele.</b></p>\n<p><b>Toitumisalane teave 100 ml kohta:</b> <span style="font-weight: 400;">Energia: 236,48 kJ / 56,52 kcal, Rasvad: 0 g, Millest küllastunud: 0 g, Süsivesikud: 14,35 g, Millest suhkrud: 13,04 g, Valgud: 0 g, Sool: 0,02 g.</span></p>\n<p><b>Müüja:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Praha 4, Tšehhi Vabariik.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefoon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>'
                }
            },
            {
                rowIndex: 20,
                code: 'AM0007',
                name: 'AM Strawberry 4D Fruit Gummy 65g CHN',
                languages: {
                    EN: '<p><i>(EN) Strawberry-flavored gummy candy.</i></p>\n<p><b>Ingredients:</b> <span style="font-weight: 400;">glucose syrup, sugar, water, gelatin, concentrated strawberry juice 3.4%, Humectant (E422), acidity regulators (E330), natural flavour, acidity regulators (E296), thickener (E440), coconut oil, concentrated fruit and vegetable juices (sweet potato concentrate, apple concentrate, purple carrot concentrate), glazing agent (E903), spirulina extract, colorant(E100), acidity regulators (E331).</span></p>\n<p><b>Nutrition facts per 100g:</b></p>\n<p><b>Energy:</b> 1345 kJ / 317 kcal</p>\n<p><b>Fat:</b> 0,2 g</p>\n<p><i>- Of which saturated: 0,2 g</i></p>\n<p><b>Carbohydrates:</b> 72 g</p>\n<p><i>- Of which sugars: 51 g</i></p>\n<p><b>Protein:</b> 6,6 g</p>\n<p><b>Salt:</b> 0,03 g</p>\n<p><strong>Storage.</strong></p>\n<p><span>Store at room temperature.</span></p>',
                    CZ: '<p><i>(CZ) Jahodový želé bonbon.</i></p>\n<p><b>Složení:</b> <span style="font-weight: 400;">glukózový sirup, cukr, voda, želatina, koncentrovaná jahodová šťáva 3,4%, zvlhčující látka (E422), regulátory kyselosti (E330), přírodní aroma, regulátory kyselosti (E296), zahušťovadlo (E440), kokosový olej, koncentrované ovocné a zeleninové šťávy (koncentrát ze sladkých brambor, koncentrát z jablek, koncentrát z fialové mrkve), leštící látka (E903), extrakt ze spiruliny, barvivo (E100), regulátory kyselosti (E331).</span></p>\n<p><b>Výživové údaje na 100g:</b></p>\n<p><b>Energie:</b> 1345 kJ / 317 kcal</p>\n<p><b>Tuky:</b> 0,2 g</p>\n<p><i>- Z toho nasycené mastné kyseliny: 0,2 g</i></p>\n<p><b>Sacharidy:</b> 72 g</p>\n<p><i>- Z toho cukry: 51 g</i></p>\n<p><b>Bílkoviny:</b> 6,6 g</p>\n<p><b>Sůl:</b> 0,03 g</p>\n<p><strong>Skladování.</strong></p>\n<p><span>Skladujte při pokojové teplotě.</span></p>',
                    DE: '<p><i>(DE) Fruchtgummi mit Erdbeergeschmack.</i></p>\n<p><b>Zutaten:</b> <span style="font-weight: 400;">Glukosesirup, Zucker, Wasser, Gelatine, konzentrierter Erdbeersaft 3,4%, Feuchthaltemittel (E422), Säureregulatoren (E330), natürliches Aroma, Säureregulatoren (E296), Verdickungsmittel (E440), Kokosöl, Überzugsmittel (E903), Spirulina-Extrakt, Farbstoff (E100), Säureregulatoren (E331).</span></p>\n<p><b>Nährwertangaben pro 100g:</b> <span style="font-weight: 400;">Energie: 1345 kJ / 317 kcal, Fett: 0,2 g, Davon gesättigte Fettsäuren: 0,2 g, Kohlenhydrate: 72 g, Davon Zucker: 51 g, Eiweiß: 6,6 g, Salz: 0,03 g.</span></p>\n<p><b>Vertrieb:</b> <span style="font-weight: 400;">PepisCR s.r.o., Leopoldova 2040/30 Chodov, 149 00 Prag 4, Tschechische Republik.</span></p>\n<p><b>Web:</b> <span style="font-weight: 400;">www.pepis.company, </span><b>Telefon:</b> <span style="font-weight: 400;">+420 608 999 488.</span></p>'
                }
            }
        ];
    }

    function init() {
        if (initialized) return;

        // Query DOM elements inside the container
        el = {
            container: document.getElementById('checker-view-root'),
            dropZone: document.getElementById('checker-drop-zone'),
            fileInput: document.getElementById('checker-file-input'),
            chooseFileBtn: document.getElementById('checker-choose-file-btn'),
            useDemoBtn: document.getElementById('checker-use-demo-btn'),
            fileStatus: document.getElementById('checker-file-status'),
            statProducts: document.getElementById('checker-stat-products'),
            statComplete: document.getElementById('checker-stat-complete'),
            statMissing: document.getElementById('checker-stat-missing'),
            searchInput: document.getElementById('checker-search-input'),
            productList: document.getElementById('checker-product-list'),
            productTitle: document.getElementById('checker-product-title'),
            productCode: document.getElementById('checker-product-code'),
            counter: document.getElementById('checker-counter'),
            prevBtn: document.getElementById('checker-prev-btn'),
            nextBtn: document.getElementById('checker-next-btn'),
            labelsGrid: document.getElementById('checker-labels-grid'),
            modal: document.getElementById('checker-modal'),
            modalTitle: document.getElementById('checker-modal-title'),
            modalRender: document.getElementById('checker-modal-render'),
            modalCode: document.getElementById('checker-modal-code'),
            modalClose: document.getElementById('checker-modal-close'),
            modalRenderTab: document.getElementById('checker-modal-render-tab'),
            modalCodeTab: document.getElementById('checker-modal-code-tab'),
            modalRenderWrap: document.getElementById('checker-modal-render-wrap'),
            modalCodeWrap: document.getElementById('checker-modal-code-wrap'),
            modalCopyBtn: document.getElementById('checker-modal-copy-btn'),
            toast: document.getElementById('checker-toast')
        };

        if (!el.container) {
            console.warn('[ProductLabelChecker] Container element #checker-view-root not found in DOM.');
            return;
        }

        // Event: File picker
        if (el.chooseFileBtn && el.fileInput) {
            el.chooseFileBtn.addEventListener('click', () => el.fileInput.click());
            el.fileInput.addEventListener('change', () => {
                const file = el.fileInput.files[0];
                if (file) handleFile(file);
            });
        }

        // Event: Drag & Drop
        if (el.dropZone) {
            ['dragenter', 'dragover'].forEach(type => {
                el.dropZone.addEventListener(type, e => {
                    e.preventDefault();
                    e.stopPropagation();
                    el.dropZone.classList.add('dragover');
                });
            });
            ['dragleave', 'drop'].forEach(type => {
                el.dropZone.addEventListener(type, e => {
                    e.preventDefault();
                    e.stopPropagation();
                    el.dropZone.classList.remove('dragover');
                });
            });
            el.dropZone.addEventListener('drop', e => {
                const file = e.dataTransfer && e.dataTransfer.files[0];
                if (file) handleFile(file);
            });
        }

        // Event: Use Demo Data
        if (el.useDemoBtn) {
            el.useDemoBtn.addEventListener('click', () => {
                loadProducts(getDemoProducts(), 'Dữ liệu mẫu (Cantabile Ade & AM Gummy)');
                showToast('Đã nạp 2 sản phẩm mẫu để kiểm tra giao diện.');
            });
        }

        // Event: Search Input
        if (el.searchInput) {
            el.searchInput.addEventListener('input', e => {
                renderProductList(e.target.value);
            });
            el.searchInput.addEventListener('keydown', e => {
                if (e.key === 'Enter') {
                    const q = el.searchInput.value.trim().toLowerCase();
                    const found = products.findIndex(p => `${p.code} ${p.name}`.toLowerCase().includes(q));
                    if (found >= 0) selectProduct(found);
                }
            });
        }

        // Event: Prev / Next buttons
        if (el.prevBtn) {
            el.prevBtn.addEventListener('click', () => selectProduct(selectedIndex - 1));
        }
        if (el.nextBtn) {
            el.nextBtn.addEventListener('click', () => selectProduct(selectedIndex + 1));
        }

        // Event: Modal interactions
        if (el.modalClose) {
            el.modalClose.addEventListener('click', closeModal);
        }
        if (el.modal) {
            el.modal.addEventListener('click', e => {
                if (e.target === el.modal) closeModal();
            });
        }
        if (el.modalRenderTab) {
            el.modalRenderTab.addEventListener('click', () => switchModalTab('render'));
        }
        if (el.modalCodeTab) {
            el.modalCodeTab.addEventListener('click', () => switchModalTab('code'));
        }
        if (el.modalCopyBtn) {
            el.modalCopyBtn.addEventListener('click', copyModalHtml);
        }

        // Event: Keyboard Shortcuts (Esc, ArrowLeft, ArrowRight)
        document.addEventListener('keydown', e => {
            if (el.container.closest('.tool-view') && el.container.closest('.tool-view').hidden) {
                return; // Tab not active
            }
            if (e.key === 'Escape' && el.modal && !el.modal.hidden) {
                closeModal();
            } else if (!el.modal || el.modal.hidden) {
                if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
                if (e.key === 'ArrowLeft' && selectedIndex > 0) {
                    selectProduct(selectedIndex - 1);
                } else if (e.key === 'ArrowRight' && selectedIndex < products.length - 1) {
                    selectProduct(selectedIndex + 1);
                }
            }
        });

        initialized = true;

        // Auto-load demo if no file loaded yet
        if (!products.length) {
            loadProducts(getDemoProducts(), 'Dữ liệu mẫu (Cantabile Ade & AM Gummy)');
        }
    }

    global.ProductLabelChecker = {
        init,
        isInitialized: () => initialized,
        loadProducts,
        handleFile
    };

})(typeof window !== 'undefined' ? window : this);
