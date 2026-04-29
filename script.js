document.addEventListener('DOMContentLoaded', () => {
    const controls = document.getElementById('controls');
    const searchInput = document.getElementById('search-input');
    const productGrid = document.getElementById('product-grid');
    const loading = document.getElementById('loading');
    const emptyState = document.getElementById('empty-state');
    const stats = document.getElementById('stats');
    const liveRefreshBtn = document.getElementById('live-refresh-btn');
    const viewListBtn = document.getElementById('view-list');
    const viewGridBtn = document.getElementById('view-grid');

    let productsData = [];
    let currentSheetName = "";
    let currentRawJson = [];
    window.globalWorkbook = null;

    // Find and Replace State
    const btnFindReplace = document.getElementById('btn-find-replace');
    const findReplaceModal = document.getElementById('find-replace-modal');
    const closeModalBtn = document.getElementById('close-modal-btn');
    const findInput = document.getElementById('find-input');
    const replaceInput = document.getElementById('replace-input');
    const btnFindPrev = document.getElementById('btn-find-prev');
    const btnFindNext = document.getElementById('btn-find-next');
    const btnReplace = document.getElementById('btn-replace');
    const btnReplaceAll = document.getElementById('btn-replace-all');
    const findStatus = document.getElementById('find-status');

    let findMatches = [];
    let currentMatchIndex = -1;

    // View Toggles
    if (viewListBtn && viewGridBtn) {
        viewListBtn.addEventListener('click', () => {
            productGrid.classList.add('list-view');
            viewListBtn.classList.add('active');
            viewGridBtn.classList.remove('active');
        });
        viewGridBtn.addEventListener('click', () => {
            productGrid.classList.remove('list-view');
            viewGridBtn.classList.add('active');
            viewListBtn.classList.remove('active');
        });
    }

    // --- HARDCODED CREDENTIALS ---
    const GOOGLE_SHEET_ID = "1TfHwy5JcQNMVLzPkkIRQcl2Q_M40ctdQynB5y2yh30Q";
    const SYNC_URL = "https://script.google.com/macros/s/AKfycbxZidSr37rM-p1-LEAM2Y15LNRqNZqY_YZ2BkPqhTCoMAC0DaRblG1hx8avyQPkfgX5/exec";

    // Mapped Headers (STRICT TO FORCE CLEAN SHEETS)
    const titleHeaders = ['product name', 'اسم المنتج'];
    const descHeaders = ['product description', 'وصف المنتج'];
    const linkHeaders = ['رابط الصورة', 'image link', 'drive link', 'link', 'url', 'google drive', 'صورة'];

    // Automatically load data on page open
    loadDefaultSheet();

    if (liveRefreshBtn) {
        liveRefreshBtn.addEventListener('click', () => {
            const icon = liveRefreshBtn.querySelector('i');
            icon.classList.add('bx-spin');
            
            // Add a timestamp to bust the cache so we instantly get new sheets/data
            const exportUrl = `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/export?format=xlsx&t=${new Date().getTime()}`;
            fetch(exportUrl, { cache: 'no-store' })
                .then(response => response.arrayBuffer())
                .then(data => {
                    const workbook = XLSX.read(data, { type: 'array' });
                    window.globalWorkbook = workbook;
                    renderTabs(workbook, currentSheetName);
                    icon.classList.remove('bx-spin');
                })
                .catch(err => {
                    console.error(err);
                    icon.classList.remove('bx-spin');
                });
        });
    }

    function loadDefaultSheet() {
        // Add a timestamp to bust the cache
        const exportUrl = `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/export?format=xlsx&t=${new Date().getTime()}`;

        showLoading("جاري جلب المستند...");
        
        fetch(exportUrl, { cache: 'no-store' })
            .then(response => {
                if (!response.ok) throw new Error('Network response was not ok');
                return response.arrayBuffer();
            })
            .then(data => {
                try {
                    // Parse ArrayBuffer via SheetJS
                    const workbook = XLSX.read(data, { type: 'array' });
                    window.globalWorkbook = workbook;
                    renderTabs(workbook);
                } catch (err) {
                    console.error(err);
                    showEmptyState("حدث خطأ أثناء معالجة البيانات.");
                }
            })
            .catch(err => {
                console.error(err);
                showEmptyState("فشل في تحميل قاعدة البيانات. تأكد من أن الرابط عام.");
            });
    }

    function renderTabs(workbook, targetSheetName = null) {
        const tabsContainer = document.getElementById('tabs-container');
        tabsContainer.innerHTML = '';
        
        if (workbook.SheetNames.length === 0) {
            showEmptyState("المستند فارغ ولا يحتوي على أي أوراق.");
            return;
        }
        
        tabsContainer.classList.remove('hidden');

        let sheetToLoad = targetSheetName && workbook.SheetNames.includes(targetSheetName) 
            ? targetSheetName 
            : workbook.SheetNames[0];

        workbook.SheetNames.forEach((sheetName) => {
            const btn = document.createElement('button');
            btn.className = `tab-btn ${sheetName === sheetToLoad ? 'active' : ''}`;
            
            // Calculate unchecked items
            const rawJson = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
            let validProducts = 0;
            let checkedProducts = 0;

            // Check if sheet is messy
            let foundTitle = false;
            let foundDesc = false;
            if (rawJson.length > 0) {
                const sheetHeaders = Object.keys(rawJson[0]);
                sheetHeaders.forEach(h => {
                    const lowerH = h.toLowerCase().trim();
                    if (titleHeaders.includes(lowerH)) foundTitle = true;
                    if (descHeaders.includes(lowerH)) foundDesc = true;
                });
            }
            const isMessy = (rawJson.length > 0) && (!foundTitle || !foundDesc);

            rawJson.forEach(row => {
                let ptitle = "";
                let pimage = "";
                for (let key in row) {
                    const lkey = key.toLowerCase().trim();
                    if (titleHeaders.includes(lkey)) ptitle = row[key];
                    if (linkHeaders.includes(lkey)) pimage = row[key];
                }
                if (ptitle && pimage) {
                    validProducts++;
                    if (localStorage.getItem(`done_${sheetName}_${ptitle}`) === 'true') {
                        checkedProducts++;
                    }
                }
            });

            const unchecked = validProducts - checkedProducts;
            let badgeHtml = '';
            
            if (isMessy) {
                badgeHtml = `<span class="tab-badge" style="background: transparent; border: none; font-size: 1.1rem; padding: 0;">⚠️</span>`;
            } else {
                badgeHtml = unchecked > 0 
                    ? `<span class="tab-badge">${unchecked}</span>` 
                    : `<span class="tab-badge done"><i class='bx bx-check'></i></span>`;
            }


            btn.innerHTML = `${sheetName} ${badgeHtml}`;
            
            btn.addEventListener('click', () => {
                // If clicking the current tab, do nothing
                if (btn.classList.contains('active')) return;

                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                
                // Add fade out effect
                productGrid.style.opacity = '0';
                productGrid.style.transform = 'translateY(10px)';
                
                setTimeout(() => {
                    loadSheet(sheetName);
                    // Fade back in
                    requestAnimationFrame(() => {
                        productGrid.style.opacity = '1';
                        productGrid.style.transform = 'translateY(0)';
                    });
                }, 200); // 200ms transition
            });
            
            tabsContainer.appendChild(btn);
        });

        // Load targeted sheet
        loadSheet(sheetToLoad);
    }

    function loadSheet(sheetName) {
        currentSheetName = sheetName;
        const worksheet = window.globalWorkbook.Sheets[sheetName];
        
        if (!worksheet) {
            showEmptyState("الورقة المطلوبة غير متوفرة.");
            return;
        }
        
        currentRawJson = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        processData(currentRawJson);
    }

    // --- LIVE BACKGROUND POLLING (AUTO-SYNC) ---
    function pollForChanges() {
        const exportUrl = `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/export?format=xlsx&t=${new Date().getTime()}`;
        
        fetch(exportUrl, { cache: 'no-store' })
            .then(response => response.arrayBuffer())
            .then(data => {
                const newWorkbook = XLSX.read(data, { type: 'array' });
                
                // Compare tabs for any newly added sheets
                const oldTabs = window.globalWorkbook ? window.globalWorkbook.SheetNames.join(',') : '';
                const newTabs = newWorkbook.SheetNames.join(',');
                
                if (oldTabs !== newTabs) {
                    window.globalWorkbook = newWorkbook;
                    renderTabs(newWorkbook, currentSheetName);
                } else if (currentSheetName && newWorkbook.Sheets[currentSheetName]) {
                    // Check if the current sheet's products have changed
                    const newRawJson = XLSX.utils.sheet_to_json(newWorkbook.Sheets[currentSheetName], { defval: "" });
                    
                    if (JSON.stringify(newRawJson) !== JSON.stringify(currentRawJson)) {
                        // Crucial: Only interrupt the user by rebuilding the UI if they are NOT actively typing!
                        const isUserTyping = document.activeElement && document.activeElement.classList.contains('shopify-content');
                        
                        if (!isUserTyping) {
                            currentRawJson = newRawJson;
                            window.globalWorkbook = newWorkbook;
                            // Silently refresh products
                            processData(currentRawJson);
                        }
                    }
                }
            })
            .catch(err => console.error("Silent sync failed:", err));
    }

    // Ping Google Sheets silently every 5 seconds to look for new tabs or edits
    setInterval(pollForChanges, 5000);

    searchInput.addEventListener('input', (e) => {
        renderProducts(e.target.value);
    });

    function processData(rawJson) {
        if (!rawJson || rawJson.length === 0) {
            showEmptyState();
            return;
        }

        const headers = Object.keys(rawJson[0]);
        let titleCol = '', descCol = '', linkCol = '';
        let missingHeaders = false;

        headers.forEach(h => {
            const lowerH = h.toLowerCase().trim();
            if (!titleCol && titleHeaders.includes(lowerH)) titleCol = h;
            if (!descCol && descHeaders.includes(lowerH)) descCol = h;
            if (!linkCol && linkHeaders.includes(lowerH)) linkCol = h;
        });

        // If strict headers were not found, banner warning
        if (!titleCol || !descCol) {
            missingHeaders = true;
        }

        // Fallbacks
        if (!titleCol && headers.length > 0) titleCol = headers[0];
        if (!descCol && headers.length > 1) descCol = headers[1];
        if (!linkCol && headers.length > 2) linkCol = headers[2];

        // UI WARNING
        let existingBanner = document.getElementById('messy-sheet-banner');
        if (missingHeaders) {
            if (!existingBanner) {
                const banner = document.createElement('div');
                banner.id = 'messy-sheet-banner';
                banner.innerHTML = `<i class='bx bx-error-circle'></i> <strong>تنبيه دقيق:</strong> لم يتم العثور على أعمدة (اسم المنتج / وصف المنتج) في هذا القسم. سيتم عرض المنتجات ولكن <strong>لن يتم حفظ أي تعديلات</strong>. يرجى تصحيح أسماء الأعمدة في جوجل شيت!`;
                banner.style.cssText = "background: #ffebee; color: #c62828; padding: 15px; text-align: center; font-weight: bold; margin: 20px; border-radius: 8px; border: 2px solid #ef9a9a; display: flex; align-items: center; justify-content: center; gap: 10px;";
                document.querySelector('#controls').insertAdjacentElement('afterend', banner);
            }
        } else if (existingBanner) {
            existingBanner.remove();
        }

        // Format data
        productsData = rawJson.map(row => ({
            title: row[titleCol] || 'منتج بدون عنوان',
            description: row[descCol] || '<p>لا يوجد وصف متاح.</p>',
            link: row[linkCol] || ''
        })).filter(p => p.title !== 'منتج بدون عنوان' || p.description !== '<p>لا يوجد وصف متاح.</p>');

        loading.classList.add('hidden');
        
        if (productsData.length === 0) {
            showEmptyState();
        } else {
            emptyState.classList.add('hidden'); // CRITICAL: explicitly hide error overlay
            controls.classList.remove('hidden');
            productGrid.classList.remove('hidden');
            searchInput.value = '';
            renderProducts();
        }
    }

    function renderProducts(query = '') {
        const lowerQuery = query.toLowerCase();
        const filtered = productsData.filter(p => {
            const strippedDesc = p.description.replace(/<[^>]*>?/gm, '');
            return p.title.toLowerCase().includes(lowerQuery) || 
                   strippedDesc.toLowerCase().includes(lowerQuery);
        });

        stats.textContent = `تم العثور على ${filtered.length} منتج`;
        productGrid.innerHTML = '';

        if (filtered.length === 0) {
            productGrid.innerHTML = `<div class="empty-state glass-panel" style="grid-column: 1/-1;">
                <i class='bx bx-search'></i>
                <h2>لا يوجد نتائج مطابقة</h2>
                <p>حاول تعديل البحث.</p>
            </div>`;
            return;
        }

        filtered.forEach(product => {
            const card = document.createElement('div');
            card.className = 'product-card';
            
            let validLink = product.link;
            if (validLink && validLink !== '#' && !validLink.startsWith('http')) {
                validLink = 'https://' + validLink;
            }

            let isImage = false;
            let displaySrc = validLink;
            
            if (validLink) {
                let fileId = null;
                const idMatch1 = validLink.match(/[?&]id=([a-zA-Z0-9_-]+)/);
                const idMatch2 = validLink.match(/\/d\/([a-zA-Z0-9_-]+)/);
                
                if (idMatch1) fileId = idMatch1[1];
                else if (idMatch2) fileId = idMatch2[1];

                if (fileId && (validLink.includes('drive.google.com') || validLink.includes('docs.google.com'))) {
                    isImage = true;
                    displaySrc = `https://lh3.googleusercontent.com/d/${fileId}`;
                } else if (validLink.match(/\.(jpeg|jpg|gif|png|webp)$/i)) {
                    isImage = true;
                }
            }

            let imageHtml = '';
            if (isImage) {
                imageHtml = `<img src="${displaySrc}" alt="${escapeHTML(product.title)}" class="product-image" loading="lazy">`;
            }

            let footerHtml = '';
            if (validLink) {
                if (isImage) {
                    footerHtml = `
                    <div class="product-footer">
                        <a href="${validLink}" target="_blank" rel="noopener noreferrer" class="drive-btn image-link-btn">
                            <i class='bx bx-image'></i> عرض الصورة الأصلية
                        </a>
                    </div>`;
                } else {
                    footerHtml = `
                    <div class="product-footer">
                        <a href="${validLink}" target="_blank" rel="noopener noreferrer" class="drive-btn">
                            <i class='bx bx-link-external'></i> عرض الرابط
                        </a>
                    </div>`;
                }
            }

            card.innerHTML = `
                ${imageHtml}
                <div class="product-header" style="display: flex; justify-content: space-between; align-items: flex-start; gap: 1rem;">
                    <h2 class="editable-title" contenteditable="true" data-original-title="${escapeHTML(product.title)}" style="flex: 1; outline: none; padding: 4px; border-radius: 4px; transition: background 0.2s;" title="انقر لتعديل العنوان" onfocus="this.style.background='rgba(255,255,255,0.05)'" onblur="this.style.background='transparent'">${escapeHTML(product.title)}</h2>
                    <button class="mark-done-btn" title="تحديد كمكتمل">
                        <i class='bx bx-check-circle'></i>
                    </button>
                </div>
                <div class="product-body">
                    <p class="shopify-label">الوصف (Description)</p>
                    <div class="shopify-editor">
                        <div class="shopify-toolbar">
                            <button title="AI Generation"><i class='bx bx-sparkles'></i></button>
                            <div class="toolbar-divider"></div>
                            <div class="toolbar-select">Paragraph <i class='bx bx-chevron-down'></i></div>
                            <div class="toolbar-divider"></div>
                            <button title="Bold"><b>B</b></button>
                            <button title="Italic"><i style="font-family: serif;">I</i></button>
                            <button title="Underline"><u>U</u></button>
                            <button title="Text Color"><span style="border-bottom: 2px solid #000; line-height: 1.1;">A</span> <i class='bx bx-chevron-down' style="font-size: 12px; margin-right: 2px;"></i></button>
                            <div class="toolbar-divider"></div>
                            <button title="Alignment" style="width: auto; padding: 0 8px;"><i class='bx bx-align-right'></i> <i class='bx bx-chevron-down' style="font-size: 12px; margin-right: 4px;"></i></button>
                            <div class="toolbar-divider"></div>
                            <button title="Link"><i class='bx bx-link'></i></button>
                            <button title="Image"><i class='bx bx-image-alt'></i></button>
                            <button title="Video"><i class='bx bx-play-circle'></i></button>
                            <button title="Table" style="width: auto; padding: 0 8px;"><i class='bx bx-table'></i> <i class='bx bx-chevron-down' style="font-size: 12px; margin-right: 4px;"></i></button>
                            <div class="toolbar-divider"></div>
                            <button title="More"><i class='bx bx-dots-horizontal-rounded'></i></button>
                            <div class="toolbar-spacer"></div>
                            <span class="save-status" style="font-size: 12px; margin-left: 8px;"></span>
                            <button title="Code"><i class='bx bx-code-alt'></i></button>
                        </div>
                        <div class="shopify-content html-content" contenteditable="true" data-title="${escapeHTML(product.title)}">
                            ${product.description}
                        </div>
                    </div>
                </div>
                ${footerHtml}
            `;
            // Check if marked as done previously in localStorage
            const doneKey = `done_${currentSheetName}_${product.title}`;
            if (localStorage.getItem(doneKey) === 'true') {
                card.classList.add('is-done');
                const btn = card.querySelector('.mark-done-btn');
                if(btn) btn.classList.add('active');
            }

            const doneBtn = card.querySelector('.mark-done-btn');
            if (doneBtn) {
                doneBtn.addEventListener('click', () => {
                    const isCurrentlyDone = card.classList.toggle('is-done');
                    doneBtn.classList.toggle('active');
                    localStorage.setItem(doneKey, isCurrentlyDone);
                    updateTabBadge(currentSheetName);
                    
                    // Fetch the live updated title in case it was just edited
                    const activeTitle = card.querySelector('.editable-title').getAttribute('data-original-title') || product.title;

                    // Immediately dispatch to Google Sheet to color the row
                    const params = new URLSearchParams();
                    params.append('sheetName', currentSheetName);
                    params.append('title', activeTitle);
                    params.append('markDone', isCurrentlyDone.toString());

                    fetch(SYNC_URL, {
                        method: 'POST',
                        mode: 'no-cors',
                        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                        body: params.toString()
                    }).catch(err => console.error("Could not sync done status", err));
                });
            }

            productGrid.appendChild(card);
        });
    }

    function updateTabBadge(sheetName) {
        if (!window.globalWorkbook || !window.globalWorkbook.Sheets[sheetName]) return;
        const rawJson = XLSX.utils.sheet_to_json(window.globalWorkbook.Sheets[sheetName], { defval: "" });
        let valid = 0, checked = 0;
        rawJson.forEach(row => {
            let ptitle = "", pimage = "";
            for (let key in row) {
                const lkey = key.toLowerCase().trim();
                if (titleHeaders.includes(lkey)) ptitle = row[key];
                if (linkHeaders.includes(lkey)) pimage = row[key];
            }
            if (ptitle && pimage) {
                valid++;
                if (localStorage.getItem(`done_${sheetName}_${ptitle}`) === 'true') checked++;
            }
        });
        
        const unchecked = valid - checked;
        const tabs = document.querySelectorAll('.tab-btn');
        tabs.forEach(tab => {
            if (tab.textContent.includes(sheetName)) {
                // Keep the text name but replace the span badge
                const badgeHtml = unchecked > 0 
                    ? `<span class="tab-badge">${unchecked}</span>` 
                    : `<span class="tab-badge done"><i class='bx bx-check'></i></span>`;
                tab.innerHTML = `${sheetName} ${badgeHtml}`;
            }
        });
    }

    function showLoading(msg) {
        emptyState.classList.add('hidden');
        productGrid.classList.add('hidden');
        controls.classList.add('hidden');
        loading.classList.remove('hidden');
        if (msg) loading.querySelector('p').textContent = msg;
    }

    function showEmptyState(msg) {
        loading.classList.add('hidden');
        productGrid.classList.add('hidden');
        emptyState.classList.remove('hidden');
        if (msg) emptyState.querySelector('p').textContent = msg;
    }

    function escapeHTML(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    let autoSaveTimeouts = {};

    productGrid.addEventListener('input', (e) => {
        if (e.target.classList.contains('shopify-content') || e.target.classList.contains('editable-title')) {
            const card = e.target.closest('.product-card');
            
            const titleEl = card.querySelector('.editable-title');
            const descEl = card.querySelector('.shopify-content');
            
            const originalTitle = titleEl.getAttribute('data-original-title');
            const newTitle = titleEl.innerText.trim() || originalTitle;
            
            // Get innerHTML and Minify it (removes the raw \n newlines that result in gaps in Google Sheets)
            let newDescription = descEl.innerHTML;
            newDescription = newDescription.replace(/(\r\n|\n|\r)/gm, ""); // Remove all raw newlines
            newDescription = newDescription.replace(/>\s+</g, "><"); // Remove empty spaces between HTML tags
            newDescription = newDescription.trim();
            
            const statusSpan = card.querySelector('.save-status');
            
            // Visual feedback for typing
            statusSpan.innerHTML = "<i class='bx bx-pencil'></i> جاري الكتابة...";
            statusSpan.style.color = '#8b5cf6'; // Primary color
            
            // Clear existing timeout
            clearTimeout(autoSaveTimeouts[originalTitle]);
            
            // Debounce for 1 second of inactivity before saving
            autoSaveTimeouts[originalTitle] = setTimeout(() => {
                // Block saving if headers are messy
                if (document.getElementById('messy-sheet-banner')) {
                    statusSpan.innerHTML = "<i class='bx bx-error'></i> خطأ في جدول البيانات!";
                    statusSpan.style.color = '#dc2626'; // Red
                    alert("لا يمكن الحفظ! يرجى إصلاح أسماء الأعمدة في Google Sheet أولاً.");
                    return;
                }
                
                statusSpan.innerHTML = "<i class='bx bx-loader-alt bx-spin'></i> جاري الحفظ...";
                statusSpan.style.color = '#eab308'; // Yellow

                // Use URLSearchParams with x-www-form-urlencoded to skip CORS
                const params = new URLSearchParams();
                params.append('action', 'updateProductDescription');
                params.append('sheetName', currentSheetName);
                params.append('title', originalTitle); // Original title finds the target row in the Google Sheet
                params.append('newTitle', newTitle);   // Only used if they upgrade the Apps Script
                params.append('description', newDescription);

                fetch(SYNC_URL, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: {
                        'Content-Type': 'application/x-www-form-urlencoded',
                    },
                    body: params.toString()
                }).then(() => {
                    statusSpan.innerHTML = "<i class='bx bx-check-circle'></i> تم الحفظ";
                    statusSpan.style.color = '#10b981'; // Green
                    
                    // Update the raw JSON so our live-polling script doesn't revert the title back instantly
                    const rawProd = currentRawJson.find(p => p.title === originalTitle);
                    if(rawProd) {
                        rawProd.title = newTitle;
                        rawProd.description = newDescription;
                    }
                    // Update dataset to maintain tracking over consecutive edits
                    titleEl.setAttribute('data-original-title', newTitle);
                    descEl.setAttribute('data-title', newTitle);
                    
                    setTimeout(() => {
                        statusSpan.textContent = '';
                    }, 3000);
                }).catch(error => {
                    console.error('Error saving:', error);
                    statusSpan.innerHTML = "<i class='bx bx-error-circle'></i> خطأ";
                    statusSpan.style.color = '#ef4444'; // Red
                });
            }, 1000);
        }
    });

    // ==========================================
    // FIND AND REPLACE LOGIC
    // ==========================================
    if (btnFindReplace && findReplaceModal) {
        const modalContent = findReplaceModal.querySelector('.modal-content');
        
        btnFindReplace.addEventListener('click', () => {
            findReplaceModal.classList.remove('hidden');
            findInput.focus();
        });

        closeModalBtn.addEventListener('click', () => {
            findReplaceModal.classList.add('hidden');
            modalContent.classList.remove('docked');
            clearHighlights();
        });

        findReplaceModal.addEventListener('click', (e) => {
            if (e.target === findReplaceModal) {
                findReplaceModal.classList.add('hidden');
                modalContent.classList.remove('docked');
                clearHighlights();
            }
        });
        
        function clearHighlights() {
            document.querySelectorAll('.highlight-card').forEach(el => el.classList.remove('highlight-card'));
            const sel = window.getSelection();
            if (sel) sel.removeAllRanges();
        }
        
        function escapeRegExp(string) {
            return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        }

        // Safely replace only in text nodes to prevent breaking HTML
        function replaceTextInElement(element, regex, replacement) {
            const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null, false);
            const nodes = [];
            while(walker.nextNode()) nodes.push(walker.currentNode);
            
            nodes.forEach(node => {
                if (node.nodeValue.match(regex)) {
                    node.nodeValue = node.nodeValue.replace(regex, replacement);
                }
            });
        }

        function performSearch() {
            const query = findInput.value;
            findMatches = [];
            currentMatchIndex = -1;
            
            if (!query) {
                findStatus.textContent = "أدخل نصاً للبحث";
                clearHighlights();
                return;
            }

            const regex = new RegExp(escapeRegExp(query), 'gi');
            const cards = document.querySelectorAll('.product-card');
            
            cards.forEach((card) => {
                const elements = [card.querySelector('.editable-title'), card.querySelector('.shopify-content')];
                
                elements.forEach(el => {
                    if (!el) return;
                    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null, false);
                    let node;
                    while(node = walker.nextNode()) {
                        regex.lastIndex = 0;
                        let match;
                        while ((match = regex.exec(node.nodeValue)) !== null) {
                            findMatches.push({
                                card,
                                element: el,
                                node: node,
                                startIndex: match.index,
                                endIndex: match.index + match[0].length
                            });
                        }
                    }
                });
            });

            if (findMatches.length > 0) {
                findStatus.textContent = `تم العثور على ${findMatches.length} نتيجة`;
                currentMatchIndex = 0;
                
                // Dock the modal to the right side so it doesn't block the view
                const modalContent = findReplaceModal.querySelector('.modal-content');
                modalContent.classList.add('docked');
                
                highlightCurrentMatch();
            } else {
                findStatus.textContent = "لم يتم العثور على نتائج";
                clearHighlights();
            }
        }

        function highlightCurrentMatch() {
            clearHighlights();
            if (findMatches.length === 0 || currentMatchIndex < 0) return;
            
            const match = findMatches[currentMatchIndex];
            match.card.classList.add('highlight-card');
            match.card.scrollIntoView({ behavior: 'smooth', block: 'center' });
            
            try {
                const range = document.createRange();
                range.setStart(match.node, match.startIndex);
                range.setEnd(match.node, match.endIndex);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
            } catch (e) {
                console.error("Selection error:", e);
            }
            
            findStatus.textContent = `نتيجة ${currentMatchIndex + 1} من ${findMatches.length}`;
        }

        
        
        if(btnFindPrev) {
            btnFindPrev.addEventListener('click', () => {
                if (findMatches.length === 0) {
                    performSearch();
                } else {
                    currentMatchIndex = (currentMatchIndex - 1 + findMatches.length) % findMatches.length;
                    highlightCurrentMatch();
                }
            });
        }
        
        btnFindNext.addEventListener('click', () => {
            if (findMatches.length === 0) {
                performSearch();
            } else {
                currentMatchIndex = (currentMatchIndex + 1) % findMatches.length;
                highlightCurrentMatch();
            }
        });

        findInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                if (findMatches.length > 0 && findInput.value === findInput.dataset.lastQuery) {
                    btnFindNext.click();
                } else {
                    findInput.dataset.lastQuery = findInput.value;
                    performSearch();
                }
            }
        });

        function replaceCurrent() {
            if (findMatches.length === 0 || currentMatchIndex < 0) return;
            
            const replacement = replaceInput.value;
            const match = findMatches[currentMatchIndex];
            
            const originalText = match.node.nodeValue;
            if (originalText === null) {
                // Node detached, resync
                performSearch();
                return;
            }

            const before = originalText.substring(0, match.startIndex);
            const after = originalText.substring(match.endIndex);
            match.node.nodeValue = before + replacement + after;
            
            // Adjust subsequent match offsets in the same text node
            const lengthDiff = replacement.length - (match.endIndex - match.startIndex);
            for (let i = currentMatchIndex + 1; i < findMatches.length; i++) {
                if (findMatches[i].node === match.node) {
                    findMatches[i].startIndex += lengthDiff;
                    findMatches[i].endIndex += lengthDiff;
                }
            }
            
            match.element.dispatchEvent(new Event('input', { bubbles: true }));
            findMatches.splice(currentMatchIndex, 1);
            
            if (findMatches.length > 0) {
                if (currentMatchIndex >= findMatches.length) currentMatchIndex = 0;
                highlightCurrentMatch();
            } else {
                currentMatchIndex = -1;
                findStatus.textContent = "لا يوجد المزيد من النتائج";
                clearHighlights();
            }
        }

        btnReplace.addEventListener('click', () => {
            if (findMatches.length === 0) performSearch();
            replaceCurrent();
        });

        btnReplaceAll.addEventListener('click', () => {
            performSearch();
            if (findMatches.length === 0) return;
            
            const query = findInput.value;
            const replacement = replaceInput.value;
            if (!query) return;
            
            const regex = new RegExp(escapeRegExp(query), 'gi');
            const uniqueElements = [...new Set(findMatches.map(m => m.element))];
            const totalMatches = findMatches.length;
            
            findStatus.textContent = `جاري استبدال ${totalMatches} نتيجة...`;
            
            let i = 0;
            function processNext() {
                if (i >= uniqueElements.length) {
                    findStatus.textContent = `تم استبدال ${totalMatches} بنجاح!`;
                    findMatches = [];
                    currentMatchIndex = -1;
                    clearHighlights();
                    return;
                }
                
                const el = uniqueElements[i];
                replaceTextInElement(el, regex, replacement);
                
                const inputEvent = new Event('input', { bubbles: true });
                el.dispatchEvent(inputEvent);
                
                i++;
                setTimeout(processNext, 50);
            }
            
            processNext();
        });
    }

});
