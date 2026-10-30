// ── Summary image panel (pannello immagini riepilogo / lightbox) · client ───
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, pushModalStack/removeModalStack (dichiarati
// nell'inline script).
function registerSummaryImageData(title, items) {
    if (!Array.isArray(items) || items.length === 0) return '';
    summaryImageRegistryCounter += 1;
    const key = 'img-' + summaryImageRegistryCounter;
    summaryImageRegistry[key] = {
        title: title || 'Billeder',
        items: items
    };
    return key;
}

function getSummaryImageSrc(item) {
    if (!item) return '';
    if (item.type === 'url') return item.value;
    return '/image-file?path=' + encodeURIComponent(item.value || '');
}

function selectCompactImageItems(items) {
    const source = Array.isArray(items) ? items : [];
    const clean = source.filter(item => item && String(item.value || '').trim());
    if (clean.length <= 2) return clean;

    const pickByLabel = (pattern) => clean.find(item => pattern.test(String(item.label || '')));
    const picked = [];
    const pushUnique = (item) => {
        if (!item) return;
        const value = String(item.value || '').trim().toLowerCase();
        if (!value) return;
        if (picked.some(x => String(x.value || '').trim().toLowerCase() === value)) return;
        picked.push(item);
    };

    pushUnique(pickByLabel(/webpg|nesting/i));
    pushUnique(pickByLabel(/pictfnm|icon/i));
    for (const item of clean) {
        pushUnique(item);
        if (picked.length >= 2) break;
    }
    return picked.slice(0, 2);
}

function openCompactImageModal(imageKey) {
    const entry = summaryImageRegistry[imageKey];
    const modal = document.getElementById('compactImageModal');
    const titleEl = document.getElementById('compactImageTitle');
    const subtitleEl = document.getElementById('compactImageSubtitle');
    const bodyEl = document.getElementById('compactImageBody');
    if (!entry || !modal || !titleEl || !bodyEl) return;

    const items = selectCompactImageItems(entry.items);
    if (!items.length) return;

    titleEl.textContent = entry.title || 'Billeder';
    subtitleEl.textContent = 'Viser nesting + hovedbilleder';

    let html = '<div class="compact-image-grid">';
    for (const item of items) {
        const src = getSummaryImageSrc(item);
        html += '<div class="compact-image-card">';
        html += '<div class="compact-image-label">' + escapeHtml(item.label || 'Billede') + '</div>';
        html += '<img class="image-preview-zoomable" src="' + escapeHtml(src) + '" alt="' + escapeHtml(item.label || entry.title || 'Billede') + '" loading="lazy" data-fullsrc="' + escapeHtml(src) + '" data-title="' + escapeHtml(item.label || entry.title || 'Billede') + '" data-path="' + escapeHtml(item.value || '') + '" />';
        html += '<div class="compact-image-path">' + escapeHtml(item.value || '') + '</div>';
        html += '</div>';
    }
    html += '</div>';
    bodyEl.innerHTML = html;
    modal.classList.add('show');
    pushModalStack('compactImageModal');
}

function closeCompactImageModal(event) {
    if (event && event.target && event.target.id !== 'compactImageModal') return;
    const modal = document.getElementById('compactImageModal');
    const bodyEl = document.getElementById('compactImageBody');
    if (!modal) return;
    modal.classList.remove('show');
    if (bodyEl) bodyEl.innerHTML = '';
    removeModalStack('compactImageModal');
}

function closeSummaryImagePanel() {
    const panels = [
        document.getElementById('summaryImagePanel'),
        document.getElementById('laserImagePanel')
    ];
    for (const panel of panels) {
        if (!panel) continue;
        panel.innerHTML = '';
        panel.classList.add('hidden');
    }
    updateSummaryImagePanelLayout();
}

function updateSummaryImagePanelLayout() {
    const wrap = document.querySelector('#summaryModal .modal-content-wrap');
    const summaryPanel = document.getElementById('summaryImagePanel');
    if (!wrap || !summaryPanel) return;
    const hasImages = !summaryPanel.classList.contains('hidden') && summaryPanel.innerHTML.trim() !== '';
    const shouldFocus = hasImages && window.matchMedia('(max-width: 1440px)').matches;
    wrap.classList.toggle('image-focus', shouldFocus);
}

function openSummaryImagePanel(imageKey, preferredPanelId) {
    const modal = document.getElementById('summaryModal');
    const title = document.getElementById('summaryModalTitle');
    const laserPanelWrap = document.getElementById('laserOrderSummaryPanel');
    const laserPanel = document.getElementById('laserImagePanel');
    const summaryPanel = document.getElementById('summaryImagePanel');
    const isLaserVisible = laserPanelWrap && laserPanelWrap.style.display !== 'none';
    const isVisible = (el) => {
        if (!el) return false;
        const s = getComputedStyle(el);
        return s.display !== 'none' && s.visibility !== 'hidden' && el.getClientRects().length > 0;
    };
    let panel = null;
    if (preferredPanelId === 'laserImagePanel') {
        panel = isVisible(laserPanel) ? laserPanel : summaryPanel;
    } else if (preferredPanelId === 'summaryImagePanel') {
        panel = summaryPanel;
    } else {
        panel = (isLaserVisible && laserPanel) ? laserPanel : summaryPanel;
    }
    const entry = summaryImageRegistry[imageKey];
    if (!panel || !entry || !Array.isArray(entry.items) || entry.items.length === 0) {
        closeSummaryImagePanel();
        return;
    }

    if (panel.id === 'summaryImagePanel' && title) {
        title.textContent = entry.title || 'Billeder';
    }
    if (panel.id === 'summaryImagePanel' && modal && modal.style.display !== 'flex') {
        modal.style.display = 'flex';
        pushModalStack('summaryModal');
    }

    let html = '<div class="summary-image-panel-header">';
    html += '<div class="summary-image-panel-title">' + escapeHtml(entry.title) + '</div>';
    html += '<button class="summary-image-close" onclick="closeSummaryImagePanel()">Luk</button>';
    html += '</div>';
    html += '<div class="image-preview-gallery">';

    for (const item of entry.items) {
        const src = getSummaryImageSrc(item);
        html += '<div class="image-preview-card">';
        html += '<div class="image-preview-label">' + escapeHtml(item.label || 'Billede') + '</div>';
        html += '<img class="image-preview-zoomable" src="' + escapeHtml(src) + '" alt="' + escapeHtml(entry.title) + '" loading="lazy" data-fullsrc="' + escapeHtml(src) + '" data-title="' + escapeHtml(item.label || entry.title || 'Billede') + '" data-path="' + escapeHtml(item.value || '') + '" />';
        html += '<div class="image-preview-path">' + escapeHtml(item.value || '') + '</div>';
        html += '</div>';
    }

    html += '</div>';
    panel.innerHTML = html;
    panel.classList.remove('hidden');
    updateSummaryImagePanelLayout();
}

function openImageLightbox(src, title, pathText) {
    const lightbox = document.getElementById('imageLightbox');
    const img = document.getElementById('imageLightboxImg');
    const titleEl = document.getElementById('imageLightboxTitle');
    const pathEl = document.getElementById('imageLightboxPath');
    if (!lightbox || !img) return;

    img.src = src || '';
    img.alt = title || 'Billede';
    if (titleEl) titleEl.textContent = title || 'Billede';
    if (pathEl) pathEl.textContent = pathText || '';
    lightbox.classList.remove('hidden');
    pushModalStack('imageLightbox');
}

function closeImageLightbox(event) {
    if (event && event.target && event.target.id !== 'imageLightbox') return;
    const lightbox = document.getElementById('imageLightbox');
    const img = document.getElementById('imageLightboxImg');
    const pathEl = document.getElementById('imageLightboxPath');
    if (!lightbox || lightbox.classList.contains('hidden')) return;

    lightbox.classList.add('hidden');
    if (img) {
        img.src = '';
        img.alt = '';
    }
    if (pathEl) pathEl.textContent = '';
    removeModalStack('imageLightbox');
}

function updateSummaryModalBackBtn() {
    const backBtn = document.getElementById('summaryModalBackBtn');
    if (!backBtn) return;
    backBtn.classList.toggle('hidden', summaryModalHistory.length === 0);
}

function pushSummaryModalState() {
    const title = document.getElementById('summaryModalTitle');
    const body = document.getElementById('summaryModalBody');
    const imagePanel = document.getElementById('summaryImagePanel');
    if (!title || !body) return;
    summaryModalHistory.push({
        title: title.textContent,
        bodyHtml: body.innerHTML,
        imageHtml: imagePanel ? imagePanel.innerHTML : '',
        imageHidden: imagePanel ? imagePanel.classList.contains('hidden') : true
    });
    updateSummaryModalBackBtn();
}

function goSummaryModalBack() {
    if (summaryModalHistory.length === 0) return;
    const prev = summaryModalHistory.pop();
    const title = document.getElementById('summaryModalTitle');
    const body = document.getElementById('summaryModalBody');
    const imagePanel = document.getElementById('summaryImagePanel');
    if (title) title.textContent = prev.title;
    if (body) body.innerHTML = prev.bodyHtml;
    if (imagePanel) {
        imagePanel.innerHTML = prev.imageHtml || '';
        imagePanel.classList.toggle('hidden', prev.imageHidden !== false);
    }
    updateSummaryImagePanelLayout();
    updateSummaryModalBackBtn();
}
