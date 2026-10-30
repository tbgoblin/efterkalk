// ── Order detail report toggle + anteprima di stampa (ordine/lista) · client ──
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, pushModalStack/removeModalStack, applyMicroTablePolish,
// goToDashboard, openModule (dichiarati nell'inline script) e
// buildModulePrintCssText/buildModulePreviewReportHtml (assets/js/ordreindgang.js).
function toggleOrderDetailReport() {
    const report = document.getElementById('orderDetailReport');
    if (!report) return;
    const isClosed = report.style.display === 'none';
    report.style.display = isClosed ? '' : 'none';
}

let currentPrintPreviewMode = null;
let modulePrintPreviewState = null;
let reportPrintRestoreState = null;

function isOrderDetailReportViewActive() {
    const bodyEl = document.getElementById('orderDetailModalBody');
    return Boolean(bodyEl && bodyEl.querySelector('#orderDetailReport'));
}

function updateOrderDetailModalBackButton() {
    const btn = document.getElementById('orderDetailModalBackBtn');
    if (!btn) return;
    const show = Boolean(reportOriginState && isOrderDetailReportViewActive());
    btn.style.display = show ? '' : 'none';
}

function restoreOrderDetailFromReport() {
    if (!reportOriginState) return false;
    const snapshot = reportOriginState;
    reportOriginState = null;
    openOrderDetailModal(snapshot.html, snapshot.title, snapshot.subtitle);
    return true;
}

function goBackFromReportToOrder() {
    restoreOrderDetailFromReport();
}

function openOrderDetailModal(html, titleText, subtitleText) {
    const overlay = document.getElementById('orderDetailModal');
    const titleEl = document.getElementById('orderDetailModalTitle');
    const subtitleEl = document.getElementById('orderDetailModalSubtitle');
    const bodyEl = document.getElementById('orderDetailModalBody');
    if (!overlay || !titleEl || !subtitleEl || !bodyEl) return;
    titleEl.textContent = titleText || 'Ordre-rapport';
    subtitleEl.textContent = subtitleText || 'Manager-oversigt med produktion, cost og sporbarhed';
    bodyEl.innerHTML = html || '';
    applyMicroTablePolish(bodyEl);
    overlay.style.display = 'flex';
    pushModalStack('orderDetailModal');
    document.body.classList.add('report-modal-open');
    updateOrderDetailModalBackButton();
}

function closeOrderDetailModal(event) {
    if (event && event.target && event.target.id !== 'orderDetailModal') return;
    if (isOrderDetailReportViewActive() && reportOriginState) {
        restoreOrderDetailFromReport();
        return;
    }
    const overlay = document.getElementById('orderDetailModal');
    const bodyEl = document.getElementById('orderDetailModalBody');
    if (overlay) overlay.style.display = 'none';
    if (bodyEl) bodyEl.innerHTML = '';
    document.body.classList.remove('report-modal-open');
    reportOriginState = null;
    updateOrderDetailModalBackButton();
    removeModalStack('orderDetailModal');
    const returnModule = orderDetailReturnModule;
    orderDetailReturnModule = null;
    if (returnModule === 'dashboard') {
        goToDashboard();
        return;
    }
    if (returnModule === 'salgordre-via') {
        openModule('salgordre-via');
        return;
    }
    goBackToList();
}

function buildStandaloneReportPrintCss() {
    return [
        '@page { size: A4 portrait; margin: 12mm; }',
        'body { margin: 0; background: #fff; color: #10253f; font-family: Arial, sans-serif; }',
        '.order-detail-report { display:block !important; }',
        '.order-report-toolbar,',
        '.order-report-actions,',
        '.list-toggle-btn,',
        'button { display:none !important; }',
        '.section { break-inside: avoid; page-break-inside: avoid; margin-bottom: 14px; }',
        '.order-report-grid { grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; }',
        '.order-report-card { break-inside: avoid; page-break-inside: avoid; }',
        '.order-report-table { width:100%; border-collapse:collapse; }',
        '.order-report-table th, .order-report-table td { border-bottom:1px solid #dfe8f3; padding:8px 10px; }',
        '.order-report-table th { background:#eef5ff; }',
        '.summary-box { break-inside: avoid; page-break-inside: avoid; }'
    ].join('\n');
}

function buildStandaloneListPrintCss() {
    return [
        '@page { size: A4 portrait; margin: 7mm; }',
        'body { margin:0; font-family: Arial, sans-serif; color:#1f2937; font-size:10px; line-height:1.2; }',
        '.order-list-summary-actions, .order-report-actions, .search-box, .header-banner-wrapper { display:none !important; }',
        '.order-list-section { margin:0 !important; box-shadow:none !important; border:none !important; padding:0 !important; }',
        '.order-list-summary { margin:0 0 8px 0 !important; padding:7px 9px !important; font-size:10px !important; }',
        '.order-list-section h3 { margin:0 0 7px 0 !important; padding:0 0 5px 0 !important; font-size:12px !important; }',
        '.order-list-table { width:100%; font-size:9.3px; border-collapse:collapse; table-layout:auto; }',
        '.order-list-table th, .order-list-table td { padding:3px 4px; line-height:1.15; }',
        '.order-list-table th:nth-child(8), .order-list-table td:nth-child(8), .order-list-table th:nth-child(9), .order-list-table td:nth-child(9) { display:none; }',
        '.order-list-table tr { break-inside: avoid; page-break-inside: avoid; }'
    ].join('\n');
}

function printStandaloneHtml(title, html, cssText) {
    const iframe = document.createElement('iframe');
    iframe.setAttribute('aria-hidden', 'true');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    iframe.style.opacity = '0';
    document.body.appendChild(iframe);

    const styles = Array.from(document.querySelectorAll('style, link[rel="stylesheet"]'))
        .map(el => el.outerHTML)
        .join('\n');
    const safeTitle = escapeHtml(title || 'Udskrift');
    const doc = iframe.contentWindow.document;
    doc.open();
    doc.write('<!doctype html><html><head><meta charset="utf-8"><title>' + safeTitle + '</title>' + styles + '<style>' + (cssText || '') + '</style></head><body>' + (html || '') + '</body></html>');
    doc.close();

    setTimeout(() => {
        try {
            iframe.contentWindow.focus();
            iframe.contentWindow.print();
        } finally {
            setTimeout(() => iframe.remove(), 1500);
        }
    }, 250);
}

function printOrderDetailReport() {
    const bodyEl = document.getElementById('orderDetailModalBody');
    if (!bodyEl) return;
    const titleEl = document.getElementById('orderDetailModalTitle');
    const reportTitle = titleEl ? titleEl.textContent : 'Ordre-rapport';
    printStandaloneHtml(reportTitle, bodyEl.innerHTML, buildStandaloneReportPrintCss());
}

function renderPrintPreview(title, html, mode) {
    const overlay = document.getElementById('printPreviewOverlay');
    const titleEl = document.getElementById('printPreviewTitle');
    const bodyEl = document.getElementById('printPreviewBody');
    if (!overlay || !titleEl || !bodyEl) return;
    currentPrintPreviewMode = mode;
    titleEl.textContent = title;
    bodyEl.innerHTML = html;
    overlay.style.display = 'flex';
    pushModalStack('printPreviewOverlay');
    document.body.classList.add('print-preview-lock');
    document.body.classList.add('print-preview-mode');
    const dialog = overlay.querySelector('.print-preview-dialog');
    if (dialog && dialog.focus) dialog.focus();
}

function closePrintPreview(event) {
    if (event && event.target && event.target.id !== 'printPreviewOverlay') return;
    const overlay = document.getElementById('printPreviewOverlay');
    if (overlay) overlay.style.display = 'none';
    const bodyEl = document.getElementById('printPreviewBody');
    if (bodyEl) bodyEl.innerHTML = '';
    document.body.classList.remove('print-preview-lock');
    document.body.classList.remove('print-preview-mode');
    currentPrintPreviewMode = null;
    modulePrintPreviewState = null;
    removeModalStack('printPreviewOverlay');
}

function confirmPrintFromPreview() {
    const bodyEl = document.getElementById('printPreviewBody');
    if (!bodyEl) return;
    const titleEl = document.getElementById('printPreviewTitle');
    const title = titleEl ? titleEl.textContent : 'Forhåndsvisning';
    if (currentPrintPreviewMode === 'module' && modulePrintPreviewState) {
        const selected = modulePrintPreviewState.selected instanceof Set
            ? modulePrintPreviewState.selected
            : new Set();
        const html = buildModulePreviewReportHtml(modulePrintPreviewState, selected);
        const cssText = buildModulePrintCssText({ orientation: modulePrintPreviewState.orientation });
        printStandaloneHtml(title, html, cssText);
        return;
    }
    const cssText = currentPrintPreviewMode === 'list'
        ? buildStandaloneListPrintCss()
        : buildStandaloneReportPrintCss();
    printStandaloneHtml(title, bodyEl.innerHTML, cssText);
}

function openOrderListPrintPreview() {
    if (!orderListVisible) {
        toggleOrderList();
    }
    const listEl = document.getElementById('orderList');
    if (!listEl) return;
    renderPrintPreview('Forhåndsvisning - ordreliste', listEl.innerHTML, 'list');
}

function openOrderDetailPrintPreview() {
    const report = document.getElementById('orderDetailReport');
    if (report) {
        renderPrintPreview('Forhåndsvisning - rapport', report.outerHTML, 'report');
        return;
    }
    if (lastOrderReportHtml) {
        renderPrintPreview(lastOrderReportTitle || 'Forhåndsvisning - rapport', lastOrderReportHtml, 'report');
    }
}

function openLatestOrderReportPreview() {
    if (!lastOrderReportHtml) {
        alert('Rapporten er ikke klar endnu. Søg efter en ordre først.');
        return;
    }

    const overlay = document.getElementById('orderDetailModal');
    const bodyEl = document.getElementById('orderDetailModalBody');
    const titleEl = document.getElementById('orderDetailModalTitle');
    const subtitleEl = document.getElementById('orderDetailModalSubtitle');
    const modalOpen = overlay && getComputedStyle(overlay).display === 'flex';
    if (modalOpen && bodyEl && !isOrderDetailReportViewActive() && !reportOriginState) {
        reportOriginState = {
            html: bodyEl.innerHTML,
            title: titleEl ? titleEl.textContent : 'Ordre-rapport',
            subtitle: subtitleEl ? subtitleEl.textContent : 'Manager-oversigt med produktion, cost og sporbarhed'
        };
    }

    openOrderDetailModal(
        lastOrderReportHtml,
        lastOrderReportTitle || 'Rapport 2.0',
        'Manager-oversigt med produktion, cost og sporbarhed'
    );
}

function updateReportOpenButtonState(isReady, ordNo) {
    const btn = document.getElementById('openReportBtn');
    if (!btn) return;
    const ready = Boolean(isReady && lastOrderReportHtml);
    btn.disabled = !ready;
    if (ready) {
        btn.title = 'Åbn seneste rapport i separat visning';
        btn.textContent = ordNo ? ('Rapport ' + ordNo) : 'Rapport 2.0';
    } else {
        btn.title = 'Søg efter en ordre for at aktivere rapport';
        btn.textContent = 'Rapport 2.0';
    }
}

window.addEventListener('afterprint', function() {
    document.body.classList.remove('print-report-mode');
    document.body.classList.remove('print-list-mode');
    document.body.classList.remove('print-preview-mode');
    const report = document.getElementById('orderDetailReport');
    if (report && reportPrintRestoreState !== null) {
        report.style.display = reportPrintRestoreState;
        reportPrintRestoreState = null;
    }
    const overlay = document.getElementById('printPreviewOverlay');
    if (overlay) overlay.style.display = 'none';
    const bodyEl = document.getElementById('printPreviewBody');
    if (bodyEl) bodyEl.innerHTML = '';
    currentPrintPreviewMode = null;
    modulePrintPreviewState = null;
});

function updateOrderListSummaryPanel() {
    const summaryEl = document.getElementById('orderListSummary');
    if (!summaryEl || !orderListVisible) return;
    const orders = getFilteredOrders();
    summaryEl.innerHTML = buildOrderListSummaryHtml(orders);
}
