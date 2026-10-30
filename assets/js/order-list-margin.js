// ── System status, warmup polling, margini e filtro/ordinamento Order List · client ──
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, formatNumber, scheduleDashboardWidgets,
// renderOrderList, openOrderListPrintPreview (dichiarati nell'inline script, i
// due ultimi più avanti) e isOrderMarkedCreditNote/isOrderMarkedUB
// (assets/js/order-notes.js).
function setSystemStatus(text, bgColor, textColor) {
    // Header right area now shows user greeting instead of system status.
}

// Warmup progress bar polling
let warmupPollTimer = null;
let warmupTopBarHideScheduled = false;
let warmupCombinedReady = false;
let warmupCombinedPct = 0;
let warmupCombinedDone = 0;
let warmupCombinedTotal = 0;
let showDashboardWarmupNotice = false;
function startWarmupPolling() {
    if (warmupPollTimer) return;
    const wrap = document.getElementById('warmupBarWrap');
    const fill = document.getElementById('warmupBarFill');
    const txt  = document.getElementById('warmupBarText');
    const dashText = document.getElementById('dashboardWarmupText');
    const dashMeta = document.getElementById('dashboardWarmupMeta');
    const dashFill = document.getElementById('dashboardWarmupFill');
    const dashPct = document.getElementById('dashboardWarmupPct');
    const dashWrap = document.getElementById('dashboardWarmupNotice');
    if (!wrap && !dashText) return;

    warmupPollTimer = setInterval(async () => {
        try {
            const r = await fetch('/warmup-status');
            if (!r.ok) return;
            const d = await r.json();

            const totalCombined = Number(d.combinedTotal || d.total || 0);
            const doneCombined = Number(d.combinedDone || d.done || 0);
            const pctCombined = Number(d.combinedPct || d.pct || 0);
            const readyCombined = d.ready === true;

            warmupCombinedPct = Math.max(0, Math.min(100, pctCombined));
            warmupCombinedDone = Math.max(0, doneCombined);
            warmupCombinedTotal = Math.max(0, totalCombined);
            warmupCombinedReady = readyCombined || (!d.running && warmupCombinedTotal === 0);

            if (dashFill) dashFill.style.width = String(Math.max(0, Math.min(100, pctCombined))) + '%';
            if (dashPct) dashPct.textContent = String(Math.max(0, Math.min(100, pctCombined))) + '%';

            // Keep warmup hidden on initial dashboard screen, unless user explicitly triggered cache reset.
            if (dashWrap) {
                const shouldShowDashWarmup = d.running || (!readyCombined && totalCombined > 0) || showDashboardWarmupNotice;
                dashWrap.classList.toggle('hidden', !shouldShowDashWarmup);
            }

            if (dashText) {
                if (d.running) {
                    dashText.textContent = 'Forbereder ' + doneCombined + '/' + totalCombined + ' ordredata...';
                    if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                } else if (readyCombined && totalCombined > 0) {
                    dashText.textContent = 'Klar! Efterkalk-data er forberedt.';
                    if (dashMeta) dashMeta.textContent = 'Åbn Efterkalk når som helst.';
                    if (dashWrap) {
                        setTimeout(() => {
                            dashWrap.classList.add('hidden');
                            showDashboardWarmupNotice = false;
                        }, 1800);
                    }
                } else if (totalCombined > 0) {
                    dashText.textContent = 'Afventer baggrundsjob...';
                    if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                } else {
                    dashText.textContent = 'Venter på warmup-status...';
                    if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                }
            }

            if (d.total === 0) {
                if (wrap) wrap.classList.remove('active');
                return;
            }

            if (wrap) wrap.classList.add('active');
            if (fill) fill.style.width = d.pct + '%';

            if (d.running) {
                if (txt) txt.textContent = 'Forberegner ' + d.done + '/' + d.total + ' ordrer...';
                warmupTopBarHideScheduled = false;
            } else {
                if (txt) txt.textContent = 'Klar! ' + d.loaded + ' nye + ' + d.cached + ' fra cache';
                if (fill) fill.style.width = '100%';
                if (!warmupTopBarHideScheduled) {
                    warmupTopBarHideScheduled = true;
                    setTimeout(() => {
                        if (wrap) wrap.classList.remove('active');
                        warmupTopBarHideScheduled = false;
                    }, 3000);
                }
            }
        } catch(e) {
            // ignore polling errors silently
        }
    }, 800);
}
function updateSystemStatusFromOrders(orders) {
    if (!orders || orders.length === 0) {
        setSystemStatus('System klar', '#e8f5e9', '#1b5e20');
        return;
    }

    const visibleOrders = orders.slice(0, MARGIN_PREFETCH_ROWS);
    const total = visibleOrders.length;
    let completed = 0;

    for (const o of visibleOrders) {
        const state = getMarginState(o.OrdNo);
        if (state && (state.status === 'success' || state.status === 'error')) {
            completed += 1;
        }
    }

    if (completed >= total) {
        setSystemStatus('System klar', '#e8f5e9', '#1b5e20');
        return;
    }

    setSystemStatus('System indlæser... ' + completed + '/' + total, '#fff3cd', '#8a6d3b');
}

function getMarginModeLabel() {
    return currentMarginMode === 'new'
        ? 'Ny (Salg/Kost x 100)'
        : 'Klassisk ((Salg-Kost)/Salg x 100)';
}

function calculateOrderMarginPercent(revenue, cost) {
    if (currentMarginMode === 'new') {
        return cost > 0 ? ((revenue / cost) * 100) : 0;
    }
    return revenue > 0 ? (((revenue - cost) / revenue) * 100) : 0;
}

function calculateLineMarginPercent(salesPrice, lineCost) {
    if (currentMarginMode === 'new') {
        return lineCost > 0 ? ((salesPrice / lineCost) * 100) : 0;
    }
    return salesPrice > 0 ? (((salesPrice - lineCost) / salesPrice) * 100) : 0;
}

function toggleMarginMode() {
    currentMarginMode = currentMarginMode === 'new' ? 'classic' : 'new';
    renderOrderList();
}

function scheduleOrderListRerender() {
    if (orderListRerenderTimer) return;
    orderListRerenderTimer = setTimeout(() => {
        orderListRerenderTimer = null;
        renderOrderList();
    }, 120);
}

function getMarginState(ordNo) {
    return marginStateByOrdNo[String(ordNo)] || null;
}

function getOrderInvoiceStatusHtml(ordNo) {
    const marginState = getMarginState(ordNo);
    if (!marginState || marginState.status !== 'success') return '<span style="color:#999;">-</span>';
    if (marginState.hasInvoiceWarning) {
        return '<span title="En eller flere linjer mangler faktura (NoInvo=0); kostberegning bruger NoFin som fallback." style="display:inline-flex;align-items:center;gap:4px;background:#fff3e0;color:#e65100;font-size:12px;font-weight:600;padding:2px 7px;border-radius:10px;border:1px solid #ffcc80;">🧾 Mangler</span>';
    }
    const orderRow = (Array.isArray(orderListData) ? orderListData : []).find(o => Number(o.OrdNo || 0) === Number(ordNo || 0));
    const invoicedAmount = Number((orderRow && orderRow.InvoAm) || 0);
    const remainingInvoiceAmount = Number((orderRow && orderRow.DInvoIF) || 0);
    if (invoicedAmount > 0 && remainingInvoiceAmount > 0) {
        return '<span title="Ordren er delvist faktureret." style="display:inline-flex;align-items:center;gap:4px;background:#fff3e0;color:#e65100;font-size:12px;font-weight:600;padding:2px 7px;border-radius:10px;border:1px solid #ffcc80;">⏳ Delvis faktura</span>';
    }
    return '<span style="color:#388e3c;font-size:13px;" title="Alle fakturaer registreret.">✓</span>';
}

function updateOrderInvoiceCell(ordNo) {
    const listEl = document.getElementById('orderList');
    if (!listEl) return;
    const cells = listEl.querySelectorAll('.order-invoice-cell[data-ordno="' + ordNo + '"]');
    const html = getOrderInvoiceStatusHtml(ordNo);
    for (const cell of cells) { cell.innerHTML = html; }
}

function getOrderMarginHtml(ordNo) {
    const marginState = getMarginState(ordNo);
    let marginHtml = '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>';
    let toneClass = 'na';
    if (marginState && marginState.status === 'loading') {
        marginHtml = '<span style="background:#546e7a; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">...</span>';
        toneClass = 'na';
    } else if (marginState && marginState.status === 'success') {
        const marginValue = calculateOrderMarginPercent(marginState.totalRevenue || 0, marginState.totalCost || 0);
        const margin = marginValue.toFixed(2);
        marginHtml = getMarginBadge(margin);
        if (currentMarginMode === 'new') {
            toneClass = marginValue >= 125 ? 'ok' : (marginValue >= 105 ? 'warn' : 'bad');
        } else {
            toneClass = marginValue > 20 ? 'ok' : (marginValue >= 5 ? 'warn' : 'bad');
        }
    }
    return '<span class="order-margin-wrap"><span class="order-kpi-tone ' + toneClass + '" title="KPI tone fra Rapport 2.0">↗</span>' + marginHtml + '</span>';
}

function updateOrderMarginCell(ordNo) {
    scheduleDashboardWidgets();
    const listEl = document.getElementById('orderList');
    if (!listEl) return;
    const cells = listEl.querySelectorAll('.order-margin-cell[data-ordno="' + ordNo + '"]');
    if (!cells || cells.length === 0) return;
    const marginHtml = getOrderMarginHtml(ordNo);
    for (const cell of cells) {
        cell.innerHTML = marginHtml;
    }
    updateOrderListSummaryPanel();
}

function refreshOrderListStatus() {
    if (!orderListVisible) return;
    const visibleOrders = getFilteredOrders().slice(0, MARGIN_PREFETCH_ROWS);
    updateSystemStatusFromOrders(visibleOrders);
}

function scheduleMarginSortRefresh() {
    if (orderListSortField !== 'margin') return;
    if (marginSortRefreshTimer) return;
    marginSortRefreshTimer = setTimeout(() => {
        marginSortRefreshTimer = null;
        renderOrderList();
    }, 350);
}

function hydrateMarginStateFromOrderList(orders) {
    marginStateByOrdNo = {};
    for (const o of orders) {
        const ordNo = Number(o.OrdNo);
        if (!Number.isFinite(ordNo)) continue;

        if (o.TotalCost !== null && o.TotalCost !== undefined) {
            marginStateByOrdNo[String(ordNo)] = {
                status: 'success',
                totalRevenue: Number(o.InvoAm || 0),
                // Kun efterkalk-listen: kost inkl. stykliste-fallback
                totalCost: Number(o.TotalCost || 0) + Number(o.StyklisteFallbackCost || 0)
            };
        }
    }
}

function queueMarginLoad(ordNos) {
    for (const ordNo of ordNos) {
        const key = String(ordNo);
        const existing = marginStateByOrdNo[key];
        if (existing && (existing.status === 'success' || existing.status === 'loading')) {
            continue;
        }

        marginStateByOrdNo[key] = { status: 'loading' };
        marginJobQueue.push(Number(ordNo));
        updateOrderMarginCell(ordNo);
    }
    pumpMarginQueue();
    refreshOrderListStatus();
}

async function loadSingleOrderMargin(ordNo) {
    const key = String(ordNo);
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), MARGIN_FETCH_TIMEOUT_MS);
    try {
        const response = await fetch('/order-margin/' + ordNo, { signal: controller.signal });
        let data = null;
        try {
            data = await response.json();
        } catch {
            data = { error: 'Invalid JSON response' };
        }
        if (!response.ok || data.error) {
            marginStateByOrdNo[key] = { status: 'error' };
            updateOrderMarginCell(ordNo);
            refreshOrderListStatus();
            scheduleMarginSortRefresh();
            return;
        }

        marginStateByOrdNo[key] = {
            status: 'success',
            totalRevenue: Number(data.totalRevenue || 0),
            // Kun efterkalk-listen: kost inkl. stykliste-fallback
            totalCost: Number(data.totalCost || 0) + Number(data.styklisteFallbackCost || 0),
            hasInvoiceWarning: Boolean(data.hasInvoiceWarning)
        };
        updateOrderMarginCell(ordNo);
        updateOrderInvoiceCell(ordNo);
        refreshOrderListStatus();
        scheduleMarginSortRefresh();
    } catch (err) {
        marginStateByOrdNo[key] = { status: 'error' };
        updateOrderMarginCell(ordNo);
        refreshOrderListStatus();
        scheduleMarginSortRefresh();
    } finally {
        clearTimeout(timeoutId);
    }
}

function pumpMarginQueue() {
    while (marginWorkerActiveCount < MARGIN_MAX_CONCURRENT && marginJobQueue.length > 0) {
        const ordNo = marginJobQueue.shift();
        marginWorkerActiveCount += 1;

        loadSingleOrderMargin(ordNo)
            .finally(() => {
                marginWorkerActiveCount -= 1;
                setTimeout(pumpMarginQueue, MARGIN_QUEUE_DELAY_MS);
            });
    }
}

function toggleOrderList() {
    orderListVisible = !orderListVisible;
    renderOrderList();
}

function setOrderListFilter() {
    const input = document.getElementById('customerFilterInput');
    orderListFilter = (input && input.value ? input.value : '').trim().toLowerCase();
    if (!orderListVisible && orderListFilter) {
        orderListVisible = true;
    }
    renderOrderList();
}

function setBrugerFilter() {
    const input = document.getElementById('brugerFilterSelect');
    orderListBrugerFilter = (input && input.value ? input.value : '').trim();
    if (!orderListVisible && orderListBrugerFilter) {
        orderListVisible = true;
    }
    renderOrderList();
}

function setOrderValueFilter() {
    const enabledInput = document.getElementById('orderMinDkkEnabled');
    const thresholdInput = document.getElementById('orderMinDkkInput');
    orderListMinDkkEnabled = !!(enabledInput && enabledInput.checked);

    const raw = Number(thresholdInput && thresholdInput.value || 0);
    orderListMinDkkValue = Number.isFinite(raw) ? Math.max(0, Math.round(raw)) : 0;

    if (thresholdInput) {
        thresholdInput.disabled = !orderListMinDkkEnabled;
        if (!thresholdInput.disabled && String(thresholdInput.value || '') !== String(orderListMinDkkValue)) {
            thresholdInput.value = String(orderListMinDkkValue);
        }
    }

    if (!orderListVisible && orderListMinDkkEnabled && orderListMinDkkValue > 0) {
        orderListVisible = true;
    }
    renderOrderList();
}

function populateBrugerFilterOptions() {
    const select = document.getElementById('brugerFilterSelect');
    if (!select) return;

    const selectedValue = orderListBrugerFilter;
    const users = Array.from(new Set(
        orderListData
            .map(o => String(o.SellerUsr || '').trim())
            .filter(v => v)
    )).sort((a, b) => a.localeCompare(b));

    let html = '<option value="">Alle brugere</option>';
    for (const user of users) {
        html += '<option value="' + user + '">' + user + '</option>';
    }
    select.innerHTML = html;
    select.value = selectedValue;
}

function setOrderListSort(field) {
    if (orderListSortField === field) {
        orderListSortDir = orderListSortDir === 'asc' ? 'desc' : 'asc';
    } else {
        orderListSortField = field;
        orderListSortDir = field === 'date' || field === 'ordno' || field === 'belob' || field === 'margin' ? 'desc' : 'asc';
    }
    renderOrderList();
}

function getMarginValue(ordNo) {
    const state = getMarginState(ordNo);
    if (!state || state.status !== 'success') return null;
    return calculateOrderMarginPercent(state.totalRevenue || 0, state.totalCost || 0);
}

function getFilteredOrders() {
    const filtered = orderListData.filter(o => {
        const bruger = String(o.SellerUsr || '').trim();
        const customer = String(o.CustomerName || '').toLowerCase();
        const ord = String(o.OrdNo || '');
        const matchesText = !orderListFilter || customer.includes(orderListFilter) || ord.includes(orderListFilter);
        const matchesBruger = !orderListBrugerFilter || bruger === orderListBrugerFilter;
        const invoDkk = Number(o.InvoAm || 0);
        const matchesMinDkk = !orderListMinDkkEnabled || invoDkk >= orderListMinDkkValue;
        return matchesText && matchesBruger && matchesMinDkk;
    });

    const dir = orderListSortDir === 'asc' ? 1 : -1;
    filtered.sort((a, b) => {
        switch (orderListSortField) {
            case 'bruger': {
                const cmp = String(a.SellerUsr || '').localeCompare(String(b.SellerUsr || ''));
                return cmp * dir || Number(b.LstInvDt || 0) - Number(a.LstInvDt || 0);
            }
            case 'ordno':
                return (Number(a.OrdNo || 0) - Number(b.OrdNo || 0)) * dir;
            case 'kunde': {
                const cmp = String(a.CustomerName || '').localeCompare(String(b.CustomerName || ''));
                return cmp * dir || Number(b.OrdNo || 0) - Number(a.OrdNo || 0);
            }
            case 'date': {
                const d = (Number(a.LstInvDt || 0) - Number(b.LstInvDt || 0)) * dir;
                return d || (Number(b.OrdNo || 0) - Number(a.OrdNo || 0)) * dir;
            }
            case 'belob':
                return (Number(a.InvoAm || 0) - Number(b.InvoAm || 0)) * dir;
            case 'margin': {
                const ma = getMarginValue(a.OrdNo);
                const mb = getMarginValue(b.OrdNo);
                if (ma === null && mb === null) return 0;
                if (ma === null) return 1;
                if (mb === null) return -1;
                return (ma - mb) * dir;
            }
            default:
                return Number(b.LstInvDt || 0) - Number(a.LstInvDt || 0);
        }
    });
    return filtered;
}

function getMarginBadge(marginPercent) {
    const margin = parseFloat(marginPercent);
    if (currentMarginMode === 'new') {
        if (margin >= 125) {
            return '<span style="background:#2e7d32; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">✅ ' + marginPercent + '%</span>';
        } else if (margin >= 105) {
            return '<span style="background:#ff9800; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">⚠️ ' + marginPercent + '%</span>';
        }
        return '<span style="background:#d32f2f; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">❌ ' + marginPercent + '%</span>';
    }

    if (margin > 20) {
        return '<span style="background:#2e7d32; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">✅ ' + marginPercent + '%</span>';
    } else if (margin >= 5) {
        return '<span style="background:#ff9800; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">⚠️ ' + marginPercent + '%</span>';
    }
    return '<span style="background:#d32f2f; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">❌ ' + marginPercent + '%</span>';
}

function getFilteredOrderSummary(orders) {
    const safeOrders = Array.isArray(orders) ? orders : [];
    let considered = 0;
    let excludedCredit = 0;
    let excludedUB = 0;
    let pendingMargin = 0;
    let totalRevenue = 0;
    let totalCost = 0;

    for (const o of safeOrders) {
        if (isOrderMarkedUB(o.OrdNo)) {
            excludedUB += 1;
            excludedCredit += 1;
            continue;
        }
        if (isOrderMarkedCreditNote(o.OrdNo)) {
            excludedCredit += 1;
            continue;
        }
        const state = getMarginState(o.OrdNo);
        if (!state || state.status !== 'success') {
            pendingMargin += 1;
            continue;
        }
        considered += 1;
        totalRevenue += Number(state.totalRevenue || 0);
        totalCost += Number(state.totalCost || 0);
    }

    const marginAmount = totalRevenue - totalCost;
    const marginPct = totalCost > 0 ? calculateOrderMarginPercent(totalRevenue, totalCost).toFixed(2) : '0.00';
    return {
        considered,
        excludedCredit,
        excludedUB,
        pendingMargin,
        totalRevenue,
        totalCost,
        marginAmount,
        marginPct
    };
}

function buildOrderListSummaryHtml(orders) {
    const listSummary = getFilteredOrderSummary(orders);
    const activeFilters = [];
    if (orderListFilter) activeFilters.push('kunde/søgning: "' + escapeHtml(orderListFilter) + '"');
    if (orderListBrugerFilter) activeFilters.push('bruger: "' + escapeHtml(orderListBrugerFilter) + '"');
    if (orderListMinDkkEnabled) activeFilters.push('minimum fakturabeløb: ' + formatNumber(orderListMinDkkValue) + ' DKK');
    const filterText = activeFilters.length > 0 ? activeFilters.join(', ') : 'ingen aktive filtre';
    const excludedCreditOnly = listSummary.excludedCredit - listSummary.excludedUB;
    let html = '<div><strong>Filtreret ordrelisteoversigt</strong> (vist: ' + orders.length + ', medtaget: ' + listSummary.considered + ', kreditnota udelukket: ' + excludedCreditOnly + (listSummary.excludedUB > 0 ? ', u/b udelukket: ' + listSummary.excludedUB : '') + ', mangler margin: ' + listSummary.pendingMargin + ')</div>';
    html += '<div style="margin-top:4px; font-size:12px; color:#57718f;">Genereret: ' + escapeHtml(new Date().toLocaleString('da-DK')) + ' • Filtre: ' + filterText + '</div>';
    html += '<div style="margin-top:6px;display:flex;gap:18px;flex-wrap:wrap;">';
    html += '<span>Samlet omsætning: <strong>' + formatNumber(listSummary.totalRevenue) + ' DKK</strong></span>';
    html += '<span>Samlet kost: <strong>' + formatNumber(listSummary.totalCost) + ' DKK</strong></span>';
    html += '<span>Margin: <strong>' + formatNumber(listSummary.marginAmount) + ' DKK (' + listSummary.marginPct + '%)</strong></span>';
    html += '</div>';
    html += '<div class="order-list-summary-actions">';
    html += '<button class="list-toggle-btn" onclick="openOrderListPrintPreview()" title="Vis forhåndsvisning af den filtrerede ordreliste">Forhåndsvisning / PDF</button>';
    html += '</div>';
    return html;
}

