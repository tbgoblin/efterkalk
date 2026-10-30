// ── Omsaetning (Omsætning/revenue dashboard) · client ────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: formatNumber, escapeHtml, escapeHtmlFE, canAccessModule,
// searchOrderByNo, pushModalStack/removeModalStack, dashboardSourceKey,
// fetchOrderFlowCached, scheduleDashboardWidgets (dichiarati nell'inline
// script) e getOrdreindgangHolidayWeeksSet/shouldIgnoreOrdreindgangHolidayWeeks
// (dichiarati in assets/js/ordreindgang.js — caricato prima di questo file).
function setOmsaetningCacheEntry(cacheMap, key, value) {
    cacheMap.set(key, {
        ts: Date.now(),
        value
    });
    if (cacheMap.size > OMSAETNING_CACHE_MAX_ITEMS) {
        const oldestKey = cacheMap.keys().next().value;
        if (oldestKey !== undefined) cacheMap.delete(oldestKey);
    }
}

function getOmsaetningCacheEntry(cacheMap, key, ttlMs) {
    const hit = cacheMap.get(key);
    if (!hit) return null;
    if ((Date.now() - Number(hit.ts || 0)) > ttlMs) {
        cacheMap.delete(key);
        return null;
    }
    return hit.value;
}

function buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, selectedCustomers) {
    const accounts = Array.from(new Set((Array.isArray(selectedAccounts) ? selectedAccounts : [])
        .map(v => String(v || '').trim())
        .filter(Boolean))).sort();
    const customers = Array.from(new Set((Array.isArray(selectedCustomers) ? selectedCustomers : [])
        .map(v => String(v || '').trim())
        .filter(Boolean))).sort();
    return JSON.stringify({ database: _settingsActiveId, username: loggedUsername.toLowerCase(), fra, til, accounts, customers });
}

async function fetchOmsaetningSummaryCached(fra, til, selectedAccounts, selectedCustomers, options) {
    const safeOptions = options && typeof options === 'object' ? options : {};
    const forceRefresh = safeOptions.forceRefresh === true;
    const cacheKey = buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, selectedCustomers);
    if (!forceRefresh) {
        const cached = getOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, OMSAETNING_SUMMARY_CACHE_TTL_MS);
        if (cached) {
            return cached;
        }
    }

    const customerFilters = Array.isArray(selectedCustomers)
        ? Array.from(new Set(selectedCustomers.map(v => String(v || '').trim()).filter(Boolean))).sort()
        : [];
    if (!forceRefresh && customerFilters.length > 0) {
        const allCustomersKey = buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, []);
        const allCustomersCached = getOmsaetningCacheEntry(omsaetningSummaryCache, allCustomersKey, OMSAETNING_SUMMARY_CACHE_TTL_MS);
        if (allCustomersCached && Array.isArray(allCustomersCached.rows)) {
            const selectedSet = new Set(customerFilters);
            const filteredRows = allCustomersCached.rows.filter(row => {
                const custNo = row && row.custNo !== null && row.custNo !== undefined
                    ? String(row.custNo).trim()
                    : '';
                return custNo && selectedSet.has(custNo);
            });
            const derivedPayload = {
                ok: true,
                filters: {
                    fra,
                    til,
                    accounts: Array.isArray(allCustomersCached.filters && allCustomersCached.filters.accounts)
                        ? allCustomersCached.filters.accounts
                        : [],
                    customers: customerFilters
                },
                totalRevenueMio: filteredRows.reduce((sum, row) => sum + Number((row && row.revenueMio) || 0), 0),
                rows: filteredRows
            };
            setOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, derivedPayload);
            return derivedPayload;
        }
    }

    if (!forceRefresh) {
        const inFlight = omsaetningSummaryInFlight.get(cacheKey);
        if (inFlight) {
            return await inFlight;
        }
    }

    const query = new URLSearchParams({
        fra,
        til,
        accounts: selectedAccounts.join(','),
        customers: selectedCustomers.join(',')
    });

    const reqPromise = (async () => {
        const response = await fetch('/omsaetning/summary?' + query.toString());
        if (!response.ok) throw new Error('HTTP ' + response.status);
        const payload = await response.json();
        setOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, payload);
        if (dashboardSourceCache && payload.ok !== false) {
            dashboardSourceCache.put(dashboardSourceKey('omsaetning', cacheKey), payload);
            scheduleDashboardWidgets();
        }
        return payload;
    })();

    omsaetningSummaryInFlight.set(cacheKey, reqPromise);
    try {
        return await reqPromise;
    } finally {
        omsaetningSummaryInFlight.delete(cacheKey);
    }
}

function formatOmsaetningVismaDate(value) {
    const raw = String(value || '').replace(/\D+/g, '');
    if (!/^\d{8}$/.test(raw)) return '-';
    return raw.slice(6, 8) + '-' + raw.slice(4, 6) + '-' + raw.slice(0, 4);
}

function buildOmsaetningMonthDetailCacheKey(month, selectedAccounts, selectedCustomers) {
    return 'month-detail:' + buildOmsaetningSummaryCacheKey(
        String(month || ''),
        String(month || ''),
        selectedAccounts,
        selectedCustomers
    );
}

function renderOmsaetningMonthDetail(payload) {
    const wrap = document.getElementById('omsaetningMonthDetailWrap');
    const title = document.getElementById('omsaetningMonthDetailTitle');
    const body = document.getElementById('omsaetningMonthDetailBody');
    if (!wrap || !title || !body) return;

    const month = String(payload && payload.month || omsaetningSelectedMonthKey || '').trim();
    const rows = Array.isArray(payload && payload.rows) ? payload.rows : [];
    const ordreindgang = payload && payload.ordreindgang ? payload.ordreindgang : {};
    const receivedOrders = Array.isArray(payload && payload.receivedOrders) ? payload.receivedOrders : [];
    const receivedSummary = payload && payload.receivedSummary ? payload.receivedSummary : {};
    const weeklyRows = Array.isArray(ordreindgang.weeklyRows) ? ordreindgang.weeklyRows : [];
    const totalRevenueDkk = Number(payload && payload.totalRevenueDkk || 0);
    const linkedRevenueDkk = Number(payload && payload.linkedRevenueDkk || 0);
    const unresolvedRevenueDkk = Number(payload && payload.unresolvedRevenueDkk || 0);
    const linkedCount = Number(payload && payload.linkedCount || 0);
    const unresolvedCount = Number(payload && payload.unresolvedCount || 0);
    const totalOrdDkk = Number(ordreindgang.totalOrdK || 0) * 1000;
    const displayedWeeklyOrdDkk = weeklyRows.reduce((sum, row) => sum + (Number(row && row.totalOrd || 0) * 1000), 0);
    const remainingByWeek = new Map();
    const currentRemainingByWeek = new Map();
    receivedOrders.forEach(row => {
        const weekKey = String(row && row.weekKey || '');
        remainingByWeek.set(weekKey, (remainingByWeek.get(weekKey) || 0) + Number(row && row.historicalRemainingDkk || 0));
        currentRemainingByWeek.set(weekKey, (currentRemainingByWeek.get(weekKey) || 0) + Math.max(0, Number(row && row.remainingDkk || 0)));
    });
    const displayedWeeklyRemainingDkk = weeklyRows.reduce((sum, row) => sum + Number(remainingByWeek.get(String(row && row.weekKey || '')) || 0), 0);
    const displayedWeeklyCurrentRemainingDkk = weeklyRows.reduce((sum, row) => sum + Number(currentRemainingByWeek.get(String(row && row.weekKey || '')) || 0), 0);

    title.textContent = formatMonthDa(month + '-01') + ' · bogførte ordrer og Ordreindgang';
    let html = '<div class="omsaetning-month-detail-kpis">' +
        '<div class="omsaetning-month-detail-kpi"><span>Bogført omsætning</span><strong>' + escapeHtmlFE(formatDkkDa(totalRevenueDkk)) + ' DKK</strong></div>' +
        '<div class="omsaetning-month-detail-kpi"><span>Koblede ordrer</span><strong>' + escapeHtmlFE(formatCount(linkedCount)) + ' · ' + escapeHtmlFE(formatDkkDa(linkedRevenueDkk)) + ' DKK</strong></div>' +
        '<div class="omsaetning-month-detail-kpi"><span>Ordreindgang, månedens uger</span><strong>' + escapeHtmlFE(formatDkkDa(totalOrdDkk)) + ' DKK</strong></div>' +
        '<div class="omsaetning-month-detail-kpi"><span>Ordreindgang, valgte uger</span><strong>' + escapeHtmlFE(formatCount(Number(receivedSummary.receivedCount || 0))) + ' ordrer · ' + escapeHtmlFE(formatDkkDa(Number(receivedSummary.receivedValueDkk || 0))) + ' DKK</strong></div>' +
        '<div class="omsaetning-month-detail-kpi"><span>Heraf faktureret i måneden</span><strong>' + escapeHtmlFE(formatCount(Number(receivedSummary.invoicedThisMonthCount || 0))) + ' ordrer · ' + escapeHtmlFE(formatDkkDa(Number(receivedSummary.invoicedThisMonthDkk || 0))) + ' DKK</strong></div>' +
        '<div class="omsaetning-month-detail-kpi"><span>Ikke afsluttet</span><strong>' + escapeHtmlFE(formatCount(Number(receivedSummary.openCount || 0))) + ' ordrer · mangler ' + escapeHtmlFE(formatDkkDa(Number(receivedSummary.remainingDkk || 0))) + ' DKK</strong></div>' +
        '</div>';

    html += '<section class="omsaetning-month-detail-section"><h4>Efterkalkulationsordrer fra bogføringen</h4>';
    if (rows.length === 0) {
        html += '<div class="omsaetning-month-detail-empty">Ingen bogførte omsætningsbevægelser med de valgte konto- og kundefiltre i denne måned.</div>';
    } else {
        html += '<div class="omsaetning-table-wrap omsaetning-compact-table-wrap" style="margin:0;border-radius:8px;max-height:390px;">' +
            '<table class="omsaetning-table omsaetning-month-orders-table"><thead><tr><th>Ordre</th><th>Ordredato</th><th>Kunde</th><th>Faktura / dato</th><th class="omsaetning-cell-right">Faktureret ordre</th></tr></thead><tbody>';
        for (const row of rows) {
            const ordNo = Number(row && row.ordNo || 0);
            const linkStatus = String(row && row.linkStatus || 'unmatched');
            let orderCell;
            if (ordNo > 0 && linkStatus === 'matched') {
                orderCell = '<button type="button" class="omsaetning-month-order-link" onclick="searchOrderByNo(' + ordNo + ')">' + escapeHtmlFE(String(ordNo)) + '</button>';
            } else if (linkStatus === 'ambiguous') {
                orderCell = '<span class="omsaetning-month-link-state ambiguous">Flere mulige</span>';
            } else {
                orderCell = '<span class="omsaetning-month-link-state unmatched">Ikke koblet</span>';
            }
            const customerLabel = String(row && row.customerName || '').trim() || String(row && row.custNo || '-');
            const invoiceLabel = String(row && row.invoiceNo || '').trim() || ('Bilag ' + String(row && row.voucherNo || '-'));
            html += '<tr><td>' + orderCell + '</td>' +
                '<td>' + escapeHtmlFE(formatOmsaetningVismaDate(row && row.orderDate)) + '</td>' +
                '<td>' + escapeHtmlFE(customerLabel) + '</td>' +
                '<td>' + escapeHtmlFE(invoiceLabel) + '<div class="omsaetning-month-detail-note">' + escapeHtmlFE(formatOmsaetningVismaDate(row && row.invoiceDate)) + '</div></td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(row && row.revenueDkk)) + '</td></tr>';
        }
        html += '</tbody></table></div>';
    }
    if (unresolvedCount > 0) {
        html += '<p class="omsaetning-month-detail-note"><strong>' + escapeHtmlFE(formatCount(unresolvedCount)) + ' bevægelse(r)</strong> svarende til ' + escapeHtmlFE(formatDkkDa(unresolvedRevenueDkk)) + ' DKK kunne ikke kobles entydigt via fakturanummer. De er bevaret i totalen og vises som kontrolrækker.</p>';
    } else {
        html += '<p class="omsaetning-month-detail-note">Beløbene kommer direkte fra de samme AcTr-rækker som månedens Omsætning. Fakturanummeret bruges kun til at åbne den tilhørende efterkalkulation.</p>';
    }
    html += '</section>';

    html += '<section class="omsaetning-month-detail-section"><h4>Ordrer modtaget i måneden — fakturering og rest</h4>';
    if (!receivedOrders.length) {
        html += '<div class="omsaetning-month-detail-empty">Ingen salgsordrer med ordredato i den valgte måned og de valgte kundefiltre.</div>';
    } else {
        html += '<div class="omsaetning-table-wrap omsaetning-compact-table-wrap" style="margin:0;border-radius:8px;max-height:390px;">' +
            '<table class="omsaetning-table omsaetning-received-orders-table"><thead><tr>' +
            '<th>Ordre</th><th>Ordredato</th><th>Kunde</th><th>Status</th><th>Faktura / dato</th>' +
            '<th class="omsaetning-cell-right">Ordreværdi</th><th class="omsaetning-cell-right">Faktureret i måneden</th><th class="omsaetning-cell-right">Faktureret i alt</th><th class="omsaetning-cell-right">Rest ved månedens udgang</th><th class="omsaetning-cell-right">Rest i dag</th>' +
            '</tr></thead><tbody>';
        for (const row of receivedOrders) {
            const status = row.completedAfterMonth ? 'Faktureret senere' : (row.complete ? 'Afsluttet' : (row.invoicedDkk > 0 ? 'Delvist faktureret' : 'Ikke faktureret'));
            const invoice = String(row.invoiceNo || '').trim() || '—';
            html += '<tr><td><button type="button" class="omsaetning-month-order-link" onclick="searchOrderByNo(' + Number(row.ordNo || 0) + ')">' + escapeHtmlFE(String(row.ordNo || '—')) + '</button></td>' +
                '<td>' + escapeHtmlFE(formatOmsaetningVismaDate(row.orderDate)) + '</td>' +
                '<td>' + escapeHtmlFE(String(row.customerName || row.custNo || '—')) + '</td>' +
                '<td>' + escapeHtmlFE(status) + (row.invoicedThisMonth ? '<div class="omsaetning-month-detail-note">Faktureret samme måned</div>' : '') + '</td>' +
                '<td>' + escapeHtmlFE(invoice) + '<div class="omsaetning-month-detail-note">' + escapeHtmlFE(formatOmsaetningVismaDate(row.invoiceDate)) + '</div></td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(row.orderValueDkk)) + '</td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(row.invoicedInMonthDkk)) + '</td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(row.invoicedDkk)) + '</td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(row.historicalRemainingDkk)) + '</td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(Math.max(0, Number(row.remainingDkk || 0)))) + '</td></tr>';
        }
        html += '</tbody></table></div>';
        html += '<p class="omsaetning-month-detail-note">Status og rest vises ved udgangen af den valgte måned ud fra finansposteringerne i AcTr til og med månedens sidste dag. Seneste fakturadato bruges også til at markere ordrer, der blev faktureret senere.</p>';
    }
    html += '</section>';

    html += '<section class="omsaetning-month-detail-section"><h4>Ordreindgang for uger knyttet til måneden</h4>';
    if (weeklyRows.length === 0) {
        html += '<div class="omsaetning-month-detail-empty">Ingen Ordreindgang-uger blev fundet for måneden.</div>';
    } else {
        html += '<div class="omsaetning-table-wrap omsaetning-compact-table-wrap" style="margin:0;border-radius:8px;">' +
            '<table class="omsaetning-table omsaetning-month-weeks-table"><thead><tr><th>Uge</th><th class="omsaetning-cell-right">Ordreindgang</th><th class="omsaetning-cell-right">Rest ved månedens udgang</th><th class="omsaetning-cell-right">Rest i dag</th></tr></thead><tbody>';
        for (const weekRow of weeklyRows) {
            html += '<tr><td>' + escapeHtmlFE(formatWeekLabel(weekRow.weekKey)) + '</td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(Number(weekRow.totalOrd || 0) * 1000)) + ' DKK</td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(Number(remainingByWeek.get(String(weekRow.weekKey || '')) || 0))) + ' DKK</td>' +
                '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(Number(currentRemainingByWeek.get(String(weekRow.weekKey || '')) || 0))) + ' DKK</td></tr>';
        }
        html += '</tbody><tfoot><tr><td>I alt</td>' +
            '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(displayedWeeklyOrdDkk)) + ' DKK</td>' +
            '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(displayedWeeklyRemainingDkk)) + ' DKK</td>' +
            '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatDkkDa(displayedWeeklyCurrentRemainingDkk)) + ' DKK</td>' +
            '</tr></tfoot></table></div>';
        html += '<p class="omsaetning-month-detail-note">For uger på tværs af to måneder medregnes kun ordrer med ordredato i den valgte måned. Det forhindrer, at samme ordrebeløb tælles i begge måneder.</p>';
    }
    html += '</section>';

    body.innerHTML = '<section id="omsaetningOrderFlow" aria-label="Ordreflow"></section>' + html;
    if (window.GohOrderFlow && canAccessModule('omsaetning')) {
        const now = new Date();
        const today = String(now.getFullYear()) + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
        const flowScope = dashboardSourceKey('order-flow', 'detail');
        const customers = Array.from(omsaetningSelectedCustomers.keys());
        window.GohOrderFlow.mount(document.getElementById('omsaetningOrderFlow'), {
            scope: flowScope, month, today, fixedMonth: true,
            load: (selectedMonth, force) => fetchOrderFlowCached(selectedMonth, customers, force),
            isActive: () => accessGranted && canAccessModule('omsaetning') && flowScope === dashboardSourceKey('order-flow', 'detail'),
            openOrder: canAccessModule('efterkalk') ? ordNo => searchOrderByNo(ordNo) : null
        });
    }
    wrap.style.display = 'block';
}

async function loadOmsaetningMonthDetail(monthKey, options) {
    const month = String(monthKey || '').trim().slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(month)) return;
    const safeOptions = options && typeof options === 'object' ? options : {};
    const selectedAccounts = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean);
    const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys()).filter(Boolean);
    if (selectedAccounts.length === 0) return;

    omsaetningSelectedMonthKey = month;
    const token = ++omsaetningMonthDetailToken;
    document.querySelectorAll('.omsaetning-threshold-row').forEach(row => {
        row.classList.toggle('selected', String(row.getAttribute('data-month') || '') === month);
    });

    const wrap = document.getElementById('omsaetningMonthDetailWrap');
    const title = document.getElementById('omsaetningMonthDetailTitle');
    const body = document.getElementById('omsaetningMonthDetailBody');
    if (wrap) wrap.style.display = 'block';
    if (title) title.textContent = formatMonthDa(month + '-01') + ' · indlæser';
    if (body) body.innerHTML = '<div class="omsaetning-month-loading">Henter bogførte ordrer og Ordreindgang...</div>';

    const cacheKey = buildOmsaetningMonthDetailCacheKey(month, selectedAccounts, selectedCustomers);
    if (safeOptions.forceRefresh !== true) {
        const cached = getOmsaetningCacheEntry(omsaetningMonthDetailCache, cacheKey, OMSAETNING_SUMMARY_CACHE_TTL_MS);
        if (cached) {
            if (token === omsaetningMonthDetailToken) renderOmsaetningMonthDetail(cached);
            return;
        }
    }

    try {
        const query = new URLSearchParams({
            month,
            accounts: selectedAccounts.join(','),
            customers: selectedCustomers.join(',')
        });
        const response = await fetch('/omsaetning/month-detail?' + query.toString());
        const payload = await response.json();
        if (!response.ok || !payload.ok) throw new Error(payload.error || ('HTTP ' + response.status));
        setOmsaetningCacheEntry(omsaetningMonthDetailCache, cacheKey, payload);
        if (token === omsaetningMonthDetailToken) renderOmsaetningMonthDetail(payload);
    } catch (error) {
        if (token !== omsaetningMonthDetailToken) return;
        if (title) title.textContent = formatMonthDa(month + '-01') + ' · fejl';
        if (body) body.innerHTML = '<div class="omsaetning-month-detail-empty">Kunne ikke hente månedsdetaljen: ' + escapeHtmlFE(String(error.message || error)) + '</div>';
    }
}

async function searchOmsaetningCustomersCached(q) {
    const key = String(q || '').trim().toLowerCase();
    if (!key) return { customers: [] };

    const cached = getOmsaetningCacheEntry(omsaetningCustomerSearchCache, key, OMSAETNING_CUSTOMER_SEARCH_CACHE_TTL_MS);
    if (cached) {
        return cached;
    }

    const inFlight = omsaetningCustomerSearchInFlight.get(key);
    if (inFlight) {
        return await inFlight;
    }

    const reqPromise = (async () => {
        const response = await fetch('/omsaetning/customers?q=' + encodeURIComponent(key) + '&limit=25');
        if (!response.ok) throw new Error('HTTP ' + response.status);
        const payload = await response.json();
        setOmsaetningCacheEntry(omsaetningCustomerSearchCache, key, payload);
        return payload;
    })();

    omsaetningCustomerSearchInFlight.set(key, reqPromise);
    try {
        return await reqPromise;
    } finally {
        omsaetningCustomerSearchInFlight.delete(key);
    }
}

function applyOmsaetningDetailsCollapsedState() {
    const tableWrap = document.getElementById('omsaetningTableWrap');
    const toggleBtn = document.getElementById('omsaetningDetailsToggleBtn');
    if (!tableWrap || !toggleBtn) return;
    tableWrap.style.display = omsaetningDetailsCollapsed ? 'none' : 'block';
    toggleBtn.textContent = omsaetningDetailsCollapsed ? 'Vis detaljer' : 'Skjul detaljer';
}

function renderOmsaetningAccountsSummary() {
    const summaryEl = document.getElementById('omsaetningAccountsSummary');
    const activeEl = document.getElementById('omsaetningAccountsActive');
    if (!summaryEl) return;
    const selectedCount = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean).length;
    const totalCount = Array.isArray(omsaetningAccounts) ? omsaetningAccounts.length : 0;
    summaryEl.textContent = String(selectedCount) + '/' + String(totalCount) + ' valgt';

    if (activeEl) {
        const selectedValues = Array.from(omsaetningSelectedAccounts.values())
            .map(v => String(v || '').trim())
            .filter(Boolean)
            .sort((a, b) => a.localeCompare(b));

        if (selectedValues.length === 0) {
            activeEl.innerHTML = '<span class="chip more">Ingen konti valgt</span>';
        } else {
            const visible = selectedValues.slice(0, 5).map(acNo => {
                const account = Array.isArray(omsaetningAccounts)
                    ? omsaetningAccounts.find(a => String(a && a.acNo || '').trim() === acNo)
                    : null;
                const label = account ? (acNo + ' ' + String(account.name || '').trim()) : acNo;
                return '<span class="chip" title="' + escapeHtmlFE(label) + '">' + escapeHtmlFE(acNo) + '</span>';
            });
            if (selectedValues.length > 5) {
                visible.push('<span class="chip more">+' + escapeHtmlFE(String(selectedValues.length - 5)) + '</span>');
            }
            activeEl.innerHTML = visible.join('');
        }
    }
}

function formatSigned(value, digits) {
    const n = Number(value || 0);
    const fixed = n.toFixed(Number.isFinite(digits) ? digits : 1);
    return (n > 0 ? '+' : '') + fixed;
}

function buildOmsaetningGaugeData(amountMio, warnThreshold, goodThreshold) {
    const amount = Number(amountMio || 0);
    const warn = Number(warnThreshold || 0);
    const good = Math.max(warn + 0.0001, Number(goodThreshold || 0));
    const span = Math.max(0.0001, good - warn);

    const marginPct = ((amount - warn) / span) * 30;
    const scaleMin = -30;
    const scaleMax = 60;
    const toLeft = value => ((value - scaleMin) / (scaleMax - scaleMin)) * 100;

    const zeroLeft = toLeft(0);
    const targetLeft = toLeft(30);
    const clamped = Math.max(scaleMin, Math.min(scaleMax, marginPct));
    const pointLeft = toLeft(clamped);

    return {
        marginPct,
        pointLeft,
        zeroLeft,
        targetLeft,
        fillLeft: Math.min(pointLeft, zeroLeft),
        fillWidth: Math.abs(pointLeft - zeroLeft),
        fillClass: pointLeft >= zeroLeft ? 'pos' : 'neg',
        deltaWarn: amount - warn,
        deltaGood: amount - good
    };
}

function applyOmsaetningAccountsPanelState() {
    const panel = document.getElementById('omsaetningAccountsPanel');
    const search = document.getElementById('omsaetningAccountSearch');
    const btn = document.getElementById('omsaetningAccountsToggleBtn');
    if (panel) panel.style.display = omsaetningAccountsPanelOpen ? 'block' : 'none';
    if (search) search.style.display = omsaetningAccountsPanelOpen ? 'block' : 'none';
    if (btn) btn.textContent = omsaetningAccountsPanelOpen ? 'Skjul konti' : 'Vis konti';
}

function toggleOmsaetningAccountsPanel() {
    omsaetningAccountsPanelOpen = !omsaetningAccountsPanelOpen;
    applyOmsaetningAccountsPanelState();
}

function toggleOmsaetningDetails() {
    omsaetningDetailsCollapsed = !omsaetningDetailsCollapsed;
    applyOmsaetningDetailsCollapsedState();
}

function formatMio(value) {
    const numeric = Number(value || 0);
    return numeric.toLocaleString('da-DK', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
}

function formatDkkFromMio(valueMio) {
    const numeric = Number(valueMio || 0) * 1000000;
    return numeric.toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatMonthDa(dateValue) {
    if (!dateValue) return '-';
    const dt = new Date(dateValue);
    if (Number.isNaN(dt.getTime())) return String(dateValue);
    return dt.toLocaleDateString('da-DK', { month: 'short', year: 'numeric' });
}

function normalizeOmsaetningMonthKey(dateValue) {
    const dt = new Date(dateValue);
    if (Number.isNaN(dt.getTime())) return String(dateValue || '').trim();
    const year = dt.getFullYear();
    const month = String(dt.getMonth() + 1).padStart(2, '0');
    return String(year) + '-' + month + '-01';
}

function parseMonthInputToPeriod(monthValue) {
    const raw = String(monthValue || '').trim();
    const match = raw.match(/^(\d{4})-(\d{2})$/);
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return null;
    return {
        year,
        month,
        period: year * 100 + month
    };
}

function calendarMonthToFiscalYrPr(year, month) {
    const y = Number(year);
    const m = Number(month);
    if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) return null;
    if (m >= 7) {
        return (y * 100) + (m - 6);
    }
    return ((y - 1) * 100) + (m + 6);
}

function getCurrentFiscalYearStart(referenceDate) {
    const now = referenceDate instanceof Date ? referenceDate : new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    return month >= 7 ? year : (year - 1);
}

function getFiscalYearRange(yearValue) {
    const startYear = Number(yearValue);
    if (!Number.isFinite(startYear)) return null;
    return {
        fromMonth: String(startYear) + '-07',
        toMonth: String(startYear + 1) + '-06',
        fra: String(startYear * 100 + 7),
        til: String((startYear + 1) * 100 + 7)
    };
}

function applySelectedFiscalYearsToInputs() {
    const fromEl = document.getElementById('omsaetningFraMonth');
    const toEl = document.getElementById('omsaetningTilMonth');
    const selectedYears = Array.from(omsaetningSelectedFiscalYears.values())
        .map(y => Number(y))
        .filter(y => Number.isFinite(y))
        .sort((a, b) => a - b);
    if (selectedYears.length === 0) return;

    const firstRange = getFiscalYearRange(selectedYears[0]);
    const lastRange = getFiscalYearRange(selectedYears[selectedYears.length - 1]);
    if (!firstRange || !lastRange) return;

    if (fromEl) fromEl.value = firstRange.fromMonth;
    if (toEl) toEl.value = lastRange.toMonth;
}

function buildOmsaetningPeriodRange() {
    const fromEl = document.getElementById('omsaetningFraMonth');
    const toEl = document.getElementById('omsaetningTilMonth');
    const fromMeta = parseMonthInputToPeriod(fromEl ? fromEl.value : '');
    const toMeta = parseMonthInputToPeriod(toEl ? toEl.value : '');
    if (!fromMeta || !toMeta) return null;

    const fromDate = new Date(fromMeta.year, fromMeta.month - 1, 1);
    const toDate = new Date(toMeta.year, toMeta.month - 1, 1);
    if (fromDate.getTime() > toDate.getTime()) return null;

    const exclusiveToDate = new Date(toMeta.year, toMeta.month, 1);
    const fraFiscal = calendarMonthToFiscalYrPr(fromMeta.year, fromMeta.month);
    const tilFiscal = calendarMonthToFiscalYrPr(exclusiveToDate.getFullYear(), exclusiveToDate.getMonth() + 1);
    if (!Number.isFinite(fraFiscal) || !Number.isFinite(tilFiscal)) return null;
    return {
        fra: String(fraFiscal),
        til: String(tilFiscal)
    };
}

function buildOmsaetningMonthKeys(periodRange) {
    if (!periodRange) return [];
    const fraNum = Number(periodRange.fra);
    const tilNum = Number(periodRange.til);
    if (!Number.isFinite(fraNum) || !Number.isFinite(tilNum) || fraNum >= tilNum) return [];

    let year = Math.floor(fraNum / 100);
    let month = fraNum % 100;
    const keys = [];

    while ((year * 100 + month) < tilNum) {
        const mm = String(month).padStart(2, '0');
        keys.push(String(year) + '-' + mm + '-01');
        month += 1;
        if (month > 12) {
            month = 1;
            year += 1;
        }
    }
    return keys;
}

function renderOmsaetningYearChips(centerYear) {
    const wrap = document.getElementById('omsaetningYears');
    if (!wrap) return;

    const currentFiscalYear = Number(centerYear || getCurrentFiscalYearStart());
    const years = [currentFiscalYear - 3, currentFiscalYear - 2, currentFiscalYear - 1, currentFiscalYear, currentFiscalYear + 1];
    wrap.innerHTML = years.map(fiscalYearStart => {
        const activeCls = omsaetningSelectedFiscalYears.has(fiscalYearStart) ? ' active' : '';
        const label = String(fiscalYearStart) + '/' + String(fiscalYearStart + 1).slice(-2);
        return '<button type="button" class="omsaetning-year-btn' + activeCls + '" onclick="toggleOmsaetningFiscalYear(' + fiscalYearStart + ')">' +
            (fiscalYearStart === currentFiscalYear ? 'Nu ' : '') + escapeHtmlFE(label) +
            '</button>';
    }).join('');
}

function toggleOmsaetningFiscalYear(year) {
    const yearValue = Number(year);
    if (!Number.isFinite(yearValue)) return;
    if (omsaetningSelectedFiscalYears.has(yearValue) && omsaetningSelectedFiscalYears.size > 1) {
        omsaetningSelectedFiscalYears.delete(yearValue);
    } else {
        omsaetningSelectedFiscalYears.add(yearValue);
    }
    applySelectedFiscalYearsToInputs();
    renderOmsaetningYearChips(getCurrentFiscalYearStart());
    loadOmsaetningSummary();
}

function filterOmsaetningAccounts() {
    const qEl = document.getElementById('omsaetningAccountSearch');
    const q = String((qEl && qEl.value) || '').trim().toLowerCase();
    const rows = document.querySelectorAll('#omsaetningAccountsList .omsaetning-account-item');
    for (const row of rows) {
        const text = String(row.getAttribute('data-search') || '').toLowerCase();
        row.style.display = (!q || text.includes(q)) ? '' : 'none';
    }
}

function setAllOmsaetningAccounts(checked) {
    const list = document.getElementById('omsaetningAccountsList');
    if (!list) return;
    const boxes = list.querySelectorAll('input[type="checkbox"][data-accno]');
    for (const box of boxes) {
        box.checked = !!checked;
        const value = String(box.getAttribute('data-accno') || '').trim();
        if (!value) continue;
        if (checked) omsaetningSelectedAccounts.add(value);
        else omsaetningSelectedAccounts.delete(value);
    }
    renderOmsaetningAccountsSummary();
    scheduleOmsaetningAutoReload();
}

function renderOmsaetningAccountsList() {
    const list = document.getElementById('omsaetningAccountsList');
    if (!list) return;
    if (!Array.isArray(omsaetningAccounts) || omsaetningAccounts.length === 0) {
        list.innerHTML = '<div class="omsaetning-account-item"><span>Ingen konti fundet</span></div>';
        return;
    }

    list.innerHTML = omsaetningAccounts.map(acc => {
        const value = String(acc.acNo || '').trim();
        const checked = omsaetningSelectedAccounts.has(value) ? ' checked' : '';
        const search = (value + ' ' + String(acc.name || '')).replace(/"/g, '&quot;');
        return '<label class="omsaetning-account-item" data-search="' + search + '">' +
            '<input type="checkbox" data-accno="' + escapeHtmlFE(value) + '"' + checked + ' onchange="toggleOmsaetningAccount(this)" />' +
            '<span>' + escapeHtmlFE(value + ' - ' + String(acc.name || '')) + '</span>' +
            '</label>';
    }).join('');
    renderOmsaetningAccountsSummary();
    applyOmsaetningAccountsPanelState();
}

function toggleOmsaetningAccount(inputEl) {
    if (!inputEl) return;
    const value = String(inputEl.getAttribute('data-accno') || '').trim();
    if (!value) return;
    if (inputEl.checked) omsaetningSelectedAccounts.add(value);
    else omsaetningSelectedAccounts.delete(value);
    renderOmsaetningAccountsSummary();
    scheduleOmsaetningAutoReload();
}

function renderOmsaetningSelectedCustomers() {
    const wrap = document.getElementById('omsaetningSelectedCustomers');
    if (!wrap) return;

    const entries = Array.from(omsaetningSelectedCustomers.entries());
    if (entries.length === 0) {
        wrap.innerHTML = '';
        return;
    }

    wrap.innerHTML = entries.map(([custNo, name]) => (
        '<span class="omsaetning-selected-chip">' +
        escapeHtmlFE(String(custNo) + ' - ' + String(name || '')) +
        '<button type="button" title="Fjern kunde" onclick="removeOmsaetningCustomer(' + Number(custNo) + ')">×</button>' +
        '</span>'
    )).join('');
}

function removeOmsaetningCustomer(custNo) {
    const key = String(custNo || '').trim();
    if (!key) return;
    omsaetningSelectedCustomers.delete(key);
    renderOmsaetningSelectedCustomers();
    onOmsaetningSelectedCustomersChanged();
    renderOmsaetningCustomerResults();
}

function clearOmsaetningCustomerSearchUi() {
    const qEl = document.getElementById('omsaetningCustomerSearch');
    if (qEl) {
        qEl.value = '';
        qEl.blur();
    }
    omsaetningCustomerResults = [];
    renderOmsaetningCustomerResults();
}

function renderOmsaetningCustomerMode() {
    const modeEl = document.getElementById('omsaetningCustomerMode');
    if (!modeEl) return;

    const selectedCount = omsaetningSelectedCustomers.size;
    if (selectedCount === 0) {
        modeEl.textContent = 'Ingen kunde valgt: viser normal visning for valgte år og konti.';
        return;
    }

    if (selectedCount === 1) {
        modeEl.textContent = '1 kunde valgt: rapporten filtreres på kunden.';
        return;
    }

    modeEl.textContent = String(selectedCount) + ' kunder valgt: søjlediagram sammenligner kunder måned for måned.';
}

function toggleOmsaetningCustomer(custNo, name) {
    const key = String(custNo || '').trim();
    if (!key) return;
    const wasSelected = omsaetningSelectedCustomers.has(key);
    if (omsaetningSelectedCustomers.has(key)) {
        omsaetningSelectedCustomers.delete(key);
    } else {
        omsaetningSelectedCustomers.set(key, String(name || '').trim());
    }
    renderOmsaetningSelectedCustomers();
    onOmsaetningSelectedCustomersChanged();
    if (!wasSelected) clearOmsaetningCustomerSearchUi();
    renderOmsaetningCustomerResults();
}

function toggleOmsaetningCustomerByButton(buttonEl) {
    if (!buttonEl) return;
    const custNo = String(buttonEl.getAttribute('data-custno') || '').trim();
    const custName = String(buttonEl.getAttribute('data-custname') || '').trim();
    toggleOmsaetningCustomer(custNo, custName);
}

function renderOmsaetningCustomerResults() {
    const wrap = document.getElementById('omsaetningCustomerResults');
    if (!wrap) return;
    const qEl = document.getElementById('omsaetningCustomerSearch');
    const q = String((qEl && qEl.value) || '').trim();
    const selectedCount = omsaetningSelectedCustomers.size;

    if (q.length < 2) {
        if (selectedCount === 0) {
            wrap.innerHTML = '<div class="omsaetning-customer-empty">Ingen kunde valgt: viser normal visning for valgte år og konti.</div>';
        } else {
            wrap.innerHTML = '<div class="omsaetning-customer-empty">Skriv mindst 2 tegn for at tilføje flere kunder.</div>';
        }
        return;
    }

    if (!Array.isArray(omsaetningCustomerResults) || omsaetningCustomerResults.length === 0) {
        wrap.innerHTML = '<div class="omsaetning-customer-empty">Ingen kunder matcher søgningen.</div>';
        return;
    }

    wrap.innerHTML = omsaetningCustomerResults.map(row => {
        const custNo = String(row.custNo || '').trim();
        const custName = String(row.name || '').trim();
        const selected = omsaetningSelectedCustomers.has(custNo);
        return '<div class="omsaetning-customer-item">' +
            '<div class="meta"><strong>' + escapeHtmlFE(custName || '(uden navn)') + '</strong><span>' + escapeHtmlFE(custNo) + '</span></div>' +
            '<button type="button" class="pick' + (selected ? ' remove' : '') + '" data-custno="' + escapeHtmlFE(custNo) + '" data-custname="' + escapeHtmlFE(custName) + '" onclick="toggleOmsaetningCustomerByButton(this)">' + (selected ? 'Valgt' : 'Vælg') + '</button>' +
            '</div>';
    }).join('');
}

function scheduleOmsaetningCustomerSearch() {
    if (omsaetningCustomerSearchTimer) {
        clearTimeout(omsaetningCustomerSearchTimer);
    }
    omsaetningCustomerSearchTimer = setTimeout(() => {
        searchOmsaetningCustomers();
    }, 220);
}

async function searchOmsaetningCustomers() {
    const qEl = document.getElementById('omsaetningCustomerSearch');
    const q = String((qEl && qEl.value) || '').trim();
    const token = ++omsaetningCustomerSearchToken;

    if (q.length < 2) {
        omsaetningCustomerResults = [];
        renderOmsaetningCustomerResults();
        return;
    }

    try {
        const payload = await searchOmsaetningCustomersCached(q);
        if (token !== omsaetningCustomerSearchToken) return;
        const rows = Array.isArray(payload.customers) ? payload.customers : [];
        omsaetningCustomerResults = rows;
        renderOmsaetningCustomerResults();
    } catch (err) {
        if (token !== omsaetningCustomerSearchToken) return;
        omsaetningCustomerResults = [];
        const wrap = document.getElementById('omsaetningCustomerResults');
        if (wrap) {
            wrap.innerHTML = '<div class="omsaetning-customer-empty">Fejl ved kundesøgning.</div>';
        }
    }
}

function getOmsaetningStatusClass(valueMio, warnThreshold, goodThreshold) {
    const n = Number(valueMio || 0);
    if (n >= goodThreshold) return 'good';
    if (n >= warnThreshold) return 'mid';
    return 'low';
}

function sanitizeOmsaetningThresholdConfig(rawConfig) {
    const raw = rawConfig && typeof rawConfig === 'object' ? rawConfig : {};
    const warnRaw = Number(raw.warnThreshold);
    const warnThreshold = Number.isFinite(warnRaw) ? Math.max(0, warnRaw) : OMSAETNING_DEFAULT_WARN_THRESHOLD;
    const goodRaw = Number(raw.goodThreshold);
    const goodThreshold = Number.isFinite(goodRaw)
        ? Math.max(warnThreshold, goodRaw)
        : Math.max(warnThreshold, OMSAETNING_DEFAULT_GOOD_THRESHOLD);
    const dailyCalculator = typeof OmsaetningDailyThresholds !== 'undefined' ? OmsaetningDailyThresholds : null;
    const daily = dailyCalculator
        ? dailyCalculator.sanitizeConfig(raw)
        : {
            useDailyBudget: raw.useDailyBudget === true,
            dailyBreakEvenDkk: Math.max(0, Number(raw.dailyBreakEvenDkk) || OMSAETNING_DEFAULT_DAILY_BREAK_EVEN_DKK),
            dailyBudgetDkk: Math.max(Number(raw.dailyBreakEvenDkk) || OMSAETNING_DEFAULT_DAILY_BREAK_EVEN_DKK, Number(raw.dailyBudgetDkk) || OMSAETNING_DEFAULT_DAILY_BUDGET_DKK)
        };
    return {
        warnThreshold,
        goodThreshold,
        useDailyBudget: daily.useDailyBudget,
        dailyBreakEvenDkk: daily.dailyBreakEvenDkk,
        dailyBudgetDkk: daily.dailyBudgetDkk
    };
}

function getOmsaetningThresholdInputs() {
    const warnThresholdInput = document.getElementById('omsaetningWarnThreshold');
    const goodThresholdInput = document.getElementById('omsaetningGoodThreshold');
    const useDailyBudgetInput = document.getElementById('omsaetningUseDailyBudget');
    const dailyBreakEvenInput = document.getElementById('omsaetningDailyBreakEven');
    const dailyBudgetInput = document.getElementById('omsaetningDailyBudget');
    return sanitizeOmsaetningThresholdConfig({
        warnThreshold: warnThresholdInput ? warnThresholdInput.value : OMSAETNING_DEFAULT_WARN_THRESHOLD,
        goodThreshold: goodThresholdInput ? goodThresholdInput.value : OMSAETNING_DEFAULT_GOOD_THRESHOLD,
        useDailyBudget: !!(useDailyBudgetInput && useDailyBudgetInput.checked),
        dailyBreakEvenDkk: dailyBreakEvenInput ? dailyBreakEvenInput.value : OMSAETNING_DEFAULT_DAILY_BREAK_EVEN_DKK,
        dailyBudgetDkk: dailyBudgetInput ? dailyBudgetInput.value : OMSAETNING_DEFAULT_DAILY_BUDGET_DKK
    });
}

function applyOmsaetningThresholdInputs(configOrWarnThreshold, legacyGoodThreshold) {
    const warnThresholdInput = document.getElementById('omsaetningWarnThreshold');
    const goodThresholdInput = document.getElementById('omsaetningGoodThreshold');
    const useDailyBudgetInput = document.getElementById('omsaetningUseDailyBudget');
    const dailyBreakEvenInput = document.getElementById('omsaetningDailyBreakEven');
    const dailyBudgetInput = document.getElementById('omsaetningDailyBudget');
    const raw = configOrWarnThreshold && typeof configOrWarnThreshold === 'object'
        ? configOrWarnThreshold
        : { warnThreshold: configOrWarnThreshold, goodThreshold: legacyGoodThreshold };
    const safe = sanitizeOmsaetningThresholdConfig(raw);
    if (warnThresholdInput) {
        warnThresholdInput.value = safe.warnThreshold.toFixed(1);
        warnThresholdInput.disabled = safe.useDailyBudget;
    }
    if (goodThresholdInput) {
        goodThresholdInput.value = safe.goodThreshold.toFixed(1);
        goodThresholdInput.disabled = safe.useDailyBudget;
    }
    if (useDailyBudgetInput) useDailyBudgetInput.checked = safe.useDailyBudget;
    if (dailyBreakEvenInput) {
        dailyBreakEvenInput.value = String(Math.round(safe.dailyBreakEvenDkk));
        dailyBreakEvenInput.disabled = !safe.useDailyBudget;
    }
    if (dailyBudgetInput) {
        dailyBudgetInput.value = String(Math.round(safe.dailyBudgetDkk));
        dailyBudgetInput.disabled = !safe.useDailyBudget;
    }
    renderOmsaetningThresholdHelp(safe);
}

function loadOmsaetningThresholdUiConfig() {
    try {
        const raw = localStorage.getItem(OMSAETNING_THRESHOLD_UI_STORAGE_KEY);
        return raw ? sanitizeOmsaetningThresholdConfig(JSON.parse(raw)) : sanitizeOmsaetningThresholdConfig(null);
    } catch {
        return sanitizeOmsaetningThresholdConfig(null);
    }
}

function saveOmsaetningThresholdUiConfig(config) {
    const safe = sanitizeOmsaetningThresholdConfig(config);
    try { localStorage.setItem(OMSAETNING_THRESHOLD_UI_STORAGE_KEY, JSON.stringify(safe)); } catch {}
    return safe;
}

function getOmsaetningThresholdsForMonth(monthKey, config) {
    const safe = sanitizeOmsaetningThresholdConfig(config);
    if (!safe.useDailyBudget || typeof OmsaetningDailyThresholds === 'undefined') {
        return {
            warnThreshold: safe.warnThreshold,
            goodThreshold: safe.goodThreshold,
            workingDays: null,
            useDailyBudget: false
        };
    }
    const targets = OmsaetningDailyThresholds.getMonthTargets(monthKey, safe, {
        manualWorkingDays: omsaetningWorkingDaysByMonth[monthKey],
        holidayWeekKeys: getOrdreindgangHolidayWeeksSet(),
        excludeCompanyHolidayWeeks: shouldIgnoreOrdreindgangHolidayWeeks()
    });
    return {
        warnThreshold: targets.warnThresholdMio,
        goodThreshold: targets.goodThresholdMio,
        workingDays: targets.workingDays,
        useDailyBudget: true
    };
}

function renderOmsaetningThresholdHelp(config) {
    const help = document.getElementById('omsaetningThresholdHelp');
    if (!help) return;
    const safe = sanitizeOmsaetningThresholdConfig(config);
    if (!safe.useDailyBudget) {
        help.textContent = 'Manuelle månedsgrænser: ' + safe.warnThreshold.toFixed(1).replace('.', ',') + ' / ' + safe.goodThreshold.toFixed(1).replace('.', ',') + ' Mio.';
        return;
    }
    help.textContent = 'Aktiv: ' + Math.round(safe.dailyBreakEvenDkk).toLocaleString('da-DK') + ' DKK nulpunkt og ' + Math.round(safe.dailyBudgetDkk).toLocaleString('da-DK') + ' DKK budget pr. arbejdsdag. Weekender, danske helligdage og registrerede ferieuger trækkes fra.';
}

async function onOmsaetningThresholdConfigChanged() {
    const safe = saveOmsaetningThresholdUiConfig(getOmsaetningThresholdInputs());
    applyOmsaetningThresholdInputs(safe);
    scheduleOmsaetningAutoReload();
    try {
        const response = await fetch('/omsaetning/daily-budget-settings', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                useDailyBudget: safe.useDailyBudget,
                dailyBreakEvenDkk: safe.dailyBreakEvenDkk,
                dailyBudgetDkk: safe.dailyBudgetDkk
            })
        });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'Dagsmål kunne ikke gemmes');
    } catch (error) {
        const help = document.getElementById('omsaetningThresholdHelp');
        if (help) help.textContent = 'Fejl: ' + String(error.message || error);
    }
}

async function loadOmsaetningThresholdForCustomer(custNo) {
    const key = String(custNo || '').trim();
    if (!/^\d{1,20}$/.test(key)) return;
    const token = ++omsaetningThresholdLoadToken;

    try {
        const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(key));
        if (!response.ok) throw new Error('HTTP ' + response.status);
        const payload = await response.json();
        if (token !== omsaetningThresholdLoadToken) return;
        const config = sanitizeOmsaetningThresholdConfig(payload);
        omsaetningThresholdsByCustomer.set(key, config);
        applyOmsaetningThresholdInputs(config);
        renderOmsaetningCustomerThresholds();
    } catch (err) {
        if (token !== omsaetningThresholdLoadToken) return;
        console.warn('loadOmsaetningThresholdForCustomer failed:', err && err.message ? err.message : err);
    }
}

function renderOmsaetningCustomerThresholds() {
    const wrap = document.getElementById('omsaetningCustomerThresholds');
    if (!wrap) return;

    const selectedEntries = Array.from(omsaetningSelectedCustomers.entries());
    if (selectedEntries.length === 0) {
        wrap.innerHTML = '';
        return;
    }

    const rows = selectedEntries.map(([custNo, custName]) => {
        const key = String(custNo || '').trim();
        const threshold = omsaetningThresholdsByCustomer.get(key);
        if (!threshold) {
            return '<div class="omsaetning-customer-threshold-row"><span class="cust">' +
                escapeHtmlFE(String(custName || key)) + ' (' + escapeHtmlFE(key) + ')' +
                '</span><span class="thr">tærskler: indlæser...</span></div>';
        }
        const thresholdText = threshold.useDailyBudget
            ? ('dagsmål: ' + Math.round(threshold.dailyBreakEvenDkk).toLocaleString('da-DK') + ' / ' + Math.round(threshold.dailyBudgetDkk).toLocaleString('da-DK') + ' DKK')
            : ('tærskler: ' + Number(threshold.warnThreshold).toFixed(1) + ' / ' + Number(threshold.goodThreshold).toFixed(1) + ' Mio');
        return '<div class="omsaetning-customer-threshold-row"><span class="cust">' +
            escapeHtmlFE(String(custName || key)) + ' (' + escapeHtmlFE(key) + ')' +
            '</span><span class="thr">' + escapeHtmlFE(thresholdText) + '</span></div>';
    });

    wrap.innerHTML = rows.join('');
}

async function refreshOmsaetningThresholdsForSelectedCustomers(options) {
    const safeOptions = options && typeof options === 'object' ? options : {};
    const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys())
        .map(v => String(v || '').trim())
        .filter(v => /^\d{1,20}$/.test(v));

    if (selectedCustomers.length === 0) {
        omsaetningThresholdLoadToken += 1;
        omsaetningThresholdsByCustomer = new Map();
        renderOmsaetningCustomerThresholds();
        return;
    }

    const token = ++omsaetningThresholdLoadToken;
    renderOmsaetningCustomerThresholds();

    const results = await Promise.all(selectedCustomers.map(async custNo => {
        try {
            const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(custNo));
            if (!response.ok) throw new Error('HTTP ' + response.status);
            const payload = await response.json();
            return { custNo, ...sanitizeOmsaetningThresholdConfig(payload) };
        } catch (err) {
            console.warn('refresh threshold failed for', custNo, err && err.message ? err.message : err);
            return {
                custNo,
                warnThreshold: OMSAETNING_DEFAULT_WARN_THRESHOLD,
                goodThreshold: OMSAETNING_DEFAULT_GOOD_THRESHOLD,
                useDailyBudget: false,
                dailyBreakEvenDkk: OMSAETNING_DEFAULT_DAILY_BREAK_EVEN_DKK,
                dailyBudgetDkk: OMSAETNING_DEFAULT_DAILY_BUDGET_DKK
            };
        }
    }));

    if (token !== omsaetningThresholdLoadToken) return;

    omsaetningThresholdsByCustomer = new Map(results.map(item => [item.custNo, {
        warnThreshold: item.warnThreshold,
        goodThreshold: item.goodThreshold,
        useDailyBudget: item.useDailyBudget,
        dailyBreakEvenDkk: item.dailyBreakEvenDkk,
        dailyBudgetDkk: item.dailyBudgetDkk
    }]));

    if (safeOptions.applySingleSelectionToInputs === true && selectedCustomers.length === 1) {
        const single = omsaetningThresholdsByCustomer.get(selectedCustomers[0]);
        if (single) {
            applyOmsaetningThresholdInputs(single);
            scheduleOmsaetningAutoReload();
        }
    }

    renderOmsaetningCustomerThresholds();
}

async function persistOmsaetningThresholdsForCustomers(customerNos, config) {
    const customerKeys = Array.from(new Set((Array.isArray(customerNos) ? customerNos : [])
        .map(v => String(v || '').trim())
        .filter(v => /^\d{1,20}$/.test(v))));
    if (customerKeys.length === 0) return;
    const safeConfig = sanitizeOmsaetningThresholdConfig(config);

    await Promise.all(customerKeys.map(async custNo => {
        const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(custNo), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(safeConfig)
        });
        if (!response.ok) throw new Error('HTTP ' + response.status);
    }));
}

function showOmsaetningThresholdPersistDialog(customerKeys, config) {
    return new Promise(resolve => {
        const safeConfig = sanitizeOmsaetningThresholdConfig(config);
        const listHtml = customerKeys.map(custNo => {
            const name = String(omsaetningSelectedCustomers.get(custNo) || '').trim();
            const label = custNo + (name ? (' - ' + name) : '');
            return '<button type="button" class="omsaetning-persist-customer-option" data-cust="' + escapeHtmlFE(custNo) + '">' + escapeHtmlFE(label) + '</button>';
        }).join('');

        const defaultCustomer = String(customerKeys[0] || '');
        let selectedCustomer = defaultCustomer;
        const overlay = document.createElement('div');
        overlay.className = 'omsaetning-persist-overlay';
        overlay.innerHTML =
            '<div class="omsaetning-persist-dialog" role="dialog" aria-modal="true" aria-label="Gem tærskler">' +
                '<div class="omsaetning-persist-head"><h4>Gem tærskler for valgte kunder</h4></div>' +
                '<div class="omsaetning-persist-body">' +
                    '<div>Flere kunder er valgt. Vælg hvor de nye tærskler skal gemmes.</div>' +
                    '<div class="omsaetning-persist-customer-list">' + listHtml + '</div>' +
                    '<div class="omsaetning-persist-thr">' + (safeConfig.useDailyBudget
                        ? ('Nye dagsmål: ' + escapeHtmlFE(Math.round(safeConfig.dailyBreakEvenDkk).toLocaleString('da-DK')) + ' / ' + escapeHtmlFE(Math.round(safeConfig.dailyBudgetDkk).toLocaleString('da-DK')) + ' DKK')
                        : ('Nye tærskler: ' + escapeHtmlFE(Number(safeConfig.warnThreshold).toFixed(1)) + ' / ' + escapeHtmlFE(Number(safeConfig.goodThreshold).toFixed(1)) + ' Mio')) + '</div>' +
                    '<div class="omsaetning-persist-pick">Valgt kunde: <span id="omsaetningPersistPicked" class="omsaetning-persist-picked">' + escapeHtmlFE(defaultCustomer) + '</span></div>' +
                    '</div>' +
                    '<div class="omsaetning-persist-actions">' +
                        '<button type="button" class="primary" data-action="all">Gem alle</button>' +
                        '<button type="button" class="ghost" data-action="single">Gem valgt kunde</button>' +
                        '<button type="button" class="danger" data-action="none">Gem ikke</button>' +
                    '</div>' +
                '</div>' +
            '</div>';

        const closeWith = value => {
            overlay.remove();
            resolve(value);
        };

        overlay.addEventListener('click', ev => {
            if (ev.target === overlay) closeWith('NONE');
        });

        overlay.querySelector('[data-action="all"]').addEventListener('click', () => closeWith('ALL'));
        overlay.querySelector('[data-action="none"]').addEventListener('click', () => closeWith('NONE'));
        overlay.querySelector('[data-action="single"]').addEventListener('click', () => {
            closeWith(String(selectedCustomer || '').trim());
        });

        const pickedEl = overlay.querySelector('#omsaetningPersistPicked');
        const customerButtons = Array.from(overlay.querySelectorAll('.omsaetning-persist-customer-option'));

        const renderPicked = () => {
            for (const btn of customerButtons) {
                const btnCust = String(btn.getAttribute('data-cust') || '');
                btn.classList.toggle('active', btnCust === selectedCustomer);
            }
            if (pickedEl) pickedEl.textContent = selectedCustomer || '-';
        };

        for (const btn of customerButtons) {
            btn.addEventListener('click', () => {
                selectedCustomer = String(btn.getAttribute('data-cust') || '').trim();
                renderPicked();
            });
        }

        overlay.addEventListener('keydown', ev => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                closeWith(String(selectedCustomer || '').trim());
            }
            if (ev.key === 'Escape') {
                ev.preventDefault();
                closeWith('NONE');
            }
        });

        document.body.appendChild(overlay);
        renderPicked();
        if (customerButtons.length > 0) {
            customerButtons[0].focus();
        }
    });
}

async function resolveOmsaetningThresholdPersistTargets(selectedCustomers, config, options) {
    const safeOptions = options && typeof options === 'object' ? options : {};
    const customerKeys = Array.from(new Set((Array.isArray(selectedCustomers) ? selectedCustomers : [])
        .map(v => String(v || '').trim())
        .filter(v => /^\d{1,20}$/.test(v))));

    if (customerKeys.length === 0) return [];
    if (safeOptions.silentValidation === true) return [];
    if (safeOptions.persistThresholdsOnUpdate !== true) return [];

    if (customerKeys.length === 1) {
        return customerKeys;
    }

    const answerRaw = await showOmsaetningThresholdPersistDialog(customerKeys, config);
    const answer = String(answerRaw || '').trim().toUpperCase();

    if (!answer || answer === 'NONE' || answer === 'NO' || answer === 'N') {
        return [];
    }
    if (answer === 'ALL' || answer === 'A') {
        return customerKeys;
    }

    const exact = customerKeys.find(c => c.toUpperCase() === answer);
    if (exact) return [exact];

    alert('Ugyldigt valg for tærskel-gemning. Ingen tærskler blev gemt.');
    return [];
}

function onOmsaetningSelectedCustomersChanged() {
    const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys()).filter(Boolean);
    renderOmsaetningCustomerMode();
    if (selectedCustomers.length === 0) {
        omsaetningThresholdLoadToken += 1;
        omsaetningThresholdsByCustomer = new Map();
        applyOmsaetningThresholdInputs(loadOmsaetningThresholdUiConfig());
        renderOmsaetningCustomerThresholds();
        scheduleOmsaetningAutoReload();
        return;
    }

    refreshOmsaetningThresholdsForSelectedCustomers({ applySingleSelectionToInputs: true });
    scheduleOmsaetningAutoReload();
}

function scheduleOmsaetningAutoReload() {
    if (omsaetningAutoReloadTimer) {
        clearTimeout(omsaetningAutoReloadTimer);
    }
    omsaetningAutoReloadTimer = setTimeout(() => {
        omsaetningAutoReloadTimer = null;
        loadOmsaetningSummary({ silentValidation: true });
    }, OMSAETNING_AUTO_RELOAD_DELAY_MS);
}

function getOmsaetningStatusLabel(statusClass) {
    if (statusClass === 'good') return 'Over mål';
    if (statusClass === 'mid') return 'Nær mål';
    return 'Under mål';
}

function getOmsaetningColor(index) {
    const palette = ['#1565c0', '#00acc1', '#00897b', '#7b1fa2', '#ef6c00', '#5e35b1', '#43a047', '#c62828'];
    return palette[index % palette.length];
}

function renderOmsaetningCharts(rows, forcedMonthKeys) {
    const chartsWrap = document.getElementById('omsaetningChartsWrap');
    const stackedSvg = document.getElementById('omsaetningStackedChart');
    const trendSvg = document.getElementById('omsaetningTrendChart');
    const legend = document.getElementById('omsaetningLegend');
    const stackedTitle = document.getElementById('omsaetningStackedTitle');
    if (!chartsWrap || !stackedSvg || !trendSvg || !legend) return;

    const safeRows = Array.isArray(rows) ? rows : [];
    const safeForcedMonths = Array.isArray(forcedMonthKeys) ? forcedMonthKeys.map(v => String(v || '').trim()).filter(Boolean) : [];

    if (safeRows.length === 0 && safeForcedMonths.length === 0) {
        chartsWrap.style.display = 'none';
        stackedSvg.innerHTML = '';
        trendSvg.innerHTML = '';
        legend.innerHTML = '';
        if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';
        return;
    }

    const monthMap = new Map();
    const accountOrder = [];
    const seenAccounts = new Set();
    for (const row of safeRows) {
        const monthKey = normalizeOmsaetningMonthKey(row.date);
        if (!monthMap.has(monthKey)) monthMap.set(monthKey, new Map());
        const accountKey = String(row.acNo || '');
        if (!seenAccounts.has(accountKey)) {
            seenAccounts.add(accountKey);
            accountOrder.push({ acNo: accountKey, name: String(row.name || '') });
        }
        const monthAcc = monthMap.get(monthKey);
        monthAcc.set(accountKey, (monthAcc.get(accountKey) || 0) + Number(row.revenueMio || 0));
    }

    const monthKeys = (safeForcedMonths.length > 0 ? safeForcedMonths : Array.from(monthMap.keys())).sort((a, b) => String(a).localeCompare(String(b)));
    if (monthKeys.length === 0) {
        chartsWrap.style.display = 'none';
        stackedSvg.innerHTML = '';
        trendSvg.innerHTML = '';
        legend.innerHTML = '';
        if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';
        return;
    }
    const monthlyTotals = monthKeys.map(key => {
        const m = monthMap.get(key) || new Map();
        let t = 0;
        for (const value of m.values()) t += Number(value || 0);
        return t;
    });

    function buildScale(values, fallbackAbs) {
        const safeValues = Array.isArray(values) ? values : [];
        let min = 0;
        let max = 0;
        for (const raw of safeValues) {
            const value = Number(raw);
            if (!Number.isFinite(value)) continue;
            if (value < min) min = value;
            if (value > max) max = value;
        }
        if (min === 0 && max === 0) {
            max = Number.isFinite(fallbackAbs) && fallbackAbs > 0 ? fallbackAbs : 0.1;
        }
        if (Math.abs(max - min) < 0.000001) {
            if (max > 0) min = 0;
            else if (min < 0) max = 0;
            else max = 0.1;
        }
        return { min, max, span: max - min };
    }

    const compareCustomers = Array.from(omsaetningSelectedCustomers.entries())
        .map(([custNo, name]) => ({ custNo: String(custNo || '').trim(), name: String(name || '').trim() }))
        .filter(c => c.custNo);
    const showCustomerComparison = compareCustomers.length > 1;

    const leftPad = 48;
    const topPad = 30;
    const bottomPad = 42;
    const chartHeight = 214;
    const innerHeight = chartHeight - topPad - bottomPad;
    const barWidth = 34;
    const barGap = 18;
    const innerWidth = Math.max(560, monthKeys.length * (barWidth + barGap));
    const viewWidth = leftPad + innerWidth + 20;
    const viewHeight = chartHeight;

    function toY(value, scale) {
        const safeScale = scale || { min: 0, max: 1, span: 1 };
        const safeValue = Number.isFinite(Number(value)) ? Number(value) : 0;
        const ratio = (safeScale.max - safeValue) / (safeScale.span || 1);
        return topPad + (ratio * innerHeight);
    }

    function appendMonthTotalLabel(svgHtml, x, width, monthKey, total, scale) {
        const value = Number(total || 0);
        const anchorY = value >= 0 ? toY(value, scale) - 6 : toY(value, scale) + 14;
        const boundedY = Math.max(12, Math.min(chartHeight - bottomPad - 2, anchorY));
        const label = formatMio(value) + ' Mio';
        const title = formatMonthDa(monthKey) + ': ' + label + ' DKK';
        return svgHtml + '<text class="omsaetning-month-total" x="' + (x + width / 2) + '" y="' + boundedY + '" text-anchor="middle" font-size="10" font-weight="700" fill="#173452"><title>' + escapeHtmlFE(title) + '</title>' + escapeHtmlFE(label) + '</text>';
    }

    function appendGrid(svgHtml, width, scale) {
        let out = svgHtml;
        const ticks = 4;
        for (let i = 0; i <= ticks; i++) {
            const ratio = i / ticks;
            const tickValue = scale.max - (ratio * scale.span);
            const y = topPad + (innerHeight * ratio);
            out += '<line x1="' + leftPad + '" y1="' + y + '" x2="' + (leftPad + width) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
            out += '<text x="' + (leftPad - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="#5f7892">' + escapeHtmlFE(formatMio(tickValue)) + '</text>';
        }
        const zeroY = toY(0, scale);
        out += '<line x1="' + leftPad + '" y1="' + zeroY + '" x2="' + (leftPad + width) + '" y2="' + zeroY + '" stroke="#8fa8c2" stroke-width="1.2" />';
        return out;
    }

    let stackedSvgHtml = '<g>';
    if (showCustomerComparison) {
        if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (kunde-sammenligning)';

        const byMonthCustomer = new Map();
        for (const row of safeRows) {
            const monthKey = normalizeOmsaetningMonthKey(row.date);
            const custNo = String(row.custNo || '').trim();
            if (!monthKey || !custNo) continue;
            if (!byMonthCustomer.has(monthKey)) byMonthCustomer.set(monthKey, new Map());
            const map = byMonthCustomer.get(monthKey);
            map.set(custNo, (map.get(custNo) || 0) + Number(row.revenueMio || 0));
        }

        const groupedValues = [];
        for (const monthKey of monthKeys) {
            const values = byMonthCustomer.get(monthKey) || new Map();
            for (const customer of compareCustomers) {
                const v = Number(values.get(customer.custNo) || 0);
                groupedValues.push(v);
            }
        }
        const groupedScale = buildScale(groupedValues, 0.1);
        const groupedZeroY = toY(0, groupedScale);

        const groupedBarWidth = 14;
        const groupedBarGap = 5;
        const monthGroupGap = 18;
        const perMonthGroupWidth = (compareCustomers.length * groupedBarWidth) + ((compareCustomers.length - 1) * groupedBarGap);
        const groupedInnerWidth = Math.max(560, monthKeys.length * (perMonthGroupWidth + monthGroupGap));

        stackedSvgHtml = appendGrid(stackedSvgHtml, groupedInnerWidth, groupedScale);

        monthKeys.forEach((monthKey, monthIndex) => {
            const monthX = leftPad + monthIndex * (perMonthGroupWidth + monthGroupGap);
            const values = byMonthCustomer.get(monthKey) || new Map();
            compareCustomers.forEach((customer, customerIndex) => {
                const value = Number(values.get(customer.custNo) || 0);
                const yValue = toY(value, groupedScale);
                const y = value >= 0 ? yValue : groupedZeroY;
                const h = Math.max(1, Math.abs(groupedZeroY - yValue));
                const x = monthX + customerIndex * (groupedBarWidth + groupedBarGap);
                const titleText = formatMonthDa(monthKey) + ' - ' + String(customer.name || customer.custNo) + ' (' + customer.custNo + '): ' + formatMio(value) + ' Mio DKK (' + formatDkkFromMio(value) + ' DKK)';
                stackedSvgHtml += '<rect x="' + x + '" y="' + y + '" width="' + groupedBarWidth + '" height="' + h + '" fill="' + getOmsaetningColor(customerIndex) + '" rx="2"><title>' + escapeHtmlFE(titleText) + '</title></rect>';
            });

            stackedSvgHtml = appendMonthTotalLabel(stackedSvgHtml, monthX, perMonthGroupWidth, monthKey, monthlyTotals[monthIndex], groupedScale);

            const labelX = monthX + (perMonthGroupWidth / 2);
            stackedSvgHtml += '<text x="' + labelX + '" y="' + (topPad + innerHeight + 14) + '" text-anchor="middle" font-size="10" fill="#47617c">' + escapeHtmlFE(formatMonthDa(monthKey)) + '</text>';
        });

        stackedSvgHtml += '</g>';
        stackedSvg.setAttribute('viewBox', '0 0 ' + (leftPad + groupedInnerWidth + 20) + ' ' + viewHeight);
        stackedSvg.innerHTML = stackedSvgHtml;

        legend.innerHTML = compareCustomers.map((customer, idx) =>
            '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:' + getOmsaetningColor(idx) + ';"></span>' +
            escapeHtmlFE(String(customer.name || customer.custNo)) + ' (' + escapeHtmlFE(customer.custNo) + ')</span>'
        ).join('');
    } else {
        if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';
        const chart = window.GohReportCharts.revenueStacked(safeRows, monthKeys);
        stackedSvg.setAttribute('viewBox', chart.viewBox);
        stackedSvg.innerHTML = chart.html;
        legend.innerHTML = chart.legend;
    }

    const trendLeftPad = 42;
    const trendTopPad = 16;
    const trendBottomPad = 28;
    const trendHeight = 190;
    const trendInnerHeight = trendHeight - trendTopPad - trendBottomPad;
    const trendInnerWidth = Math.max(560, monthKeys.length * 54);
    const trendViewWidth = trendLeftPad + trendInnerWidth + 16;
    const trendViewHeight = trendHeight;
    const trendScale = buildScale(monthlyTotals, 0.1);

    function toTrendY(value) {
        const safeValue = Number.isFinite(Number(value)) ? Number(value) : 0;
        const ratio = (trendScale.max - safeValue) / (trendScale.span || 1);
        return trendTopPad + (ratio * trendInnerHeight);
    }

    let trendSvgHtml = '<g>';
    for (let i = 0; i <= 4; i++) {
        const ratio = i / 4;
        const y = trendTopPad + (trendInnerHeight * ratio);
        const tickValue = trendScale.max - (ratio * trendScale.span);
        trendSvgHtml += '<line x1="' + trendLeftPad + '" y1="' + y + '" x2="' + (trendLeftPad + trendInnerWidth) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
        trendSvgHtml += '<text x="' + (trendLeftPad - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="#5f7892">' + escapeHtmlFE(formatMio(tickValue)) + '</text>';
        }

    const trendZeroY = toTrendY(0);
    trendSvgHtml += '<line x1="' + trendLeftPad + '" y1="' + trendZeroY + '" x2="' + (trendLeftPad + trendInnerWidth) + '" y2="' + trendZeroY + '" stroke="#8fa8c2" stroke-width="1.2" />';

    const points = monthKeys.map((monthKey, idx) => {
        const x = trendLeftPad + (trendInnerWidth * (monthKeys.length === 1 ? 0.5 : (idx / (monthKeys.length - 1))));
        const y = toTrendY(monthlyTotals[idx]);
        return { x, y, monthKey, total: monthlyTotals[idx] };
    });
    if (points.length === 0) {
        chartsWrap.style.display = 'none';
        trendSvg.innerHTML = '';
        return;
    }
    const linePath = points.map((p, idx) => (idx === 0 ? 'M' : 'L') + p.x + ' ' + p.y).join(' ');
    const areaPath = linePath + ' L ' + points[points.length - 1].x + ' ' + trendZeroY + ' L ' + points[0].x + ' ' + trendZeroY + ' Z';
    trendSvgHtml += '<path d="' + areaPath + '" fill="rgba(21,101,192,0.12)" />';
    trendSvgHtml += '<path d="' + linePath + '" fill="none" stroke="#1565c0" stroke-width="3" />';
    points.forEach(p => {
        const pointTitle = formatMonthDa(p.monthKey) + ': ' + formatMio(p.total) + ' Mio DKK (' + formatDkkFromMio(p.total) + ' DKK)';
        trendSvgHtml += '<circle cx="' + p.x + '" cy="' + p.y + '" r="3.5" fill="#0f3560"><title>' + escapeHtmlFE(pointTitle) + '</title></circle>';
    });
    trendSvgHtml += '</g>';

    trendSvg.setAttribute('viewBox', '0 0 ' + trendViewWidth + ' ' + trendViewHeight);
    trendSvg.innerHTML = trendSvgHtml;
    chartsWrap.style.display = 'grid';
}
