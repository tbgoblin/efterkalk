// ── Kundefaktura-oversigt / "Månedens DB" · client ───────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, formatNumber, pushModalStack, removeModalStack,
// searchOrderByNo (dichiarato nell'inline script, chiamato solo da onclick
// generati in HTML — risolto al momento del click, non del caricamento).
// ── Kundefaktura-oversigt ─────────────────────────────────────
let _kfSelectedCustNo  = null;
let _kfSelectedCustNm  = '';
let _kfInvoiceRows     = [];
let _kfMarginMap       = {};    // OrdNo → { totalRevenue, totalCost, styklisteFallbackCost }
let _kfCostMode        = 'adjusted';
let _kfCustomerTimer   = null;
let _kfMarginAbort     = false;
let _kfMarginFetchRunning = false;
let _kfAllCustomers   = false;
let _kfTrendRequestId = 0;
let _kfBackfillRequested = false;
let _kfVareforbrugGL = null; // Vareforbrug fra bogføringen (kontogruppe 12_Vareforbrug) for den valgte måned
let _kfPurchaseGap = null; // { missingValue, missingOrderCount, partialValue, partialOrderCount, totalValue }
let _kfVarelagerDelta = null; // { currentValue, previousValue, currentMonth, previousMonth, delta }
let _kfPlateMode = 'fifo'; // Prisgrundlag for Pladelager i Δ Varelager — 'fifo' eller 'standard'
let _kfVareforbrugMonth = null; // Måneden sidst sendt til /efterkalk/vareforbrug, så plateMode-skift kan genhente uden fuld reload
let _kfLoadSeq = 0; // Stiger ved hver ny indlæsning; ældre asynkrone svar må ikke skrive over den nyere
let _kfLoadedMonth = null; // Måneden de viste rækker tilhører — ikke input-feltet, som kan have ændret sig under en igangværende hentning
let _kfUbRows = [];
let _kfUbRequest = 0;
let _kfUbLoaded = false;
let _kfSummaryInvoicesReady = false;
let _kfVareforbrugRequest = 0;

function _kfBillableRows() {
    return _kfInvoiceRows.filter(row => !row.IsUB);
}

function _refreshKfConfigurableSummary() {
    if (!_kfAllCustomers || !window.GohMonthlySummary) return;
    const rows=_kfBillableRows(),ready=_kfSummaryInvoicesReady;
    const complete=ready && rows.every(row => Boolean(_kfMarginMap[String(row.OrdNo)]));
    const breakdown=complete && rows.every(row => _kfMarginMap[String(row.OrdNo)].costBreakdownAvailable);
    const sum=field => rows.reduce((total,row) => total + Number(_kfMarginMap[String(row.OrdNo)]?.[field] || 0),0);
    const invoiced=ready ? rows.reduce((total,row) => total + Number(row.InvoAm || 0),0) : null;
    const cost=complete ? rows.reduce((total,row) => total + _kfCost(_kfMarginMap[String(row.OrdNo)]),0) : null;
    const credits=rows.filter(row => Number(row.InvoAm)<0);
    window.GohMonthlySummary.update({
        orders:ready ? rows.length : null,invoiced,cost,
        db:complete ? invoiced-cost : null,dg:complete && invoiced!==0 ? (invoiced-cost)/invoiced*100 : null,
        credits:ready ? credits.reduce((total,row) => total+Number(row.InvoAm),0) : null,
        credit_count:ready ? credits.length : null,
        material:breakdown ? sum('materialCost')+sum('stangCost')+sum('purchasedPartCost')+sum('underleverandorCost') : null,
        material_only:breakdown ? sum('materialCost') : null,stang:breakdown ? sum('stangCost') : null,
        time:breakdown ? sum('timeCost') : null,parts:breakdown ? sum('purchasedPartCost') : null,
        subcontract:breakdown ? sum('underleverandorCost') : null,
        warehouse_delta:_kfVarelagerDelta?.delta ?? null,purchase_gap:_kfPurchaseGap?.totalValue ?? null,
        ub_orders:_kfUbLoaded ? _kfUbRows.length : null
    });
}

async function _loadKfUbOrders(month, seq = _kfLoadSeq) {
    const request = ++_kfUbRequest;
    const section = document.getElementById('kfUbSection');
    if (section) section.style.display = _kfAllCustomers ? 'block' : 'none';
    if (!_kfAllCustomers) return;
    _kfUbLoaded=false;
    const status = document.getElementById('kfUbStatus');
    const tbody = document.getElementById('kfUbBody');
    if (status) status.textContent = 'Henter U/B-ordrer…';
    if (tbody) tbody.innerHTML = '';
    try {
        const response = await fetch('/efterkalk/ub-orders?month=' + encodeURIComponent(month));
        const payload = await response.json();
        if (seq !== _kfLoadSeq || request !== _kfUbRequest) return;
        if (!response.ok || !payload.ok) throw new Error(payload.error || 'Kunne ikke hente U/B-ordrer.');
        _kfUbRows = payload.rows || [];
        _kfUbLoaded=true;
        _refreshKfConfigurableSummary();
        if (status) status.textContent = _kfUbRows.length + ' U/B-ordrer · ' + _kfMonthLabel(month);
        if (tbody) tbody.innerHTML = _kfUbRows.map(row => '<tr>'
            + '<td><button type="button" class="kf-trend-btn" onclick="searchOrderByNo(' + Number(row.OrdNo) + ')">' + Number(row.OrdNo) + '</button></td>'
            + '<td>' + escapeHtml(_fmtKfDateFromInt(row.OrdDt)) + '</td>'
            + '<td>' + escapeHtml(String(row.CustomerName || row.CustNo || '—')) + '</td>'
            + '<td class="kf-ub-note">' + escapeHtml(String(row.text || '')) + '</td>'
            + '<td><button type="button" class="kf-trend-btn" onclick="removeKfUbOrder(' + Number(row.OrdNo) + ')">Fjern fra U/B</button></td></tr>').join('')
            || '<tr><td colspan="5">Ingen U/B-ordrer registreret i måneden.</td></tr>';
    } catch (error) {
        if (seq === _kfLoadSeq && request === _kfUbRequest && status) status.textContent = error.message;
    }
}

async function addKfUbOrder(event) {
    if (event) event.preventDefault();
    const month = _kfLoadedMonth;
    if (!month || !_kfAllCustomers) return;
    const seq = _kfLoadSeq;
    const orderInput = document.getElementById('kfUbOrdNo');
    const noteInput = document.getElementById('kfUbNote');
    const status = document.getElementById('kfUbStatus');
    const button = document.getElementById('kfUbAdd');
    if (button) button.disabled = true;
    try {
        const response = await fetch('/efterkalk/ub-orders', { method:'POST', headers:{ 'Content-Type':'application/json' },
            body:JSON.stringify({ ordNo:Number(orderInput.value), month, text:noteInput.value.trim() }) });
        const payload = await response.json();
        if (!response.ok || !payload.ok) throw new Error(payload.error || 'U/B-ordren kunne ikke gemmes.');
        if (seq !== _kfLoadSeq) return;
        if (typeof orderNotesCache !== 'undefined') orderNotesCache[String(payload.order.OrdNo)] = payload.note;
        _kfInvoiceRows.forEach(row => { if (Number(row.OrdNo) === Number(payload.order.OrdNo)) row.IsUB = true; });
        orderInput.value = ''; noteInput.value = '';
        _renderKfTable(); _updateKfKpis(); _addKfTotalRow();
        await _loadKfUbOrders(month, seq);
    } catch (error) { if (seq === _kfLoadSeq && status) status.textContent = error.message; }
    finally { if (button) button.disabled = false; }
}

async function removeKfUbOrder(ordNo) {
    const month = _kfLoadedMonth;
    const seq = _kfLoadSeq;
    const status = document.getElementById('kfUbStatus');
    try {
        const response = await fetch('/efterkalk/ub-orders/' + Number(ordNo) + '?month=' + encodeURIComponent(month), { method:'DELETE' });
        const payload = await response.json();
        if (!response.ok || !payload.ok) throw new Error(payload.error || 'U/B-registreringen kunne ikke fjernes.');
        if (seq !== _kfLoadSeq) return;
        if (typeof orderNotesCache !== 'undefined') orderNotesCache[String(ordNo)] = payload.note || {};
        _kfInvoiceRows.forEach(row => { if (Number(row.OrdNo) === Number(ordNo)) row.IsUB = false; });
        _renderKfTable(); _updateKfKpis(); _addKfTotalRow();
        await _loadKfUbOrders(month, seq);
    } catch (error) { if (seq === _kfLoadSeq && status) status.textContent = error.message; }
}

function refreshKfOrderClassification(ordNo, note) {
    _kfInvoiceRows.forEach(row => { if (Number(row.OrdNo) === Number(ordNo)) row.IsUB = Boolean(note && note.isUB); });
    if (!_kfLoadedMonth && !_kfInvoiceRows.length) return;
    _renderKfTable(); _updateKfKpis(); _addKfTotalRow();
    if (_kfAllCustomers && _kfLoadedMonth) _loadKfUbOrders(_kfLoadedMonth);
}

function _kfCost(margin, mode) {
    if (!margin) return null;
    const registered = Number(margin.totalCost || 0);
    const fallback = Number(margin.styklisteFallbackCost || 0);
    return registered + ((mode || _kfCostMode) === 'adjusted' ? fallback : 0);
}

// Kompakt: ét felt der folder sig ud til materiale/stang/tid/indkøbte dele/underleverandør,
// i stedet for fem altid-synlige kolonner (sparer bredde, jf. kundenavn-kolonnen).
function _kfBreakdownHtml(complete, v) {
    if (!complete) return '<span style="color:#aaa">…</span>';
    const parts = [
        'Materiale ' + formatNumber(v.materialCost),
        'Stang ' + formatNumber(v.stangCost),
        'Tid ' + formatNumber(v.timeCost),
        'Indkøbte dele ' + formatNumber(v.purchasedPartCost),
        'Underleverandør ' + formatNumber(v.underleverandorCost)
    ];
    return '<details><summary>Detaljer</summary><span class="kf-breakdown-detail">' + escapeHtml(parts.join(' · ')) + '</span></details>';
}

function setKfCostMode(mode) {
    _kfCostMode = mode === 'registered' ? 'registered' : 'adjusted';
    closeKfCustomerTrend();
    _syncKfCostModeUi();
    _renderKfTable();
    _updateKfKpis();
    _addKfTotalRow();
}

function _syncKfCostModeUi() {
    const adjusted = _kfCostMode === 'adjusted';
    const regButton = document.getElementById('kfCostModeRegistered');
    const adjButton = document.getElementById('kfCostModeAdjusted');
    if (regButton) regButton.disabled = !adjusted;
    if (adjButton) adjButton.disabled = adjusted;
    const costLabel = adjusted ? 'Kost inkl. stykliste' : 'Registreret kost';
    const customerHead = document.getElementById('kfCustomerCostHead');
    const orderHead = document.getElementById('kfOrderCostHead');
    if (customerHead) customerHead.textContent = costLabel;
    if (orderHead) orderHead.textContent = costLabel;
    const hint = document.getElementById('kfCostModeHint');
    if (hint) {
        hint.style.display = _kfInvoiceRows.length ? 'block' : 'none';
        const fallbackTotal = _kfBillableRows().reduce((sum, row) => sum + Number(_kfMarginMap[String(row.OrdNo)]?.styklisteFallbackCost || 0), 0);
        hint.textContent = adjusted
            ? 'Korrigeret visning: lægger planlagt kost fra stykliste til, når en færdigmeldt operation mangler registreret tid. Tillæg: ' + formatNumber(fallbackTotal) + ' DKK.'
            : 'Registreret visning: viser kun de omkostninger, der faktisk er registreret på ordren.';
    }
}

function openKundefakturaModal() {
    _kfAllCustomers = false;
    const overlay = document.getElementById('kfModalOverlay');
    if (overlay) overlay.classList.add('open');
    // Default: last 365 days → today
    const toDate = new Date();
    const frDate = new Date(toDate);
    frDate.setFullYear(frDate.getFullYear() - 1);
    const fmt = d => d.toISOString().slice(0,10);
    const kfFrom = document.getElementById('kfFromDate');
    const kfTo   = document.getElementById('kfToDate');
    if (kfFrom && !kfFrom.value) kfFrom.value = fmt(frDate);
    if (kfTo   && !kfTo.value)   kfTo.value   = fmt(toDate);
    _setKfModalMode();

    // Always start in editable customer-search mode.
    _kfSelectedCustNo = null;
    _kfSelectedCustNm = '';
    _kfInvoiceRows = [];
    _kfMarginMap = {};
    _resetKfKpis();
    kfSetStatus('');
    const inp = document.getElementById('kfCustomerInput');
    if (inp) {
        inp.disabled = false;
        inp.readOnly = false;
        inp.value = '';
        setTimeout(() => {
            try { inp.focus({ preventScroll: true }); } catch { inp.focus(); }
        }, 0);
    }
    const btn = document.getElementById('kfLoadBtn');
    if (btn) btn.disabled = true;

    pushModalStack('kfModalOverlay');
}

function _kfPreviousMonthKey(date = new Date()) {
    // Forrige kalendermåned, ikke indeværende: på dag 1 i en ny måned er indeværende
    // måned praktisk talt tom (ingen fakturaer, intet Vareforbrug endnu bogført), så den
    // er næsten aldrig den relevante standardmåned for et månedligt fakturaoverblik.
    const year = date.getFullYear(), month = date.getMonth(); // 0-indexed
    const prevYear = month === 0 ? year - 1 : year;
    const prevMonth = month === 0 ? 12 : month;
    return prevYear + '-' + String(prevMonth).padStart(2, '0');
}

function openMaanedsfakturaModal() {
    _kfAllCustomers = true;
    const overlay = document.getElementById('kfModalOverlay');
    if (overlay) overlay.classList.add('open');
    const month = document.getElementById('kfMonth');
    if (month && !month.value) month.value = _kfPreviousMonthKey();
    _kfSelectedCustNo = null;
    _kfSelectedCustNm = '';
    _kfInvoiceRows = [];
    _kfMarginMap = {};
    _resetKfKpis();
    _setKfModalMode();
    const btn = document.getElementById('kfLoadBtn');
    if (btn) btn.disabled = false;
    pushModalStack('kfModalOverlay');
    loadKundefakturaInvoices();
}

function _setKfModalMode() {
    if (window.GohMonthlySummary) window.GohMonthlySummary.mode(_kfAllCustomers);
    const show = (id, visible) => { const el = document.getElementById(id); if (el) el.style.display = visible ? '' : 'none'; };
    show('kfCustomerField', !_kfAllCustomers);
    show('kfMonthField', _kfAllCustomers);
    show('kfFromField', !_kfAllCustomers);
    show('kfToField', !_kfAllCustomers);
    const title = document.getElementById('kfModalTitle');
    if (title) title.textContent = _kfAllCustomers ? '📅 Månedsfakturaer og dækningsbidrag' : '📋 Fakturaoversigt — Kunde';
    const btn = document.getElementById('kfLoadBtn');
    if (btn) btn.textContent = _kfAllCustomers ? 'Hent måned' : 'Hent fakturaer';
}

function closeKundefakturaModal() {
    const overlay = document.getElementById('kfModalOverlay');
    if (overlay) overlay.classList.remove('open');
    _kfMarginAbort = true;
    _kfTrendRequestId++;
    removeModalStack('kfModalOverlay');
}

function scheduleKfCustomerSearch() {
    if (_kfCustomerTimer) clearTimeout(_kfCustomerTimer);
    _kfCustomerTimer = setTimeout(_doKfCustomerSearch, 220);
}

function onKfCustomerInputChanged() {
    const inp = document.getElementById('kfCustomerInput');
    const currentText = String((inp && inp.value) || '').trim();
    const selectedText = (_kfSelectedCustNm && _kfSelectedCustNo)
        ? (_kfSelectedCustNm + ' (' + _kfSelectedCustNo + ')')
        : '';

    if (!currentText || currentText !== selectedText) {
        _kfSelectedCustNo = null;
        _kfSelectedCustNm = '';
        const btn = document.getElementById('kfLoadBtn');
        if (btn) btn.disabled = true;
    }

    scheduleKfCustomerSearch();
}

async function _doKfCustomerSearch() {
    const inp = document.getElementById('kfCustomerInput');
    const q   = inp ? inp.value.trim() : '';
    const dd  = document.getElementById('kfCustomerDropdown');
    if (!dd) return;
    if (q.length < 1) { dd.style.display = 'none'; return; }
    try {
        const r = await fetch('/efterkalk/customers?q=' + encodeURIComponent(q));
        const d = await r.json();
        const rows = d.rows || [];
        if (!rows.length) { dd.style.display = 'none'; return; }
        dd.innerHTML = rows.map((c, i) =>
            '<div class="kf-customer-option" onmousedown="_kfPickCustomer(' + i + ',event)">' +
            escapeHtml(c.Nm) + '<span class="kf-cno">' + escapeHtml(String(c.CustNo)) + '</span></div>'
        ).join('');
        dd._rows = rows;
        dd.style.display = 'block';
    } catch { dd.style.display = 'none'; }
}

function _kfPickCustomer(idx, e) {
    if (e) e.preventDefault();
    const dd = document.getElementById('kfCustomerDropdown');
    const rows = dd && dd._rows ? dd._rows : [];
    const c = rows[idx];
    if (!c) return;
    _kfSelectedCustNo = c.CustNo;
    _kfSelectedCustNm = c.Nm;
    const inp = document.getElementById('kfCustomerInput');
    if (inp) inp.value = c.Nm + ' (' + c.CustNo + ')';
    dd.style.display = 'none';
    const btn = document.getElementById('kfLoadBtn');
    if (btn) btn.disabled = false;
}

function hideKfDropdown() {
    setTimeout(() => {
        const dd = document.getElementById('kfCustomerDropdown');
        if (dd) dd.style.display = 'none';
    }, 200);
}

async function loadKundefakturaInvoices() {
    if (!_kfAllCustomers && !_kfSelectedCustNo) { kfSetStatus('Vælg en kunde først.'); return; }
    let from = (document.getElementById('kfFromDate') || {}).value || '';
    let to   = (document.getElementById('kfToDate')   || {}).value || '';
    let selectedMonth = '';
    if (_kfAllCustomers) {
        const month = (document.getElementById('kfMonth') || {}).value || '';
        if (month.length !== 7 || month.charAt(4) !== '-') { kfSetStatus('Vælg en måned.'); return; }
        selectedMonth = month;
        const parts = month.split('-').map(Number);
        from = month + '-01';
        to = month + '-' + String(new Date(parts[0], parts[1], 0).getDate()).padStart(2, '0');
    }
    if (!from) { kfSetStatus('Vælg en fra-dato.'); return; }
    const btn = document.getElementById('kfLoadBtn');
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Henter…'; }
    kfSetStatus(_kfAllCustomers ? 'Henter alle fakturaer for den valgte måned…' : 'Henter fakturaer for ' + _kfSelectedCustNm + '…');
    _kfMarginAbort = false;
    _kfMarginFetchRunning = false;
    const seq = ++_kfLoadSeq;
    _kfLoadedMonth = _kfAllCustomers ? selectedMonth : null;
    _kfInvoiceRows = [];
    _kfMarginMap   = {};
    _resetKfKpis();
    if (_kfAllCustomers) {
        if (window.GohMonthlySummary) window.GohMonthlySummary.load(selectedMonth);
        _refreshKfConfigurableSummary();
        _fetchKfVareforbrug(selectedMonth, seq);
        _loadKfUbOrders(selectedMonth, seq);
    }

    try {
        if (_kfAllCustomers) await _loadKfStoredMonth(selectedMonth, seq);
        if (seq !== _kfLoadSeq) return;
        const url = '/efterkalk/customer-invoices?' + (_kfAllCustomers ? 'scope=all' : 'custno=' + _kfSelectedCustNo) +
            '&includeAllAmounts=0&from=' + encodeURIComponent(from) + '&to=' + encodeURIComponent(to || new Date().toISOString().slice(0,10));
        const r = await fetch(url);
        const d = await r.json();
        if (seq !== _kfLoadSeq) return;
        if (!d.ok) throw new Error(d.error || 'Fejl');
        _kfInvoiceRows = d.rows || [];
        _kfSummaryInvoicesReady=true;
        kfSetStatus(_kfInvoiceRows.length + ' faktura- og kreditnotaordrer hentet' + (_kfAllCustomers ? '' : ' for ' + _kfSelectedCustNm) +
            ' (' + _fmtKfDate(d.fromInt) + ' – ' + _fmtKfDate(d.toInt) + ') · U/B vises separat');
        _renderKfTable();
        _updateKfKpis();
        _addKfTotalRow();
        if (_kfInvoiceRows.length > 0) _startKfMarginFetch(false, seq);
    } catch (err) {
        if (seq === _kfLoadSeq) kfSetStatus('Fejl: ' + err.message);
    } finally {
        if (btn && seq === _kfLoadSeq) { btn.disabled = false; btn.textContent = _kfAllCustomers ? 'Hent måned' : 'Hent fakturaer'; }
    }
}

function kfSetStatus(txt) {
    const el = document.getElementById('kfStatus');
    if (el) el.textContent = txt;
}

function _fmtKfDate(intDate) {
    const s = String(intDate || '');
    if (s.length !== 8) return s;
    return s.slice(6,8) + '/' + s.slice(4,6) + '-' + s.slice(0,4);
}

function _fmtKfDateFromInt(intDate) {
    const s = String(intDate || '');
    if (s.length !== 8) return s;
    return s.slice(6,8) + '.' + s.slice(4,6) + '.' + s.slice(0,4);
}

function _kfMarginClass(pct) {
    if (!Number.isFinite(pct)) return 'na';
    if (pct > 20) return 'ok';
    if (pct >= 5)  return 'warn';
    return 'bad';
}

function _kfIsoDateFromInt(intDate) {
    const s = String(intDate || '');
    return s.length === 8 ? s.slice(0,4) + '-' + s.slice(4,6) + '-' + s.slice(6,8) : null;
}

function _kfCalcMarginPct(rev, cost) {
    if (rev == null || cost == null) return null;
    const revenue = Number(rev);
    const orderCost = Number(cost);
    return Number.isFinite(revenue) && Number.isFinite(orderCost) && revenue !== 0
        ? ((revenue - orderCost) / revenue) * 100
        : null;
}

async function _fetchKfVareforbrug(month, seq = _kfLoadSeq) {
    const request=++_kfVareforbrugRequest;
    _kfVareforbrugMonth = month;
    try {
        const r = await fetch('/efterkalk/vareforbrug?month=' + encodeURIComponent(month) + '&plateMode=' + encodeURIComponent(_kfPlateMode));
        const d = await r.json();
        if (seq !== _kfLoadSeq || request!==_kfVareforbrugRequest) return;
        if (d.ok && d.month !== month) throw new Error('Forkert måned i svar: ' + d.month);
        _kfVareforbrugGL = (d.ok && d.vareforbrug !== null && d.vareforbrug !== undefined) ? Number(d.vareforbrug) : null;
        _kfPurchaseGap = (d.ok && d.purchaseGap) ? d.purchaseGap : null;
        _kfVarelagerDelta = (d.ok && d.varelager) ? d.varelager : null;
    } catch { if (seq !== _kfLoadSeq || request!==_kfVareforbrugRequest) return; _kfVareforbrugGL = null; _kfPurchaseGap = null; _kfVarelagerDelta = null; }
    _updateKfKpis();
}

function setKfPlateMode(mode) {
    _kfPlateMode = mode === 'standard' ? 'standard' : 'fifo';
    const selection=document.getElementById('kfSummaryPlateMode');
    if (selection) selection.value=_kfPlateMode;
    if (_kfLoadedMonth) _fetchKfVareforbrug(_kfLoadedMonth);
}

function _resetKfKpis() {
    _kfSummaryInvoicesReady=false;
    _kfUbLoaded=false;
    _kfUbRows=[];
    const ids = ['kfKpis','kfExportActions','kfCostModeHint','kfCustomerSummary','kfCustomerTrend','kfTableWrap','kfMarginProgress','kfUbSection'];
    ids.forEach(id => { const el = document.getElementById(id); if(el) el.style.display='none'; });
    _kfVareforbrugGL = null;
    _kfPurchaseGap = null;
    _kfVarelagerDelta = null;
    _kfVareforbrugMonth = null;
    const tile = document.getElementById('kfKpiVareforbrugTile');
    if (tile) tile.style.display = 'none';
}

function _renderKfCustomerSummary() {
    const invoiceRows = _kfBillableRows();
    const section = document.getElementById('kfCustomerSummary');
    const tbody = document.getElementById('kfCustomerSummaryBody');
    const tfoot = document.getElementById('kfCustomerSummaryFoot');
    if (!section || !tbody || !tfoot) return;
    if (!_kfAllCustomers || !invoiceRows.length) {
        section.style.display = 'none';
        tbody.innerHTML = '';
        tfoot.innerHTML = '';
        return;
    }

    const totalInvoice = invoiceRows.reduce((sum, row) => sum + Number(row.InvoAm || 0), 0);
    const groups = new Map();
    for (const row of invoiceRows) {
        const key = String(row.CustNo || row.CustomerName || row.CustomerShrt || 'ukendt');
        if (!groups.has(key)) {
            groups.set(key, {
                custNo: row.CustNo,
                name: row.CustomerName || row.CustomerShrt || 'Ukendt kunde',
                orders: 0,
                invoice: 0,
                cost: 0,
                loaded: 0,
                materialCost: 0, stangCost: 0, timeCost: 0, purchasedPartCost: 0, underleverandorCost: 0,
                breakdownLoaded: 0
            });
        }
        const group = groups.get(key);
        group.orders++;
        group.invoice += Number(row.InvoAm || 0);
        const margin = _kfMarginMap[String(row.OrdNo)];
        if (margin) {
            group.cost += _kfCost(margin);
            group.loaded++;
            if (margin.costBreakdownAvailable) {
                group.materialCost += Number(margin.materialCost || 0);
                group.stangCost += Number(margin.stangCost || 0);
                group.timeCost += Number(margin.timeCost || 0);
                group.purchasedPartCost += Number(margin.purchasedPartCost || 0);
                group.underleverandorCost += Number(margin.underleverandorCost || 0);
                group.breakdownLoaded++;
            }
        }
    }

    const rows = Array.from(groups.values()).sort((a, b) => b.invoice - a.invoice);
    const totalCost = rows.reduce((sum, group) => sum + group.cost, 0);
    const allComplete = rows.every(group => group.loaded === group.orders);
    const totalDb = allComplete ? totalInvoice - totalCost : null;
    const totalDg = allComplete ? _kfCalcMarginPct(totalInvoice, totalCost) : null;
    const pending = '<span style="color:#8a9bad" title="Beregner kost…">…</span>';

    tbody.innerHTML = rows.map(group => {
        const complete = group.loaded === group.orders;
        const share = totalInvoice !== 0 ? (group.invoice / totalInvoice) * 100 : null;
        const costPct = complete && group.invoice !== 0 ? (group.cost / group.invoice) * 100 : null;
        const db = complete ? group.invoice - group.cost : null;
        const dg = complete ? _kfCalcMarginPct(group.invoice, group.cost) : null;
        const cls = _kfMarginClass(dg);
        const customerLabel = escapeHtml(String(group.name)) + (group.custNo ? ' <span style="color:#789;font-size:11px;">(' + escapeHtml(String(group.custNo)) + ')</span>' : '');
        const trendButton = group.custNo ? '<button class="kf-trend-btn" onclick="_loadKfCustomerTrend(' + Number(group.custNo) + ')">Vis trend</button>' : '';
        const breakdownComplete = group.breakdownLoaded === group.orders;
        return '<tr data-kf-customer-row="1" data-kf-orders="' + group.orders + '" data-kf-invoice="' + group.invoice + '" data-kf-cost="' + group.cost + '" data-kf-loaded="' + group.loaded + '"' +
            ' data-kf-material="' + group.materialCost + '" data-kf-stang="' + group.stangCost + '" data-kf-time="' + group.timeCost + '"' +
            ' data-kf-parts="' + group.purchasedPartCost + '" data-kf-ulev="' + group.underleverandorCost + '" data-kf-breakdown-loaded="' + group.breakdownLoaded + '">' +
            '<td><strong><span class="kf-cust-name" title="' + escapeHtml(String(group.name)) + '">' + customerLabel + '</span></strong>' + trendButton + '</td>' +
            '<td class="r">' + group.orders + '</td>' +
            '<td class="r">' + formatNumber(group.invoice) + '</td>' +
            '<td class="r">' + (share !== null ? share.toFixed(1) + '%' : '—') + '</td>' +
            '<td class="r">' + (complete ? formatNumber(group.cost) : pending) + '</td>' +
            '<td class="r">' + (costPct !== null ? costPct.toFixed(1) + '%' : pending) + '</td>' +
            '<td class="r kf-margin-cell ' + (complete ? cls : 'na') + '">' + (db !== null ? formatNumber(db) : pending) + '</td>' +
            '<td class="r kf-margin-cell ' + (complete ? cls : 'na') + '">' + (dg !== null ? dg.toFixed(1) + '%' : pending) + '</td>' +
            '<td class="kf-breakdown-cell">' + _kfBreakdownHtml(breakdownComplete, group) + '</td>' +
            '</tr>';
    }).join('');
    section.style.display = 'block';
    filterKfCustomerSummary();
}

function _kfCustomerTotalRow(label, customerCount, orders, invoice, cost, complete, monthlyInvoice, breakdown) {
    const pending = '<span style="color:#8a9bad" title="Beregner kost…">…</span>';
    const share = monthlyInvoice !== 0 ? (invoice / monthlyInvoice) * 100 : 0;
    const db = complete ? invoice - cost : null;
    const dg = complete && invoice !== 0 ? _kfCalcMarginPct(invoice, cost) : null;
    const costPct = complete && invoice !== 0 ? (cost / invoice) * 100 : null;
    const breakdownComplete = !!breakdown && orders > 0 && breakdown.breakdownLoaded === orders;
    const breakdownValues = breakdown ? {
        materialCost: breakdown.material, stangCost: breakdown.stang, timeCost: breakdown.time,
        purchasedPartCost: breakdown.parts, underleverandorCost: breakdown.ulev
    } : {};
    return '<tr class="total-row">' +
        '<td><strong>' + label + ' (' + customerCount + ' kunder)</strong></td>' +
        '<td class="r"><strong>' + orders + '</strong></td>' +
        '<td class="r"><strong>' + formatNumber(invoice) + '</strong></td>' +
        '<td class="r"><strong>' + share.toFixed(1) + '%</strong></td>' +
        '<td class="r"><strong>' + (complete ? formatNumber(cost) : pending) + '</strong></td>' +
        '<td class="r"><strong>' + (costPct !== null ? costPct.toFixed(1) + '%' : (orders ? pending : '—')) + '</strong></td>' +
        '<td class="r"><strong>' + (db !== null ? formatNumber(db) : (orders ? pending : '—')) + '</strong></td>' +
        '<td class="r"><strong>' + (dg !== null ? dg.toFixed(1) + '%' : (orders ? pending : '—')) + '</strong></td>' +
        '<td class="kf-breakdown-cell">' + (orders ? _kfBreakdownHtml(breakdownComplete, breakdownValues) : '—') + '</td>' +
        '</tr>';
}

function filterKfCustomerSummary() {
    const tbody = document.getElementById('kfCustomerSummaryBody');
    const tfoot = document.getElementById('kfCustomerSummaryFoot');
    const input = document.getElementById('kfCustomerSummaryFilter');
    const count = document.getElementById('kfCustomerSummaryCount');
    if (!tbody || !tfoot) return;
    const query = String(input ? input.value : '').trim().toLowerCase();
    const tableRows = Array.from(tbody.querySelectorAll('tr[data-kf-customer-row]'));
    const blank = () => ({ customers:0, orders:0, invoice:0, cost:0, loaded:0, material:0, stang:0, time:0, parts:0, ulev:0, breakdownLoaded:0 });
    const total = blank();
    const visible = blank();
    tableRows.forEach(row => {
        const values = {
            orders:Number(row.dataset.kfOrders || 0), invoice:Number(row.dataset.kfInvoice || 0),
            cost:Number(row.dataset.kfCost || 0), loaded:Number(row.dataset.kfLoaded || 0),
            material:Number(row.dataset.kfMaterial || 0), stang:Number(row.dataset.kfStang || 0),
            time:Number(row.dataset.kfTime || 0), parts:Number(row.dataset.kfParts || 0),
            ulev:Number(row.dataset.kfUlev || 0), breakdownLoaded:Number(row.dataset.kfBreakdownLoaded || 0)
        };
        const matches = !query || String(row.cells[0] ? row.cells[0].textContent : '').toLowerCase().includes(query);
        row.style.display = matches ? '' : 'none';
        for (const key of Object.keys(values)) total[key] += values[key];
        total.customers++;
        if (matches) { for (const key of Object.keys(values)) visible[key] += values[key]; visible.customers++; }
    });
    const visibleComplete = visible.loaded === visible.orders;
    const totalComplete = total.loaded === total.orders;
    const visibleRow = _kfCustomerTotalRow('SYNLIG TOTAL', visible.customers, visible.orders, visible.invoice, visible.cost, visibleComplete, total.invoice, visible);
    const monthlyRow = _kfCustomerTotalRow('MÅNED TOTAL', total.customers, total.orders, total.invoice, total.cost, totalComplete, total.invoice, total);
    tfoot.innerHTML = query ? visibleRow + monthlyRow : monthlyRow;
    if (count) count.textContent = query ? visible.customers + ' af ' + total.customers + ' kunder' : total.customers + ' kunder';
}

function closeKfCustomerTrend() {
    _kfTrendRequestId++;
    const panel = document.getElementById('kfCustomerTrend');
    if (panel) panel.style.display = 'none';
}

function _kfMonthKeyFromInt(value) {
    const digits = String(value || '');
    return digits.length === 8 ? digits.slice(0, 4) + '-' + digits.slice(4, 6) : '';
}

function _kfMonthLabel(monthKey) {
    const parts = String(monthKey || '').split('-').map(Number);
    if (parts.length !== 2 || !parts[0] || !parts[1]) return monthKey;
    return new Intl.DateTimeFormat('da-DK', { month:'long', year:'numeric' }).format(new Date(parts[0], parts[1] - 1, 1));
}

async function _loadKfCustomerTrend(custNo) {
    const customerNo = Number(custNo);
    if (!Number.isInteger(customerNo) || customerNo <= 0) return;
    const selectedRow = _kfInvoiceRows.find(row => Number(row.CustNo) === customerNo);
    const customerName = selectedRow ? (selectedRow.CustomerName || selectedRow.CustomerShrt || String(customerNo)) : String(customerNo);
    const selectedMonth = _kfLoadedMonth || _kfPreviousMonthKey();
    const monthParts = selectedMonth.split('-').map(Number);
    const endDate = new Date(monthParts[0], monthParts[1], 0);
    const startDate = new Date(monthParts[0], monthParts[1] - 12, 1);
    const dateText = date => date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
    const requestId = ++_kfTrendRequestId;
    const panel = document.getElementById('kfCustomerTrend');
    const title = document.getElementById('kfCustomerTrendTitle');
    const status = document.getElementById('kfCustomerTrendStatus');
    const tbody = document.getElementById('kfCustomerTrendBody');
    if (!panel || !tbody) return;
    panel.style.display = 'block';
    if (title) title.textContent = 'DG-udvikling — ' + customerName + ' (' + customerNo + ')';
    if (status) status.textContent = 'Henter 12 måneders fakturaer…';
    tbody.innerHTML = '<tr><td colspan="7">Henter data…</td></tr>';
    panel.scrollIntoView({ behavior:'smooth', block:'nearest' });

    try {
        const url = '/efterkalk/customer-invoices?custno=' + customerNo + '&includeAllAmounts=0&from=' + dateText(startDate) + '&to=' + dateText(endDate);
        const response = await fetch(url);
        const payload = await response.json();
        if (!payload.ok) throw new Error(payload.error || 'Fejl');
        if (requestId !== _kfTrendRequestId) return;
        const invoiceRows = (payload.rows || []).filter(row => !row.IsUB);
        const costs = {};
        let done = 0;
        const batchSize = 4;
        for (let i = 0; i < invoiceRows.length; i += batchSize) {
            if (requestId !== _kfTrendRequestId) return;
            const batch = invoiceRows.slice(i, i + batchSize);
            await Promise.all(batch.map(async row => {
                const cached = _kfMarginMap[String(row.OrdNo)];
                if (cached) {
                    costs[String(row.OrdNo)] = _kfCost(cached);
                } else {
                    try {
                        const marginResponse = await fetch('/order-margin/' + row.OrdNo);
                        if (marginResponse.ok) {
                            const margin = await marginResponse.json();
                            if (margin && margin.totalCost !== null && margin.totalCost !== undefined) costs[String(row.OrdNo)] = _kfCost(margin);
                        }
                    } catch { /* keep missing costs visible */ }
                }
                done++;
            }));
            if (status) status.textContent = done + '/' + invoiceRows.length + ' ordrer beregnet…';
        }
        if (requestId !== _kfTrendRequestId) return;
        _renderKfCustomerTrendRows(selectedMonth, invoiceRows, costs);
        if (status) status.textContent = invoiceRows.length + ' fakturaordrer fordelt på de seneste 12 måneder.';
    } catch (err) {
        if (requestId !== _kfTrendRequestId) return;
        if (status) status.textContent = 'Fejl: ' + err.message;
        tbody.innerHTML = '<tr><td colspan="7">Kunne ikke hente udviklingen.</td></tr>';
    }
}

async function _loadKfStoredMonth(month, seq) {
    try {
        const response = await fetch('/efterkalk/month-snapshot?month=' + encodeURIComponent(month));
        if (!response.ok) return false;
        const payload = await response.json();
        if (seq !== _kfLoadSeq) return false;
        if (!payload.ok || !payload.found || !Array.isArray(payload.rows) || payload.rows.length === 0) return false;
        const missingFallback = payload.rows.some(row => row.CostComplete
            && (row.StyklisteFallbackCost === null || row.StyklisteFallbackCost === undefined));
        if (missingFallback) return false;
        // Vis straks den gemte værdi; den friske hentning fra bogføringen (allerede i gang) opdaterer den bagefter.
        if (payload.vareforbrugGL !== null && payload.vareforbrugGL !== undefined) _kfVareforbrugGL = Number(payload.vareforbrugGL);
        if (payload.purchaseMissingInvoiceValue !== null && payload.purchaseMissingInvoiceValue !== undefined) {
            _kfPurchaseGap = {
                missingValue: Number(payload.purchaseMissingInvoiceValue || 0),
                partialValue: Number(payload.purchasePartialInvoiceValue || 0),
                totalValue: Number(payload.purchaseMissingInvoiceValue || 0) + Number(payload.purchasePartialInvoiceValue || 0)
            };
        }
        const intFromDate = value => Number(String(value || '').slice(0, 10).replace(/-/g, '')) || 0;
        _kfInvoiceRows = (payload.rows || []).map(row => ({
            OrdNo:Number(row.OrdNo), OrdDt:intFromDate(row.OrderDate), CustNo:row.CustNo,
            LstInvDt:intFromDate(row.InvoiceDate), InvoAm:Number(row.Revenue || 0), InvoNo:row.InvoiceNo,
            CustomerName:row.CustomerName, CustomerShrt:row.CustomerName, SellerUsr:row.Seller
        }));
        _kfMarginMap = {};
        (payload.rows || []).forEach(row => {
            if (row.CostComplete && row.Cost !== null && row.Cost !== undefined) {
                const breakdownAvailable = !!row.CostBreakdownComplete;
                _kfMarginMap[String(row.OrdNo)] = {
                    totalRevenue:Number(row.Revenue || 0),
                    totalCost:Number(row.Cost || 0),
                    styklisteFallbackCost:Number(row.StyklisteFallbackCost || 0),
                    costBreakdownAvailable:breakdownAvailable,
                    materialCost:breakdownAvailable ? Number(row.MaterialCost || 0) : null,
                    stangCost:breakdownAvailable ? Number(row.StangCost || 0) : null,
                    timeCost:breakdownAvailable ? Number(row.TimeCost || 0) : null,
                    purchasedPartCost:breakdownAvailable ? Number(row.PurchasedPartCost || 0) : null,
                    underleverandorCost:breakdownAvailable ? Number(row.UnderleverandorCost || 0) : null
                };
            }
        });
        // Stored costs are reused, but invoices are always reconciled with Visma:
        // older snapshots can omit credit notes or contain out-of-date U/B flags.
        kfSetStatus('Afstemmer fakturaer og kreditnotaer med Visma…');
        return true;
    } catch {
        return false;
    }
}

function _renderKfStoredCustomerTrendRows(selectedMonth, storedRows) {
    const tbody = document.getElementById('kfCustomerTrendBody');
    if (!tbody) return;
    const byMonth = new Map((storedRows || []).map(row => {
        const date = new Date(row.PeriodStart);
        const key = date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
        return [key, row];
    }));
    const selectedParts = selectedMonth.split('-').map(Number);
    const months = [];
    for (let offset = 11; offset >= 0; offset--) {
        const date = new Date(selectedParts[0], selectedParts[1] - 1 - offset, 1);
        months.push(date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0'));
    }
    let previousDg = null;
    tbody.innerHTML = months.map(key => {
        const row = byMonth.get(key);
        const complete = !!(row && row.CostComplete && (_kfCostMode === 'registered' || row.FallbackComplete));
        const storedCost = complete
            ? Number(_kfCostMode === 'adjusted' ? row.AdjustedCost : row.Cost)
            : null;
        const dg = complete ? _kfCalcMarginPct(Number(row.Revenue || 0), storedCost) : null;
        const delta = dg !== null && previousDg !== null ? dg - previousDg : null;
        if (dg !== null) previousDg = dg;
        const deltaClass = delta === null ? '' : (delta >= 0 ? 'up' : 'down');
        const cls = _kfMarginClass(dg);
        return '<tr>' +
            '<td>' + escapeHtml(_kfMonthLabel(key)) + '</td>' +
            '<td class="r">' + (row ? Number(row.InvoiceOrderCount || 0) : 0) + '</td>' +
            '<td class="r">' + (row ? formatNumber(Number(row.Revenue || 0)) : '0,00') + '</td>' +
            '<td class="r">' + (complete ? formatNumber(storedCost) : '—') + '</td>' +
            '<td class="r kf-margin-cell ' + (complete ? cls : 'na') + '">' + (complete ? formatNumber(Number(row.Revenue || 0) - storedCost) : '—') + '</td>' +
            '<td class="r kf-margin-cell ' + (dg !== null ? cls : 'na') + '">' + (dg !== null ? dg.toFixed(1) + '%' : '—') + '</td>' +
            '<td class="r kf-trend-delta ' + deltaClass + '">' + (delta === null ? '—' : (delta >= 0 ? '+' : '') + delta.toFixed(1) + ' pp') + '</td>' +
            '</tr>';
    }).join('');
}

function _renderKfCustomerTrendRows(selectedMonth, invoiceRows, costs) {
    const tbody = document.getElementById('kfCustomerTrendBody');
    if (!tbody) return;
    const selectedParts = selectedMonth.split('-').map(Number);
    const months = [];
    for (let offset = 11; offset >= 0; offset--) {
        const date = new Date(selectedParts[0], selectedParts[1] - 1 - offset, 1);
        months.push(date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0'));
    }
    const summary = new Map(months.map(key => [key, { key, orders:0, invoice:0, cost:0, loaded:0 }]));
    for (const row of invoiceRows) {
        const group = summary.get(_kfMonthKeyFromInt(row.LstInvDt));
        if (!group) continue;
        group.orders++;
        group.invoice += Number(row.InvoAm || 0);
        if (Object.prototype.hasOwnProperty.call(costs, String(row.OrdNo))) {
            group.cost += Number(costs[String(row.OrdNo)] || 0);
            group.loaded++;
        }
    }
    let previousDg = null;
    tbody.innerHTML = months.map(key => {
        const month = summary.get(key);
        const complete = month.loaded === month.orders;
        const db = complete ? month.invoice - month.cost : null;
        const dg = complete && month.invoice !== 0 ? _kfCalcMarginPct(month.invoice, month.cost) : null;
        const delta = dg !== null && previousDg !== null ? dg - previousDg : null;
        if (dg !== null) previousDg = dg;
        const deltaClass = delta === null ? '' : (delta >= 0 ? 'up' : 'down');
        const deltaText = delta === null ? '—' : (delta >= 0 ? '+' : '') + delta.toFixed(1) + ' pp';
        const cls = _kfMarginClass(dg);
        return '<tr>' +
            '<td>' + escapeHtml(_kfMonthLabel(key)) + '</td>' +
            '<td class="r">' + month.orders + '</td>' +
            '<td class="r">' + formatNumber(month.invoice) + '</td>' +
            '<td class="r">' + (complete ? formatNumber(month.cost) : '—') + '</td>' +
            '<td class="r kf-margin-cell ' + (db !== null ? cls : 'na') + '">' + (db !== null ? formatNumber(db) : '—') + '</td>' +
            '<td class="r kf-margin-cell ' + (dg !== null ? cls : 'na') + '">' + (dg !== null ? dg.toFixed(1) + '%' : '—') + '</td>' +
            '<td class="r kf-trend-delta ' + deltaClass + '">' + deltaText + '</td>' +
            '</tr>';
    }).join('');
}

function _renderKfTable() {
    const invoiceRows = _kfBillableRows();
    const wrap = document.getElementById('kfTableWrap');
    const tbody = document.getElementById('kfTableBody');
    if (!wrap || !tbody) return;
    if (!invoiceRows.length) { wrap.style.display = 'none'; _renderKfCustomerSummary(); return; }

    tbody.innerHTML = invoiceRows.map(row => {
        const m = _kfMarginMap[String(row.OrdNo)];
        const cost     = _kfCost(m);
        const rev      = Number(row.InvoAm || 0);
        const margDkk  = (cost !== null && rev !== null) ? (rev - cost) : null;
        const margPct  = _kfCalcMarginPct(rev, cost);
        const cls      = _kfMarginClass(margPct);
        const pending  = '<span style="color:#aaa">…</span>';
        return '<tr>' +
            '<td class="ordno-link" onclick="searchOrderByNo(' + row.OrdNo + ')">' + row.OrdNo + '</td>' +
            '<td>' + _fmtKfDateFromInt(row.LstInvDt) + '</td>' +
            '<td>' + escapeHtml(String(row.InvoNo || '')) + (rev < 0 ? '<span class="kf-credit-label">Kreditnota</span>' : '') + '</td>' +
            '<td class="kf-cust-name" title="' + escapeHtml(String(row.CustomerName || row.CustomerShrt || '—')) + '">' + escapeHtml(String(row.CustomerName || row.CustomerShrt || '—')) + '</td>' +
            '<td>' + escapeHtml(String(row.SellerUsr || '—')) + '</td>' +
            '<td class="r">' + formatNumber(row.InvoAm) + '</td>' +
            '<td class="r kf-margin-cell ' + (cost !== null ? cls : 'na') + '" id="kf-cost-' + row.OrdNo + '">' + (cost !== null ? formatNumber(cost) : pending) + '</td>' +
            '<td class="r kf-margin-cell ' + (margDkk !== null ? cls : 'na') + '" id="kf-mdkk-' + row.OrdNo + '">' + (margDkk !== null ? formatNumber(margDkk) : pending) + '</td>' +
            '<td class="r kf-margin-cell ' + cls + '" id="kf-mpct-' + row.OrdNo + '">' + (margPct !== null ? margPct.toFixed(1) + '%' : pending) + '</td>' +
            '<td class="kf-breakdown-cell" id="kf-breakdown-' + row.OrdNo + '">' + _kfBreakdownHtml(m && m.costBreakdownAvailable, m || {}) + '</td>' +
            '</tr>';
    }).join('');
    wrap.style.display = 'block';
    _renderKfCustomerSummary();
}

function _updateKfKpis() {
    const invoiceRows = _kfBillableRows();
    const kpis = document.getElementById('kfKpis');
    if (!kpis) return;
    kpis.style.display = 'grid';

    const totalInvo = invoiceRows.reduce((s, r) => s + Number(r.InvoAm || 0), 0);
    let totalCost = 0; let totalRev = 0; let countWithMargin = 0;
    for (const row of invoiceRows) {
        const m = _kfMarginMap[String(row.OrdNo)];
        if (m) { totalCost += _kfCost(m); totalRev += Number(row.InvoAm || 0); countWithMargin++; }
    }
    const hasCost  = countWithMargin > 0;
    const margDkk  = hasCost ? (totalRev - totalCost) : null;
    const margPct  = hasCost ? _kfCalcMarginPct(totalRev, totalCost) : null;

    const setText = (id, v) => { const el = document.getElementById(id); if(el) el.textContent = v; };
    setText('kfKpiOrders',  invoiceRows.length);
    setText('kfKpiTotal',   formatNumber(totalInvo));
    const credits = invoiceRows.filter(row => Number(row.InvoAm) < 0);
    setText('kfKpiTotalSub', credits.length ? 'DKK · inkl. ' + credits.length + ' kreditnotaer: ' + formatNumber(credits.reduce((sum, row) => sum + Number(row.InvoAm), 0)) + ' DKK' : 'DKK');
    setText('kfKpiCost',    hasCost ? formatNumber(totalCost) : '—');
    setText('kfKpiCostSub', hasCost ? countWithMargin + '/' + invoiceRows.length + ' ordrer' : 'beregner…');
    setText('kfKpiMargin',  margPct !== null ? margPct.toFixed(1) + '%' : '—');
    setText('kfKpiMarginSub', margDkk !== null ? formatNumber(margDkk) + ' DKK' : 'beregner…');
    const exportActions = document.getElementById('kfExportActions');
    const exportCustomers = document.getElementById('kfExportCustomersBtn');
    const exportOrders = document.getElementById('kfExportOrdersBtn');
    const exportReady = invoiceRows.length > 0 && countWithMargin === invoiceRows.length;
    if (exportActions) exportActions.style.display = _kfInvoiceRows.length ? 'flex' : 'none';
    const hint = document.getElementById('kfCostModeHint');
    if (hint) hint.style.display = _kfInvoiceRows.length ? 'block' : 'none';
    _syncKfCostModeUi();
    if (exportCustomers) {
        exportCustomers.style.display = _kfAllCustomers ? '' : 'none';
        exportCustomers.disabled = !exportReady;
    }
    if (exportOrders) exportOrders.disabled = !exportReady;
    const refreshCosts = document.getElementById('kfRefreshCostsBtn');
    if (refreshCosts) refreshCosts.disabled = !_kfInvoiceRows.length || _kfMarginFetchRunning;
    const vareforbrugTile = document.getElementById('kfKpiVareforbrugTile');
    if (vareforbrugTile) {
        if (_kfAllCustomers && invoiceRows.length) {
            vareforbrugTile.style.display = '';
            let materialSum = 0, breakdownN = 0;
            for (const row of invoiceRows) {
                const m = _kfMarginMap[String(row.OrdNo)];
                if (m && m.costBreakdownAvailable) {
                    // Tid (løn/tid-kost) holdes udenfor: det er ikke en del af Vareforbrug, det hører til lønomkostningerne.
                    materialSum += Number(m.materialCost || 0) + Number(m.stangCost || 0) + Number(m.purchasedPartCost || 0) + Number(m.underleverandorCost || 0);
                    breakdownN++;
                }
            }
            const breakdownComplete = breakdownN === invoiceRows.length;
            const rowsBlock = document.getElementById('kfVfRows');
            const explainBlock = document.getElementById('kfVfExplain');
            const gapDetailRow = document.getElementById('kfVfGapDetailRow');
            const residualRow = document.getElementById('kfVfResidualRow');
            if (_kfVareforbrugGL === null) {
                if (rowsBlock) rowsBlock.style.display = 'none';
                if (explainBlock) explainBlock.style.display = 'none';
                setText('kfKpiVareforbrugSub', 'henter Vareforbrug fra bogføringen…');
            } else if (!breakdownComplete) {
                if (rowsBlock) rowsBlock.style.display = 'none';
                if (explainBlock) explainBlock.style.display = 'none';
                setText('kfKpiVareforbrugSub', 'beregner materialesum for ordrerne…');
            } else {
                const delta = _kfVareforbrugGL - materialSum;
                setText('kfKpiVareforbrugSub', '');
                if (rowsBlock) rowsBlock.style.display = 'flex';
                if (explainBlock) explainBlock.style.display = 'block';
                setText('kfVfBooked', formatNumber(_kfVareforbrugGL) + ' DKK');
                setText('kfVfCalc', formatNumber(materialSum) + ' DKK');
                setText('kfVfDelta', (delta >= 0 ? '+' : '') + formatNumber(delta) + ' DKK');

                const gapTotal = _kfPurchaseGap ? Number(_kfPurchaseGap.totalValue || 0) : 0;
                setText('kfVfGapTotal', _kfPurchaseGap ? (formatNumber(gapTotal) + ' DKK') : 'henter…');
                if (gapDetailRow) {
                    if (_kfPurchaseGap && gapTotal > 0) {
                        gapDetailRow.style.display = 'flex';
                        setText('kfVfGapDetail', 'mangler ' + formatNumber(_kfPurchaseGap.missingValue) + ' (' + _kfPurchaseGap.missingOrderCount + ' ordrer) · delvis ' + formatNumber(_kfPurchaseGap.partialValue) + ' (' + _kfPurchaseGap.partialOrderCount + ' ordrer)');
                    } else {
                        gapDetailRow.style.display = 'none';
                    }
                }

                const varelagerReady = _kfVarelagerDelta && _kfVarelagerDelta.delta !== null && _kfVarelagerDelta.delta !== undefined;
                setText('kfVfVarelagerMonths', _kfVarelagerDelta ? ('(' + _kfVarelagerDelta.previousMonth + ' → ' + _kfVarelagerDelta.currentMonth + ')') : '');
                if (!_kfVarelagerDelta) {
                    setText('kfVfVarelager', 'henter…');
                } else if (!varelagerReady) {
                    setText('kfVfVarelager', 'ikke tilgængelig (mangler lagerlukning for ' + _kfVarelagerDelta.previousMonth + ')');
                } else {
                    const vd = Number(_kfVarelagerDelta.delta);
                    setText('kfVfVarelager', (vd >= 0 ? '+' : '') + formatNumber(vd) + ' DKK');
                }

                // Uforklaret rest = Forskel - ΔVarelager + Gap: hvad regnestykket ikke kan
                // forklare med lagerændring og udestående indkøbsfakturaer alene. Bogført er
                // allerede kun konto 12070 "Varekøb", så der er intet ikke-indkøb at trække fra.
                if (residualRow) {
                    if (_kfPurchaseGap && varelagerReady) {
                        residualRow.style.display = 'flex';
                        const residual = delta - Number(_kfVarelagerDelta.delta) + gapTotal;
                        setText('kfVfResidual', (residual >= 0 ? '+' : '') + formatNumber(residual) + ' DKK');
                    } else {
                        residualRow.style.display = 'none';
                    }
                }
            }
        } else {
            vareforbrugTile.style.display = 'none';
        }
    }
    _renderKfCustomerSummary();
    _refreshKfConfigurableSummary();
}

function _kfCsvText(value) {
    let text = String(value == null ? '' : value);
    if (/^[=+@-]/.test(text)) text = "'" + text;
    return '"' + text.replace(/"/g, '""') + '"';
}

function _kfCsvNumber(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number.toFixed(2).replace('.', ',') : '';
}

function _kfCsvPercent(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number.toFixed(2).replace('.', ',') + '%' : '';
}

function _downloadKfCsv(lines, suffix) {
    const month = (_kfLoadedMonth || _kfPreviousMonthKey()).replace(/[^0-9-]/g, '');
    const blob = new Blob(['\uFEFFsep=;\r\n' + lines.join('\r\n')], { type:'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'maanedens-db_' + month + '_' + suffix + '.csv';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function exportKfCustomersCsv() {
    const invoiceRows = _kfBillableRows();
    if (!_kfAllCustomers || !invoiceRows.length) return;
    const totalInvoice = invoiceRows.reduce((sum, row) => sum + Number(row.InvoAm || 0), 0);
    const groups = new Map();
    for (const row of invoiceRows) {
        const key = String(row.CustNo || row.CustomerName || row.CustomerShrt || 'ukendt');
        if (!groups.has(key)) groups.set(key, { custNo:row.CustNo || '', name:row.CustomerName || row.CustomerShrt || 'Ukendt kunde', orders:0, invoice:0, registeredCost:0, fallbackCost:0, materialCost:0, stangCost:0, timeCost:0, purchasedPartCost:0, underleverandorCost:0 });
        const group = groups.get(key);
        group.orders++;
        group.invoice += Number(row.InvoAm || 0);
        const margin = _kfMarginMap[String(row.OrdNo)];
        if (margin) {
            group.registeredCost += Number(margin.totalCost || 0);
            group.fallbackCost += Number(margin.styklisteFallbackCost || 0);
            if (margin.costBreakdownAvailable) {
                group.materialCost += Number(margin.materialCost || 0);
                group.stangCost += Number(margin.stangCost || 0);
                group.timeCost += Number(margin.timeCost || 0);
                group.purchasedPartCost += Number(margin.purchasedPartCost || 0);
                group.underleverandorCost += Number(margin.underleverandorCost || 0);
            }
        }
    }
    const lines = ['Kundenr;Kunde;Ordrer;Faktureret DKK;Andel af månedsomsætning %;Registreret kost DKK;Registreret DB DKK;Registreret DG %;Stykliste fallback DKK;Korrigeret kost DKK;Korrigeret DB DKK;Korrigeret DG %;Materiale DKK;Stang DKK;Tid DKK;Indkøbte dele DKK;Underleverandør DKK'];
    Array.from(groups.values()).sort((a, b) => b.invoice - a.invoice).forEach(group => {
        const registeredDb = group.invoice - group.registeredCost;
        const adjustedCost = group.registeredCost + group.fallbackCost;
        const adjustedDb = group.invoice - adjustedCost;
        lines.push([
            _kfCsvText(group.custNo), _kfCsvText(group.name), group.orders,
            _kfCsvNumber(group.invoice), _kfCsvPercent(totalInvoice ? group.invoice / totalInvoice * 100 : 0),
            _kfCsvNumber(group.registeredCost), _kfCsvNumber(registeredDb), _kfCsvPercent(group.invoice ? registeredDb / group.invoice * 100 : 0),
            _kfCsvNumber(group.fallbackCost), _kfCsvNumber(adjustedCost), _kfCsvNumber(adjustedDb), _kfCsvPercent(group.invoice ? adjustedDb / group.invoice * 100 : 0),
            _kfCsvNumber(group.materialCost), _kfCsvNumber(group.stangCost), _kfCsvNumber(group.timeCost), _kfCsvNumber(group.purchasedPartCost), _kfCsvNumber(group.underleverandorCost)
        ].join(';'));
    });
    _downloadKfCsv(lines, 'kunder');
}

function exportKfOrdersCsv() {
    const invoiceRows = _kfBillableRows();
    if (!invoiceRows.length) return;
    const lines = ['Ordrenr;Fakturadato;Fakturanr;Kundenr;Kunde;Sælger;Faktureret DKK;Registreret kost DKK;Registreret DB DKK;Registreret DG %;Stykliste fallback DKK;Korrigeret kost DKK;Korrigeret DB DKK;Korrigeret DG %;Materiale DKK;Stang DKK;Tid DKK;Indkøbte dele DKK;Underleverandør DKK'];
    for (const row of invoiceRows) {
        const margin = _kfMarginMap[String(row.OrdNo)];
        const invoice = Number(row.InvoAm || 0);
        const registeredCost = margin ? Number(margin.totalCost || 0) : 0;
        const fallbackCost = margin ? Number(margin.styklisteFallbackCost || 0) : 0;
        const registeredDb = invoice - registeredCost;
        const adjustedCost = registeredCost + fallbackCost;
        const adjustedDb = invoice - adjustedCost;
        const breakdownAvailable = !!(margin && margin.costBreakdownAvailable);
        lines.push([
            _kfCsvText(row.OrdNo), _kfCsvText(_fmtKfDateFromInt(row.LstInvDt)), _kfCsvText(row.InvoNo),
            _kfCsvText(row.CustNo), _kfCsvText(row.CustomerName || row.CustomerShrt || ''), _kfCsvText(row.SellerUsr || ''),
            _kfCsvNumber(invoice), _kfCsvNumber(registeredCost), _kfCsvNumber(registeredDb), _kfCsvPercent(invoice ? registeredDb / invoice * 100 : 0),
            _kfCsvNumber(fallbackCost), _kfCsvNumber(adjustedCost), _kfCsvNumber(adjustedDb), _kfCsvPercent(invoice ? adjustedDb / invoice * 100 : 0),
            _kfCsvNumber(breakdownAvailable ? margin.materialCost : ''), _kfCsvNumber(breakdownAvailable ? margin.stangCost : ''),
            _kfCsvNumber(breakdownAvailable ? margin.timeCost : ''), _kfCsvNumber(breakdownAvailable ? margin.purchasedPartCost : ''),
            _kfCsvNumber(breakdownAvailable ? margin.underleverandorCost : '')
        ].join(';'));
    }
    _downloadKfCsv(lines, 'ordrer');
}

async function _startKfMarginFetch(forceRefresh = false, seq = _kfLoadSeq) {
    const progress  = document.getElementById('kfMarginProgress');
    const fill      = document.getElementById('kfMarginProgressFill');
    const pctTxt    = document.getElementById('kfMarginProgressPct');
    const progTxt   = document.getElementById('kfMarginProgressText');
    if (progress) progress.style.display = 'flex';
    const rows = _kfInvoiceRows;
    const total = rows.length;
    let done = 0;
    const isStale = () => seq !== _kfLoadSeq || _kfMarginAbort;
    _kfMarginFetchRunning = true;

    const updateProgress = () => {
        const pct = total > 0 ? Math.round((done / total) * 100) : 100;
        if (fill)   fill.style.width = pct + '%';
        if (pctTxt) pctTxt.textContent = pct + '%';
        if (progTxt) progTxt.textContent = done + '/' + total + ' marginer hentet';
        if (done >= total && progress) progress.style.display = 'none';
    };

    // Process in batches of 3 (same as existing margin queue)
    const BATCH = 3;
    for (let i = 0; i < rows.length; i += BATCH) {
        if (isStale()) break;
        const batch = rows.slice(i, i + BATCH);
        await Promise.all(batch.map(async row => {
            if (isStale()) return;
            try {
                const cached = _kfMarginMap[String(row.OrdNo)];
                if (!forceRefresh && cached && cached.costBreakdownAvailable) {
                    _updateKfRowMargin(row.OrdNo, cached);
                    return;
                }
                let marginData = null;
                const r = await fetch('/order-margin/' + row.OrdNo + '?breakdown=1' + (forceRefresh ? '&force=1' : ''));
                if (r.ok) {
                    const d = await r.json();
                    if (d && d.totalCost !== null && d.totalCost !== undefined) {
                        marginData = {
                            totalRevenue: Number(row.InvoAm || 0),
                            totalCost: Number(d.totalCost || 0),
                            styklisteFallbackCost: Number(d.styklisteFallbackCost || 0),
                            costBreakdownAvailable: Boolean(d.costBreakdownAvailable),
                            materialCost: d.costBreakdownAvailable ? Number(d.materialCost || 0) : null,
                            stangCost: d.costBreakdownAvailable ? Number(d.stangCost || 0) : null,
                            timeCost: d.costBreakdownAvailable ? Number(d.timeCost || 0) : null,
                            purchasedPartCost: d.costBreakdownAvailable ? Number(d.purchasedPartCost || 0) : null,
                            underleverandorCost: d.costBreakdownAvailable ? Number(d.underleverandorCost || 0) : null
                        };
                    }
                }
                if (marginData && !isStale()) {
                    _kfMarginMap[String(row.OrdNo)] = marginData;
                    _updateKfRowMargin(row.OrdNo, marginData);
                }
            } catch { /* silent */ }
            finally { done++; }
            if (isStale()) return;
            updateProgress();
            _updateKfKpis();
        }));
    }
    if (isStale()) {
        if (seq === _kfLoadSeq) _kfMarginFetchRunning = false;
        return;
    }
    updateProgress();
    _updateKfKpis();
    _addKfTotalRow();
    _kfMarginFetchRunning = false;
    if (_kfAllCustomers) await _saveKfMonthSnapshot();
}

async function refreshKfCostBreakdown() {
    if (_kfMarginFetchRunning || !_kfInvoiceRows.length) return;
    const btn = document.getElementById('kfRefreshCostsBtn');
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Opdaterer…'; }
    kfSetStatus('Opdaterer omkostninger fra Visma…');
    try {
        await _startKfMarginFetch(true);
        kfSetStatus('Omkostninger opdateret.');
    } catch (err) {
        kfSetStatus('Fejl ved opdatering: ' + err.message);
    } finally {
        if (btn) { btn.disabled = false; btn.textContent = '🔄 Opdater omkostninger'; }
        _updateKfKpis();
    }
}

async function _saveKfMonthSnapshot() {
    // Måneden kommer fra de indlæste data, ikke fra input-feltet: feltet kan være ændret
    // siden hentningen startede, og så ville dataene gemmes under forkert måned i GOHCache.
    const month = _kfLoadedMonth || '';
    if (!/^\d{4}-\d{2}$/.test(month) || !_kfInvoiceRows.length) return;
    const parts = month.split('-').map(Number);
    const periodEnd = month + '-' + String(new Date(parts[0], parts[1], 0).getDate()).padStart(2, '0');
    const rows = _kfInvoiceRows.map(row => {
        const margin = _kfMarginMap[String(row.OrdNo)];
        return {
            OrdNo:Number(row.OrdNo), OrderDate:_kfIsoDateFromInt(row.OrdDt),
            InvoNo:String(row.InvoNo || ''), InvoiceDate:_kfIsoDateFromInt(row.LstInvDt),
            CustNo:String(row.CustNo || ''), CustomerName:row.CustomerName || row.CustomerShrt || '',
            SellerUsr:row.SellerUsr || '', InvoAm:Number(row.InvoAm || 0),
            Cost:margin ? Number(margin.totalCost || 0) : null,
            StyklisteFallbackCost:margin ? Number(margin.styklisteFallbackCost || 0) : null,
            CostComplete:!!margin,
            MaterialCost:margin && margin.costBreakdownAvailable ? Number(margin.materialCost || 0) : null,
            StangCost:margin && margin.costBreakdownAvailable ? Number(margin.stangCost || 0) : null,
            TimeCost:margin && margin.costBreakdownAvailable ? Number(margin.timeCost || 0) : null,
            PurchasedPartCost:margin && margin.costBreakdownAvailable ? Number(margin.purchasedPartCost || 0) : null,
            UnderleverandorCost:margin && margin.costBreakdownAvailable ? Number(margin.underleverandorCost || 0) : null,
            CostBreakdownComplete:!!(margin && margin.costBreakdownAvailable)
        };
    });
    try {
        const response = await fetch('/efterkalk/month-snapshot', {
            method:'POST', headers:{ 'Content-Type':'application/json' },
            body:JSON.stringify({ periodStart:month + '-01', periodEnd, rows, vareforbrugGL:_kfVareforbrugGL,
                purchaseMissingInvoiceValue: _kfPurchaseGap ? _kfPurchaseGap.missingValue : null,
                purchasePartialInvoiceValue: _kfPurchaseGap ? _kfPurchaseGap.partialValue : null })
        });
        const result = await response.json();
        if (response.ok && result.ok) {
            kfSetStatus((document.getElementById('kfStatus').textContent || '') + ' · Gemt i GOHCache');
            _startKfSnapshotBackfill();
        }
    } catch { /* GOHCache er fail-soft; visningen forbliver brugbar */ }
}

function _startKfSnapshotBackfill() {
    if (_kfBackfillRequested) return;
    _kfBackfillRequested = true;
    fetch('/efterkalk/snapshot-backfill', { method:'POST' }).catch(() => {
        _kfBackfillRequested = false;
    });
}

function _updateKfRowMargin(ordNo, m) {
    const cost    = _kfCost(m);
    const row     = _kfInvoiceRows.find(item => Number(item.OrdNo) === Number(ordNo));
    const rev     = Number((row && row.InvoAm) || 0);
    const margDkk = rev - cost;
    const margPct = _kfCalcMarginPct(rev, cost);
    const cls     = _kfMarginClass(margPct);
    const setCellClass = (id, txt, c) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.textContent = txt;
        el.className = 'r kf-margin-cell ' + c;
    };
    setCellClass('kf-cost-' + ordNo, formatNumber(cost), cls);
    setCellClass('kf-mdkk-' + ordNo, formatNumber(margDkk), cls);
    setCellClass('kf-mpct-' + ordNo, margPct !== null ? margPct.toFixed(1) + '%' : '—', cls);
    const breakdownCell = document.getElementById('kf-breakdown-' + ordNo);
    if (breakdownCell) breakdownCell.innerHTML = _kfBreakdownHtml(m.costBreakdownAvailable, m);
}

function _addKfTotalRow() {
    const invoiceRows = _kfBillableRows();
    const tbody = document.getElementById('kfTableBody');
    if (!tbody || !invoiceRows.length) return;
    // Remove existing total row
    const existing = tbody.querySelector('.total-row');
    if (existing) existing.remove();

    const totalInvo = invoiceRows.reduce((s, r) => s + Number(r.InvoAm || 0), 0);
    let totalCost = 0; let totalRev = 0; let n = 0;
    let totalMaterial = 0, totalStang = 0, totalTime = 0, totalParts = 0, totalUlev = 0, breakdownN = 0;
    for (const row of invoiceRows) {
        const m = _kfMarginMap[String(row.OrdNo)];
        if (m) {
            totalCost += _kfCost(m); totalRev += Number(row.InvoAm || 0); n++;
            if (m.costBreakdownAvailable) {
                totalMaterial += Number(m.materialCost || 0); totalStang += Number(m.stangCost || 0);
                totalTime += Number(m.timeCost || 0); totalParts += Number(m.purchasedPartCost || 0);
                totalUlev += Number(m.underleverandorCost || 0); breakdownN++;
            }
        }
    }
    const margDkk = n > 0 ? (totalRev - totalCost) : null;
    const margPct = n > 0 ? _kfCalcMarginPct(totalRev, totalCost) : null;
    const cls     = _kfMarginClass(margPct);
    const breakdownComplete = invoiceRows.length > 0 && breakdownN === invoiceRows.length;
    const breakdownValues = { materialCost: totalMaterial, stangCost: totalStang, timeCost: totalTime, purchasedPartCost: totalParts, underleverandorCost: totalUlev };

    const tr = document.createElement('tr');
    tr.className = 'total-row';
    tr.innerHTML =
        '<td colspan="5"><strong>TOTAL (' + invoiceRows.length + ' ordrer)</strong></td>' +
        '<td class="r"><strong>' + formatNumber(totalInvo) + '</strong></td>' +
        '<td class="r kf-margin-cell ' + (n > 0 ? cls : 'na') + '"><strong>' + (n > 0 ? formatNumber(totalCost) : '—') + '</strong></td>' +
        '<td class="r kf-margin-cell ' + (margDkk !== null ? cls : 'na') + '"><strong>' + (margDkk !== null ? formatNumber(margDkk) : '—') + '</strong></td>' +
        '<td class="r kf-margin-cell ' + (margPct !== null ? cls : 'na') + '"><strong>' + (margPct !== null ? margPct.toFixed(1) + '%' : '—') + '</strong></td>' +
        '<td class="kf-breakdown-cell"><strong>' + _kfBreakdownHtml(breakdownComplete, breakdownValues) + '</strong></td>';
    tbody.appendChild(tr);
}
