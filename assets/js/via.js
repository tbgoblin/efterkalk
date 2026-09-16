// ── SalgOrdre VIA · client ──────────────────────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, formatNumber, authToken, openModule, selectOrder,
// orderDetailReturnModule (dichiarato nell'inline script).
let salgordreViaRows = [];
let salgordreViaLoadState = 'idle';
let salgordreViaReservations = [];
let salgordreViaOrderBacklogValue = null;
let salgordreViaMeta = {};
let salgordreViaContextKey = '';
let salgordreViaRequestVersion = 0;
let salgordreViaPending = null;
let salgordreViaMessage = '';
let salgordreViaReservationsVersion = 0;
let salgordreViaSortField = 'deliveryDate';
let salgordreViaSortDirection = 'asc';
let salgordreViaColumnWidths = {};
try {
    salgordreViaColumnWidths = JSON.parse(localStorage.getItem('salgordreViaColumnWidths') || '{}') || {};
} catch (_) { salgordreViaColumnWidths = {}; }

function applySalgordreViaColumnWidths() {
    const table = document.querySelector('#viaResults .order-list-table');
    if (!table) return;
    table.style.tableLayout = 'fixed';
    table.querySelectorAll('th[data-column-field]').forEach((header, index) => {
        const field = header.getAttribute('data-column-field');
        const width = Number(salgordreViaColumnWidths[field] || 0);
        if (width > 0) {
            header.style.width = width + 'px';
            table.querySelectorAll('tr').forEach(row => {
                if (row.cells[index]) row.cells[index].style.width = width + 'px';
            });
        }
    });
}

function startSalgordreViaColumnResize(event, field) {
    event.preventDefault();
    event.stopPropagation();
    const handle = event.currentTarget;
    const header = handle.closest('th');
    const table = handle.closest('table');
    if (!header || !table) return;
    const columnIndex = header.cellIndex;
    const startX = event.clientX;
    const startWidth = header.getBoundingClientRect().width;
    handle.classList.add('active');
    const onMove = moveEvent => {
        const width = Math.max(72, Math.round(startWidth + moveEvent.clientX - startX));
        salgordreViaColumnWidths[field] = width;
        table.querySelectorAll('tr').forEach(row => {
            if (row.cells[columnIndex]) row.cells[columnIndex].style.width = width + 'px';
        });
    };
    const onUp = () => {
        handle.classList.remove('active');
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        try { localStorage.setItem('salgordreViaColumnWidths', JSON.stringify(salgordreViaColumnWidths)); } catch (_) {}
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
}

function viaDateDigits(value) {
    return Array.from(String(value == null ? '' : value))
        .filter(character => character >= '0' && character <= '9')
        .join('');
}

function formatViaDate(value) {
    const digits = viaDateDigits(value);
    if (digits.length === 8) return digits.slice(6, 8) + '-' + digits.slice(4, 6) + '-' + digits.slice(0, 4);
    const iso = String(value || '').slice(0, 10);
    const parts = iso.split('-');
    return parts.length === 3 ? (parts[2] + '-' + parts[1] + '-' + parts[0]) : '-';
}

function getSalgordreViaProgress(row) {
    const completedMinutes = Number(row.CompletedResourceMinutes || 0);
    const effectiveMinutes = Number(row.EffectiveResourceMinutes || 0);
    const percentage = effectiveMinutes > 0
        ? Math.max(0, Math.min(100, Math.round((completedMinutes / effectiveMinutes) * 100)))
        : 0;
    return { completedMinutes, effectiveMinutes, percentage };
}

function setSalgordreViaSort(field) {
    if (salgordreViaSortField === field) {
        salgordreViaSortDirection = salgordreViaSortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        salgordreViaSortField = field;
        salgordreViaSortDirection = field === 'progress' ? 'desc' : 'asc';
    }
    renderSalgordreVia();
}

function setSalgordreViaSortFromElement(element) {
    setSalgordreViaSort(String(element && element.dataset && element.dataset.sortField || 'deliveryDate'));
}

function getSalgordreViaVisibleRows() {
    const query = String((document.getElementById('viaSearchInput') || {}).value || '').trim().toLowerCase();
    const rows = salgordreViaRows.filter(row => !query
        || String(row.OrdNo || '').includes(query)
        || String(row.CustomerName || '').toLowerCase().includes(query)).slice();
    rows.sort((left, right) => {
        const leftProgress = getSalgordreViaProgress(left);
        const rightProgress = getSalgordreViaProgress(right);
        let leftValue;
        let rightValue;
        if (salgordreViaSortField === 'order') {
            leftValue = Number(left.OrdNo || 0);
            rightValue = Number(right.OrdNo || 0);
        } else if (salgordreViaSortField === 'deliveryDate' || salgordreViaSortField === 'plannedDate') {
            const key = salgordreViaSortField === 'deliveryDate' ? 'DeliveryDate' : 'PlannedDate';
            leftValue = Number(viaDateDigits(left[key])) || 99991231;
            rightValue = Number(viaDateDigits(right[key])) || 99991231;
        } else if (salgordreViaSortField === 'progress') {
            leftValue = leftProgress.percentage;
            rightValue = rightProgress.percentage;
        } else if (salgordreViaSortField === 'remainingSalesValue') {
            leftValue = Number(left.RemainingSalesValue || 0);
            rightValue = Number(right.RemainingSalesValue || 0);
        } else if (salgordreViaSortField === 'materialCost' || salgordreViaSortField === 'stangCost' || salgordreViaSortField === 'purchasedPartCost' || salgordreViaSortField === 'timeCost' || salgordreViaSortField === 'totalCost') {
            const leftMaterialCost = Number(left.MaterialCost || 0);
            const rightMaterialCost = Number(right.MaterialCost || 0);
            const leftStangCost = Number(left.StangCost || 0);
            const rightStangCost = Number(right.StangCost || 0);
            const leftPurchasedPartCost = Number(left.PurchasedPartCost || 0);
            const rightPurchasedPartCost = Number(right.PurchasedPartCost || 0);
            const leftTimeCost = Number(left.TimeCost || 0);
            const rightTimeCost = Number(right.TimeCost || 0);
            leftValue = salgordreViaSortField === 'materialCost'
                ? leftMaterialCost
                : (salgordreViaSortField === 'stangCost' ? leftStangCost
                    : (salgordreViaSortField === 'purchasedPartCost' ? leftPurchasedPartCost
                        : (salgordreViaSortField === 'timeCost' ? leftTimeCost : leftMaterialCost + leftStangCost + leftPurchasedPartCost + leftTimeCost)));
            rightValue = salgordreViaSortField === 'materialCost'
                ? rightMaterialCost
                : (salgordreViaSortField === 'stangCost' ? rightStangCost
                    : (salgordreViaSortField === 'purchasedPartCost' ? rightPurchasedPartCost
                        : (salgordreViaSortField === 'timeCost' ? rightTimeCost : rightMaterialCost + rightStangCost + rightPurchasedPartCost + rightTimeCost)));
        } else {
            const key = salgordreViaSortField === 'customer' ? 'CustomerName'
                : (salgordreViaSortField === 'seller' ? 'SellerUsr' : 'ResourceName');
            leftValue = String(left[key] || '').toLocaleLowerCase('da-DK');
            rightValue = String(right[key] || '').toLocaleLowerCase('da-DK');
        }
        const comparison = leftValue < rightValue ? -1 : (leftValue > rightValue ? 1 : 0);
        return salgordreViaSortDirection === 'asc' ? comparison : -comparison;
    });
    return rows;
}

function exportSalgordreViaCsv() {
    const rows = getSalgordreViaVisibleRows();
    if (!rows.length) {
        alert('Ingen rækker at eksportere.');
        return;
    }
    const csvNumber = value => value === null ? '' : Number(value || 0).toFixed(2).replace('.', ',');
    const csvText = value => '"' + String(value == null ? '' : value).replace(/"/g, '""') + '"';
    const hasBacklog = salgordreViaMeta.scope === 'open-backlog';
    const costsComplete = rows.every(row => row.CostDataAvailable !== false);
    const lines = ['Salgsordre;Kunde;Levdato;Ansvarlig;Materialekost;Stangkost;Indkøbte dele til ordre;Tidskost;Total kost' + (hasBacklog ? ';Restsalgsværdi' : '') + ';Koststatus'];
    let sumMaterial = 0;
    let sumStang = 0;
    let sumPurchased = 0;
    let sumTime = 0;
    let sumRemaining = 0;
    for (const row of rows) {
        const materialCost = Number(row.MaterialCost || 0);
        const stangCost = Number(row.StangCost || 0);
        const purchasedPartCost = Number(row.PurchasedPartCost || 0);
        const timeCost = Number(row.TimeCost || 0);
        sumMaterial += materialCost;
        sumStang += stangCost;
        sumPurchased += purchasedPartCost;
        sumTime += timeCost;
        sumRemaining += Number(row.RemainingSalesValue || 0);
        const costNumber = value => csvNumber(row.CostDataAvailable === false ? null : value);
        lines.push([
            csvText(row.OrdNo || ''),
            csvText(row.CustomerName || ''),
            csvText(formatViaDate(row.DeliveryDate)),
            csvText(row.SellerUsr || ''),
            costNumber(materialCost),
            costNumber(stangCost),
            costNumber(purchasedPartCost),
            costNumber(timeCost),
            costNumber(materialCost + stangCost + purchasedPartCost + timeCost),
            ...(hasBacklog ? [csvNumber(row.RemainingSalesValue)] : []),
            csvText(row.CostDataAvailable === false ? 'Ikke hentet' : 'Hentet')
        ].join(';'));
    }
    const totalNumber = value => csvNumber(costsComplete ? value : null);
    lines.push(['"I alt"', '""', '""', '""', totalNumber(sumMaterial), totalNumber(sumStang), totalNumber(sumPurchased), totalNumber(sumTime), totalNumber(sumMaterial + sumStang + sumPurchased + sumTime), ...(hasBacklog ? [csvNumber(sumRemaining)] : []), csvText(costsComplete ? 'Hentet' : 'Ufuldstændige kostdata')].join(';'));
    const blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'salgordre-via_' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}

function renderSalgordreVia() {
    const target = document.getElementById('viaResults');
    const kpis = document.getElementById('viaKpis');
    if (!target) return;
    renderSalgordreViaReservations();
    const rows = getSalgordreViaVisibleRows();
    const materialCost = rows.reduce((sum, row) => sum + Number(row.MaterialCost || 0), 0);
    const stangCost = rows.reduce((sum, row) => sum + Number(row.StangCost || 0), 0);
    const purchasedPartCost = rows.reduce((sum, row) => sum + Number(row.PurchasedPartCost || 0), 0);
    const timeCost = rows.reduce((sum, row) => sum + Number(row.TimeCost || 0), 0);
    const totalCost = materialCost + stangCost + purchasedPartCost + timeCost;
    const hasBacklog = salgordreViaMeta.scope === 'open-backlog';
    const remainingSalesValue = rows.reduce((sum, row) => sum + Number(row.RemainingSalesValue || 0), 0);
    const costsComplete = rows.every(row => row.CostDataAvailable !== false);
    const costTotal = value => costsComplete ? formatNumber(value) + ' DKK' : 'Ufuldstændige kostdata';
    if (kpis) {
        kpis.innerHTML = '<div class="via-kpi"><span>Materialekost</span><strong>' + costTotal(materialCost) + '</strong></div>'
            + '<div class="via-kpi"><span>Stangkost</span><strong>' + costTotal(stangCost) + '</strong></div>'
            + '<div class="via-kpi"><span>Indkøbte dele til ordre</span><strong>' + costTotal(purchasedPartCost) + '</strong></div>'
            + '<div class="via-kpi"><span>Tidskost (færdigmeldt)</span><strong>' + costTotal(timeCost) + '</strong></div>'
            + '<div class="via-kpi"><span>Samlet kost</span><strong>' + costTotal(totalCost) + '</strong></div>'
            + (hasBacklog ? '<div class="via-kpi via-kpi-backlog"><span>Ordrebeholdning' + (rows.length !== salgordreViaRows.length ? ' (filtreret)' : '') + '</span><strong>' + formatNumber(remainingSalesValue) + ' DKK</strong></div>' : '');
    }
    updateSalgordreViaStatus();
    if (!rows.length) {
        target.innerHTML = '<div class="omsaetning-empty">Ingen aktive salgsordrer matcher søgningen.</div>';
        return;
    }
    const sortHeader = (field, label) => '<th data-sort-field="' + field + '" data-column-field="' + field + '" onclick="event.stopPropagation();setSalgordreViaSortFromElement(this)" style="cursor:pointer;user-select:none;">'
        + label + (salgordreViaSortField === field ? (salgordreViaSortDirection === 'asc' ? ' ▲' : ' ▼') : ' ↕')
        + '<span class="via-col-resizer" onmousedown="startSalgordreViaColumnResize(event, &#39;' + field + '&#39;)"></span></th>';
    const plainHeader = (field, label) => '<th data-column-field="' + field + '">' + label
        + '<span class="via-col-resizer" onmousedown="startSalgordreViaColumnResize(event, &#39;' + field + '&#39;)"></span></th>';
    let html = '<div class="order-list-section"><table class="order-list-table' + (hasBacklog ? ' via-backlog-table' : '') + '"><thead><tr>'
        + sortHeader('order', 'Salgsordre')
        + sortHeader('customer', 'Kunde')
        + sortHeader('deliveryDate', 'Levdato')
        + sortHeader('seller', 'Ansvarlig')
        + sortHeader('materialCost', 'Materiale')
        + sortHeader('stangCost', 'Stang')
        + sortHeader('purchasedPartCost', 'Indkøbte dele')
        + sortHeader('timeCost', 'Tid')
        + sortHeader('totalCost', 'Total kost')
        + (hasBacklog ? sortHeader('remainingSalesValue', 'Restsalgsværdi') : '')
        + sortHeader('progress', 'Procesfremskridt')
        + sortHeader('resource', 'Næste ressource')
        + plainHeader('refresh', 'Opdater') + '</tr></thead><tbody>';
    for (const row of rows) {
        const progress = getSalgordreViaProgress(row);
        const rowMaterialCost = Number(row.MaterialCost || 0);
        const rowStangCost = Number(row.StangCost || 0);
        const rowPurchasedPartCost = Number(row.PurchasedPartCost || 0);
        const rowTimeCost = Number(row.TimeCost || 0);
        const rowTotalCost = rowMaterialCost + rowStangCost + rowPurchasedPartCost + rowTimeCost;
        const costText = value => row.CostDataAvailable === false ? 'Ikke hentet' : formatNumber(value) + ' DKK';
        html += '<tr onclick="openSalgordreViaOrder(' + Number(row.OrdNo) + ')">'
            + '<td><strong>' + escapeHtml(String(row.OrdNo || '-')) + '</strong></td>'
            + '<td>' + escapeHtml(String(row.CustomerName || '-')) + '</td>'
            + '<td>' + escapeHtml(formatViaDate(row.DeliveryDate)) + '</td>'
            + '<td>' + escapeHtml(String(row.SellerUsr || '-')) + '</td>'
            + '<td>' + costText(rowMaterialCost) + '</td>'
            + '<td>' + costText(rowStangCost) + '</td>'
            + '<td title="Forbrugt samt modtaget og verificeret reserveret mængde medregnes">' + costText(rowPurchasedPartCost) + '</td>'
            + '<td>' + costText(rowTimeCost) + '</td>'
            + '<td><strong>' + costText(rowTotalCost) + '</strong></td>'
            + (hasBacklog ? '<td>' + formatNumber(row.RemainingSalesValue) + ' DKK</td>' : '')
            + '<td>' + (row.CostDataAvailable === false ? 'Ikke hentet' : '<div class="via-progress">' + formatNumber(progress.completedMinutes) + ' af ' + formatNumber(progress.effectiveMinutes) + ' min (' + progress.percentage + '%)<div class="via-progress-bar"><span style="width:' + progress.percentage + '%"></span></div></div>') + '</td>'
            + '<td>' + escapeHtml(String(row.ResourceName || '-')) + '<br><small>' + escapeHtml(formatViaDate(row.PlannedDate)) + '</small></td>'
            + '<td><button class="list-toggle-btn" type="button" onclick="event.stopPropagation();refreshSalgordreViaOrder(' + Number(row.OrdNo) + ', this)" style="padding:4px 8px;margin:0;">Opdater</button></td>'
            + '</tr>';
        const purchasedDetails = Array.isArray(row.PurchasedPartDetails) ? row.PurchasedPartDetails : [];
        if (purchasedDetails.length) {
            html += '<tr class="via-purchased-detail-row"><td colspan="' + (hasBacklog ? 13 : 12) + '"><details onclick="event.stopPropagation()"><summary>Indkøbte dele til ordre · ' + purchasedDetails.length + ' linjer</summary>'
                + '<div class="order-list-section"><table class="order-list-table"><thead><tr><th>Indkøbsordre</th><th>Produkt</th><th>Bestilt</th><th>Modtaget</th><th>Forbrugt</th><th>Medregnet antal</th><th>Enhedspris</th><th>Medregnet VIA</th></tr></thead><tbody>'
                + purchasedDetails.map(detail => '<tr><td>' + escapeHtml(String(detail.purchaseOrderNo || '-')) + '</td><td><strong>' + escapeHtml(String(detail.prodNo || '-')) + '</strong><br><small>' + escapeHtml(String(detail.descr || '')) + '</small></td>'
                    + '<td>' + formatNumber(detail.orderedQty) + '</td><td>' + formatNumber(detail.receivedQty) + '</td><td>' + formatNumber(detail.consumedQty) + '</td>'
                    + '<td>' + formatNumber(detail.countedQty) + '</td><td>' + formatNumber(detail.unitPrice) + ' DKK</td><td><strong>' + formatNumber(detail.countedValue) + ' DKK</strong></td></tr>').join('')
                + '</tbody></table></div></details></td></tr>';
        }
    }
    target.innerHTML = html + '</tbody></table></div>';
    applySalgordreViaColumnWidths();
}

function renderSalgordreViaReservations() {
    const target = document.getElementById('viaReservationsResults');
    const totalTarget = document.getElementById('viaReservationsTotal');
    if (!target) return;
    const query = String((document.getElementById('viaSearchInput') || {}).value || '').trim().toLowerCase();
    const rows = salgordreViaReservations
        .filter(row => row.valuationEligible && Number(row.activeQty || 0) > 0.005)
        .filter(row => !query
            || String(row.salesOrderNo || '').includes(query)
            || String(row.orderNo || '').includes(query)
            || String(row.prodNo || '').toLowerCase().includes(query)
            || String(row.customerName || '').toLowerCase().includes(query))
        .sort((left, right) => Number(left.salesOrderNo || 0) - Number(right.salesOrderNo || 0)
            || String(left.prodNo || '').localeCompare(String(right.prodNo || ''), 'da-DK'));
    const total = rows.reduce((sum, row) => sum + Number(row.activeValue || 0), 0);
    if (totalTarget) totalTarget.textContent = formatNumber(total) + ' DKK';
    if (!rows.length) {
        target.innerHTML = '<div class="omsaetning-empty">Ingen verificerede åbne reservationer matcher søgningen.</div>';
        return;
    }
    let html = '<div class="order-list-section"><table class="order-list-table"><thead><tr>'
        + '<th>Salgsordre</th><th>Kunde</th><th>Produkt</th><th>Vareparti</th>'
        + '<th>Reserveret</th><th>Plukket</th><th>Aktiv mængde</th><th>Kostpris</th><th>Lagerværdi</th>'
        + '</tr></thead><tbody>';
    for (const row of rows) {
        html += '<tr onclick="openSalgordreViaOrder(' + Number(row.salesOrderNo) + ')">'
            + '<td><strong>' + escapeHtml(String(row.salesOrderNo || '-')) + '</strong>'
            + (Number(row.orderNo || 0) !== Number(row.salesOrderNo || 0) ? '<br><small>Prod. ' + escapeHtml(String(row.orderNo || '-')) + '</small>' : '') + '</td>'
            + '<td>' + escapeHtml(String(row.customerName || '-')) + '</td>'
            + '<td><strong>' + escapeHtml(String(row.prodNo || '-')) + '</strong><br><small>' + escapeHtml(String(row.descr || '')) + '</small></td>'
            + '<td>' + escapeHtml(String(row.shipmentNo || '-')) + '</td>'
            + '<td>' + formatNumber(row.reservedQty) + '</td>'
            + '<td>' + formatNumber(row.pickedQty) + '</td>'
            + '<td>' + formatNumber(row.activeQty) + '</td>'
            + '<td>' + formatNumber(row.costPrice) + ' DKK</td>'
            + '<td><strong>' + formatNumber(row.activeValue) + ' DKK</strong></td></tr>';
    }
    target.innerHTML = html + '</tbody></table></div>';
}

function resetSalgordreVia() {
    salgordreViaRequestVersion++;
    salgordreViaReservationsVersion++;
    salgordreViaPending = null;
    salgordreViaContextKey = '';
    salgordreViaRows = [];
    salgordreViaReservations = [];
    salgordreViaOrderBacklogValue = null;
    salgordreViaMeta = {};
    salgordreViaMessage = '';
    salgordreViaLoadState = 'idle';
    for (const id of ['viaResults', 'viaKpis', 'viaStatus', 'viaReservationsResults', 'viaReservationsTotal']) {
        const target = document.getElementById(id);
        if (target) target.textContent = '';
    }
}

function salgordreViaContext() {
    return JSON.stringify([accessGranted, authToken, loggedUsername, _settingsActiveId, canAccessModule('salgordre-via'), canAccessModule('omsaetning')]);
}

function updateSalgordreViaStatus() {
    const status = document.getElementById('viaStatus');
    if (!status) return;
    const count = getSalgordreViaVisibleRows().length;
    const messages = [count + ' af ' + salgordreViaRows.length + (salgordreViaMeta.scope === 'open-backlog' ? ' åbne salgsordrer' : ' salgsordrer i produktionsudvalget')];
    if (salgordreViaMeta.asOf) messages.push('pr. ' + formatViaDate(salgordreViaMeta.asOf));
    if (salgordreViaMeta.unknownCount > 0) messages.push(salgordreViaMeta.unknownCount + ' ordrer udeladt: fakturahistorik kan ikke afstemmes');
    const missingCosts = salgordreViaRows.filter(row => row.CostDataAvailable === false).length;
    if (missingCosts) messages.push(missingCosts + ' ordrer uden hentede kostdata');
    if (Math.abs(salgordreViaMeta.excludedResidualDkk || 0) >= 0.005) messages.push('Restbeløb højst 0,01 DKK pr. ordre udeladt: ' + formatNumber(salgordreViaMeta.excludedResidualDkk) + ' DKK');
    if (salgordreViaMessage) messages.push(salgordreViaMessage);
    status.textContent = messages.join(' · ');
}

function validateSalgordreViaPayload(data) {
    if (!data || !['open-backlog', 'production'].includes(data.scope)) throw new Error('VIA-serveren skal genstartes for at indlæse det nye ordregrundlag.');
    if (!Array.isArray(data.rows)) throw new Error('Ugyldigt VIA-resultat');
    const seen = new Set();
    for (const row of data.rows) {
        const ordNo = Number(row.OrdNo);
        if (!Number.isSafeInteger(ordNo) || ordNo <= 0 || seen.has(ordNo)) throw new Error('Ugyldige eller dublerede VIA-ordrer');
        seen.add(ordNo);
        if (data.scope === 'open-backlog' && (!Number.isFinite(row.RemainingSalesValue) || row.RemainingSalesValue <= 0.01)) throw new Error('Ugyldig restsaldo for ordre ' + ordNo);
    }
    if (data.scope === 'open-backlog') {
        const sum = data.rows.reduce((total, row) => total + row.RemainingSalesValue, 0);
        if (!Number.isFinite(data.orderBacklogValueDkk) || Math.abs(sum - data.orderBacklogValueDkk) > 0.01) throw new Error('VIA-restsaldo stemmer ikke med rækkerne');
    }
}

async function fetchSalgordreViaData(parameters = {}) {
    const query = new URLSearchParams(parameters);
    if (_settingsActiveId) query.set('profile', _settingsActiveId);
    const response = await fetch('/salgordre-via?' + query, {
        headers: { Authorization: 'Bearer ' + String(authToken || '') },
        signal: AbortSignal.timeout(parameters.cached ? 10000 : 130000)
    });
    const data = await response.json();
    if (!response.ok || data.error) throw new Error(data.error || ('HTTP ' + response.status));
    return data;
}

async function loadSalgordreViaReservations(forceRefresh = false) {
    const context = salgordreViaContext();
    const version = ++salgordreViaReservationsVersion;
    const current = () => accessGranted && context === salgordreViaContext() && version === salgordreViaReservationsVersion;
    const target = document.getElementById('viaReservationsResults');
    try {
        const response = await fetch('/salgordre-via/reservations' + (forceRefresh ? '?force=1' : ''), {
            headers: { Authorization: 'Bearer ' + String(authToken || '') },
            signal: AbortSignal.timeout(65000)
        });
        const data = await response.json();
        if (!response.ok || data.error) throw new Error(data.error || ('HTTP ' + response.status));
        if (!current()) return;
        salgordreViaReservations = Array.isArray(data.rows) ? data.rows : [];
        renderSalgordreViaReservations();
    } catch (err) {
        if (!current()) return;
        salgordreViaReservations = [];
        renderSalgordreViaReservations();
        if (target) target.innerHTML = '<div class="error">Kunne ikke hente reservationer: ' + escapeHtml(String(err.message || err)) + '</div>';
    }
}

async function refreshSalgordreViaOrder(ordNo, button) {
    const context = salgordreViaContext();
    const version = salgordreViaRequestVersion;
    const current = () => accessGranted && context === salgordreViaContext() && version === salgordreViaRequestVersion;
    if (button) {
        button.disabled = true;
        button.textContent = '...';
    }
    try {
        if (salgordreViaPending) await salgordreViaPending.promise;
        if (!current()) return;
        const data = await fetchSalgordreViaData({ ordNo: String(ordNo), force: '1' });
        if (!current()) return;
        validateSalgordreViaPayload(data);
        if (data.scope !== salgordreViaMeta.scope || data.rows.some(row => Number(row.OrdNo) !== Number(ordNo))) throw new Error('Svaret matcher ikke den valgte ordre');
        const refreshed = data.rows.find(row => Number(row.OrdNo) === Number(ordNo));
        salgordreViaRows = salgordreViaRows.filter(row => Number(row.OrdNo) !== Number(ordNo));
        if (refreshed) salgordreViaRows.push(refreshed);
        salgordreViaOrderBacklogValue = data.scope === 'open-backlog' ? salgordreViaRows.reduce((sum, row) => sum + row.RemainingSalesValue, 0) : null;
        salgordreViaMeta = { ...salgordreViaMeta, unknownCount: data.unknownCount, excludedResidualDkk: data.excludedResidualDkk };
        salgordreViaMessage = 'Ordre ' + ordNo + ' opdateret; øvrige ordrer uændrede';
        if (typeof scheduleDashboardWidgets === 'function') scheduleDashboardWidgets();
        renderSalgordreVia();
    } catch (err) {
        if (!current()) return;
        if (button) {
            button.disabled = false;
            button.textContent = 'Fejl';
        }
        alert('Kunne ikke opdatere ordre ' + ordNo + ': ' + String(err.message || err));
    } finally {
        if (button && button.isConnected && button.textContent === '...') {
            button.disabled = false;
            button.textContent = 'Opdater';
        }
    }
}

async function loadSalgordreVia(forceRefresh = false, options = {}) {
    if (!accessGranted || !canAccessModule('salgordre-via')) return;
    const context = salgordreViaContext();
    if (salgordreViaContextKey !== context) {
        resetSalgordreVia();
        salgordreViaContextKey = context;
    }
    if (salgordreViaPending?.context === context) return salgordreViaPending.promise;
    const version = ++salgordreViaRequestVersion;
    const promise = loadSalgordreViaSnapshot(forceRefresh, options, () => accessGranted && context === salgordreViaContext() && version === salgordreViaRequestVersion);
    salgordreViaPending = { context, promise };
    try { return await promise; }
    finally { if (salgordreViaPending?.promise === promise) salgordreViaPending = null; }
}

async function loadSalgordreViaSnapshot(forceRefresh, options, current) {
    salgordreViaLoadState = 'loading';
    salgordreViaMessage = 'Opdaterer...';
    if (typeof scheduleDashboardWidgets === 'function') scheduleDashboardWidgets();
    const target = document.getElementById('viaResults');
    if (options.loadReservations !== false) loadSalgordreViaReservations(forceRefresh);
    const applyRows = data => {
        validateSalgordreViaPayload(data);
        salgordreViaRows = data.rows;
        salgordreViaOrderBacklogValue = data.orderBacklogValueDkk != null && Number.isFinite(Number(data.orderBacklogValueDkk)) ? Number(data.orderBacklogValueDkk) : null;
        salgordreViaMeta = data;
        salgordreViaMessage = data.fresh === false ? 'Viser tidligere data; opdaterer...' : '';
        salgordreViaLoadState = 'ready';
        renderSalgordreVia();
        if (typeof renderDashboardWidgets === 'function') renderDashboardWidgets();
    };
    if (salgordreViaRows.length) renderSalgordreVia();
    else if (target) target.innerHTML = '<div class="loading">Henter åbne salgsordrer...</div>';
    try {
        let stale = false;
        if (!forceRefresh) {
            try {
                const cachedData = await fetchSalgordreViaData({ cached: '1' });
                if (!current()) return;
                if (cachedData && !cachedData.notCached) {
                    applyRows(cachedData);
                    if (cachedData.fresh === true) return;
                    stale = true;
                }
            } catch (_) {}
        }
        if (!current()) return;
        const data = await fetchSalgordreViaData(forceRefresh || stale ? { force: '1' } : {});
        if (!current()) return;
        applyRows(data);
        if (stale && typeof showOrderDetailUpdateNotice === 'function') {
            showOrderDetailUpdateNotice('SalgOrdre VIA er opdateret med nye tal');
        }
    } catch (err) {
        if (!current()) return;
        salgordreViaLoadState = 'error';
        salgordreViaMessage = 'Opdatering mislykkedes' + (salgordreViaRows.length ? '; viser tidligere data' : '') + ': ' + String(err.message || err);
        if (typeof scheduleDashboardWidgets === 'function') scheduleDashboardWidgets();
        if (salgordreViaRows.length > 0) {
            renderSalgordreVia();
            return;
        }
        updateSalgordreViaStatus();
        if (target) target.innerHTML = '<div class="error">Kunne ikke hente SalgOrdre VIA: ' + escapeHtml(String(err.message || err)) + '</div>';
    }
}

function openSalgordreViaOrder(ordNo) {
    orderDetailReturnModule = 'salgordre-via';
    openModule('efterkalk');
    selectOrder(ordNo);
}
