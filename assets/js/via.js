// ── SalgOrdre VIA · client ──────────────────────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, formatNumber, authToken, openModule, selectOrder,
// orderDetailReturnModule (dichiarato nell'inline script).
let salgordreViaRows = [];
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

function renderSalgordreVia() {
    const target = document.getElementById('viaResults');
    const kpis = document.getElementById('viaKpis');
    const query = String((document.getElementById('viaSearchInput') || {}).value || '').trim().toLowerCase();
    if (!target) return;
    const rows = salgordreViaRows.filter(row => !query
        || String(row.OrdNo || '').includes(query)
        || String(row.CustomerName || '').toLowerCase().includes(query)).slice();
    const materialCost = rows.reduce((sum, row) => sum + Number(row.MaterialCost || 0), 0);
    const timeCost = rows.reduce((sum, row) => sum + Number(row.TimeCost || 0), 0);
    const totalCost = materialCost + timeCost;
    const salesValue = rows.reduce((sum, row) => sum + Number(row.SalesValue || 0), 0);
    if (kpis) {
        kpis.innerHTML = '<div class="via-kpi"><span>Materialekost</span><strong>' + formatNumber(materialCost) + ' DKK</strong></div>'
            + '<div class="via-kpi"><span>Tidskost (færdigmeldt)</span><strong>' + formatNumber(timeCost) + ' DKK</strong></div>'
            + '<div class="via-kpi"><span>Samlet kost</span><strong>' + formatNumber(totalCost) + ' DKK</strong></div>'
            + '<div class="via-kpi"><span>Salgsværdi</span><strong>' + formatNumber(salesValue) + ' DKK</strong></div>';
    }
    if (!rows.length) {
        target.innerHTML = '<div class="omsaetning-empty">Ingen aktive salgsordrer matcher søgningen.</div>';
        return;
    }
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
        } else if (salgordreViaSortField === 'materialCost' || salgordreViaSortField === 'timeCost' || salgordreViaSortField === 'totalCost') {
            const leftMaterialCost = Number(left.MaterialCost || 0);
            const rightMaterialCost = Number(right.MaterialCost || 0);
            const leftTimeCost = Number(left.TimeCost || 0);
            const rightTimeCost = Number(right.TimeCost || 0);
            leftValue = salgordreViaSortField === 'materialCost'
                ? leftMaterialCost
                : (salgordreViaSortField === 'timeCost' ? leftTimeCost : leftMaterialCost + leftTimeCost);
            rightValue = salgordreViaSortField === 'materialCost'
                ? rightMaterialCost
                : (salgordreViaSortField === 'timeCost' ? rightTimeCost : rightMaterialCost + rightTimeCost);
        } else {
            const key = salgordreViaSortField === 'customer' ? 'CustomerName'
                : (salgordreViaSortField === 'seller' ? 'SellerUsr' : 'ResourceName');
            leftValue = String(left[key] || '').toLocaleLowerCase('da-DK');
            rightValue = String(right[key] || '').toLocaleLowerCase('da-DK');
        }
        const comparison = leftValue < rightValue ? -1 : (leftValue > rightValue ? 1 : 0);
        return salgordreViaSortDirection === 'asc' ? comparison : -comparison;
    });
    const sortHeader = (field, label) => '<th data-sort-field="' + field + '" data-column-field="' + field + '" onclick="event.stopPropagation();setSalgordreViaSortFromElement(this)" style="cursor:pointer;user-select:none;">'
        + label + (salgordreViaSortField === field ? (salgordreViaSortDirection === 'asc' ? ' ▲' : ' ▼') : ' ↕')
        + '<span class="via-col-resizer" onmousedown="startSalgordreViaColumnResize(event, &#39;' + field + '&#39;)"></span></th>';
    const plainHeader = (field, label) => '<th data-column-field="' + field + '">' + label
        + '<span class="via-col-resizer" onmousedown="startSalgordreViaColumnResize(event, &#39;' + field + '&#39;)"></span></th>';
    let html = '<div class="order-list-section"><table class="order-list-table"><thead><tr>'
        + sortHeader('order', 'Salgsordre')
        + sortHeader('customer', 'Kunde')
        + sortHeader('deliveryDate', 'Levdato')
        + sortHeader('seller', 'Ansvarlig')
        + sortHeader('materialCost', 'Materiale')
        + sortHeader('timeCost', 'Tid')
        + sortHeader('totalCost', 'Total kost')
        + sortHeader('progress', 'Procesfremskridt')
        + sortHeader('resource', 'Næste ressource')
        + plainHeader('refresh', 'Opdater') + '</tr></thead><tbody>';
    for (const row of rows) {
        const progress = getSalgordreViaProgress(row);
        const rowMaterialCost = Number(row.MaterialCost || 0);
        const rowTimeCost = Number(row.TimeCost || 0);
        const rowTotalCost = rowMaterialCost + rowTimeCost;
        html += '<tr onclick="openSalgordreViaOrder(' + Number(row.OrdNo) + ')">'
            + '<td><strong>' + escapeHtml(String(row.OrdNo || '-')) + '</strong></td>'
            + '<td>' + escapeHtml(String(row.CustomerName || '-')) + '</td>'
            + '<td>' + escapeHtml(formatViaDate(row.DeliveryDate)) + '</td>'
            + '<td>' + escapeHtml(String(row.SellerUsr || '-')) + '</td>'
            + '<td>' + formatNumber(rowMaterialCost) + ' DKK</td>'
            + '<td>' + formatNumber(rowTimeCost) + ' DKK</td>'
            + '<td><strong>' + formatNumber(rowTotalCost) + ' DKK</strong></td>'
            + '<td><div class="via-progress">' + formatNumber(progress.completedMinutes) + ' min registrati (' + progress.percentage + '%)<div class="via-progress-bar"><span style="width:' + progress.percentage + '%"></span></div></div></td>'
            + '<td>' + escapeHtml(String(row.ResourceName || '-')) + '<br><small>' + escapeHtml(formatViaDate(row.PlannedDate)) + '</small></td>'
            + '<td><button class="list-toggle-btn" type="button" onclick="event.stopPropagation();refreshSalgordreViaOrder(' + Number(row.OrdNo) + ', this)" style="padding:4px 8px;margin:0;">Opdater</button></td>'
            + '</tr>';
    }
    target.innerHTML = html + '</tbody></table></div>';
    applySalgordreViaColumnWidths();
}

async function refreshSalgordreViaOrder(ordNo, button) {
    if (button) {
        button.disabled = true;
        button.textContent = '...';
    }
    try {
        const response = await fetch('/salgordre-via?ordNo=' + encodeURIComponent(String(ordNo)) + '&force=1', { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || data.error) throw new Error(data.error || ('HTTP ' + response.status));
        const refreshed = Array.isArray(data.rows) ? data.rows[0] : null;
        salgordreViaRows = salgordreViaRows.filter(row => Number(row.OrdNo) !== Number(ordNo));
        if (refreshed) salgordreViaRows.push(refreshed);
        const status = document.getElementById('viaStatus');
        if (status) status.textContent = salgordreViaRows.length + ' aktive salgsordrer';
        renderSalgordreVia();
    } catch (err) {
        if (button) {
            button.disabled = false;
            button.textContent = 'Fejl';
        }
        alert('Kunne ikke opdatere ordre ' + ordNo + ': ' + String(err.message || err));
    }
}

async function loadSalgordreVia(forceRefresh = false) {
    const target = document.getElementById('viaResults');
    const status = document.getElementById('viaStatus');
    if (target) target.innerHTML = '<div class="loading">Henter aktive salgsordrer...</div>';
    try {
        const response = await fetch('/salgordre-via' + (forceRefresh ? '?force=1' : ''), { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || data.error) throw new Error(data.error || ('HTTP ' + response.status));
        salgordreViaRows = Array.isArray(data.rows) ? data.rows : [];
        if (status) status.textContent = salgordreViaRows.length + ' aktive salgsordrer';
        renderSalgordreVia();
    } catch (err) {
        if (target) target.innerHTML = '<div class="error">Kunne ikke hente SalgOrdre VIA: ' + escapeHtml(String(err.message || err)) + '</div>';
    }
}

function openSalgordreViaOrder(ordNo) {
    orderDetailReturnModule = 'salgordre-via';
    openModule('efterkalk');
    selectOrder(ordNo);
}
