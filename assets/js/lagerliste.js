// ── Lagerliste · client ─────────────────────────────────────────────────────
let lagerlisteCurrent = null;

function lagerlisteFormat(value) {
    return new Intl.NumberFormat('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value || 0)) + ' DKK';
}

function lagerlisteEscape(value) {
    return String(value == null ? '' : value)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function lagerlisteRowsTable(rows, columns) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen data.</div>';
    return '<div class="omsaetning-table-wrap"><table class="order-list-table"><thead><tr>'
        + columns.map(column => '<th>' + lagerlisteEscape(column.label) + '</th>').join('')
        + '</tr></thead><tbody>'
        + safeRows.map(row => '<tr>' + columns.map(column => '<td>' + lagerlisteEscape(column.format ? column.format(row[column.key], row) : (row[column.key] ?? '-')) + '</td>').join('') + '</tr>').join('')
        + '</tbody></table></div>';
}

function lagerlisteTogglePlateGroup(button) {
    const target = document.getElementById(button.getAttribute('data-target'));
    if (!target) return;
    const open = !target.classList.contains('is-open');
    target.classList.toggle('is-open', open);
    button.textContent = open ? '−' : '+';
}

function lagerlisteFormatGroupsTable(details, parentIndex) {
    const groups = new Map();
    for (const row of details || []) {
        const format = String(row.Format || 'Andet');
        const group = groups.get(format) || { Format: format, Quantity: 0, PlateCount: 0, Value: 0, FifoValue: 0, details: [] };
        group.Quantity += Number(row.Quantity || 0);
        group.PlateCount += Number(row.PlateCount || 0);
        group.Value += Number(row.Value || 0);
        group.FifoValue += Number(row.FifoValue || 0);
        group.details.push(row);
        groups.set(format, group);
    }
    const rows = Array.from(groups.values()).map((group, index) => {
        const targetId = 'lagerliste-format-' + parentIndex + '-' + index;
        const detail = lagerlisteRowsTable(group.details, [
            { key: 'ProdNo', label: 'Produkt' }, { key: 'Descr', label: 'Beskrivelse' },
            { key: 'Thickness', label: 'Tyk.' }, { key: 'WidthM', label: 'Bredde m' }, { key: 'LengthM', label: 'Længde m' },
            { key: 'Quantity', label: 'TOT kg' }, { key: 'PlateCount', label: 'Plader' },
            { key: 'StandardPrice', label: 'Pris', format: lagerlisteFormat }, { key: 'Value', label: 'Værdi', format: lagerlisteFormat },
            { key: 'UnitCost', label: 'FIFO-pris', format: lagerlisteFormat }, { key: 'FifoValue', label: 'FIFO', format: lagerlisteFormat }
        ]);
        return '<tr><td><button type="button" data-target="' + targetId + '" onclick="lagerlisteTogglePlateGroup(this)">+</button></td>'
            + '<td><strong>' + lagerlisteEscape(group.Format) + '</strong></td><td>' + lagerlisteEscape(Math.round(group.PlateCount)) + '</td>'
            + '<td>' + lagerlisteEscape(lagerlisteFormat(group.Value)) + '</td><td>' + lagerlisteEscape(lagerlisteFormat(group.FifoValue)) + '</td>'
            + '<td>' + lagerlisteEscape(lagerlisteFormat(group.Quantity > 0 ? group.FifoValue / group.Quantity : 0)) + '</td></tr>'
            + '<tr id="' + targetId + '" class="lagerliste-plate-detail-row"><td colspan="6">' + detail + '</td></tr>';
    }).join('');
    return '<table class="order-list-table lagerliste-format-table"><thead><tr><th></th><th>Format</th><th>Plader</th><th>Sum af Værdi</th><th>Sum af FIFO</th><th>FIFO-pris</th></tr></thead><tbody>' + rows + '</tbody></table>';
}

function lagerlistePlateGroupsTable(groups) {
    const safeGroups = Array.isArray(groups) ? groups : [];
    if (!safeGroups.length) return '<div class="omsaetning-empty">Ingen plader.</div>';
    return '<div class="omsaetning-table-wrap"><table class="order-list-table"><thead><tr><th></th><th>Rækkemærkater</th><th>Plader</th><th>TOT kg</th><th>Sum af Værdi</th><th>Sum af FIFO</th><th>FIFO-pris</th></tr></thead><tbody>'
        + safeGroups.map((group, index) => {
            const targetId = 'lagerliste-plate-group-' + index;
            const detail = lagerlisteFormatGroupsTable(group.details, index);
            return '<tr><td><button type="button" data-target="' + targetId + '" onclick="lagerlisteTogglePlateGroup(this)">+</button></td>'
                + '<td><strong>' + lagerlisteEscape(group.PlateType + ' - ' + group.PlateTypeLabel) + '</strong></td>'
                + '<td>' + lagerlisteEscape(group.PlateCount) + '</td>'
                + '<td>' + lagerlisteEscape(group.Quantity) + '</td>'
                + '<td>' + lagerlisteEscape(lagerlisteFormat(group.Value)) + '</td>'
                + '<td>' + lagerlisteEscape(lagerlisteFormat(group.FifoValue)) + '</td>'
                + '<td>' + lagerlisteEscape(lagerlisteFormat(group.FifoPrice)) + '</td></tr>'
                + '<tr id="' + targetId + '" class="lagerliste-plate-detail-row"><td colspan="7">' + detail + '</td></tr>';
        }).join('')
        + '</tbody></table></div>';
}

function lagerlisteRestGroupsTable(groups) {
    const typeMap = new Map();
    for (const group of groups || []) {
        const typeGroup = typeMap.get(group.PlateType) || { PlateType: group.PlateType, Weight: 0, Value: 0, materials: [] };
        typeGroup.Weight += Number(group.Weight || 0);
        typeGroup.Value += Number(group.Value || 0);
        typeGroup.materials.push(group);
        typeMap.set(group.PlateType, typeGroup);
    }
    const rows = Array.from(typeMap.values()).map((typeGroup, typeIndex) => {
        const typeTargetId = 'lagerliste-rest-type-' + typeIndex;
        const materialRows = typeGroup.materials.map((group, index) => {
            const targetId = 'lagerliste-rest-group-' + typeIndex + '-' + index;
        const detail = lagerlisteRowsTable(group.details, [
            { key: 'ProdNo', label: 'Produkt' }, { key: 'OrdNo', label: 'Ordre' }, { key: 'Txt2', label: 'Dimension' },
            { key: 'Weight', label: 'Vægt kg' }, { key: 'PricePerKg', label: 'Pris/kg', format: value => Number(value || 0).toFixed(2) + ' DKK' },
            { key: 'Value', label: 'Værdi', format: lagerlisteFormat }
        ]);
        const price = group.Weight > 0 ? group.Value / group.Weight : 0;
            return '<tr><td><button type="button" data-target="' + targetId + '" onclick="lagerlisteTogglePlateGroup(this)">+</button></td>'
            + '<td><strong>' + lagerlisteEscape(group.PlateType) + '</strong></td><td>' + lagerlisteEscape(group.Material) + '</td>'
            + '<td>' + lagerlisteEscape(group.Weight.toFixed(2)) + '</td><td>' + lagerlisteEscape(lagerlisteFormat(group.Value)) + '</td>'
            + '<td>' + lagerlisteEscape(Number(price).toFixed(2) + ' DKK') + '</td></tr>'
            + '<tr id="' + targetId + '" class="lagerliste-plate-detail-row"><td colspan="6">' + detail + '</td></tr>';
            }).join('');
            const typePrice = typeGroup.Weight > 0 ? typeGroup.Value / typeGroup.Weight : 0;
            return '<tr><td><button type="button" data-target="' + typeTargetId + '" onclick="lagerlisteTogglePlateGroup(this)">+</button></td>'
                + '<td colspan="2"><strong>' + lagerlisteEscape(typeGroup.PlateType) + '</strong></td>'
                + '<td>' + lagerlisteEscape(typeGroup.Weight.toFixed(2)) + '</td><td>' + lagerlisteEscape(lagerlisteFormat(typeGroup.Value)) + '</td>'
                + '<td>' + lagerlisteEscape(Number(typePrice).toFixed(2) + ' DKK') + '</td></tr>'
                + '<tr id="' + typeTargetId + '" class="lagerliste-plate-detail-row"><td colspan="6"><table class="order-list-table lagerliste-format-table"><thead><tr><th></th><th>Type</th><th>Materiale</th><th>Vægt kg</th><th>Værdi</th><th>Pris/kg</th></tr></thead><tbody>' + materialRows + '</tbody></table></td></tr>';
            });
    if (!rows) return '<div class="omsaetning-empty">Ingen restplader.</div>';
    return '<div class="omsaetning-table-wrap"><table class="order-list-table"><thead><tr><th></th><th colspan="2">Type</th><th>Vægt kg</th><th>Værdi</th><th>Pris/kg</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
}

function lagerlisteRender(payload) {
    const root = document.getElementById('lagerlisteResults');
    if (!root) return;
    lagerlisteCurrent = payload;
    const categories = payload.categories || {};
    const totals = payload.totals || {};
    const plateGroups = categories.plateGroups || [];
    const total = plateGroups.reduce((sum, group) => sum + Number(group.Value || 0), 0);
    root.innerHTML = '<div class="via-kpis lagerliste-kpis">'
        + '<div class="via-kpi"><span>Pladelager i alt</span><strong>' + lagerlisteFormat(total) + '</strong></div>'
        + '<div class="via-kpi"><span>Antal typer</span><strong>' + String(plateGroups.length) + '</strong></div>'
        + '<div class="via-kpi"><span>Rest Plader</span><strong>' + lagerlisteFormat(totals.restPlates) + '</strong></div>'
        + '</div>'
        + '<section class="lagerliste-section"><h4>Pladelager</h4>'
        + lagerlistePlateGroupsTable(plateGroups) + '</section>'
        + '<section class="lagerliste-section"><h4>Rest Plader</h4>'
        + lagerlisteRestGroupsTable(categories.restPlateGroups || []) + '</section>';
}

async function loadLagerliste() {
    const root = document.getElementById('lagerlisteResults');
    if (root) root.innerHTML = '<div class="loading">Henter lagerdata...</div>';
    try {
        const response = await fetch('/lagerliste/current', { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        lagerlisteRender(data);
    } catch (err) {
        if (root) root.innerHTML = '<div class="error">Kunne ikke hente Lagerliste: ' + lagerlisteEscape(err.message || err) + '</div>';
    }
}

async function saveLagerlisteSnapshot() {
    const monthInput = document.getElementById('lagerlisteMonth');
    const month = String(monthInput && monthInput.value || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) {
        alert('Vælg en måned.');
        return;
    }
    try {
        const response = await fetch('/lagerliste/snapshot/' + encodeURIComponent(month), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + String(authToken || '') },
            body: JSON.stringify({ diverse: [] })
        });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        const status = document.getElementById('lagerlisteSnapshotStatus');
        if (status) status.textContent = 'Snapshot gemt for ' + month;
    } catch (err) {
        const status = document.getElementById('lagerlisteSnapshotStatus');
        if (status) status.textContent = 'Fejl: ' + String(err.message || err);
    }
}
