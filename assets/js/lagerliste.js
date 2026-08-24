// ── Lagerliste · client ─────────────────────────────────────────────────────
let lagerlisteCurrent = null;
let lagerlisteSnapshotRows = [];
let lagerlistePreviousMonth = null;
let lagerlistePreviousMonthLabel = '';

function lagerlisteFormat(value) {
    return new Intl.NumberFormat('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value || 0)) + ' DKK';
}

function lagerlisteFormatDateTime(value) {
    const date = new Date(value || '');
    if (Number.isNaN(date.getTime())) return '-';
    return date.toLocaleString('da-DK');
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
        + safeRows.map(row => '<tr>' + columns.map(column => {
            const rendered = column.format ? column.format(row[column.key], row) : (row[column.key] ?? '-');
            return '<td>' + (column.allowHtml ? String(rendered) : lagerlisteEscape(rendered)) + '</td>';
        }).join('') + '</tr>').join('')
        + '</tbody></table></div>';
}

function lagerlisteParseSortValue(value) {
    const text = String(value == null ? '' : value).trim();
    const numeric = Number(text.replace(/[^0-9,.-]/g, '').replace(/\./g, '').replace(',', '.'));
    return text && Number.isFinite(numeric) ? numeric : text.toLocaleLowerCase();
}

function lagerlisteEnhanceTables() {
    const root = document.getElementById('lagerlisteResults');
    if (!root) return;
    root.querySelectorAll('table').forEach((table, tableIndex) => {
        if (table.dataset.lagerlisteEnhanced === '1') return;
        table.dataset.lagerlisteEnhanced = '1';
        const headerCells = Array.from(table.querySelectorAll(':scope > thead > tr:first-child > th, :scope > tbody > tr:first-child > th'));
        const headerRow = headerCells.length ? headerCells[0].parentElement : null;
        if (!headerRow || !headerCells.length) return;
        const toolbar = document.createElement('div');
        toolbar.className = 'lagerliste-table-tools';
        toolbar.innerHTML = '<label>Filtrer tabel <input type="search" placeholder="Søg i denne tabel..." aria-label="Filtrer tabel" /></label><span class="lagerliste-table-hint">Klik på en kolonne for at sortere</span>';
        table.parentNode.insertBefore(toolbar, table);
        const filterInput = toolbar.querySelector('input');
        let sortIndex = -1;
        let sortDirection = 1;
        const directRows = () => Array.from(table.querySelectorAll(':scope > tbody > tr'));
        const rowGroups = () => {
            const rows = directRows();
            const groups = [];
            for (let index = 0; index < rows.length; index += 1) {
                const row = rows[index];
                if (row.classList.contains('lagerliste-plate-detail-row') && groups.length) {
                    groups[groups.length - 1].push(row);
                } else {
                    groups.push([row]);
                }
            }
            return groups;
        };
        const applyFilter = () => {
            const query = String(filterInput && filterInput.value || '').trim().toLowerCase();
            rowGroups().forEach(group => {
                const visible = !query || group.some(row => String(row.textContent || '').toLowerCase().includes(query));
                group.forEach(row => { row.style.display = visible ? '' : 'none'; });
            });
        };
        const applySort = (index) => {
            const body = table.querySelector(':scope > tbody');
            if (!body) return;
            if (sortIndex === index) sortDirection *= -1;
            else { sortIndex = index; sortDirection = 1; }
            rowGroups().sort((left, right) => {
                const leftCell = left[0].children[index];
                const rightCell = right[0].children[index];
                const leftValue = lagerlisteParseSortValue(leftCell && leftCell.textContent);
                const rightValue = lagerlisteParseSortValue(rightCell && rightCell.textContent);
                if (leftValue < rightValue) return -1 * sortDirection;
                if (leftValue > rightValue) return 1 * sortDirection;
                return 0;
            }).forEach(group => group.forEach(row => body.appendChild(row)));
            headerCells.forEach((cell, cellIndex) => {
                cell.classList.toggle('lagerliste-sort-active', cellIndex === sortIndex);
                cell.setAttribute('aria-sort', cellIndex === sortIndex ? (sortDirection === 1 ? 'ascending' : 'descending') : 'none');
            });
            applyFilter();
        };
        headerCells.forEach((cell, index) => {
            cell.tabIndex = 0;
            cell.setAttribute('role', 'button');
            cell.setAttribute('aria-sort', 'none');
            cell.addEventListener('click', () => applySort(index));
            cell.addEventListener('keydown', event => {
                if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); applySort(index); }
            });
        });
        if (filterInput) filterInput.addEventListener('input', applyFilter);
        void tableIndex;
    });
}

function lagerlisteOpenOrder(ordNo) {
    const normalized = Number(ordNo || 0);
    if (!Number.isFinite(normalized) || normalized <= 0) return;
    const input = document.getElementById('orderInput');
    if (input) input.value = String(Math.round(normalized));
    if (typeof searchOrder === 'function') {
        searchOrder();
    }
}

function lagerlisteTogglePlateGroup(button) {
    const target = document.getElementById(button.getAttribute('data-target'));
    if (!target) return;
    const open = !target.classList.contains('is-open');
    target.classList.toggle('is-open', open);
    button.textContent = open ? '−' : '+';
}

function lagerlisteToggleSection(button) {
    const target = document.getElementById(button.getAttribute('data-target'));
    if (!target) return;
    const open = target.style.display === 'none';
    target.style.display = open ? 'block' : 'none';
    button.textContent = open ? '−' : '+';
    button.setAttribute('aria-expanded', open ? 'true' : 'false');
}

function lagerlisteOpenSection(targetId) {
    const target = document.getElementById(String(targetId || ''));
    if (!target) return;
    if (target.style.display === 'none') {
        target.style.display = 'block';
        const toggle = document.querySelector('button[data-target="' + String(targetId || '').replace(/"/g, '&quot;') + '"]');
        if (toggle) {
            toggle.textContent = '−';
            toggle.setAttribute('aria-expanded', 'true');
        }
    }
    if (typeof target.scrollIntoView === 'function') {
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function lagerlisteOverviewCard({ title, total, detailTarget, emphasize = false }) {
    return '<article class="lagerliste-overview-card">'
    + '<div class="lagerliste-overview-title">' + lagerlisteEscape(title) + '</div>'
    + '<div class="lagerliste-overview-value' + (emphasize ? ' is-emphasis' : '') + '">' + lagerlisteEscape(lagerlisteFormat(total)) + '</div>'
        + '<button type="button" class="lagerliste-overview-action" onclick="lagerlisteOpenSection(\'' + lagerlisteEscape(detailTarget) + '\')">Se detaljer</button>'
        + '</article>';
}

function lagerlisteComparisonCell(value, previousValue, isChange = false) {
    const numericValue = Number(value || 0);
    if (!isChange) return lagerlisteEscape(lagerlisteFormat(numericValue));
    const difference = numericValue - Number(previousValue || 0);
    const cls = difference > 0 ? 'lagerliste-diff-pos' : (difference < 0 ? 'lagerliste-diff-neg' : 'lagerliste-diff-zero');
    return '<span class="' + cls + '">' + lagerlisteEscape(lagerlisteFormat(difference)) + '</span>';
}

function lagerlisteSummaryTable({ generatedAt, totals, categories, comparison = null }) {
    const previousTotals = comparison && comparison.totals ? comparison.totals : null;
    const viaRows = Array.isArray(categories && categories.salgordreVia) ? categories.salgordreVia : [];
    const previousViaRows = Array.isArray(comparison && comparison.categories && comparison.categories.salgordreVia)
        ? comparison.categories.salgordreVia
        : [];
    const viaTid = viaRows.reduce((sum, row) => sum + Number(row.TimeCost || 0), 0);
    const viaLaser = viaRows.reduce((sum, row) => sum + Number(row.MaterialCost || 0), 0);
    const viaStang = viaRows.reduce((sum, row) => sum + Number(row.StangCost || 0), 0);
    const nestingCuttingRows = Array.isArray(categories && categories.nestingCutting) ? categories.nestingCutting : [];
    const previousNestingCuttingRows = Array.isArray(comparison && comparison.categories && comparison.categories.nestingCutting)
        ? comparison.categories.nestingCutting
        : [];
    const viaPlader = nestingCuttingRows.reduce((sum, row) => sum + Number(row.Value || 0), 0);
    const previousViaTid = previousViaRows.reduce((sum, row) => sum + Number(row.TimeCost || 0), 0);
    const previousViaLaser = previousViaRows.reduce((sum, row) => sum + Number(row.MaterialCost || 0), 0);
    const previousViaStang = previousViaRows.reduce((sum, row) => sum + Number(row.StangCost || 0), 0);
    const previousViaPlader = previousNestingCuttingRows.reduce((sum, row) => sum + Number(row.Value || 0), 0);
    const workInProgress = Number(totals.finishedNotInvoiced || 0) + viaTid + viaLaser + viaStang + viaPlader;
    const previousWorkInProgress = previousTotals
        ? Number(previousTotals.finishedNotInvoiced || 0) + previousViaTid + previousViaLaser + previousViaStang + previousViaPlader
        : null;
    const warehouseWithoutRest = Number(totals.plates || 0) + Number(totals.opfolgningvare || 0) + Number(totals.stang || 0);
    const warehouseWithRest = warehouseWithoutRest + Number(totals.restPlates || 0);
    const previousWarehouseWithoutRest = previousTotals
        ? Number(previousTotals.plates || 0) + Number(previousTotals.opfolgningvare || 0) + Number(previousTotals.stang || 0)
        : null;
    const previousWarehouseWithRest = previousWarehouseWithoutRest === null
        ? null
        : previousWarehouseWithoutRest + Number(previousTotals.restPlates || 0);
    const rows = [
        ['Pladelager', totals.plates, previousTotals && previousTotals.plates, 'lagerliste-plates-section'],
        ['Rest plader', totals.restPlates, previousTotals && previousTotals.restPlates, 'lagerliste-rest-section'],
        ['Stang materiale', totals.stang, previousTotals && previousTotals.stang, 'lagerliste-stang-section'],
        ['Opfølgningsvarer', totals.opfolgningvare, previousTotals && previousTotals.opfolgningvare, 'lagerliste-opfolgning-section'],
        ['Varelager uden rest', warehouseWithoutRest, previousWarehouseWithoutRest, null],
        ['Varelager', warehouseWithRest, previousWarehouseWithRest, null],
        ['Færdige SO kostpris', totals.finishedNotInvoiced, previousTotals && previousTotals.finishedNotInvoiced, 'lagerliste-ready-invoice-section'],
        ['VIA Tid', viaTid, previousViaRows.length ? previousViaTid : null, 'lagerliste-salgordre-via-section', 'lagerliste-summary-subrow'],
        ['VIA Laser', viaLaser, previousViaRows.length ? previousViaLaser : null, 'lagerliste-salgordre-via-section', 'lagerliste-summary-subrow'],
        ['VIA Stang', viaStang, previousViaRows.length ? previousViaStang : null, 'lagerliste-salgordre-via-section', 'lagerliste-summary-subrow'],
        ['VIA Plader (Værdi i skæring)', viaPlader, previousNestingCuttingRows.length ? previousViaPlader : null, 'lagerliste-nesting-cutting-section', 'lagerliste-summary-subrow'],
        ['Vare i arbejde', workInProgress, previousWorkInProgress, null],
        ['TOTAL', warehouseWithRest + workInProgress, previousWarehouseWithRest === null || previousWorkInProgress === null ? null : previousWarehouseWithRest + previousWorkInProgress, null]
    ];
    const previousCell = value => value === null || value === undefined ? '-' : lagerlisteFormat(value);
    return '<section class="lagerliste-summary-board">'
        + '<div class="lagerliste-summary-head"><h4>Oversigt</h4><span class="lagerliste-generated-at">Aktuel: ' + lagerlisteEscape(generatedAt) + (comparison ? ' · Forrige måned: ' + lagerlisteEscape(comparison.label) : '') + '</span></div>'
        + '<div class="lagerliste-summary-table-wrap"><table class="lagerliste-sheet-table lagerliste-overview-table"><thead><tr><th>Post</th><th>Aktuel</th><th>Forrige måned</th><th>Ændring</th></tr></thead><tbody>'
        + rows.map(row => '<tr class="' + (row[4] || (row[0] === 'TOTAL' ? 'lagerliste-sheet-grand' : (row[0] === 'Varelager' ? 'lagerliste-sheet-total' : ''))) + '"><td>'
            + (row[3] ? '<button type="button" class="lagerliste-sheet-link" onclick="lagerlisteOpenSection(\'' + row[3] + '\')">' + lagerlisteEscape(row[0]) + '</button>' : lagerlisteEscape(row[0]))
            + '</td><td>' + lagerlisteEscape(lagerlisteFormat(row[1])) + '</td><td>' + lagerlisteEscape(previousCell(row[2])) + '</td><td>' + (comparison ? lagerlisteComparisonCell(row[1], row[2], true) : '-') + '</td></tr>').join('')
        + '</tbody></table></div></section>';
}

function lagerlisteSummaryBoard({ generatedAt, totals, categories, comparison = null }) {
    const plateGroups = Array.isArray(categories && categories.plateGroups) ? categories.plateGroups : [];
    const restGroups = Array.isArray(categories && categories.restPlateGroups) ? categories.restPlateGroups : [];
    const viaRows = Array.isArray(categories && categories.salgordreVia) ? categories.salgordreVia : [];

    const restByType = new Map();
    for (const row of restGroups) {
        const key = String(row.PlateType || '').trim() || '?';
        const item = restByType.get(key) || { PlateType: key, Value: 0 };
        item.Value += Number(row.Value || 0);
        restByType.set(key, item);
    }
    const restTypeRows = Array.from(restByType.values()).sort((a, b) => String(a.PlateType).localeCompare(String(b.PlateType)));

    const viaMaterial = viaRows.reduce((sum, row) => sum + Number(row.MaterialCost || 0), 0);
    const viaStang = viaRows.reduce((sum, row) => sum + Number(row.StangCost || 0), 0);
    const viaTime = viaRows.reduce((sum, row) => sum + Number(row.TimeCost || 0), 0);
    const viaTotal = viaMaterial + viaStang + viaTime;

    const plateFifoTotal = Array.isArray(categories && categories.plates)
        ? categories.plates.reduce((sum, row) => sum + Number(row.FifoValue || 0), 0)
        : 0;

    const warehouseWithoutRest = Number(totals.plates || 0) + Number(totals.opfolgningvare || 0) + Number(totals.stang || 0);
    const warehouseWithRest = warehouseWithoutRest + Number(totals.restPlates || 0);
    const previousTotals = comparison && comparison.totals ? comparison.totals : null;
    const previousWarehouseWithRest = previousTotals
        ? Number(previousTotals.plates || 0) + Number(previousTotals.opfolgningvare || 0) + Number(previousTotals.stang || 0) + Number(previousTotals.restPlates || 0)
        : null;
    const comparisonRows = [
        ['Pladelager', totals.plates, previousTotals && previousTotals.plates],
        ['Rest plader', totals.restPlates, previousTotals && previousTotals.restPlates],
        ['Stang materiale', totals.stang, previousTotals && previousTotals.stang],
        ['Opfølgningsvarer', totals.opfolgningvare, previousTotals && previousTotals.opfolgningvare],
        ['Klar til fakturering', totals.finishedNotInvoiced, previousTotals && previousTotals.finishedNotInvoiced],
        ['Varelager', warehouseWithRest, previousWarehouseWithRest],
        ['TOTAL', totals.total, previousTotals && previousTotals.total]
    ];

    return '<section class="lagerliste-summary-board">'
        + '<div class="lagerliste-summary-head">'
        + '<h4>Oversigt</h4>'
        + '<span class="lagerliste-generated-at">Senest opdateret: ' + lagerlisteEscape(generatedAt) + '</span>'
        + '</div>'
        + '<div class="lagerliste-comparison-block">'
        + '<div class="lagerliste-comparison-title">Sammenligning med forrige måned: ' + lagerlisteEscape(comparison ? comparison.label : 'ikke indlæst') + '</div>'
        + (comparison
            ? '<table class="lagerliste-sheet-table lagerliste-comparison-table"><thead><tr><th>Post</th><th>Aktuel</th><th>Forrige måned</th><th>Ændring</th></tr></thead><tbody>'
                + comparisonRows.map(row => '<tr' + (row[0] === 'TOTAL' ? ' class="lagerliste-sheet-grand"' : '') + '><td>' + lagerlisteEscape(row[0]) + '</td><td>' + lagerlisteComparisonCell(row[1], row[2]) + '</td><td>' + lagerlisteComparisonCell(row[2], row[1]) + '</td><td>' + lagerlisteComparisonCell(row[1], row[2], true) + '</td></tr>').join('')
                + '</tbody></table>'
            : '<div class="lagerliste-comparison-empty">Tryk “Sammenlign med forrige måned” for at hente månedslukningen.</div>')
        + '</div>'
        + '<div class="lagerliste-sheet-block">'
        + '<table class="lagerliste-sheet-table"><thead><tr><th>Pladelager</th><th>Sum af Værdi</th><th>Sum af FIFO</th></tr></thead><tbody>'
        + plateGroups.map(group => '<tr><td><button type="button" class="lagerliste-sheet-link" onclick="lagerlisteOpenSection(\'lagerliste-plates-section\')">' + lagerlisteEscape(group.PlateType + ' - ' + group.PlateTypeLabel) + '</button></td><td>' + lagerlisteEscape(lagerlisteFormat(group.Value)) + '</td><td>' + lagerlisteEscape(lagerlisteFormat(group.FifoValue)) + '</td></tr>').join('')
        + '<tr class="lagerliste-sheet-total"><td>Hovedtotal</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(Number(totals.plates || 0))) + '</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(plateFifoTotal)) + '</td></tr>'
        + '</tbody></table></div>'
        + '<div class="lagerliste-sheet-block">'
        + '<table class="lagerliste-sheet-table"><thead><tr><th>RestLager (plader)</th><th>Sum af Pris</th></tr></thead><tbody>'
        + restTypeRows.map(row => '<tr><td><button type="button" class="lagerliste-sheet-link" onclick="lagerlisteOpenSection(\'lagerliste-rest-section\')">' + lagerlisteEscape(row.PlateType) + '</button></td><td>' + lagerlisteEscape(lagerlisteFormat(row.Value)) + '</td></tr>').join('')
        + '<tr class="lagerliste-sheet-total"><td>Hovedtotal</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(Number(totals.restPlates || 0))) + '</td></tr>'
        + '</tbody></table></div>'
        + '<div class="lagerliste-sheet-block">'
        + '<table class="lagerliste-sheet-table"><tbody>'
        + '<tr><td>Div. stangmatr.</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(Number(totals.stang || 0))) + '</td></tr>'
        + '<tr><td>Opfølgningsvarer (-L)</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(Number(totals.opfolgningvare || 0))) + '</td></tr>'
        + '<tr class="lagerliste-sheet-total"><td>Varelager (uden resten)</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(warehouseWithoutRest)) + '</td></tr>'
        + '<tr class="lagerliste-sheet-total"><td>Varelager</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(warehouseWithRest)) + '</td></tr>'
        + '<tr><td>Færdige SO kostpris</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(Number(totals.finishedNotInvoiced || 0))) + '</td></tr>'
        + '<tr><td>Vare i arbejde (VIA) tider</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(viaTime)) + '</td></tr>'
        + '<tr><td>Laser materiale VIA</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(viaMaterial)) + '</td></tr>'
        + '<tr><td>Stang materiale VIA</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(viaStang)) + '</td></tr>'
        + '<tr class="lagerliste-sheet-total"><td>VIA lager</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(viaTotal)) + '</td></tr>'
        + '<tr class="lagerliste-sheet-grand"><td>TOTAL</td><td class="lagerliste-sheet-value">' + lagerlisteEscape(lagerlisteFormat(Number(totals.total || 0))) + '</td></tr>'
        + '</tbody></table></div>'
        + '</section>';
}

function lagerlisteOverviewLayout({ generatedAt, totals, categories }) {
    const cards = [
        {
            title: 'Pladelager - Værdi',
            total: Number((totals && totals.plates) || 0),
            detailTarget: 'lagerliste-plates-section'
        },
        {
            title: 'Pladelager - FIFO',
            total: Array.isArray(categories && categories.plates)
                ? categories.plates.reduce((sum, row) => sum + Number(row.FifoValue || 0), 0)
                : 0,
            detailTarget: 'lagerliste-plates-section'
        },
        {
            title: 'Rest Plader',
            total: Number((totals && totals.restPlates) || 0),
            detailTarget: 'lagerliste-rest-section'
        },
        {
            title: 'Stang materiale',
            total: Number((totals && totals.stang) || 0),
            detailTarget: 'lagerliste-stang-section'
        },
        {
            title: 'Opfølgningsvarer',
            total: Number((totals && totals.opfolgningvare) || 0),
            detailTarget: 'lagerliste-opfolgning-section'
        },
        {
            title: 'Klar til fakturering',
            total: Number((totals && totals.finishedNotInvoiced) || 0),
            detailTarget: 'lagerliste-ready-invoice-section'
        },
        {
            title: 'Igangværende arbejde (VIA)',
            total: Number((totals && totals.salgordreVia) || 0),
            detailTarget: 'lagerliste-ready-invoice-section'
        },
        {
            title: 'TOTAL',
            total: Number((totals && totals.total) || 0),
            detailTarget: 'lagerliste-plates-section',
            emphasize: true
        }
    ];
    return '<section class="lagerliste-overview">'
        + '<div class="lagerliste-overview-head">'
        + '<h4>Oversigt</h4>'
        + '<span class="lagerliste-generated-at">Senest opdateret: ' + lagerlisteEscape(generatedAt) + '</span>'
        + '</div>'
        + '<div class="lagerliste-overview-grid">' + cards.map(lagerlisteOverviewCard).join('') + '</div>'
        + '</section>';
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
    const totalValue = safeGroups.reduce((sum, group) => sum + Number(group.Value || 0), 0);
    const totalFifo = safeGroups.reduce((sum, group) => sum + Number(group.FifoValue || 0), 0);
    const table = '<div class="omsaetning-table-wrap"><table class="order-list-table"><thead><tr><th></th><th>Rækkemærkater</th><th>Plader</th><th>TOT kg</th><th>Sum af Værdi</th><th>Sum af FIFO</th><th>FIFO-pris</th></tr></thead><tbody>'
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
    return table + '<div class="lagerliste-total-row"><strong>Sum af Værdi</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalValue)) + '</strong><strong>Sum af FIFO</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalFifo)) + '</strong></div>';
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

function lagerlisteStangTable(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen stangmateriale.</div>';
    const totalValue = safeRows.reduce((sum, row) => sum + Number(row.Value || 0), 0);
    const totalFifo = safeRows.reduce((sum, row) => sum + Number(row.FifoValue || 0), 0);
    const table = lagerlisteRowsTable(safeRows, [
        { key: 'ProdNo', label: 'Produkt' },
        { key: 'Descr', label: 'Beskrivelse' },
        { key: 'Quantity', label: 'TOT kg', format: value => Number(value || 0).toFixed(2) },
        { key: 'StandardPrice', label: 'Pris', format: lagerlisteFormat },
        { key: 'UnitCost', label: 'FIFO-pris', format: lagerlisteFormat },
        { key: 'Value', label: 'Værdi', format: lagerlisteFormat },
        { key: 'FifoValue', label: 'FIFO', format: lagerlisteFormat }
    ]);
    return table + '<div class="lagerliste-total-row"><strong>Sum af Værdi</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalValue)) + '</strong><strong>Sum af FIFO</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalFifo)) + '</strong></div>';
}

function lagerlisteNestingCuttingTable(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen plader i aktivt snit denne måned.</div>';
    const totalValue = safeRows.reduce((sum, row) => sum + Number(row.Value || 0), 0);
    return lagerlisteRowsTable(safeRows, [
        { key: 'OrdNo', label: 'Nestingordre' },
        { key: 'OrdDt', label: 'Ordredato' },
        { key: 'Route', label: 'Rute (TrInf4)' },
        { key: 'ProdNo', label: 'Plade' },
        { key: 'Quantity', label: 'Færdigmeldt plade', format: value => Number(value || 0).toFixed(2) },
        { key: 'ProductCount', label: 'Ikke færdigmeldte produkter' },
        { key: 'Value', label: 'Værdi i skæring', format: lagerlisteFormat }
    ]) + '<div class="lagerliste-total-row"><strong>Samlet værdi af Plader VIA</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalValue)) + '</strong></div>';
}

function lagerlisteOpfolgningTable(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen opfølgningsvarer.</div>';
    const totalValue = safeRows.reduce((sum, row) => sum + Number(row.Value || 0), 0);
    const totalPoPhStBValue = safeRows.reduce((sum, row) => sum + Number(row.PoPhStBValue || 0), 0);
    const totalDiff = totalPoPhStBValue - totalValue;
    const table = lagerlisteRowsTable(safeRows, [
        { key: 'ProdNo', label: 'Produkt' },
        { key: 'Descr', label: 'Beskrivelse' },
        { key: 'PoPhStB', label: 'PoPhStB', format: value => Number(value || 0).toFixed(2) },
        { key: 'Beholdning', label: 'Beholdning', format: value => Number(value || 0).toFixed(2) },
        { key: 'PhCstPr', label: 'FIFO-pris', format: lagerlisteFormat },
        { key: 'Value', label: 'Værdi (Beholdning)', format: lagerlisteFormat },
        { key: 'PoPhStBValue', label: 'Værdi (PoPhStB)', format: lagerlisteFormat },
        { key: 'Diff', label: 'Dif.', format: lagerlisteFormat }
    ]);
    return table
        + '<div class="lagerliste-total-row"><strong>Sum Beholdning Værdi</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalValue)) + '</strong><strong>Sum PoPhStB Værdi</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalPoPhStBValue)) + '</strong></div>'
        + '<div class="lagerliste-total-row"><strong>Samlet difference</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalDiff)) + '</strong></div>';
}

function lagerlisteReadyToInvoiceTable(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen ordrer klar til fakturering.</div>';
    const totalValue = safeRows.reduce((sum, row) => sum + Number(row.Value || 0), 0);
    const totalLegacy = safeRows.reduce((sum, row) => sum + Number(row.LegacyValue || 0), 0);
    const totalDiff = totalValue - totalLegacy;
    const table = lagerlisteRowsTable(safeRows, [
        {
            key: 'OrdNo',
            label: 'Ordre',
            allowHtml: true,
            format: value => {
                const ordNo = Number(value || 0);
                if (!Number.isFinite(ordNo) || ordNo <= 0) return '-';
                return '<button type="button" class="lagerliste-order-link" onclick="lagerlisteOpenOrder(' + ordNo + ')">' + ordNo + '</button>';
            }
        },
        { key: 'CustomerName', label: 'Kunde' },
        { key: 'LineCount', label: 'Linjer', format: value => String(Math.round(Number(value || 0))) },
        { key: 'LegacyValue', label: 'Legacy', format: lagerlisteFormat },
        { key: 'Value', label: 'Kostpris (Efterkalk)', format: lagerlisteFormat },
        {
            key: 'Diff',
            label: 'Dif.',
            allowHtml: true,
            format: (_value, row) => {
                const diff = Number(row.Value || 0) - Number(row.LegacyValue || 0);
                const cls = diff > 0 ? 'lagerliste-diff-pos' : (diff < 0 ? 'lagerliste-diff-neg' : 'lagerliste-diff-zero');
                return '<span class="' + cls + '">' + lagerlisteEscape(lagerlisteFormat(diff)) + '</span>';
            }
        }
    ]);
    return table
        + '<div class="lagerliste-total-row"><strong>Sum Legacy</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalLegacy)) + '</strong><strong>Sum Kostpris (Efterkalk)</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalValue)) + '</strong></div>'
        + '<div class="lagerliste-total-row"><strong>Samlet difference</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalDiff)) + '</strong></div>';
}

function lagerlisteSalgordreViaTable(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen salgsordrer med aktivt arbejde.</div>';
    const totalTime = safeRows.reduce((sum, row) => sum + Number(row.TimeCost || 0), 0);
    const totalMaterial = safeRows.reduce((sum, row) => sum + Number(row.MaterialCost || 0), 0);
    const totalStang = safeRows.reduce((sum, row) => sum + Number(row.StangCost || 0), 0);
    const totalValue = safeRows.reduce((sum, row) => sum + Number(row.Value || 0), 0);
    const table = lagerlisteRowsTable(safeRows, [
        {
            key: 'OrdNo',
            label: 'Salgsordre',
            allowHtml: true,
            format: value => {
                const ordNo = Number(value || 0);
                return ordNo > 0
                    ? '<button type="button" class="lagerliste-order-link" onclick="lagerlisteOpenOrder(' + ordNo + ')">' + ordNo + '</button>'
                    : '-';
            }
        },
        { key: 'CustomerName', label: 'Kunde' },
        { key: 'TimeCost', label: 'VIA Tid', format: lagerlisteFormat },
        { key: 'MaterialCost', label: 'VIA Laser', format: lagerlisteFormat },
        { key: 'StangCost', label: 'VIA Stang', format: lagerlisteFormat },
        { key: 'Value', label: 'Vare i arbejde', format: lagerlisteFormat }
    ]);
    return table + '<div class="lagerliste-total-row"><strong>VIA Tid</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalTime))
        + '</strong><strong>VIA Laser</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalMaterial))
        + '</strong><strong>VIA Stang</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalStang))
        + '</strong><strong>Vare i arbejde</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalValue)) + '</strong></div>';
}

function lagerlisteCollapsibleSection(title, content, targetId, value = null) {
    return '<section class="lagerliste-section">'
    + '<h4><button type="button" data-target="' + lagerlisteEscape(targetId) + '" onclick="lagerlisteToggleSection(this)" aria-expanded="false">+</button> '
        + lagerlisteEscape(title) + (value === null ? '' : '<span class="lagerliste-section-value">' + lagerlisteEscape(lagerlisteFormat(value)) + '</span>') + '</h4>'
    + '<div id="' + lagerlisteEscape(targetId) + '" style="display:none">' + content + '</div>'
        + '</section>';
}

function lagerlisteRender(payload, comparison = lagerlistePreviousMonth) {
    const root = document.getElementById('lagerlisteResults');
    if (!root) return;
    lagerlisteCurrent = payload;
    const categories = payload.categories || {};
    const totals = payload.totals || {};
    const plateGroups = categories.plateGroups || [];
    const stangRows = categories.stang || [];
    const opfolgningRows = categories.opfolgningvare || [];
    const nestingCuttingRows = categories.nestingCutting || [];
    const readyToInvoiceRows = categories.finishedNotInvoiced || [];
    const viaRows = categories.salgordreVia || [];
    const sumRows = (rows, key = 'Value') => (Array.isArray(rows) ? rows : []).reduce((sum, row) => sum + Number(row[key] || 0), 0);
    const generatedAt = lagerlisteFormatDateTime(payload.generatedAt);
    root.innerHTML = lagerlisteSummaryTable({ generatedAt, totals, categories, comparison })
        + lagerlisteCollapsibleSection('Pladelager', lagerlistePlateGroupsTable(plateGroups), 'lagerliste-plates-section', totals.plates)
        + lagerlisteCollapsibleSection('Rest Plader', lagerlisteRestGroupsTable(categories.restPlateGroups || []), 'lagerliste-rest-section', totals.restPlates)
        + lagerlisteCollapsibleSection('Stang materiale', lagerlisteStangTable(stangRows), 'lagerliste-stang-section', totals.stang)
        + lagerlisteCollapsibleSection('Plader VIA', lagerlisteNestingCuttingTable(nestingCuttingRows), 'lagerliste-nesting-cutting-section', sumRows(nestingCuttingRows))
        + lagerlisteCollapsibleSection('Opfølgningsvarer', lagerlisteOpfolgningTable(opfolgningRows), 'lagerliste-opfolgning-section', totals.opfolgningvare)
        + lagerlisteCollapsibleSection('Ordrer klar til fakturering', lagerlisteReadyToInvoiceTable(readyToInvoiceRows), 'lagerliste-ready-invoice-section', totals.finishedNotInvoiced)
        + lagerlisteCollapsibleSection('Salgsordre VIA', lagerlisteSalgordreViaTable(viaRows), 'lagerliste-salgordre-via-section', sumRows(viaRows));
    lagerlisteEnhanceTables();
}

async function loadLagerliste() {
    const root = document.getElementById('lagerlisteResults');
    if (root) root.innerHTML = '<div class="loading">Henter lagerdata...</div>';
    refreshLagerlisteSnapshotList().catch(() => {});
    try {
        const response = await fetch('/lagerliste/current?force=1', { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        lagerlistePreviousMonth = null;
        lagerlistePreviousMonthLabel = '';
        lagerlisteRender(data);
    } catch (err) {
        if (root) root.innerHTML = '<div class="error">Kunne ikke hente Lagerliste: ' + lagerlisteEscape(err.message || err) + '</div>';
    }
}

function lagerlistePreviousMonthKey() {
    const date = new Date();
    date.setDate(1);
    date.setMonth(date.getMonth() - 1);
    return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
}

async function compareLagerlistePreviousMonth() {
    const status = document.getElementById('lagerlisteSnapshotStatus');
    const month = lagerlistePreviousMonthKey();
    try {
        if (status) status.textContent = 'Henter månedslukning ' + month + '...';
        const response = await fetch('/lagerliste/snapshot/' + encodeURIComponent(month), {
            headers: { Authorization: 'Bearer ' + String(authToken || '') }
        });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        lagerlistePreviousMonth = { ...(data.current || {}), label: month };
        lagerlistePreviousMonthLabel = month;
        lagerlisteRender(lagerlisteCurrent, lagerlistePreviousMonth);
        if (status) status.textContent = 'Sammenligner med ' + month;
    } catch (err) {
        lagerlistePreviousMonth = null;
        lagerlistePreviousMonthLabel = '';
        if (lagerlisteCurrent) lagerlisteRender(lagerlisteCurrent, null);
        if (status) status.textContent = 'Ingen månedslukning for ' + month;
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

function lagerlisteDownloadJson(filename, data) {
    const blob = new Blob([JSON.stringify(data, null, 2) + '\n'], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
}

function exportLagerlisteJson() {
    if (!lagerlisteCurrent) {
        alert('Ingen lagerdata at eksportere endnu.');
        return;
    }
    const stamp = new Date().toISOString().replace(/[:]/g, '-').replace(/\..+$/, '');
    lagerlisteDownloadJson('lagerliste-' + stamp + '.json', {
        exportedAt: new Date().toISOString(),
        payload: lagerlisteCurrent
    });
}

function exportLagerlistePdf() {
    const root = document.getElementById('lagerlisteResults');
    if (!root) return;
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
        alert('Kunne ikke åbne print-vindue. Tillad popups og prøv igen.');
        return;
    }
    const printRoot = root.cloneNode(true);
    printRoot.querySelectorAll('.lagerliste-section > div[id]').forEach(sectionBody => {
        sectionBody.style.display = 'block';
    });
    printRoot.querySelectorAll('.lagerliste-plate-detail-row').forEach(detailRow => {
        detailRow.style.display = 'table-row';
    });
    printRoot.querySelectorAll('.lagerliste-table-tools').forEach(tool => tool.remove());
    printWindow.document.write('<!DOCTYPE html><html><head><title>Lagerliste snapshot</title><meta charset="UTF-8">'
        + '<style>body{font-family:Segoe UI,Arial,sans-serif;padding:12px;color:#123} h2{margin:0 0 10px} table{width:100%;border-collapse:collapse;font-size:12px} th,td{border:1px solid #ccd;padding:6px;text-align:left} th{background:#eef5ff} .lagerliste-section{margin-bottom:12px} .lagerliste-section > div[id]{display:block!important} .lagerliste-plate-detail-row{display:table-row!important} .lagerliste-total-row{display:flex;gap:10px;flex-wrap:wrap;border:1px solid #ccd;padding:6px;margin-top:6px}</style>'
        + '</head><body><h2>Lagerliste</h2>' + printRoot.innerHTML + '</body></html>');
    printWindow.document.close();
    printWindow.focus();
    setTimeout(() => printWindow.print(), 150);
}

async function refreshLagerlisteSnapshotList() {
    try {
        const response = await fetch('/lagerliste/snapshots', { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        lagerlisteSnapshotRows = Array.isArray(data.rows) ? data.rows : [];
        const select = document.getElementById('lagerlisteSnapshotSelect');
        if (!select) return;
        select.innerHTML = '<option value="">Vælg snapshot...</option>'
            + lagerlisteSnapshotRows.map(row => {
                const id = String(row.snapshotId || '');
                const dateLabel = lagerlisteFormatDateTime(row.capturedAt || row.createdAt || id.replace('_', 'T').replace(/-/g, ':'));
                return '<option value="' + lagerlisteEscape(id) + '">' + lagerlisteEscape(id + (dateLabel ? ' (' + dateLabel + ')' : '')) + '</option>';
            }).join('');
    } catch (_err) {
        // Silent: snapshot list is optional.
    }
}

async function saveLagerlistePointSnapshot() {
    const status = document.getElementById('lagerlisteSnapshotStatus');
    try {
        if (status) status.textContent = 'Gemmer punkt-snapshot...';
        if (!lagerlisteCurrent) {
            throw new Error('Ingen lagerdata indlæst endnu');
        }
        const basePayload = {
            capturedAt: new Date().toISOString(),
            note: 'Manual save from Lagerliste',
            forceRefresh: '0'
        };
        const requestBody = JSON.stringify(basePayload);
        if (status) status.textContent = 'Gemmer punkt-snapshot fra server-cache...';
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 20000);
        const response = await fetch('/lagerliste/snapshots', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + String(authToken || '') },
            body: requestBody,
            signal: controller.signal
        }).finally(() => clearTimeout(timeoutId));
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        if (status) status.textContent = 'Snapshot gemt: ' + String(data.snapshotId || '-');
        setTimeout(() => refreshLagerlisteSnapshotList().catch(() => {}), 800);
    } catch (err) {
        const message = err && err.name === 'AbortError'
            ? 'Fejl: Snapshot timeout efter 20 sekunder'
            : 'Fejl: ' + String(err.message || err);
        if (status) status.textContent = message;
    }
}

async function openSelectedLagerlisteSnapshot() {
    const select = document.getElementById('lagerlisteSnapshotSelect');
    const snapshotId = String(select && select.value || '').trim();
    if (!snapshotId) return;
    const status = document.getElementById('lagerlisteSnapshotStatus');
    try {
        if (status) status.textContent = 'Henter snapshot ' + snapshotId + '...';
        const response = await fetch('/lagerliste/snapshots/' + encodeURIComponent(snapshotId), {
            headers: { Authorization: 'Bearer ' + String(authToken || '') }
        });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        const payload = data.snapshot && data.snapshot.current;
        if (!payload) throw new Error('Snapshot indeholder ingen lagerdata');
        lagerlisteRender(payload);
        if (status) status.textContent = 'Viser snapshot: ' + snapshotId;
    } catch (err) {
        if (status) status.textContent = 'Fejl: ' + String(err.message || err);
    }
}
