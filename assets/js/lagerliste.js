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

function lagerlisteNestingCountedValue(row) {
    const value = Number(row && row.Value || 0);
    if (row && row.CountedValue !== undefined && row.CountedValue !== null) return Number(row.CountedValue || 0);
    return value < 0 ? 0 : value;
}

function lagerlisteRowsTable(rows, columns, rowAttributes = null) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen data.</div>';
    return '<div class="omsaetning-table-wrap"><table class="order-list-table"><thead><tr>'
        + columns.map(column => '<th>' + lagerlisteEscape(column.label) + '</th>').join('')
        + '</tr></thead><tbody>'
        + safeRows.map(row => '<tr' + (typeof rowAttributes === 'function' ? String(rowAttributes(row) || '') : '') + '>' + columns.map(column => {
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
    ['lagerlisteResults', 'lagerlisteVareopslagResults', 'lagerlisteCompareResults'].forEach(rootId => {
        const root = document.getElementById(rootId);
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
    const viaIndkobt = viaRows.reduce((sum, row) => sum + Number(row.PurchasedPartCost || 0), 0);
    const nestingCuttingRows = Array.isArray(categories && categories.nestingCutting) ? categories.nestingCutting : [];
    const previousNestingCuttingRows = Array.isArray(comparison && comparison.categories && comparison.categories.nestingCutting)
        ? comparison.categories.nestingCutting
        : [];
    const viaPlader = nestingCuttingRows.reduce((sum, row) => sum + lagerlisteNestingCountedValue(row), 0);
    const previousViaTid = previousViaRows.reduce((sum, row) => sum + Number(row.TimeCost || 0), 0);
    const previousViaLaser = previousViaRows.reduce((sum, row) => sum + Number(row.MaterialCost || 0), 0);
    const previousViaStang = previousViaRows.reduce((sum, row) => sum + Number(row.StangCost || 0), 0);
    const previousViaIndkobt = previousViaRows.reduce((sum, row) => sum + Number(row.PurchasedPartCost || 0), 0);
    const previousViaPlader = previousNestingCuttingRows.reduce((sum, row) => sum + lagerlisteNestingCountedValue(row), 0);
    const lagerKomponenterValue = (Array.isArray(categories && categories.gr5Items) ? categories.gr5Items : [])
        .reduce((sum, row) => sum + Number(row.FifoValue || 0), 0);
    const previousLagerKomponenterValue = (Array.isArray(comparison && comparison.categories && comparison.categories.gr5Items) ? comparison.categories.gr5Items : [])
        .reduce((sum, row) => sum + Number(row.FifoValue || 0), 0);
    const workInProgress = Number(totals.finishedNotInvoiced || 0) + viaTid + viaLaser + viaStang + viaIndkobt + viaPlader;
    const previousWorkInProgress = previousTotals
        ? Number(previousTotals.finishedNotInvoiced || 0) + previousViaTid + previousViaLaser + previousViaStang + previousViaIndkobt + previousViaPlader
        : null;
    const warehouseWithoutRest = Number(totals.plates || 0) + Number(totals.opfolgningvare || 0) + Number(totals.stang || 0) + lagerKomponenterValue;
    const warehouseWithRest = warehouseWithoutRest + Number(totals.restPlates || 0);
    const previousWarehouseWithoutRest = previousTotals
        ? Number(previousTotals.plates || 0) + Number(previousTotals.opfolgningvare || 0) + Number(previousTotals.stang || 0) + previousLagerKomponenterValue
        : null;
    const previousWarehouseWithRest = previousWarehouseWithoutRest === null
        ? null
        : previousWarehouseWithoutRest + Number(previousTotals.restPlates || 0);
    const rows = [
        ['Pladelager', totals.plates, previousTotals && previousTotals.plates, 'lagerliste-plates-section', '', 'Plader på lager: beholdning × standardpris.'],
        ['Rest plader', totals.restPlates, previousTotals && previousTotals.restPlates, 'lagerliste-rest-section', '', 'Restplader: vægt × fast pris pr. kg.'],
        ['Stang materiale', totals.stang, previousTotals && previousTotals.stang, 'lagerliste-stang-section', '', 'Stangmateriale: lagerbevægelse til og med i dag × pris/FIFO-pris.'],
        ['Opfølgningsvarer', totals.opfolgningvare, previousTotals && previousTotals.opfolgningvare, 'lagerliste-opfolgning-section', '', 'Opfølgningsvarer: (Bal + StcInc − ShpRsv) × FIFO-pris.'],
        ['Lager Komponenter (FIFO)', lagerKomponenterValue, previousLagerKomponenterValue, 'lagerliste-gr5-section', '', 'Komponenter med Prod.Gr5 = 11: beholdning × FIFO-pris.'],
        ['Varelager uden rest', warehouseWithoutRest, previousWarehouseWithoutRest, null, '', 'Plader + stangmateriale + opfølgningsvarer + lagerkomponenter, uden restplader.'],
        ['Varelager', warehouseWithRest, previousWarehouseWithRest, null, '', 'Varelager uden rest + Rest plader.'],
        ['Færdige SO kostpris', totals.finishedNotInvoiced, previousTotals && previousTotals.finishedNotInvoiced, 'lagerliste-ready-invoice-section', '', 'Færdigmeldte salgsordrer, der endnu ikke er faktureret, beregnet med Efterkalk.'],
        ['VIA Tid', viaTid, previousViaRows.length ? previousViaTid : null, 'lagerliste-salgordre-via-section', 'lagerliste-summary-subrow', 'Aktive salgsordrer VIA: registrerede minutter × operationspris.'],
        ['VIA Laser', viaLaser, previousViaRows.length ? previousViaLaser : null, 'lagerliste-salgordre-via-section', 'lagerliste-summary-subrow', 'Aktive salgsordrer VIA: registreret laser/materialeforbrug.'],
        ['VIA Stang', viaStang, previousViaRows.length ? previousViaStang : null, 'lagerliste-salgordre-via-section', 'lagerliste-summary-subrow', 'Aktive salgsordrer VIA: registreret stangmateriale.'],
        ['Indkøbt dele', viaIndkobt, previousViaRows.length ? previousViaIndkobt : null, 'lagerliste-salgordre-via-section', 'lagerliste-summary-subrow', 'Aktive salgsordrer VIA: købte dele fra den tilknyttede købsordre.'],
        ['VIA Plader (Værdi i skæring)', viaPlader, previousNestingCuttingRows.length ? previousViaPlader : null, 'lagerliste-nesting-cutting-section', 'lagerliste-summary-subrow', 'Plader i aktiv skæring: pladen er færdigmeldt, mens alle produkter på samme rute stadig ikke er færdigmeldt.'],
        ['Vare i arbejde', workInProgress, previousWorkInProgress, null, '', 'Færdige SO kostpris + VIA Tid + VIA Laser + VIA Stang + Indkøbt dele + VIA Plader.'],
        ['TOTAL', warehouseWithRest + workInProgress, previousWarehouseWithRest === null || previousWorkInProgress === null ? null : previousWarehouseWithRest + previousWorkInProgress, null, '', 'Varelager + Vare i arbejde.']
    ];
    const previousCell = value => value === null || value === undefined ? '-' : lagerlisteFormat(value);
    return '<section class="lagerliste-summary-board">'
        + '<div class="lagerliste-summary-head"><h4>Oversigt</h4><span class="lagerliste-generated-at">Aktuel: ' + lagerlisteEscape(generatedAt) + (comparison ? ' · Forrige måned: ' + lagerlisteEscape(comparison.label) : '') + '</span></div>'
        + '<div class="lagerliste-summary-table-wrap"><table class="lagerliste-sheet-table lagerliste-overview-table"><thead><tr><th>Post</th><th>Aktuel</th><th>Forrige måned</th><th>Ændring</th><th>Info</th></tr></thead><tbody>'
        + rows.map(row => '<tr class="' + (row[4] || (row[0] === 'TOTAL' ? 'lagerliste-sheet-grand' : (row[0] === 'Varelager' ? 'lagerliste-sheet-total' : ''))) + '"><td>'
            + (row[3] ? '<button type="button" class="lagerliste-sheet-link" onclick="lagerlisteOpenSection(\'' + row[3] + '\')">' + lagerlisteEscape(row[0]) + '</button>' : lagerlisteEscape(row[0]))
            + '</td><td>' + lagerlisteEscape(lagerlisteFormat(row[1])) + '</td><td>' + lagerlisteEscape(previousCell(row[2])) + '</td><td>' + (comparison ? lagerlisteComparisonCell(row[1], row[2], true) : '-') + '</td><td class="lagerliste-explanation"><span class="lagerliste-info-icon" title="' + lagerlisteEscape(row[5]) + '" aria-label="' + lagerlisteEscape(row[5]) + '" role="img">i</span></td></tr>').join('')
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
    const totalFifo = safeRows.reduce((sum, row) => sum + Number(row.FifoValue || 0), 0);
    const table = lagerlisteRowsTable(safeRows, [
        { key: 'ProdNo', label: 'Produkt' },
        { key: 'Descr', label: 'Beskrivelse' },
        { key: 'Quantity', label: 'TOT kg', format: value => Number(value || 0).toFixed(2) },
        { key: 'StandardPrice', label: 'Pris', format: lagerlisteFormat },
        { key: 'UnitCost', label: 'FIFO-pris', format: lagerlisteFormat },
        { key: 'FifoValue', label: 'Værdi (FIFO)', format: lagerlisteFormat }
    ]);
    return table + '<div class="lagerliste-total-row"><strong>Sum af Værdi (FIFO)</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalFifo)) + '</strong></div>';
}

function lagerlisteGr5Table(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen Lager Komponenter.</div>';
    const totalValue = safeRows.reduce((sum, row) => sum + Number(row.FifoValue || 0), 0);
    return lagerlisteRowsTable(safeRows, [
        { key: 'ProdNo', label: 'Produkt' },
        { key: 'Descr', label: 'Beskrivelse' },
        { key: 'Quantity', label: 'Beholdning', format: value => Number(value || 0).toFixed(2) },
        { key: 'StandardPrice', label: 'Pris', format: lagerlisteFormat },
        { key: 'UnitCost', label: 'FIFO-pris', format: lagerlisteFormat },
        { key: 'FifoValue', label: 'Værdi (FIFO)', format: lagerlisteFormat }
    ]) + '<div class="lagerliste-total-row"><strong>Sum af Værdi (FIFO)</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalValue)) + '</strong></div>';
}

function lagerlisteNestingCuttingTable(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!safeRows.length) return '<div class="omsaetning-empty">Ingen plader i aktivt snit de seneste 3 måneder.</div>';
    const totalValue = safeRows.reduce((sum, row) => sum + lagerlisteNestingCountedValue(row), 0);
    return lagerlisteRowsTable(safeRows, [
        { key: 'OrdNo', label: 'Nestingordre' },
        { key: 'OrdDt', label: 'Ordredato' },
        { key: 'Route', label: 'Rute (TrInf4)' },
        { key: 'ProdNo', label: 'Plade' },
        { key: 'Products', label: 'Produkter', format: value => String(value || '').trim() || '-' },
        { key: 'Quantity', label: 'Færdigmeldt plade', format: value => Number(value || 0).toFixed(2) },
        { key: 'ProductCount', label: 'Ikke færdigmeldte produkter' },
        { key: 'Value', label: 'Vist værdi', format: lagerlisteFormat },
        { key: 'CountedValue', label: 'Værdi i skæring', format: (_value, row) => lagerlisteFormat(lagerlisteNestingCountedValue(row)) },
        { key: 'IsEstimatedRest', label: 'Status', format: value => value ? 'Estimeret rest (ikke medregnet)' : 'Medregnet' }
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
    const totalPurchased = safeRows.reduce((sum, row) => sum + Number(row.PurchasedPartCost || 0), 0);
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
        { key: 'PurchasedPartCost', label: 'Indkøbt dele', format: lagerlisteFormat },
        { key: 'SalesValue', label: 'Salgsværdi', format: lagerlisteFormat },
        { key: 'Value', label: 'Vare i arbejde', format: lagerlisteFormat }
    ]);
    return table + '<div class="lagerliste-total-row"><strong>VIA Tid</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalTime))
        + '</strong><strong>VIA Laser</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalMaterial))
        + '</strong><strong>VIA Stang</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalStang))
        + '</strong><strong>Indkøbt dele</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalPurchased))
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
    const gr5Rows = categories.gr5Items || [];
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
        + lagerlisteCollapsibleSection('Lager Komponenter', lagerlisteGr5Table(gr5Rows), 'lagerliste-gr5-section', sumRows(gr5Rows, 'FifoValue'))
        + lagerlisteCollapsibleSection('Plader VIA', lagerlisteNestingCuttingTable(nestingCuttingRows), 'lagerliste-nesting-cutting-section', nestingCuttingRows.reduce((sum, row) => sum + lagerlisteNestingCountedValue(row), 0))
        + lagerlisteCollapsibleSection('Opfølgningsvarer', lagerlisteOpfolgningTable(opfolgningRows), 'lagerliste-opfolgning-section', totals.opfolgningvare)
        + lagerlisteCollapsibleSection('Ordrer klar til fakturering', lagerlisteReadyToInvoiceTable(readyToInvoiceRows), 'lagerliste-ready-invoice-section', totals.finishedNotInvoiced)
        + lagerlisteCollapsibleSection('Salgsordre VIA', lagerlisteSalgordreViaTable(viaRows), 'lagerliste-salgordre-via-section', sumRows(viaRows));
    lagerlisteEnhanceTables();
}

async function loadLagerliste(forceAftercalc = false) {
    const root = document.getElementById('lagerlisteResults');
    if (root) root.innerHTML = '<div class="loading">' + (forceAftercalc ? 'Genberegner Efterkalk og henter lagerdata...' : 'Henter lagerdata...') + '</div>';
    refreshLagerlisteSnapshotList().then(() => refreshLagerlisteCompareOptions()).catch(() => {});
    try {
        const response = await fetch('/lagerliste/current?force=1' + (forceAftercalc ? '&aftercalc=1' : ''), { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        lagerlistePreviousMonth = null;
        lagerlistePreviousMonthLabel = '';
        lagerlisteRender(data);
    } catch (err) {
        if (root) root.innerHTML = '<div class="error">Kunne ikke hente Lagerliste: ' + lagerlisteEscape(err.message || err) + '</div>';
    }
}

function lagerlisteFormatVismaDate(value) {
    const raw = String(Math.round(Number(value || 0)) || '');
    if (raw.length !== 8) return '-';
    return raw.slice(6, 8) + '.' + raw.slice(4, 6) + '.' + raw.slice(0, 4);
}

function lagerlisteOrderTypeLabel(trTp) {
    const map = { 1: 'Salgsordre', 5: 'Produktionsordre', 6: 'Indkøbsordre', 7: 'Produktionsordre' };
    const code = Number(trTp || 0);
    return map[code] || (code ? 'Type ' + code : '-');
}

function lagerlisteOrderLinkCell(value) {
    const ordNo = Number(value || 0);
    return ordNo > 0
        ? '<button type="button" class="lagerliste-order-link" onclick="lagerlisteOpenOrder(' + ordNo + ')">' + ordNo + '</button>'
        : '-';
}

function lagerlisteVareopslagRender(data) {
    const root = document.getElementById('lagerlisteVareopslagResults');
    if (!root) return;
    const product = data.product || {};
    const summary = data.summary || {};
    const number = value => new Intl.NumberFormat('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value || 0));
    const keyFigures = [
        ['Fysisk beholdning (PoPhStB)', number(summary.physicalStock)],
        ['Beholdning (Bal)', number(summary.bal)],
        ['Modtaget ej faktureret (StcInc)', number(summary.incoming)],
        ['Reserveret (ShpRsv)', number(summary.reserved)],
        ['Disponibel (Bal + StcInc − ShpRsv)', number(summary.available)],
        ['FIFO-pris', lagerlisteFormat(summary.fifoPrice)],
        ['Standardpris (Inf)', lagerlisteFormat(summary.standardPrice)],
        ['Lagerværdi (FIFO)', lagerlisteFormat(summary.fifoValue)],
        ['Antal varepartier', String(summary.lotCount || 0)],
        ['Lokationer', (summary.locations || []).join(', ') || '-'],
        ['Aktive reservationer', String(summary.reservationCount || 0)],
        ['Åbne ordrelinjer', String(summary.openOrderLineCount || 0)]
    ];
    const keyFigureHtml = '<div class="lagerliste-vareopslag-grid">'
        + keyFigures.map(([label, value]) => '<div class="lagerliste-vareopslag-kpi"><span>' + lagerlisteEscape(label) + '</span><strong>' + lagerlisteEscape(value) + '</strong></div>').join('')
        + '</div>';
    const lotsHtml = lagerlisteRowsTable(data.lots || [], [
        { key: 'ShpNo', label: 'Parti' },
        { key: 'Loc', label: 'Lokation', format: value => String(value || '').trim() || '-' },
        { key: 'RestBal', label: 'Restbeholdning', format: value => number(value) },
        { key: 'NoRsv', label: 'Reserveret', format: value => number(value) },
        { key: 'CstPr', label: 'Kostpris', format: lagerlisteFormat },
        { key: 'Value', label: 'Værdi', format: lagerlisteFormat },
        { key: 'RecDt', label: 'Modtaget', format: lagerlisteFormatVismaDate },
        { key: 'OrdNo', label: 'Ordre', allowHtml: true, format: lagerlisteOrderLinkCell }
    ]);
    const reservationsHtml = lagerlisteRowsTable(data.reservations || [], [
        { key: 'OrdNo', label: 'Ordre', allowHtml: true, format: lagerlisteOrderLinkCell },
        { key: 'OrderTrTp', label: 'Ordretype', format: lagerlisteOrderTypeLabel },
        { key: 'SalesOrdNo', label: 'Salgsordre', allowHtml: true, format: lagerlisteOrderLinkCell },
        { key: 'CustomerName', label: 'Kunde', format: value => String(value || '').trim() || '-' },
        { key: 'NoRsv', label: 'Reserveret', format: value => number(value) },
        { key: 'NoPic', label: 'Plukket', format: value => number(value) },
        { key: 'NoFin', label: 'Færdigmeldt', format: value => number(value) },
        { key: 'CstPr', label: 'Kostpris', format: lagerlisteFormat },
        { key: 'Value', label: 'Reserveret værdi', format: lagerlisteFormat },
        { key: 'DelDt', label: 'Leveringsdato', format: lagerlisteFormatVismaDate }
    ]);
    const orderLinesHtml = lagerlisteRowsTable(data.openOrderLines || [], [
        { key: 'OrdNo', label: 'Ordre', allowHtml: true, format: lagerlisteOrderLinkCell },
        { key: 'LnNo', label: 'Linje' },
        { key: 'OrderTrTp', label: 'Ordretype', format: lagerlisteOrderTypeLabel },
        { key: 'SalesOrdNo', label: 'Salgsordre', allowHtml: true, format: lagerlisteOrderLinkCell },
        { key: 'CustomerName', label: 'Kunde', format: value => String(value || '').trim() || '-' },
        { key: 'NoOrg', label: 'Bestilt', format: value => number(value) },
        { key: 'NoFin', label: 'Færdigmeldt', format: value => number(value) },
        { key: 'Rest', label: 'Rest', allowHtml: false, format: (_value, row) => number(Number(row.NoOrg || 0) - Number(row.NoFin || 0)) },
        { key: 'NoRsv', label: 'Reserveret', format: value => number(value) },
        { key: 'CstPr', label: 'Kostpris', format: lagerlisteFormat },
        { key: 'DelDt', label: 'Leveringsdato', format: lagerlisteFormatVismaDate }
    ]);
    root.innerHTML = '<div class="lagerliste-vareopslag-panel">'
        + '<div class="lagerliste-vareopslag-head"><h4>' + lagerlisteEscape(String(product.ProdNo || '').trim()) + ' — ' + lagerlisteEscape(String(product.Descr || '').trim()) + '</h4>'
        + '<button type="button" class="lagerliste-vareopslag-close" onclick="lagerlisteVareopslagClear()">Luk</button></div>'
        + keyFigureHtml
        + '<h5>Varepartier (' + (data.lots || []).length + ')</h5>' + lotsHtml
        + '<h5>Aktive reservationer (' + (data.reservations || []).length + ')</h5>' + reservationsHtml
        + '<h5>Åbne ordrelinjer (' + (data.openOrderLines || []).length + ')</h5>' + orderLinesHtml
        + '</div>';
    lagerlisteEnhanceTables();
}

function lagerlisteVareopslagClear() {
    const root = document.getElementById('lagerlisteVareopslagResults');
    if (root) root.innerHTML = '';
    const input = document.getElementById('lagerlisteVareopslagInput');
    if (input) input.value = '';
}

async function lagerlisteVareopslag() {
    const input = document.getElementById('lagerlisteVareopslagInput');
    const root = document.getElementById('lagerlisteVareopslagResults');
    const prodNo = String(input && input.value || '').trim();
    if (!prodNo) {
        if (root) root.innerHTML = '<div class="omsaetning-empty">Indtast et varenummer.</div>';
        return;
    }
    if (root) root.innerHTML = '<div class="loading">Søger ' + lagerlisteEscape(prodNo) + '...</div>';
    try {
        const response = await fetch('/lagerliste/vareopslag/' + encodeURIComponent(prodNo), {
            headers: { Authorization: 'Bearer ' + String(authToken || '') }
        });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        lagerlisteVareopslagRender(data);
    } catch (err) {
        if (root) root.innerHTML = '<div class="error">Vareopslag fejlede: ' + lagerlisteEscape(err.message || err) + '</div>';
    }
}

function lagerlistePreviousMonthKey() {
    const date = new Date();
    date.setDate(1);
    date.setMonth(date.getMonth() - 1);
    return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
}

// ── Sammenlign to perioder (måneder/snapshots/aktuel) ───────────────────────
function lagerlisteComputeFigures(payload) {
    const totals = (payload && payload.totals) || {};
    const categories = (payload && payload.categories) || {};
    const sumRows = (rows, key) => (Array.isArray(rows) ? rows : []).reduce((sum, row) => sum + Number(row[key] || 0), 0);
    const viaRows = categories.salgordreVia || [];
    const viaTid = sumRows(viaRows, 'TimeCost');
    const viaLaser = sumRows(viaRows, 'MaterialCost');
    const viaStang = sumRows(viaRows, 'StangCost');
    const viaIndkobt = sumRows(viaRows, 'PurchasedPartCost');
    const viaPlader = (categories.nestingCutting || []).reduce((sum, row) => sum + lagerlisteNestingCountedValue(row), 0);
    const lagerKomponenter = sumRows(categories.gr5Items || [], 'FifoValue');
    const warehouseWithoutRest = Number(totals.plates || 0) + Number(totals.opfolgningvare || 0) + Number(totals.stang || 0) + lagerKomponenter;
    const warehouseWithRest = warehouseWithoutRest + Number(totals.restPlates || 0);
    const workInProgress = Number(totals.finishedNotInvoiced || 0) + viaTid + viaLaser + viaStang + viaIndkobt + viaPlader;
    return [
        ['Pladelager', Number(totals.plates || 0), ''],
        ['Rest plader', Number(totals.restPlates || 0), ''],
        ['Stang materiale', Number(totals.stang || 0), ''],
        ['Opfølgningsvarer', Number(totals.opfolgningvare || 0), ''],
        ['Lager Komponenter (FIFO)', lagerKomponenter, ''],
        ['Varelager uden rest', warehouseWithoutRest, ''],
        ['Varelager', warehouseWithRest, 'lagerliste-sheet-total'],
        ['Færdige SO kostpris', Number(totals.finishedNotInvoiced || 0), ''],
        ['VIA Tid', viaTid, 'lagerliste-summary-subrow'],
        ['VIA Laser', viaLaser, 'lagerliste-summary-subrow'],
        ['VIA Stang', viaStang, 'lagerliste-summary-subrow'],
        ['Indkøbt dele', viaIndkobt, 'lagerliste-summary-subrow'],
        ['VIA Plader (Værdi i skæring)', viaPlader, 'lagerliste-summary-subrow'],
        ['Vare i arbejde', workInProgress, ''],
        ['TOTAL', warehouseWithRest + workInProgress, 'lagerliste-sheet-grand']
    ];
}

function lagerlisteMovementSpecs(payload) {
    const categories = (payload && payload.categories) || {};
    const flatPlateDetails = [];
    for (const group of categories.plateGroups || []) {
        for (const row of (group && group.details) || []) flatPlateDetails.push(row);
    }
    const flatRestDetails = [];
    for (const group of categories.restPlateGroups || []) {
        for (const row of (group && group.details) || []) flatRestDetails.push(row);
    }
    return [
        { category: 'Pladelager', rows: flatPlateDetails, keyOf: r => String(r.ProdNo || ''), labelOf: r => r.Descr, valueOf: r => Number(r.FifoValue ?? r.Value ?? 0), orderNoOf: () => 0 },
        { category: 'Rest plader', rows: flatRestDetails, keyOf: r => String(r.ProdNo || '') + (r.OrdNo ? '/' + String(r.OrdNo) : ''), labelOf: r => r.Txt2 || r.Descr, valueOf: r => Number(r.Value || 0), orderNoOf: r => Number(r.OrdNo || 0) },
        { category: 'Stang materiale', rows: categories.stang || [], keyOf: r => String(r.ProdNo || ''), labelOf: r => r.Descr, valueOf: r => Number(r.Value || 0), orderNoOf: () => 0 },
        { category: 'Lager Komponenter', rows: categories.gr5Items || [], keyOf: r => String(r.ProdNo || ''), labelOf: r => r.Descr, valueOf: r => Number(r.FifoValue || 0), orderNoOf: () => 0 },
        { category: 'Opfølgningsvarer', rows: categories.opfolgningvare || [], keyOf: r => String(r.ProdNo || ''), labelOf: r => r.Descr, valueOf: r => Number(r.Value || 0), orderNoOf: () => 0 },
        { category: 'Plader VIA', rows: categories.nestingCutting || [], keyOf: r => String(r.OrdNo || '') + '/' + String(r.ProdNo || ''), labelOf: r => (r.Products ? String(r.Products) + ' · ' : '') + 'Rute ' + String(r.Route || '-'), valueOf: r => lagerlisteNestingCountedValue(r), orderNoOf: r => Number(r.SalesOrdNo || r.OrdNo || 0) },
        { category: 'Færdige SO', rows: categories.finishedNotInvoiced || [], keyOf: r => String(r.OrdNo || ''), labelOf: r => r.CustomerName, valueOf: r => Number(r.Value || 0), orderNoOf: r => Number(r.OrdNo || 0) },
        { category: 'VIA Tid', rows: categories.salgordreVia || [], keyOf: r => String(r.OrdNo || ''), labelOf: r => [r.MainProdNo, r.CustomerName].filter(Boolean).join(' · '), valueOf: r => Number(r.TimeCost || 0), orderNoOf: r => Number(r.OrdNo || 0) },
        { category: 'VIA Laser', rows: categories.salgordreVia || [], keyOf: r => String(r.OrdNo || ''), labelOf: r => [r.MainProdNo, r.CustomerName].filter(Boolean).join(' · '), valueOf: r => Number(r.MaterialCost || 0), orderNoOf: r => Number(r.OrdNo || 0) },
        { category: 'VIA Stang', rows: categories.salgordreVia || [], keyOf: r => String(r.OrdNo || ''), labelOf: r => [r.MainProdNo, r.CustomerName].filter(Boolean).join(' · '), valueOf: r => Number(r.StangCost || 0), orderNoOf: r => Number(r.OrdNo || 0) },
        { category: 'Indkøbt dele', rows: categories.salgordreVia || [], keyOf: r => String(r.OrdNo || ''), labelOf: r => [r.MainProdNo, r.CustomerName].filter(Boolean).join(' · '), valueOf: r => Number(r.PurchasedPartCost || 0), orderNoOf: r => Number(r.OrdNo || 0) }
    ];
}

function lagerlisteBuildMovements(payloadA, payloadB) {
    const specsA = lagerlisteMovementSpecs(payloadA);
    const specsB = lagerlisteMovementSpecs(payloadB);
    const categoriesA = (payloadA && payloadA.categories) || {};
    const categoriesB = (payloadB && payloadB.categories) || {};
    const orderSet = rows => new Set((rows || []).map(row => Number(row.OrdNo || 0)).filter(ordNo => ordNo > 0));
    const viaBefore = orderSet(categoriesA.salgordreVia);
    const viaAfter = orderSet(categoriesB.salgordreVia);
    const finishedBefore = orderSet(categoriesA.finishedNotInvoiced);
    const finishedAfter = orderSet(categoriesB.finishedNotInvoiced);
    const viaCategories = new Set(['VIA Tid', 'VIA Laser', 'VIA Stang', 'Indkøbt dele', 'Plader VIA']);
    const movements = [];
    for (let index = 0; index < specsA.length; index += 1) {
        const specA = specsA[index];
        const specB = specsB[index];
        const map = new Map();
        for (const row of specA.rows) {
            const key = specA.keyOf(row);
            if (!key) continue;
            const entry = map.get(key) || { key, label: specA.labelOf(row), orderNo: Number(specA.orderNoOf(row) || 0), valueA: 0, valueB: 0 };
            entry.valueA += specA.valueOf(row);
            if (!entry.label) entry.label = specA.labelOf(row);
            map.set(key, entry);
        }
        for (const row of specB.rows) {
            const key = specB.keyOf(row);
            if (!key) continue;
            const entry = map.get(key) || { key, label: specB.labelOf(row), orderNo: Number(specB.orderNoOf(row) || 0), valueA: 0, valueB: 0 };
            entry.valueB += specB.valueOf(row);
            if (!entry.label) entry.label = specB.labelOf(row);
            map.set(key, entry);
        }
        for (const entry of map.values()) {
            const diff = entry.valueB - entry.valueA;
            if (Math.abs(diff) < 0.005) continue;
            movements.push({
                Category: specA.category,
                Key: entry.key,
                Label: String(entry.label || '').trim(),
                OrderNo: entry.orderNo,
                ValueA: entry.valueA,
                ValueB: entry.valueB,
                Diff: diff,
                Status: entry.valueA === 0 ? 'Ny' : (entry.valueB === 0 ? 'Udgået' : 'Ændret')
            });
        }
    }
    const movementsByOrder = new Map();
    for (const movement of movements) {
        const ordNo = Number(movement.OrderNo || 0);
        if (ordNo <= 0) continue;
        const list = movementsByOrder.get(ordNo) || [];
        list.push(movement);
        movementsByOrder.set(ordNo, list);
    }
    const materialViaCategories = new Set(['VIA Laser', 'VIA Stang', 'Indkøbt dele']);
    for (const movement of movements) {
        const ordNo = Number(movement.OrderNo || 0);
        if (!Number.isFinite(ordNo) || ordNo <= 0) continue;
        const movedFromViaToFinished = viaBefore.has(ordNo) && !viaAfter.has(ordNo)
            && !finishedBefore.has(ordNo) && finishedAfter.has(ordNo);
        if (movedFromViaToFinished && viaCategories.has(movement.Category)) {
            movement.Status = 'Overført';
            movement.Flow = 'VIA → Færdige SO';
        } else if (movedFromViaToFinished && movement.Category === 'Færdige SO') {
            movement.Status = 'Overført';
            movement.Flow = 'VIA → Færdige SO';
        }
        if (!movement.Flow) {
            const orderMovements = movementsByOrder.get(ordNo) || [];
            const nestingOut = orderMovements.filter(row => row.Category === 'Plader VIA' && row.Diff < 0);
            const nestingIn = orderMovements.filter(row => row.Category === 'Plader VIA' && row.Diff > 0);
            const materialViaIn = orderMovements.filter(row => materialViaCategories.has(row.Category) && row.Diff > 0);
            const materialViaOut = orderMovements.filter(row => materialViaCategories.has(row.Category) && row.Diff < 0);
            if (movement.Category === 'Plader VIA' && nestingOut.length && materialViaIn.length) {
                movement.Status = 'Overført';
                movement.Flow = 'Plader VIA → ' + Array.from(new Set(materialViaIn.map(row => row.Category))).join(', ');
            } else if (materialViaCategories.has(movement.Category) && movement.Diff > 0 && nestingOut.length) {
                movement.Status = 'Overført';
                movement.Flow = 'Plader VIA → ' + movement.Category;
            } else if (movement.Category === 'Plader VIA' && nestingIn.length && materialViaOut.length) {
                movement.Status = 'Overført';
                movement.Flow = Array.from(new Set(materialViaOut.map(row => row.Category))).join(', ') + ' → Plader VIA';
            } else if (materialViaCategories.has(movement.Category) && movement.Diff < 0 && nestingIn.length) {
                movement.Status = 'Overført';
                movement.Flow = movement.Category + ' → Plader VIA';
            }
        }
        if (!movement.Flow && viaCategories.has(movement.Category) && viaBefore.has(ordNo) && viaAfter.has(ordNo)) {
            movement.Flow = 'Forbliver i VIA';
        }
    }
    movements.sort((left, right) => Math.abs(right.Diff) - Math.abs(left.Diff));
    return movements;
}

function lagerlisteBuildMaterialBalance(payloadA, payloadB) {
    const categoriesA = (payloadA && payloadA.categories) || {};
    const categoriesB = (payloadB && payloadB.categories) || {};
    const flatPlateRows = payload => (payload.plateGroups || []).flatMap(group => group.details || []);
    const compareRows = (rowsA, rowsB, keyOf, labelOf, valueOf, category, direction) => {
        const values = new Map();
        for (const row of rowsA || []) {
            const key = String(keyOf(row) || '');
            if (!key) continue;
            const entry = values.get(key) || { Key: key, Label: labelOf(row), ValueA: 0, ValueB: 0 };
            entry.ValueA += Number(valueOf(row) || 0);
            values.set(key, entry);
        }
        for (const row of rowsB || []) {
            const key = String(keyOf(row) || '');
            if (!key) continue;
            const entry = values.get(key) || { Key: key, Label: labelOf(row), ValueA: 0, ValueB: 0 };
            entry.ValueB += Number(valueOf(row) || 0);
            if (!entry.Label) entry.Label = labelOf(row);
            values.set(key, entry);
        }
        return Array.from(values.values()).flatMap(entry => {
            const diff = entry.ValueB - entry.ValueA;
            const relevant = direction === 'out' ? diff < -0.005 : diff > 0.005;
            return relevant ? [{
                Side: direction === 'out' ? 'Ud af lager (FIFO)' : 'Ind i VIA',
                Category: category,
                Key: entry.Key,
                Label: String(entry.Label || '').trim(),
                ValueA: entry.ValueA,
                ValueB: entry.ValueB,
                Amount: Math.abs(diff)
            }] : [];
        });
    };
    const stockOut = [
        ...compareRows(flatPlateRows(categoriesA), flatPlateRows(categoriesB), r => r.ProdNo, r => r.Descr, r => r.FifoValue ?? r.Value, 'Pladelager', 'out'),
        ...compareRows(categoriesA.stang, categoriesB.stang, r => r.ProdNo, r => r.Descr, r => r.FifoValue ?? r.Value, 'Stang materiale', 'out'),
        ...compareRows(categoriesA.gr5Items, categoriesB.gr5Items, r => r.ProdNo, r => r.Descr, r => r.FifoValue ?? r.Value, 'Lager Komponenter', 'out'),
        ...compareRows(categoriesA.opfolgningvare, categoriesB.opfolgningvare, r => r.ProdNo, r => r.Descr, r => r.Value, 'Opfølgningsvarer', 'out')
    ];
    const viaIn = [
        ...compareRows(categoriesA.nestingCutting, categoriesB.nestingCutting, r => String(r.OrdNo || '') + '/' + String(r.ProdNo || ''), r => (r.Products ? r.Products + ' · ' : '') + 'Rute ' + String(r.Route || '-'), lagerlisteNestingCountedValue, 'Plader VIA', 'in'),
        ...compareRows(categoriesA.salgordreVia, categoriesB.salgordreVia, r => r.OrdNo, r => [r.MainProdNo, r.CustomerName].filter(Boolean).join(' · '), r => r.MaterialCost, 'VIA Laser', 'in'),
        ...compareRows(categoriesA.salgordreVia, categoriesB.salgordreVia, r => r.OrdNo, r => [r.MainProdNo, r.CustomerName].filter(Boolean).join(' · '), r => r.StangCost, 'VIA Stang', 'in'),
        ...compareRows(categoriesA.salgordreVia, categoriesB.salgordreVia, r => r.OrdNo, r => [r.MainProdNo, r.CustomerName].filter(Boolean).join(' · '), r => r.PurchasedPartCost, 'Indkøbt dele', 'in')
    ];
    const fifoOut = stockOut.reduce((sum, row) => sum + row.Amount, 0);
    const viaMaterialIn = viaIn.reduce((sum, row) => sum + row.Amount, 0);
    const rows = stockOut.concat(viaIn).sort((left, right) => right.Amount - left.Amount);
    return { fifoOut, viaMaterialIn, difference: viaMaterialIn - fifoOut, rows };
}

async function lagerlisteResolvePeriod(key) {
    const value = String(key || '').trim();
    if (!value) throw new Error('Vælg to perioder');
    if (value === 'current') {
        if (lagerlisteCurrent) return { label: 'Aktuel', payload: lagerlisteCurrent };
        const response = await fetch('/lagerliste/current', { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        return { label: 'Aktuel', payload: data };
    }
    if (value.startsWith('month:')) {
        const month = value.slice(6);
        const response = await fetch('/lagerliste/snapshot/' + encodeURIComponent(month), { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        if (!data.current) throw new Error('Månedslukning ' + month + ' indeholder ingen lagerdata');
        return { label: 'Måned ' + month, payload: data.current };
    }
    if (value.startsWith('snap:')) {
        const snapshotId = value.slice(5);
        const response = await fetch('/lagerliste/snapshots/' + encodeURIComponent(snapshotId), { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        const payload = data.snapshot && data.snapshot.current;
        if (!payload) throw new Error('Snapshot ' + snapshotId + ' indeholder ingen lagerdata');
        return { label: 'Snapshot ' + snapshotId, payload };
    }
    throw new Error('Ukendt periode: ' + value);
}

async function refreshLagerlisteCompareOptions() {
    const selects = [document.getElementById('lagerlisteCompareA'), document.getElementById('lagerlisteCompareB')];
    if (!selects[0] || !selects[1]) return;
    const options = ['<option value="">Vælg periode...</option>', '<option value="current">Aktuel (live)</option>'];
    try {
        const response = await fetch('/lagerliste/snapshot-months', { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (response.ok && data.ok) {
            for (const month of data.months || []) {
                options.push('<option value="month:' + lagerlisteEscape(month) + '">Måned ' + lagerlisteEscape(month) + '</option>');
            }
        }
    } catch (_err) { /* months optional */ }
    for (const row of lagerlisteSnapshotRows || []) {
        const id = String(row.snapshotId || '');
        options.push('<option value="snap:' + lagerlisteEscape(id) + '">Snapshot ' + lagerlisteEscape(id) + '</option>');
    }
    for (const select of selects) {
        const previous = select.value;
        select.innerHTML = options.join('');
        if (previous && Array.from(select.options).some(option => option.value === previous)) select.value = previous;
    }
}

function lagerlisteCompareClear() {
    const root = document.getElementById('lagerlisteCompareResults');
    if (root) root.innerHTML = '';
}

function lagerlisteFilterMovementCategory() {
    const select = document.getElementById('lagerlisteMovementCategory');
    const root = document.getElementById('lagerlisteCompareResults');
    if (!select || !root) return;
    const category = String(select.value || '');
    const table = root.querySelector('#lagerlisteMovementTable');
    if (!table) return;
    table.querySelectorAll(':scope > tbody > tr').forEach(row => {
        const rowCategory = String(row.dataset.category || '');
        row.style.display = !category || rowCategory === category ? '' : 'none';
    });
}

async function lagerlisteComparePeriods() {
    const root = document.getElementById('lagerlisteCompareResults');
    const keyA = document.getElementById('lagerlisteCompareA') && document.getElementById('lagerlisteCompareA').value;
    const keyB = document.getElementById('lagerlisteCompareB') && document.getElementById('lagerlisteCompareB').value;
    if (!root) return;
    if (!keyA || !keyB) {
        root.innerHTML = '<div class="omsaetning-empty">Vælg to perioder for at sammenligne.</div>';
        return;
    }
    if (keyA === keyB) {
        root.innerHTML = '<div class="omsaetning-empty">Vælg to forskellige perioder.</div>';
        return;
    }
    root.innerHTML = '<div class="loading">Henter perioder...</div>';
    try {
        const [periodA, periodB] = await Promise.all([lagerlisteResolvePeriod(keyA), lagerlisteResolvePeriod(keyB)]);
        const figuresA = lagerlisteComputeFigures(periodA.payload);
        const figuresB = lagerlisteComputeFigures(periodB.payload);
        const movements = lagerlisteBuildMovements(periodA.payload, periodB.payload);
        const materialBalance = lagerlisteBuildMaterialBalance(periodA.payload, periodB.payload);
        const movementsByCategory = new Map();
        for (const movement of movements) {
            const list = movementsByCategory.get(movement.Category) || [];
            list.push(movement);
            movementsByCategory.set(movement.Category, list);
        }
        const categoryTooltip = (categoryNames) => {
            const lines = [];
            let totalIn = 0;
            let totalOut = 0;
            let countIn = 0;
            let countOut = 0;
            const top = [];
            for (const name of categoryNames) {
                for (const movement of movementsByCategory.get(name) || []) {
                    if (movement.Diff > 0) { totalIn += movement.Diff; countIn += 1; } else { totalOut += movement.Diff; countOut += 1; }
                    top.push(movement);
                }
            }
            if (!top.length) return 'Ingen bevægelser i denne kategori.';
            lines.push('Ind: +' + lagerlisteFormat(totalIn) + ' (' + countIn + ' poster) · Ud: ' + lagerlisteFormat(totalOut) + ' (' + countOut + ' poster)');
            top.sort((left, right) => Math.abs(right.Diff) - Math.abs(left.Diff));
            lines.push('Største bevægelser:');
            for (const movement of top.slice(0, 8)) {
                const sign = movement.Diff > 0 ? '+' : '';
                const label = String(movement.Label || '').trim();
                lines.push(sign + lagerlisteFormat(movement.Diff) + ' · ' + movement.Key + (label ? ' (' + label + ')' : '') + ' [' + movement.Status + ']');
            }
            if (top.length > 8) lines.push('... og ' + (top.length - 8) + ' flere (se Bevægelser-tabellen)');
            return lines.join('\n');
        };
        const aggregateTooltip = (categoryNames) => {
            const lines = ['Sammensat af:'];
            for (const name of categoryNames) {
                const list = movementsByCategory.get(name) || [];
                const diff = list.reduce((sum, movement) => sum + movement.Diff, 0);
                lines.push((diff > 0 ? '+' : '') + lagerlisteFormat(diff) + ' · ' + name + ' (' + list.length + ' bevægelser)');
            }
            return lines.join('\n');
        };
        const stockCategories = ['Pladelager', 'Stang materiale', 'Opfølgningsvarer', 'Lager Komponenter'];
        const wipCategories = ['Færdige SO', 'VIA Tid', 'VIA Laser', 'VIA Stang', 'Indkøbt dele', 'Plader VIA'];
        const tooltipByLabel = {
            'Pladelager': () => categoryTooltip(['Pladelager']),
            'Rest plader': () => categoryTooltip(['Rest plader']),
            'Stang materiale': () => categoryTooltip(['Stang materiale']),
            'Opfølgningsvarer': () => categoryTooltip(['Opfølgningsvarer']),
            'Lager Komponenter (FIFO)': () => categoryTooltip(['Lager Komponenter']),
            'Varelager uden rest': () => aggregateTooltip(stockCategories),
            'Varelager': () => aggregateTooltip(stockCategories.concat(['Rest plader'])),
            'Færdige SO kostpris': () => categoryTooltip(['Færdige SO']),
            'VIA Tid': () => categoryTooltip(['VIA Tid']),
            'VIA Laser': () => categoryTooltip(['VIA Laser']),
            'VIA Stang': () => categoryTooltip(['VIA Stang']),
            'Indkøbt dele': () => categoryTooltip(['Indkøbt dele']),
            'VIA Plader (Værdi i skæring)': () => categoryTooltip(['Plader VIA']),
            'Vare i arbejde': () => aggregateTooltip(wipCategories),
            'TOTAL': () => aggregateTooltip(stockCategories.concat(['Rest plader']).concat(wipCategories))
        };
        const categoryRows = figuresA.map((row, index) => {
            const valueA = row[1];
            const valueB = figuresB[index][1];
            const diff = valueB - valueA;
            const cls = diff > 0 ? 'lagerliste-diff-pos' : (diff < 0 ? 'lagerliste-diff-neg' : 'lagerliste-diff-zero');
            const tooltipBuilder = tooltipByLabel[row[0]];
            const tooltip = tooltipBuilder ? tooltipBuilder() : '';
            return '<tr class="' + (row[2] || '') + '"><td>' + lagerlisteEscape(row[0]) + '</td>'
                + '<td>' + lagerlisteEscape(lagerlisteFormat(valueA)) + '</td>'
                + '<td>' + lagerlisteEscape(lagerlisteFormat(valueB)) + '</td>'
                + '<td><span class="' + cls + (tooltip ? ' lagerliste-diff-tooltip' : '') + '"' + (tooltip ? ' title="' + lagerlisteEscape(tooltip) + '"' : '') + '>' + lagerlisteEscape(lagerlisteFormat(diff)) + '</span></td></tr>';
        }).join('');
        const movementsTable = lagerlisteRowsTable(movements, [
            { key: 'Category', label: 'Kategori' },
            { key: 'Key', label: 'Produkt/Ordre' },
            { key: 'Label', label: 'Beskrivelse', format: value => String(value || '').trim() || '-' },
            { key: 'Status', label: 'Status', allowHtml: true, format: value => {
                const cls = value === 'Ny' ? 'lagerliste-diff-pos' : (value === 'Udgået' ? 'lagerliste-diff-neg' : (value === 'Overført' ? 'lagerliste-diff-pos' : 'lagerliste-diff-zero'));
                return '<span class="' + cls + '">' + lagerlisteEscape(value) + '</span>';
            } },
            { key: 'Flow', label: 'Fra → Til', format: value => String(value || '').trim() || '-' },
            { key: 'ValueA', label: periodA.label, format: lagerlisteFormat },
            { key: 'ValueB', label: periodB.label, format: lagerlisteFormat },
            { key: 'Diff', label: 'Bevægelse', allowHtml: true, format: value => {
                const diff = Number(value || 0);
                const cls = diff > 0 ? 'lagerliste-diff-pos' : (diff < 0 ? 'lagerliste-diff-neg' : 'lagerliste-diff-zero');
                return '<span class="' + cls + '">' + lagerlisteEscape(lagerlisteFormat(diff)) + '</span>';
            } }
        ], row => ' data-category="' + lagerlisteEscape(row.Category) + '"');
        const movementCategories = Array.from(new Set(movements.map(row => String(row.Category || '')).filter(Boolean))).sort((left, right) => left.localeCompare(right));
        const movementFilter = '<div class="lagerliste-movement-filter"><label for="lagerlisteMovementCategory">Kategori</label><select id="lagerlisteMovementCategory" class="filter-select" onchange="lagerlisteFilterMovementCategory()"><option value="">Alle kategorier</option>'
            + movementCategories.map(category => '<option value="' + lagerlisteEscape(category) + '">' + lagerlisteEscape(category) + '</option>').join('')
            + '</select></div>';
        const materialRowsTable = lagerlisteRowsTable(materialBalance.rows, [
            { key: 'Side', label: 'Bilagsside' },
            { key: 'Category', label: 'Kategori' },
            { key: 'Key', label: 'Produkt/Ordre' },
            { key: 'Label', label: 'Beskrivelse', format: value => String(value || '').trim() || '-' },
            { key: 'ValueA', label: periodA.label, format: lagerlisteFormat },
            { key: 'ValueB', label: periodB.label, format: lagerlisteFormat },
            { key: 'Amount', label: 'Beløb', format: lagerlisteFormat }
        ]);
        const materialDiffClass = materialBalance.difference > 0 ? 'lagerliste-diff-pos' : (materialBalance.difference < 0 ? 'lagerliste-diff-neg' : 'lagerliste-diff-zero');
        const totalIn = movements.reduce((sum, row) => sum + (row.Diff > 0 ? row.Diff : 0), 0);
        const totalOut = movements.reduce((sum, row) => sum + (row.Diff < 0 ? row.Diff : 0), 0);
        root.innerHTML = '<div class="lagerliste-vareopslag-panel">'
            + '<div class="lagerliste-vareopslag-head"><h4>Sammenligning: ' + lagerlisteEscape(periodA.label) + ' → ' + lagerlisteEscape(periodB.label) + '</h4>'
            + '<button type="button" class="lagerliste-vareopslag-close" onclick="lagerlisteCompareClear()">Luk</button></div>'
            + '<h5>Kategorier</h5>'
            + '<div class="omsaetning-table-wrap"><table class="lagerliste-sheet-table lagerliste-overview-table"><thead><tr><th>Post</th><th>' + lagerlisteEscape(periodA.label) + '</th><th>' + lagerlisteEscape(periodB.label) + '</th><th>Ændring</th></tr></thead><tbody>' + categoryRows + '</tbody></table></div>'
            + '<h5>Materialebalance (FIFO)</h5>'
            + '<div class="lagerliste-total-row"><strong>Ud af lager (FIFO)</strong><strong>' + lagerlisteEscape(lagerlisteFormat(materialBalance.fifoOut)) + '</strong><strong>Ind i VIA (uden tid)</strong><strong>' + lagerlisteEscape(lagerlisteFormat(materialBalance.viaMaterialIn)) + '</strong><strong>Difference (VIA − FIFO)</strong><strong><span class="' + materialDiffClass + '">' + lagerlisteEscape(lagerlisteFormat(materialBalance.difference)) + '</span></strong></div>'
            + '<div class="lagerliste-material-note">Lagerindgange og VIA Tid er ikke medregnet. Pladelager, stang, komponenter og opfølgningsvarer værdisættes til FIFO; VIA følger den gældende VIA-kostregel.</div>'
            + materialRowsTable
            + '<h5>Bevægelser (' + movements.length + ')</h5>'
            + movementFilter
            + movementsTable.replace('<table class="order-list-table">', '<table id="lagerlisteMovementTable" class="order-list-table">')
            + '<div class="lagerliste-total-row"><strong>Ind (tilgang)</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalIn)) + '</strong><strong>Ud (afgang)</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalOut)) + '</strong><strong>Netto</strong><strong>' + lagerlisteEscape(lagerlisteFormat(totalIn + totalOut)) + '</strong></div>'
            + '</div>';
        lagerlisteEnhanceTables();
    } catch (err) {
        root.innerHTML = '<div class="error">Sammenligning fejlede: ' + lagerlisteEscape(err.message || err) + '</div>';
    }
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

async function deleteSelectedLagerlisteSnapshot() {
    const select = document.getElementById('lagerlisteSnapshotSelect');
    const snapshotId = String(select && select.value || '').trim();
    if (!snapshotId) {
        alert('Vælg et dags-snapshot først.');
        return;
    }
    if (!confirm('Slet dags-snapshot ' + snapshotId + '? Denne handling kan ikke fortrydes.')) return;
    const status = document.getElementById('lagerlisteSnapshotStatus');
    try {
        if (status) status.textContent = 'Sletter snapshot ' + snapshotId + '...';
        const response = await fetch('/lagerliste/snapshots/' + encodeURIComponent(snapshotId), {
            method: 'DELETE',
            headers: { Authorization: 'Bearer ' + String(authToken || '') }
        });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        if (status) status.textContent = 'Snapshot slettet: ' + snapshotId;
        await refreshLagerlisteSnapshotList();
        await refreshLagerlisteCompareOptions();
    } catch (err) {
        if (status) status.textContent = 'Fejl: ' + String(err.message || err);
    }
}
