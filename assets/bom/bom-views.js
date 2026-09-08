// ── BOM Workspace · views ────────────────────────────────────────────
// Navigation, stykliste, komponenter, ressourcer, materialer,
// parametre, leverandører og lokale produktkladder.

function renderNav() {
    navList.innerHTML = visibleNavItems().map((item, idx) => {
        const active = item.key === state.view ? 'active' : '';
        return '<button class="nav-btn ' + active + '" data-view="' + item.key + '" title="Genvej: Alt+' + (idx + 1) + '"><kbd class="nav-kbd">Alt+' + (idx + 1) + '</kbd><strong>' + escapeHtml(item.title) + '</strong><span>' + escapeHtml(item.description) + '</span></button>';
    }).join('');
    navList.querySelectorAll('.nav-btn').forEach(btn => btn.addEventListener('click', () => switchView(btn.getAttribute('data-view'))));
}
function switchView(view) {
    if (!visibleNavItems().some(item => item.key === view)) return;
    state.view = view;
    renderNav();
    Object.keys(viewMeta).forEach(key => {
        const el = document.getElementById('view-' + key);
        if (el) el.classList.toggle('active', key === view);
    });
    viewTitle.textContent = viewMeta[view].title;
    viewSubtitle.textContent = viewMeta[view].subtitle;
    if (view === 'resources' && state.resources.length === 0) loadResources();
    if (view === 'materials' && state.materials.length === 0) loadMaterials();
    if (view === 'calculators') loadCalculators();
    if (view === 'komponenter' && state.components.length === 0) loadComponents();
    if (view === 'leverandorer' && state.suppliers.length === 0) loadSuppliers();
    if (view === 'beregner') primeBeregner();
}
function updateContext() {
    document.getElementById('ctxCustomer').textContent = state.selectedCustomer ? (state.selectedCustomer.CustNo + ' - ' + (state.selectedCustomer.Nm || '')) : 'Ikke valgt';
    document.getElementById('ctxCustomerCode').textContent = state.selectedCustomer ? (state.selectedCustomer.Gr || state.selectedCustomer['Varenr.'] || '-') : 'Ikke valgt';
    document.getElementById('ctxProduct').textContent = state.selectedProduct ? (state.selectedProduct.ProdNo || '-') : 'Ikke valgt';
    document.getElementById('ctxTgn').textContent = state.selectedProduct ? (state.selectedProduct.TgNo || tgnInput.value.trim() || '-') : (tgnInput.value.trim() || 'Ikke valgt');
}
function renderSimpleTable(headEl, bodyEl, rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    const columns = safeRows.length ? Object.keys(safeRows[0]) : [];
    headEl.innerHTML = columns.length ? '<tr>' + columns.map(col => '<th>' + escapeHtml(col) + '</th>').join('') + '</tr>' : '<tr><th>Ingen data</th></tr>';
    bodyEl.classList.add('reveal-stagger');
    bodyEl.innerHTML = safeRows.length ? safeRows.map(row => '<tr>' + columns.map(col => '<td>' + escapeHtml(row[col]) + '</td>').join('') + '</tr>').join('') : '<tr><td class="empty">Ingen rækker fundet.</td></tr>';
}
function currentProductRows() {
    const filter = String(productSearchInput.value || '').trim().toLowerCase();
    const rows = state.products;
    if (!filter) return rows;
    return rows.filter(row => [row.ProdNo, row.Descr, row.TgNo, row.RevNo].some(v => String(v || '').toLowerCase().includes(filter)));
}
function renderCustomers() {
    const filter = String(customerSearchInput.value || '').trim().toLowerCase();
    const rows = filter ? state.customers.filter(row => [row.CustNo, row.Nm, row.Shrt, row.PArea, row.Gr].some(v => String(v || '').toLowerCase().includes(filter))) : state.customers;
    customersMeta.textContent = rows.length + ' kunder';
    customerSelect.innerHTML = rows.map(row => {
        const selected = state.selectedCustomer && String(state.selectedCustomer.CustNo) === String(row.CustNo) ? ' selected' : '';
        return '<option value="' + escapeHtml(row.CustNo) + '"' + selected + '>' + escapeHtml((row.CustNo || '') + ' · ' + (row.Gr || '') + ' · ' + (row.Nm || '')) + '</option>';
    }).join('');
    customersList.classList.toggle('reveal-stagger', !!customersList.querySelector('.skl'));
    customersList.innerHTML = rows.length ? rows.map(row => {
        const active = state.selectedCustomer && String(state.selectedCustomer.CustNo) === String(row.CustNo) ? ' active' : '';
        return '<div class="list-item' + active + '" data-customer="' + escapeHtml(row.CustNo) + '" tabindex="0" role="button"><strong>' + escapeHtml((row.CustNo || '') + ' · ' + (row.Gr || '') + ' · ' + (row.Nm || '')) + '</strong><span>' + escapeHtml((row.Shrt || '') + ' · ' + (row.PArea || '')) + '</span></div>';
    }).join('') : '<div class="empty">Ingen kunder fundet.</div>';
    customersList.querySelectorAll('[data-customer]').forEach(el => {
        el.addEventListener('click', () => {
            const customerNo = el.getAttribute('data-customer');
            state.selectedCustomer = state.customers.find(row => String(row.CustNo) === String(customerNo)) || null;
            state.selectedProduct = null;
            state.selectedRevision = null;
            customerSelect.value = customerNo;
            updateContext();
            renderCustomers();
            loadProducts();
            loadCustomerNotes();
        });
    });
}
function renderProducts() {
    const filtered = currentProductRows();
    productsMeta.textContent = state.selectedCustomer ? ('Kundenøgler ' + state.selectedCustomer.CustNo + ' / ' + state.selectedCustomer.Gr + ' · ' + filtered.length + ' af ' + state.products.length + ' produkter') : 'Vælg kunde';
    productsList.classList.toggle('reveal-stagger', !!productsList.querySelector('.skl'));
    productsList.innerHTML = filtered.length ? filtered.map(row => {
        const active = state.selectedProduct && String(state.selectedProduct.ProdNo) === String(row.ProdNo) ? ' active' : '';
        const draftClass = row.IsLocalDraft ? ' draft' : '';
        const draftTag = row.IsLocalDraft ? '<span class="tag">LOKAL</span>' : '';
        return '<div class="list-item' + active + draftClass + '" data-product="' + escapeHtml(row.ProdNo) + '" tabindex="0" role="button"><strong>' + escapeHtml((row.ProdNo || '-') + ' · ' + (row.Descr || '')) + draftTag + '</strong><span>' + escapeHtml('TgNo: ' + (row.TgNo || '-') + ' · Rev: ' + (row.RevNo || '-')) + '</span></div>';
    }).join('') : '<div class="empty">Ingen produkter matcher søgningen.</div>';
    productsList.querySelectorAll('[data-product]').forEach(el => {
        el.addEventListener('click', () => {
            const prodNo = el.getAttribute('data-product');
            state.selectedProduct = state.products.find(row => String(row.ProdNo) === String(prodNo)) || null;
            state.selectedRevision = null;
            if (state.selectedProduct && state.selectedProduct.TgNo) tgnInput.value = state.selectedProduct.TgNo;
            updateContext();
            renderProducts();
            renderProductDetail();
            loadRevisions();
            loadProductTree();
        });
    });
}
function renderProductDetail() {
    const product = state.selectedProduct;
    const rows = product ? [
        ['ProdNo', product.ProdNo],
        ['Beskrivelse', product.Descr],
        ['TgNo', product.TgNo],
        ['Revision', product.RevNo],
        ['PosNo', product.PosNo],
        ['Inf3', product.Inf3],
        ['Inf4', product.Inf4],
        ['chck', product.chck]
    ] : [];
    productDetailGrid.innerHTML = rows.length ? rows.map(([label, value]) => '<div class="kv"><label>' + escapeHtml(label) + '</label><div>' + escapeHtml(value || '-') + '</div></div>').join('') : '<div class="empty" style="grid-column:1 / -1;">Vælg et produkt for at se detaljer.</div>';
}
function renderRevisions() {
    revisionsMeta.textContent = state.selectedProduct ? ((state.selectedProduct.ProdNo || '-') + ' · TgNo ' + (tgnInput.value.trim() || '-')) : 'Ingen valgt';
    revisionsCountMeta.textContent = state.revisions.length + ' rækker';
    renderSimpleTable(revisionsHead, revisionsBody, state.revisions);
}
function renderResources(rows) { renderSimpleTable(resourcesHead, resourcesBody, rows); }
function renderMaterials() { renderSimpleTable(materialsHead, materialsBody, state.materials); }
let editingLaserRow = null;
let editingLaserTechnicalRow = null;
let editingBendingMachineRow = null;
let editingBendingBandRow = null;

function parameterCell(row, field, editing, attributeName) {
    const value = row[field[0]] == null ? '' : row[field[0]];
    if (!editing) return '<td>' + escapeHtml(value === '' ? '-' : value) + '</td>';
    return '<td><input ' + attributeName + '="' + field[0] + '" type="' + field[2] + '" step="any" value="' + escapeHtml(value) + '" /></td>';
}

function renderLaserParameters(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    const fields = [
        ['ProdNo', 'Varenr.', 'text'], ['Descr', 'Beskrivelse', 'text'], ['Tykkelse', 'Tykkelse', 'number'],
        ['Maskine', 'Maskine', 'text'], ['Skærehast.', 'm/min', 'number'], ['Pircing', 'Piercing min.', 'number'],
        ['Tillæg', 'Tillæg %', 'number'], ['Linse', 'Linse', 'text']
    ];
    laserHead.innerHTML = '<tr>' + fields.map(field => '<th>' + field[1] + '</th>').join('') + '<th>Kilde</th><th></th></tr>';
    laserBody.innerHTML = safeRows.length ? safeRows.map((row, rowIndex) => {
        const editing = editingLaserRow === rowIndex;
        const cells = fields.map(field => parameterCell(row, field, editing, 'data-laser-field')).join('');
        const actions = editing
            ? '<button type="button" data-save-laser>Gem</button> <button type="button" class="alt" data-cancel-laser>Annuller</button>'
            : '<button type="button" class="alt" data-edit-laser>Rediger</button>';
        return '<tr data-laser-row="' + rowIndex + '">' + cells + '<td>' + escapeHtml(row.Source || '-') + '</td><td class="parameter-actions">' + actions + '</td></tr>';
    }).join('') : '<tr><td colspan="10" class="empty">Ingen rækker fundet. Opret en ny laserparameter.</td></tr>';
    laserBody.querySelectorAll('[data-save-laser]').forEach(button => button.addEventListener('click', () => saveLaserParameter(button)));
    laserBody.querySelectorAll('[data-edit-laser]').forEach(button => button.addEventListener('click', () => {
        editingLaserRow = Number(button.closest('tr').dataset.laserRow);
        renderLaserParameters(state.laserParams);
    }));
    laserBody.querySelectorAll('[data-cancel-laser]').forEach(button => button.addEventListener('click', () => {
        const index = Number(button.closest('tr').dataset.laserRow);
        if (state.laserParams[index] && state.laserParams[index]._isNew) state.laserParams.splice(index, 1);
        editingLaserRow = null;
        renderLaserParameters(state.laserParams);
    }));
}

function renderLaserTechnicalParameters(rows) {
    const head = document.getElementById('laserTechnicalHead');
    const body = document.getElementById('laserTechnicalBody');
    const safeRows = Array.isArray(rows) ? rows : [];
    const fields = [
        ['Technology', 'Teknologi', 'text'], ['Material', 'Materiale', 'text'], ['Thickness', 'mm', 'number'],
        ['Lens', 'Linse', 'text'], ['PiercingMilliseconds', 'Piercing ms', 'number'],
        ['VaporPowerW', 'Vapor W', 'number'], ['ReducedPowerW', 'Reduceret W', 'number'],
        ['FeedrateLargeMmMin', 'Stor mm/min', 'number'], ['FeedrateMediumMmMin', 'Mellem mm/min', 'number'],
        ['FeedrateSmallMmMin', 'Lille mm/min', 'number'], ['FeedrateEngravingMmMin', 'Gravering mm/min', 'number'],
        ['GasPressureBar', 'Gastryk bar', 'number'], ['NozzleSizeMm', 'Dyse mm', 'number']
    ];
    head.innerHTML = '<tr>' + fields.map(field => '<th>' + field[1] + '</th>').join('') + '<th>Kilde</th><th></th></tr>';
    body.innerHTML = safeRows.length ? safeRows.map((row, index) => {
        const editing = editingLaserTechnicalRow === index;
        const actions = editing
            ? '<button type="button" data-save-laser-technical>Gem</button> <button type="button" class="alt" data-cancel-laser-technical>Annuller</button>'
            : '<button type="button" class="alt" data-edit-laser-technical>Rediger</button>';
        return '<tr data-laser-technical-row="' + index + '">'
            + fields.map(field => parameterCell(row, field, editing, 'data-laser-technical-field')).join('')
            + '<td>' + escapeHtml(row.Source || '-') + '</td><td class="parameter-actions">' + actions + '</td></tr>';
    }).join('') : '<tr><td colspan="15" class="empty">Ingen laserteknologier konfigureret. Kør Excel-importen.</td></tr>';
    body.querySelectorAll('[data-save-laser-technical]').forEach(button => button.addEventListener('click', () => saveLaserTechnicalParameter(button)));
    body.querySelectorAll('[data-edit-laser-technical]').forEach(button => button.addEventListener('click', () => {
        editingLaserTechnicalRow = Number(button.closest('tr').dataset.laserTechnicalRow);
        renderLaserTechnicalParameters(state.laserTechnicalParams);
    }));
    body.querySelectorAll('[data-cancel-laser-technical]').forEach(button => button.addEventListener('click', () => {
        const index = Number(button.closest('tr').dataset.laserTechnicalRow);
        if (state.laserTechnicalParams[index] && state.laserTechnicalParams[index]._isNew) state.laserTechnicalParams.splice(index, 1);
        editingLaserTechnicalRow = null;
        renderLaserTechnicalParameters(state.laserTechnicalParams);
    }));
}
function renderBendingMachines(rows) {
    const head = document.getElementById('bendingMachineHead');
    const body = document.getElementById('bendingMachineBody');
    const fields = [
        ['MachineCode', 'Maskine', 'text'], ['Description', 'Beskrivelse', 'text'],
        ['MaxBendLengthMm', 'Maks. længde mm', 'number'], ['MaxForceKn', 'Maks. kN', 'number'],
        ['BaseCycleSeconds', 'Basis sek./buk', 'number'], ['SecondsPerDegree', 'Sek./grad', 'number'],
        ['BackGaugeSeconds', 'Bagstop sek.', 'number'], ['SetupMinutes', 'Opstart min.', 'number'],
        ['SafetyFactor', 'Sikkerhed', 'number']
    ];
    head.innerHTML = '<tr>' + fields.map(field => '<th>' + field[1] + '</th>').join('') + '<th></th></tr>';
    body.innerHTML = rows.length ? rows.map((row, index) => {
        const editing = editingBendingMachineRow === index;
        const actions = editing
            ? '<button type="button" data-save-bending-machine>Gem</button> <button type="button" class="alt" data-cancel-bending-machine>Annuller</button>'
            : '<button type="button" class="alt" data-edit-bending-machine>Rediger</button>';
        return '<tr data-bending-machine-row="' + index + '">'
            + fields.map(field => parameterCell(row, field, editing, 'data-bending-machine-field')).join('')
            + '<td class="parameter-actions">' + actions + '</td></tr>';
    }).join('')
        : '<tr><td colspan="10" class="empty">Ingen buk-maskiner konfigureret.</td></tr>';
    body.querySelectorAll('[data-save-bending-machine]').forEach(button => button.addEventListener('click', () => saveBendingMachine(button)));
    body.querySelectorAll('[data-edit-bending-machine]').forEach(button => button.addEventListener('click', () => {
        editingBendingMachineRow = Number(button.closest('tr').dataset.bendingMachineRow);
        renderBendingMachines(state.bendingMachines);
    }));
    body.querySelectorAll('[data-cancel-bending-machine]').forEach(button => button.addEventListener('click', () => {
        const index = Number(button.closest('tr').dataset.bendingMachineRow);
        if (state.bendingMachines[index] && state.bendingMachines[index]._isNew) state.bendingMachines.splice(index, 1);
        editingBendingMachineRow = null;
        renderBendingMachines(state.bendingMachines);
    }));
}

function renderBendingBands(rows) {
    const head = document.getElementById('bendingBandHead');
    const body = document.getElementById('bendingBandBody');
    const fields = [
        ['BandName', 'Klasse', 'text'], ['MinWeightKg', 'Min. kg', 'number'], ['MaxWeightKg', 'Maks. kg', 'number'],
        ['MinLongestSideMm', 'Min. side mm', 'number'], ['MaxLongestSideMm', 'Maks. side mm', 'number'],
        ['LoadSeconds', 'Læg på sek.', 'number'], ['UnloadSeconds', 'Tag af sek.', 'number'],
        ['Rotate90Seconds', 'Rotér 90° sek.', 'number'], ['FlipSeconds', 'Vend sek.', 'number']
    ];
    head.innerHTML = '<tr>' + fields.map(field => '<th>' + field[1] + '</th>').join('') + '<th></th></tr>';
    body.innerHTML = rows.length ? rows.map((row, index) => {
        const editing = editingBendingBandRow === index;
        const actions = editing
            ? '<button type="button" data-save-bending-band>Gem</button> <button type="button" class="alt" data-cancel-bending-band>Annuller</button>'
            : '<button type="button" class="alt" data-edit-bending-band>Rediger</button>';
        return '<tr data-bending-band-row="' + index + '" data-band-id="' + escapeHtml(row.HandlingBandId || '') + '">'
            + fields.map(field => parameterCell(row, field, editing, 'data-bending-band-field')).join('')
            + '<td class="parameter-actions">' + actions + '</td></tr>';
    }).join('')
        : '<tr><td colspan="10" class="empty">Ingen håndteringsklasser konfigureret.</td></tr>';
    body.querySelectorAll('[data-save-bending-band]').forEach(button => button.addEventListener('click', () => saveBendingBand(button)));
    body.querySelectorAll('[data-edit-bending-band]').forEach(button => button.addEventListener('click', () => {
        editingBendingBandRow = Number(button.closest('tr').dataset.bendingBandRow);
        renderBendingBands(state.bendingHandlingBands);
    }));
    body.querySelectorAll('[data-cancel-bending-band]').forEach(button => button.addEventListener('click', () => {
        const index = Number(button.closest('tr').dataset.bendingBandRow);
        if (state.bendingHandlingBands[index] && state.bendingHandlingBands[index]._isNew) state.bendingHandlingBands.splice(index, 1);
        editingBendingBandRow = null;
        renderBendingBands(state.bendingHandlingBands);
    }));
}

function renderCalculators(laserRows, laserTechnicalRows, gasPrices, bendingData, processRows, processResourceRows) {
    editingLaserRow = null;
    editingLaserTechnicalRow = null;
    editingBendingMachineRow = null;
    editingBendingBandRow = null;
    state.laserParams = laserRows;
    state.laserTechnicalParams = laserTechnicalRows;
    state.laserGasPrices = gasPrices || { nitrogenPricePerKg: 0, oxygenPricePerKg: 0,
        nitrogenSpecificVolumeM3Kg: 0.862, oxygenSpecificVolumeM3Kg: 0.7, mixLineOxygenPercent: 22 };
    document.getElementById('nitrogenPriceInput').value = state.laserGasPrices.nitrogenPricePerKg || 0;
    document.getElementById('oxygenPriceInput').value = state.laserGasPrices.oxygenPricePerKg || 0;
    document.getElementById('nitrogenSpecificVolumeInput').value = state.laserGasPrices.nitrogenSpecificVolumeM3Kg || 0.862;
    document.getElementById('oxygenSpecificVolumeInput').value = state.laserGasPrices.oxygenSpecificVolumeM3Kg || 0.7;
    document.getElementById('mixLineOxygenPercentInput').value = state.laserGasPrices.mixLineOxygenPercent == null ? 22 : state.laserGasPrices.mixLineOxygenPercent;
    state.bendingMachines = bendingData.machines || [];
    state.bendingHandlingBands = bendingData.handlingBands || [];
    renderLaserParameters(laserRows);
    renderLaserTechnicalParameters(laserTechnicalRows);
    if (typeof refreshLaserTechnologyOptions === 'function') refreshLaserTechnologyOptions();
    renderBendingMachines(state.bendingMachines);
    renderBendingBands(state.bendingHandlingBands);
    renderSimpleTable(processHead, processBody, processRows);
    renderSimpleTable(processResourceHead, processResourceBody, processResourceRows);
}

async function saveLaserGasPrices() {
    const button = document.getElementById('saveLaserGasPricesBtn');
    button.disabled = true;
    try {
        const result = await fetchJson('/bom/calculators/laser-gas-prices', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                nitrogenPricePerKg: document.getElementById('nitrogenPriceInput').value,
                oxygenPricePerKg: document.getElementById('oxygenPriceInput').value,
                nitrogenSpecificVolumeM3Kg: document.getElementById('nitrogenSpecificVolumeInput').value,
                oxygenSpecificVolumeM3Kg: document.getElementById('oxygenSpecificVolumeInput').value,
                mixLineOxygenPercent: document.getElementById('mixLineOxygenPercentInput').value
            })
        });
        state.laserGasPrices = result.prices;
        if (typeof refreshLaserTechnologyOptions === 'function') refreshLaserTechnologyOptions();
        setStatus('Globale gaspriser gemt i GOH');
        showToast('Gaspriser gemt', 'ok');
    } finally { button.disabled = false; }
}

async function saveBendingMachine(button) {
    const row = button.closest('tr');
    const value = field => row.querySelector('[data-bending-machine-field="' + field + '"]').value.trim();
    button.disabled = true;
    try {
        await fetchJson('/bom/calculators/bending-machines', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ machineCode: value('MachineCode'), description: value('Description'),
                maxBendLengthMm: value('MaxBendLengthMm'), maxForceKn: value('MaxForceKn'), baseCycleSeconds: value('BaseCycleSeconds'),
                secondsPerDegree: value('SecondsPerDegree'), backGaugeSeconds: value('BackGaugeSeconds'),
                setupMinutes: value('SetupMinutes'), safetyFactor: value('SafetyFactor') })
        });
        setStatus('Buk-maskine gemt i GOH');
        await loadCalculators();
    } finally { button.disabled = false; }
}

async function saveBendingBand(button) {
    const row = button.closest('tr');
    const value = field => row.querySelector('[data-bending-band-field="' + field + '"]').value.trim();
    button.disabled = true;
    try {
        await fetchJson('/bom/calculators/bending-handling-bands', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ handlingBandId: row.dataset.bandId || 0, bandName: value('BandName'),
                minWeightKg: value('MinWeightKg'), maxWeightKg: value('MaxWeightKg'),
                minLongestSideMm: value('MinLongestSideMm'), maxLongestSideMm: value('MaxLongestSideMm'),
                loadSeconds: value('LoadSeconds'), unloadSeconds: value('UnloadSeconds'),
                rotate90Seconds: value('Rotate90Seconds'), flipSeconds: value('FlipSeconds') })
        });
        setStatus('Håndteringsklasse gemt i GOH');
        await loadCalculators();
    } finally { button.disabled = false; }
}

function addBendingMachine() {
    state.bendingMachines.unshift({ MachineCode: '', Description: '', MaxBendLengthMm: '', MaxForceKn: '', BaseCycleSeconds: '', SecondsPerDegree: '0', BackGaugeSeconds: '0', SetupMinutes: '0', SafetyFactor: '0.8', _isNew: true });
    editingBendingMachineRow = 0;
    renderBendingMachines(state.bendingMachines);
}

function addBendingBand() {
    state.bendingHandlingBands.unshift({ HandlingBandId: '', BandName: '', MinWeightKg: '0', MaxWeightKg: '', MinLongestSideMm: '0', MaxLongestSideMm: '', LoadSeconds: '', UnloadSeconds: '', Rotate90Seconds: '', FlipSeconds: '', _isNew: true });
    editingBendingBandRow = 0;
    renderBendingBands(state.bendingHandlingBands);
}

async function saveLaserParameter(button) {
    const row = button.closest('tr');
    const value = field => row.querySelector('[data-laser-field="' + field + '"]').value.trim();
    button.disabled = true;
    try {
        await fetchJson('/bom/calculators/laser-params', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                prodNo: value('ProdNo'), description: value('Descr'), thickness: value('Tykkelse'), machine: value('Maskine'),
                cutSpeedMPerMin: value('Skærehast.'), piercingMinutes: value('Pircing'), surchargePercent: value('Tillæg'), lens: value('Linse')
            })
        });
        setStatus('Laserparameter gemt i GOH');
        await loadCalculators();
    } finally {
        button.disabled = false;
    }
}

async function saveLaserTechnicalParameter(button) {
    const row = button.closest('tr');
    const value = field => row.querySelector('[data-laser-technical-field="' + field + '"]').value.trim();
    button.disabled = true;
    try {
        await fetchJson('/bom/calculators/laser-technical-params', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                technology: value('Technology'), material: value('Material'), thickness: value('Thickness'), lens: value('Lens'),
                piercingMilliseconds: value('PiercingMilliseconds'), vaporPowerW: value('VaporPowerW'), reducedPowerW: value('ReducedPowerW'),
                feedrateLargeMmMin: value('FeedrateLargeMmMin'), feedrateMediumMmMin: value('FeedrateMediumMmMin'),
                feedrateSmallMmMin: value('FeedrateSmallMmMin'), feedrateEngravingMmMin: value('FeedrateEngravingMmMin'),
                gasPressureBar: value('GasPressureBar'), nozzleSizeMm: value('NozzleSizeMm')
            })
        });
        setStatus('Laserteknologi gemt i GOH');
        await loadCalculators();
    } finally { button.disabled = false; }
}

function addLaserParameter() {
    state.laserParams.unshift({ ProdNo: '', Descr: '', Tykkelse: '', Maskine: document.getElementById('laserMachineInput').value || '', 'Skærehast.': '', Pircing: '0', 'Tillæg': '0', Linse: '', Source: 'ny', _isNew: true });
    editingLaserRow = 0;
    renderLaserParameters(state.laserParams);
    const firstInput = laserBody.querySelector('input');
    if (firstInput) firstInput.focus();
}

function addLaserTechnicalParameter() {
    state.laserTechnicalParams.unshift({ Technology: '', Material: '', Thickness: '', Lens: '', PiercingMilliseconds: '0',
        VaporPowerW: '0', ReducedPowerW: '0', FeedrateLargeMmMin: '0', FeedrateMediumMmMin: '0', FeedrateSmallMmMin: '0',
        FeedrateEngravingMmMin: '0', GasPressureBar: '0', NozzleSizeMm: '0', Source: 'ny', _isNew: true });
    editingLaserTechnicalRow = 0;
    renderLaserTechnicalParameters(state.laserTechnicalParams);
    const firstInput = document.getElementById('laserTechnicalBody').querySelector('input');
    if (firstInput) firstInput.focus();
}

async function importLaserParametersFromExcel() {
    const button = document.getElementById('importLaserExcelBtn');
    button.disabled = true;
    setStatus('Importerer laserparametre fra BOM.xlsm...');
    try {
        const result = await fetchJson('/bom/calculators/laser-params/import-excel', {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ overwriteExisting: false })
        });
        const technical = result.technical || { inserted: 0, preserved: 0 };
        setStatus('Excel-import: ' + result.inserted + ' artikelrækker og ' + technical.inserted + ' teknologier indsat');
        showToast('Import færdig: ' + result.inserted + ' artikler, ' + technical.inserted + ' teknologier', 'ok');
        await loadCalculators();
    } catch (err) {
        setStatus('Excel-import fejlede: ' + err.message);
        showToast('Excel-import fejlede: ' + err.message, 'err');
    } finally { button.disabled = false; }
}
function renderTreeNode(row, kindLabel, kindClass) {
    return '<div class="tree-node"><span class="tree-kind ' + kindClass + '">' + escapeHtml(kindLabel) + '</span><strong>' + escapeHtml(row.ProdNo || '-') + '</strong> ' + escapeHtml(row.Descr || '') + '<div class="muted">TgNo: ' + escapeHtml(row.TgNo || '-') + ' · Rev: ' + escapeHtml(row.RevNo || '-') + ' · Pos: ' + escapeHtml(row.PosNo || '-') + '</div></div>';
}
async function loadProductTree() {
    const product = state.selectedProduct;
    if (!product || product.IsLocalDraft) {
        productTree.innerHTML = product && product.IsLocalDraft ? renderLocalDraftTree(product) : '<div class="empty">Vælg et produkt for at se træet.</div>';
        treeMeta.textContent = product && product.IsLocalDraft ? 'lokal kladde' : '-';
        return;
    }
    treeMeta.textContent = 'henter...';
    try {
        const data = await fetchJson('/bom/product-tree?prodNo=' + encodeURIComponent(product.ProdNo));
        const parts = [];
        if (data.parent) parts.push('<div class="tree-node parent"><span class="tree-kind">FAR</span><strong>' + escapeHtml(data.parent.ProdNo) + '</strong> ' + escapeHtml(data.parent.Descr || '') + '</div>');
        if (data.route) parts.push(renderTreeNode(data.route, 'RUTE', 'route'));
        if (data.laser) parts.push(renderTreeNode(data.laser, 'LASER', 'laser'));
        if (Array.isArray(data.sublevels) && data.sublevels.length) {
            parts.push('<div class="tree-children">' + data.sublevels.map(slot => {
                const inner = [];
                if (slot.main) inner.push(renderTreeNode(slot.main, 'POS ' + (slot.pos == null ? '?' : slot.pos), ''));
                if (slot.laser) inner.push(renderTreeNode(slot.laser, 'LASER', 'laser'));
                return inner.join('');
            }).join('') + '</div>');
        }
        productTree.innerHTML = parts.length ? parts.join('') : '<div class="empty">Ingen underniveauer fundet for ' + escapeHtml(product.ProdNo) + '.</div>';
        treeMeta.textContent = data.count + ' noder · ' + (data.sublevels ? data.sublevels.length : 0) + ' underniveauer';
    } catch (err) {
        productTree.innerHTML = '<div class="empty">Fejl: ' + escapeHtml(err.message) + '</div>';
        treeMeta.textContent = 'fejl';
    }
}
function renderLocalDraftTree(draft) {
    const subs = Array.isArray(draft.Sublevels) ? draft.Sublevels : [];
    const parts = ['<div class="tree-node parent"><span class="tree-kind">FAR</span><strong>' + escapeHtml(draft.ProdNo) + '</strong> ' + escapeHtml(draft.Descr || '') + ' <span class="tag">LOKAL</span></div>'];
    if (draft.MaterialProdNo) parts.push('<div class="tree-node"><span class="tree-kind laser">PLADE</span>' + escapeHtml(draft.MaterialProdNo) + ' ' + escapeHtml(draft.MaterialDescr || '') + '</div>');
    if (Array.isArray(draft.Resources) && draft.Resources.length) parts.push('<div class="tree-node"><span class="tree-kind route">RES</span>' + escapeHtml(draft.Resources.map(r => r.ProdNo + ' ' + (r.Descr || '')).join(' · ')) + '</div>');
    if (subs.length) {
        parts.push('<div class="tree-children">' + subs.map((sub, idx) => '<div class="tree-node"><span class="tree-kind">POS ' + (idx + 1) + '</span><strong>' + escapeHtml(sub.ProdNo) + '</strong> ' + escapeHtml(sub.Descr || '') + (sub.IsLaser ? ' <span class="tree-kind laser">L</span>' : '') + '</div>').join('') + '</div>');
    }
    return parts.join('');
}
async function loadCustomerNotes() {
    if (!state.selectedCustomer) {
        customerNotesList.innerHTML = '<div class="empty">Vælg en kunde.</div>';
        notesMeta.textContent = '-';
        return;
    }
    const code = state.selectedCustomer.Gr || state.selectedCustomer['Varenr.'] || '';
    if (!code) {
        customerNotesList.innerHTML = '<div class="empty">Kunden har ingen varenr-kode.</div>';
        notesMeta.textContent = '0 noter';
        return;
    }
    try {
        const data = await fetchJson('/bom/customer-notes?customerCode=' + encodeURIComponent(code));
        const rows = data.rows || [];
        notesMeta.textContent = rows.length + ' noter';
        customerNotesList.innerHTML = rows.length ? rows.map(row => '<div class="note-item">' + escapeHtml(row.Txt1) + '</div>').join('') : '<div class="empty">Ingen BOM-noter for denne kunde.</div>';
    } catch (err) {
        customerNotesList.innerHTML = '<div class="empty">Fejl: ' + escapeHtml(err.message) + '</div>';
        notesMeta.textContent = 'fejl';
    }
}
async function loadComponents() {
    setStatus('Henter komponenter...');
    tableSkeleton(componentsHead, componentsBody, 5, 8);
    const q = encodeURIComponent(String(componentsSearchInput.value || '').trim());
    const data = await fetchJson('/bom/components?q=' + q + '&limit=1000');
    state.components = data.rows || [];
    renderSimpleTable(componentsHead, componentsBody, state.components);
    setStatus('Komponenter indlæst', state.components.length);
}
async function loadSuppliers() {
    setStatus('Henter leverandører...');
    tableSkeleton(suppliersHead, suppliersBody, 5, 8);
    const q = encodeURIComponent(String(suppliersSearchInput.value || '').trim());
    const data = await fetchJson('/bom/suppliers?q=' + q);
    state.suppliers = data.rows || [];
    renderSimpleTable(suppliersHead, suppliersBody, state.suppliers);
    setStatus('Leverandører indlæst', state.suppliers.length);
}
async function loadCustomers() {
    setStatus('Henter kunder...');
    if (!state.customers.length) listSkeleton(customersList, 8);
    const q = encodeURIComponent(String(customerSearchInput.value || '').trim());
    const data = await fetchJson('/bom/customers?q=' + q);
    state.customers = data.rows || [];
    if (!state.selectedCustomer && state.customers.length) state.selectedCustomer = state.customers[0];
    renderCustomers();
    updateContext();
    setMetric('metricCustomers', formatNumber(state.customers.length));
    cachePill.textContent = 'Cache: kunder ' + (data.cached ? 'hit' : 'miss');
    setStatus('Kunder indlæst', state.customers.length);
}
async function loadProducts() {
    if (!state.selectedCustomer) {
        state.products = [];
        renderProducts();
        return;
    }
    setStatus('Henter produkter for kunde ' + state.selectedCustomer.CustNo + '...');
    listSkeleton(productsList, 6);
    const data = await fetchJson('/bom/products?customerNo=' + encodeURIComponent(state.selectedCustomer.CustNo) + '&customerCode=' + encodeURIComponent(state.selectedCustomer.Gr || state.selectedCustomer['Varenr.'] || ''));
    const backendRows = data.rows || [];
    const draftRows = loadDraftProducts(state.selectedCustomer.CustNo);
    state.draftProducts = draftRows;
    state.products = backendRows.concat(draftRows);
    if (!state.selectedProduct && state.products.length) {
        state.selectedProduct = state.products[0];
        if (state.selectedProduct.TgNo) tgnInput.value = state.selectedProduct.TgNo;
    }
    renderProducts();
    renderProductDetail();
    loadProductTree();
    updateContext();
    setMetric('metricProducts', formatNumber(state.products.length));
    cachePill.textContent = 'Cache: produkter ' + (data.cached ? 'hit' : 'miss') + ' · lokale kladder ' + state.draftProducts.length;
    setStatus('Produkter indlæst', state.products.length);
}
async function loadRevisions() {
    if (!state.selectedCustomer) return;
    const tgn = String(tgnInput.value || '').trim() || String((state.selectedProduct && state.selectedProduct.TgNo) || '').trim();
    if (!tgn) {
        state.revisions = [];
        renderRevisions();
        setMetric('metricRevisions', '-');
        return;
    }
    setStatus('Henter revisioner for TgNo ' + tgn + '...');
    const data = await fetchJson('/bom/revisions/by-drawing?customerNo=' + encodeURIComponent(state.selectedCustomer.CustNo) + '&customerCode=' + encodeURIComponent(state.selectedCustomer.Gr || state.selectedCustomer['Varenr.'] || '') + '&tgn=' + encodeURIComponent(tgn));
    state.revisions = data.rows || [];
    state.selectedRevision = state.revisions.length ? state.revisions[0] : null;
    renderRevisions();
    setMetric('metricRevisions', formatNumber(state.revisions.length));
    cachePill.textContent = 'Cache: revisioner ' + (data.cached ? 'hit' : 'miss');
    setStatus('Revisioner indlæst', state.revisions.length);
}
async function loadResources() {
    setStatus('Henter ressourcer...');
    if (!state.resources.length) tableSkeleton(resourcesHead, resourcesBody, 6, 8);
    const data = await fetchJson('/bom/resources');
    state.resources = data.rows || [];
    const filter = String(resourcesSearchInput.value || '').trim().toLowerCase();
    const filtered = filter ? state.resources.filter(row => [row.ProdNo, row.Descr, row.CustomerNo].some(v => String(v || '').toLowerCase().includes(filter))) : state.resources;
    renderResources(filtered);
    setMetric('metricResources', formatNumber(state.resources.length));
    cachePill.textContent = 'Cache: ressourcer ' + (data.cached ? 'hit' : 'miss');
    setStatus('Ressourcer indlæst', filtered.length);
}
async function loadMaterials() {
    setStatus('Henter materialer...');
    if (!state.materials.length) tableSkeleton(materialsHead, materialsBody, 6, 8);
    const q = encodeURIComponent(String(materialsSearchInput.value || '').trim());
    const data = await fetchJson('/bom/materials?q=' + q);
    state.materials = data.rows || [];
    renderMaterials();
    setMetric('metricMaterials', formatNumber(state.materials.length));
    cachePill.textContent = 'Cache: materialer ' + (data.cached ? 'hit' : 'miss');
    setStatus('Materialer indlæst', state.materials.length);
}
async function loadCalculators() {
    setStatus('Henter parametre...');
    if (!laserBody.children.length) {
        tableSkeleton(laserHead, laserBody, 5, 4);
        tableSkeleton(processHead, processBody, 5, 4);
        tableSkeleton(processResourceHead, processResourceBody, 5, 4);
    }
    const machine = encodeURIComponent(String(document.getElementById('laserMachineInput').value || '').trim());
    const family = String(document.getElementById('processFilterSelect').value || '').trim().toLowerCase();
    const familyCodes = {
        laser: ['11', '12'],
        buk: ['21'],
        svejs: ['50', '50-1', '51', '56'],
        flad: ['60', '61', '62', '63', '64']
    };
    const [laserData, bendingData, processData, resourceData] = await Promise.all([
        fetchJson('/bom/calculators/laser-params?machine=' + machine),
        fetchJson('/bom/calculators/bending-params'),
        fetchJson('/bom/calculators/process-params'),
        fetchJson('/bom/resources')
    ]);
    const familyRows = (resourceData.rows || []).filter(row => {
        const resourceFamily = String(row.R7 || '').trim();
        if (!family) return Object.values(familyCodes).some(codes => codes.includes(resourceFamily));
        return (familyCodes[family] || []).includes(resourceFamily);
    });
    renderCalculators(laserData.rows || [], laserData.technicalRows || [], laserData.gasPrices || {}, bendingData, processData.rows || [], familyRows);
    setMetric('metricCalculators', formatNumber((laserData.rows || []).length + (laserData.technicalRows || []).length + (bendingData.machines || []).length + (processData.rows || []).length));
    cachePill.textContent = 'Cache: parametre ' + ((laserData.cached && processData.cached) ? 'hit' : 'miss');
    const familyLabel = family ? document.getElementById('processFilterSelect').selectedOptions[0].textContent : 'Alle procesfamilier';
    setStatus(familyLabel + ': ressourcer indlæst', familyRows.length);
}
async function primeOverviewCounts() {
    if (state.permissions.bomStykliste || state.permissions.bomCalculator) try { const customers = await fetchJson('/bom/customers'); setMetric('metricCustomers', formatNumber(customers.count || 0)); } catch (_) {}
    if (state.permissions.bomResources || state.permissions.bomParameters || state.permissions.bomCalculator) try { const resources = await fetchJson('/bom/resources'); state.resources = resources.rows || []; setMetric('metricResources', formatNumber(resources.count || 0)); } catch (_) {}
    if (state.permissions.bomMaterials || state.permissions.bomCalculator) try { const materials = await fetchJson('/bom/materials'); state.materials = materials.rows || []; setMetric('metricMaterials', formatNumber(materials.count || 0)); } catch (_) {}
    if (state.permissions.bomParameters || state.permissions.bomCalculator) try {
        const laser = await fetchJson('/bom/calculators/laser-params?machine=R1100');
        const process = await fetchJson('/bom/calculators/process-params');
        setMetric('metricCalculators', formatNumber((laser.count || 0) + (process.count || 0)));
    } catch (_) {}
}
async function invalidateActiveCache() {
    const scopeMap = { overview: 'all', stykliste: 'customers', resources: 'resources', materials: 'materials', calculators: 'calculators' };
    const scope = scopeMap[state.view] || 'all';
    await fetchJson('/bom/cache/invalidate?scope=' + encodeURIComponent(scope), { method: 'POST' });
    cachePill.textContent = 'Cache: nulstillet';
    setStatus('Cache ryddet for ' + scope);
}

// ── Lokal produktkladde (modal) ──
function renderDraftResourceChips() {
    draftResourceChips.innerHTML = state.draftResources.length
        ? state.draftResources.map((row, idx) => '<span class="chip" data-idx="' + idx + '" style="cursor:pointer;" title="klik for at fjerne">' + escapeHtml((row.ProdNo || '') + ' ' + (row.Descr || '')) + ' ✕</span>').join('')
        : '<span class="muted">Ingen ressourcer valgt</span>';
    draftResourceChips.querySelectorAll('.chip[data-idx]').forEach(el => {
        el.addEventListener('click', () => {
            state.draftResources.splice(Number(el.getAttribute('data-idx')), 1);
            renderDraftResourceChips();
        });
    });
}
function openDraftModal() {
    if (!state.selectedCustomer) {
        setStatus('Vælg en kunde først for at oprette lokal produktkladde.');
        return;
    }
    const custCode = String(state.selectedCustomer.Gr || state.selectedCustomer.CustNo || '');
    draftCustomerText.value = (state.selectedCustomer.CustNo || '-') + ' - ' + (state.selectedCustomer.Nm || '');
    if (draftProdNoPrefix) draftProdNoPrefix.textContent = custCode ? (custCode + ' +') : '';
    if (draftProdNoSuffix) draftProdNoSuffix.value = '';
    if (draftProdNoPreview) draftProdNoPreview.textContent = '—';
    draftDescr.value = '';
    draftTgNo.value = '';
    if (draftRevNo) draftRevNo.value = '';
    if (draftTgForm) draftTgForm.value = 'A4';
    if (draftCustomerNoAlt) draftCustomerNoAlt.value = '';
    draftNote.value = '';
    sublevelRows.innerHTML = '';
    state.draftMaterial = null;
    state.draftResources = [];
    draftMaterialSearch.value = '';
    draftMaterialChosen.textContent = 'Ingen valgt';
    draftResourceSearch.value = '';
    renderDraftResourceChips();
    if (vismaPreviewPanel) vismaPreviewPanel.style.display = 'none';
    if (createVismaBtn) { createVismaBtn.style.display = 'none'; createVismaBtn.disabled = true; }
    Promise.all([ensureMaterials(), ensureResources()]).catch(() => {});
    draftProductModal.classList.add('open');
    if (draftProdNoSuffix) draftProdNoSuffix.focus();
}
function closeDraftModal() {
    draftProductModal.classList.remove('open');
}
function addSublevelRow(prefill) {
    const idx = sublevelRows.children.length + 1;
    const row = document.createElement('div');
    row.className = 'sub-row';
    row.innerHTML = '<input class="sub-no" placeholder="auto: far-' + idx + '" value="' + escapeHtml((prefill && prefill.ProdNo) || '') + '" />'
        + '<input class="sub-descr" placeholder="beskrivelse" value="' + escapeHtml((prefill && prefill.Descr) || '') + '" />'
        + '<input class="sub-tgno" placeholder="tgno" value="' + escapeHtml((prefill && prefill.TgNo) || '') + '" />'
        + '<label class="laser-check"><input type="checkbox" class="sub-laser"' + ((prefill && prefill.IsLaser) ? ' checked' : '') + ' /> L</label>'
        + '<button type="button" class="alt sub-remove">Fjern</button>';
    row.querySelector('.sub-remove').addEventListener('click', () => row.remove());
    sublevelRows.appendChild(row);
}
function collectSublevels(parentProdNo) {
    return Array.from(sublevelRows.querySelectorAll('.sub-row')).map((row, idx) => {
        const manualNo = String(row.querySelector('.sub-no').value || '').trim();
        const isLaser = row.querySelector('.sub-laser').checked;
        const autoNo = parentProdNo + '-' + (idx + 1) + (isLaser ? 'L' : '');
        return {
            ProdNo: manualNo || autoNo,
            Descr: String(row.querySelector('.sub-descr').value || '').trim(),
            TgNo: String(row.querySelector('.sub-tgno').value || '').trim(),
            IsLaser: isLaser
        };
    }).filter(sub => sub.ProdNo);
}
