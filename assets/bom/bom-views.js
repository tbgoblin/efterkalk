// ── BOM Workspace · views ────────────────────────────────────────────
// Navigation, stykliste, komponenter, ressourcer, materialer,
// parametre, leverandører og lokale produktkladder.

function renderNav() {
    navList.innerHTML = navItems.map(item => {
        const active = item.key === state.view ? 'active' : '';
        return '<button class="nav-btn ' + active + '" data-view="' + item.key + '"><strong>' + escapeHtml(item.title) + '</strong><span>' + escapeHtml(item.description) + '</span></button>';
    }).join('');
    navList.querySelectorAll('.nav-btn').forEach(btn => btn.addEventListener('click', () => switchView(btn.getAttribute('data-view'))));
}
function switchView(view) {
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
    customersList.innerHTML = rows.length ? rows.map(row => {
        const active = state.selectedCustomer && String(state.selectedCustomer.CustNo) === String(row.CustNo) ? ' active' : '';
        return '<div class="list-item' + active + '" data-customer="' + escapeHtml(row.CustNo) + '"><strong>' + escapeHtml((row.CustNo || '') + ' · ' + (row.Gr || '') + ' · ' + (row.Nm || '')) + '</strong><span>' + escapeHtml((row.Shrt || '') + ' · ' + (row.PArea || '')) + '</span></div>';
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
    productsList.innerHTML = filtered.length ? filtered.map(row => {
        const active = state.selectedProduct && String(state.selectedProduct.ProdNo) === String(row.ProdNo) ? ' active' : '';
        const draftClass = row.IsLocalDraft ? ' draft' : '';
        const draftTag = row.IsLocalDraft ? '<span class="tag">LOKAL</span>' : '';
        return '<div class="list-item' + active + draftClass + '" data-product="' + escapeHtml(row.ProdNo) + '"><strong>' + escapeHtml((row.ProdNo || '-') + ' · ' + (row.Descr || '')) + draftTag + '</strong><span>' + escapeHtml('TgNo: ' + (row.TgNo || '-') + ' · Rev: ' + (row.RevNo || '-')) + '</span></div>';
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
function renderCalculators(laserRows, processRows, processResourceRows) {
    renderSimpleTable(laserHead, laserBody, laserRows);
    renderSimpleTable(processHead, processBody, processRows);
    renderSimpleTable(processResourceHead, processResourceBody, processResourceRows);
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
    const q = encodeURIComponent(String(componentsSearchInput.value || '').trim());
    const data = await fetchJson('/bom/components?q=' + q + '&limit=1000');
    state.components = data.rows || [];
    renderSimpleTable(componentsHead, componentsBody, state.components);
    setStatus('Komponenter indlæst', state.components.length);
}
async function loadSuppliers() {
    setStatus('Henter leverandører...');
    const q = encodeURIComponent(String(suppliersSearchInput.value || '').trim());
    const data = await fetchJson('/bom/suppliers?q=' + q);
    state.suppliers = data.rows || [];
    renderSimpleTable(suppliersHead, suppliersBody, state.suppliers);
    setStatus('Leverandører indlæst', state.suppliers.length);
}
async function loadCustomers() {
    setStatus('Henter kunder...');
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
    const machine = encodeURIComponent(String(document.getElementById('laserMachineInput').value || '').trim());
    const family = String(document.getElementById('processFilterSelect').value || '').trim().toLowerCase();
    const [laserData, processData, resourceData] = await Promise.all([
        fetchJson('/bom/calculators/laser-params?machine=' + machine),
        fetchJson('/bom/calculators/process-params'),
        fetchJson('/bom/resources')
    ]);
    const familyRows = (resourceData.rows || []).filter(row => {
        if (!family) return ['laser', 'buk', 'svejs', 'flad'].some(term => String(row.Descr || '').toLowerCase().includes(term));
        return String(row.Descr || '').toLowerCase().includes(family);
    });
    renderCalculators(laserData.rows || [], processData.rows || [], familyRows);
    setMetric('metricCalculators', formatNumber((laserData.rows || []).length + (processData.rows || []).length));
    cachePill.textContent = 'Cache: parametre ' + ((laserData.cached && processData.cached) ? 'hit' : 'miss');
    setStatus('Parametre indlæst', (laserData.rows || []).length + (processData.rows || []).length + familyRows.length);
}
async function primeOverviewCounts() {
    try { const customers = await fetchJson('/bom/customers'); setMetric('metricCustomers', formatNumber(customers.count || 0)); } catch (_) {}
    try { const resources = await fetchJson('/bom/resources'); state.resources = resources.rows || []; setMetric('metricResources', formatNumber(resources.count || 0)); } catch (_) {}
    try { const materials = await fetchJson('/bom/materials'); state.materials = materials.rows || []; setMetric('metricMaterials', formatNumber(materials.count || 0)); } catch (_) {}
    try {
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
    draftCustomerText.value = (state.selectedCustomer.CustNo || '-') + ' - ' + (state.selectedCustomer.Nm || '');
    draftProdNo.value = '';
    draftDescr.value = '';
    draftTgNo.value = '';
    draftRevNo.value = 'KLADDE';
    draftNote.value = '';
    sublevelRows.innerHTML = '';
    state.draftMaterial = null;
    state.draftResources = [];
    draftMaterialSearch.value = '';
    draftMaterialChosen.textContent = 'Ingen valgt';
    draftResourceSearch.value = '';
    renderDraftResourceChips();
    Promise.all([ensureMaterials(), ensureResources()]).catch(() => {});
    draftProductModal.classList.add('open');
    draftProdNo.focus();
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
