// ── BOM Workspace · core ─────────────────────────────────────────────
// Fælles state, DOM-referencer og hjælpefunktioner.
// Læser KUN fra Visma — ingen skrivning til databasen.

const state = {
    view: 'overview',
    customers: [],
    products: [],
    revisions: [],
    resources: [],
    materials: [],
    components: [],
    laserParams: [],
    processParams: [],
    selectedCustomer: null,
    selectedProduct: null,
    selectedRevision: null,
    draftProducts: [],
    suppliers: [],
    fileAnalysis: null,
    calcMaterial: null,
    calcCustomer: null,
    calcCustomerMetaCache: {},
    calcComponents: [],
    draftMaterial: null,
    draftResources: []
};

const navItems = [
    { key: 'overview', title: 'Oversigt', description: 'Status og arbejdsgang' },
    { key: 'stykliste', title: 'Stykliste', description: 'Kunde → produkt → TgNo → revision' },
    { key: 'komponenter', title: 'Komponenter', description: 'Komp-katalog (Gr5 2/3/6/10/11)' },
    { key: 'resources', title: 'Ressourcer', description: 'Ressource- og rutekatalog' },
    { key: 'materials', title: 'Materialer', description: 'Råvarer med lagerstatus' },
    { key: 'calculators', title: 'Parametre', description: 'Skæreparametre og procesmatrix' },
    { key: 'beregner', title: 'Beregner', description: 'Fil-analyse, nesting og pris' },
    { key: 'leverandorer', title: 'Leverandører', description: 'Leverandørkatalog fra Actor' }
];

const viewMeta = {
    overview: { title: 'Oversigt', subtitle: 'Status og arbejdsgang' },
    stykliste: { title: 'Stykliste', subtitle: 'Opslag som i BOM-regnearket' },
    komponenter: { title: 'Komponenter', subtitle: 'Komp-ark katalog' },
    resources: { title: 'Ressourcer', subtitle: 'Ressource- og ruteoversigt' },
    materials: { title: 'Materialer', subtitle: 'Råvarer-forespørgsel med lager' },
    calculators: { title: 'Parametre', subtitle: 'Skæreparametre, procesmatrix og ressourcefamilier' },
    beregner: { title: 'Beregner', subtitle: 'Fil-analyse, nesting og priskalkulation' },
    leverandorer: { title: 'Leverandører', subtitle: 'Lev-ark fra Actor (SupNo)' }
};

// ── DOM-referencer ──
const navList = document.getElementById('navList');
const viewTitle = document.getElementById('viewTitle');
const viewSubtitle = document.getElementById('viewSubtitle');
const statusText = document.getElementById('statusText');
const countText = document.getElementById('countText');
const cachePill = document.getElementById('cachePill');
const customerSearchInput = document.getElementById('customerSearchInput');
const customerSelect = document.getElementById('customerSelect');
const productSearchInput = document.getElementById('productSearchInput');
const tgnInput = document.getElementById('tgnInput');
const customersList = document.getElementById('customersList');
const productsList = document.getElementById('productsList');
const draftProductModal = document.getElementById('draftProductModal');
const draftProductForm = document.getElementById('draftProductForm');
const draftCustomerText = document.getElementById('draftCustomerText');
const draftProdNo = document.getElementById('draftProdNo');
const draftDescr = document.getElementById('draftDescr');
const draftTgNo = document.getElementById('draftTgNo');
const draftRevNo = document.getElementById('draftRevNo');
const draftNote = document.getElementById('draftNote');
const customersMeta = document.getElementById('customersMeta');
const productsMeta = document.getElementById('productsMeta');
const revisionsMeta = document.getElementById('revisionsMeta');
const revisionsCountMeta = document.getElementById('revisionsCountMeta');
const productDetailGrid = document.getElementById('productDetailGrid');
const revisionsHead = document.getElementById('revisionsHead');
const revisionsBody = document.getElementById('revisionsBody');
const resourcesHead = document.getElementById('resourcesHead');
const resourcesBody = document.getElementById('resourcesBody');
const materialsHead = document.getElementById('materialsHead');
const materialsBody = document.getElementById('materialsBody');
const materialsSearchInput = document.getElementById('materialsSearchInput');
const resourcesSearchInput = document.getElementById('resourcesSearchInput');
const laserHead = document.getElementById('laserHead');
const laserBody = document.getElementById('laserBody');
const processHead = document.getElementById('processHead');
const processBody = document.getElementById('processBody');
const processResourceHead = document.getElementById('processResourceHead');
const processResourceBody = document.getElementById('processResourceBody');
const componentsHead = document.getElementById('componentsHead');
const componentsBody = document.getElementById('componentsBody');
const componentsSearchInput = document.getElementById('componentsSearchInput');
const suppliersHead = document.getElementById('suppliersHead');
const suppliersBody = document.getElementById('suppliersBody');
const suppliersSearchInput = document.getElementById('suppliersSearchInput');
const productTree = document.getElementById('productTree');
const treeMeta = document.getElementById('treeMeta');
const customerNotesList = document.getElementById('customerNotesList');
const notesMeta = document.getElementById('notesMeta');
const draftMaterialSearch = document.getElementById('draftMaterialSearch');
const draftMaterialList = document.getElementById('draftMaterialList');
const draftMaterialChosen = document.getElementById('draftMaterialChosen');
const draftResourceSearch = document.getElementById('draftResourceSearch');
const draftResourceList = document.getElementById('draftResourceList');
const draftResourceChips = document.getElementById('draftResourceChips');
const sublevelRows = document.getElementById('sublevelRows');
const calcMaterialSearch = document.getElementById('calcMaterialSearch');
const calcMaterialList = document.getElementById('calcMaterialList');
const calcMaterialChosen = document.getElementById('calcMaterialChosen');
const processCards = document.getElementById('processCards');
const fileAnalysisStatus = document.getElementById('fileAnalysisStatus');
const fileAnalysisGrid = document.getElementById('fileAnalysisGrid');
const nestingGrid = document.getElementById('nestingGrid');
const nestingMeta = document.getElementById('nestingMeta');
const quoteGrid = document.getElementById('quoteGrid');
const quoteMeta = document.getElementById('quoteMeta');
const quotePriceBig = document.getElementById('quotePriceBig');
const quoteStatus = document.getElementById('quoteStatus');

// ── Hjælpefunktioner ──
function setStatus(text, count) {
    statusText.textContent = text;
    if (typeof count === 'number') countText.textContent = count + ' rækker';
}
function escapeHtml(value) {
    return String(value == null ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
function formatNumber(value) {
    return new Intl.NumberFormat('da-DK', { maximumFractionDigits: 0 }).format(Number(value || 0));
}
function formatMoney(value) {
    return new Intl.NumberFormat('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value || 0));
}
function parseDaNumber(value) {
    return Number(String(value == null ? '0' : value).replace(',', '.')) || 0;
}
async function fetchJson(url, options) {
    const response = await fetch(url, options || {});
    if (!response.ok) {
        let msg = 'HTTP ' + response.status;
        try {
            const body = await response.json();
            if (body && body.error) msg = body.error;
        } catch (_) {}
        throw new Error(msg);
    }
    return response.json();
}
function setMetric(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

// ── Data-hentning (cache i state) ──
function materialOptionLabel(row) {
    return (row.ProdNo || '') + ' · ' + (row.Descr || row.beskrivelse || '');
}
async function ensureMaterials() {
    if (state.materials.length === 0) {
        const data = await fetchJson('/bom/materials');
        state.materials = data.rows || [];
    }
    return state.materials;
}
async function ensureResources() {
    if (state.resources.length === 0) {
        const data = await fetchJson('/bom/resources');
        state.resources = data.rows || [];
    }
    return state.resources;
}
async function ensureCustomers() {
    if (state.customers.length === 0) {
        const data = await fetchJson('/bom/customers');
        state.customers = data.rows || [];
    }
    return state.customers;
}
async function ensureComponents() {
    if (state.components.length === 0) {
        const data = await fetchJson('/bom/components?limit=3000');
        state.components = data.rows || [];
    }
    return state.components;
}

// ── Generisk søge-picker (skriv for at søge, som i Excel) ──
function attachPicker({ input, list, getRows, rowLabel, rowSub, onPick, maxItems }) {
    const max = maxItems || 40;
    function render() {
        const q = String(input.value || '').trim().toLowerCase();
        const terms = q.split(/\s+/).filter(Boolean);
        const rows = getRows().filter(row => {
            if (!terms.length) return true;
            const hay = (rowLabel(row) + ' ' + (rowSub ? rowSub(row) : '')).toLowerCase();
            return terms.every(term => hay.includes(term));
        }).slice(0, max);
        list.innerHTML = rows.length
            ? rows.map((row, idx) => '<div class="picker-item" data-i="' + idx + '"><strong>' + escapeHtml(rowLabel(row)) + '</strong>' + (rowSub ? '<small>' + escapeHtml(rowSub(row)) + '</small>' : '') + '</div>').join('')
            : '<div class="picker-item">Intet match — prøv andre ord</div>';
        list.querySelectorAll('.picker-item[data-i]').forEach(el => {
            el.addEventListener('mousedown', evt => {
                evt.preventDefault();
                onPick(rows[Number(el.getAttribute('data-i'))]);
                close();
            });
        });
        list.classList.add('open');
    }
    function close() { list.classList.remove('open'); }
    input.addEventListener('input', render);
    input.addEventListener('focus', render);
    input.addEventListener('blur', () => setTimeout(close, 150));
    input.addEventListener('keydown', evt => { if (evt.key === 'Escape') close(); });
}

// ── Række-labels til pickers ──
function materialRowLabel(row) { return row.ProdNo || ''; }
function materialRowSub(row) {
    const lager = Number(row.Bal == null ? 0 : row.Bal);
    const lagerTxt = row.Bal == null ? 'lager ?' : 'lager ' + lager;
    return (row.beskrivelse || row.Descr || '') + ' · tyk ' + (row.tykklese == null ? '-' : row.tykklese) + ' · ' + (row.Bredde || '-') + 'x' + (row['Længde'] || '-') + ' m · ' + lagerTxt;
}
function filteredCalcMaterials() {
    const th = Number(document.getElementById('calcThicknessFilter').value || 0);
    const onlyStock = document.getElementById('calcOnlyStock').checked;
    return state.materials.filter(row => {
        if (th > 0 && Math.abs(Number(row.tykklese || 0) - th) > 0.011) return false;
        if (onlyStock && !(Number(row.Bal || 0) > 0)) return false;
        return true;
    });
}
function customerRowLabel(row) { return row.Nm || String(row.CustNo || ''); }
function customerRowSub(row) {
    return 'nr ' + (row.CustNo || '-') + ' · prisliste ' + (row.CustPrGr == null ? '-' : row.CustPrGr) + (row.PArea ? ' · ' + row.PArea : '');
}
function resourceRowLabel(row) { return row.ProdNo || ''; }
function resourceRowSub(row) {
    return (row.Descr || '') + ' · kost ' + (row.CstPr == null ? '-' : row.CstPr) + ' · salg ' + (row.SalePr == null ? '-' : row.SalePr);
}
function componentRowLabel(row) { return row.ProdNo || ''; }
function componentRowSub(row) {
    return (row.Descr || '') + ' · ' + (row.Enhed || '-') + ' · pris ' + (row.Pris == null ? '-' : row.Pris);
}

// ── Kunde-indstillinger (pristype, opstart, minimum) huskes lokalt pr kunde ──
const CUST_PREFS_KEY = 'bomCalcCustomerPrefs';
function loadCustPrefs() {
    try { return JSON.parse(localStorage.getItem(CUST_PREFS_KEY) || '{}'); } catch (_) { return {}; }
}
function saveCalcCustomerPref() {
    if (!state.calcCustomer) return;
    const prefs = loadCustPrefs();
    prefs[String(state.calcCustomer.CustNo)] = {
        priceBasis: document.getElementById('calcPriceBasis').value,
        laserOpstart: document.getElementById('calcLaserOpstartChk').checked,
        laserOpstartMin: Number(document.getElementById('calcLaserOpstartMin').value || 0),
        minOrderAmount: Number(document.getElementById('calcMinAmount').value || 0),
        minQty: Number(document.getElementById('calcMinQty').value || 0)
    };
    try { localStorage.setItem(CUST_PREFS_KEY, JSON.stringify(prefs)); } catch (_) {}
}
function applyCalcCustomerPref(row) {
    const pref = loadCustPrefs()[String(row.CustNo)];
    if (!pref) {
        const custPrGr = Number(row && row.CustPrGr);
        document.getElementById('calcPriceBasis').value = Number.isFinite(custPrGr) && custPrGr <= 0 ? 'cost' : 'sale';
        return;
    }
    document.getElementById('calcPriceBasis').value = pref.priceBasis || 'sale';
    document.getElementById('calcLaserOpstartChk').checked = pref.laserOpstart !== false;
    if (pref.laserOpstartMin != null) document.getElementById('calcLaserOpstartMin').value = pref.laserOpstartMin;
    if (pref.minOrderAmount != null) document.getElementById('calcMinAmount').value = pref.minOrderAmount;
    if (pref.minQty != null) document.getElementById('calcMinQty').value = pref.minQty;
}

// ── Lokale produktkladder (kun browser, aldrig Visma) ──
function getDraftStorageKey(customerNo) {
    return 'bom_local_product_drafts:' + String(customerNo || 'none');
}
function loadDraftProducts(customerNo) {
    try {
        const raw = localStorage.getItem(getDraftStorageKey(customerNo));
        const rows = raw ? JSON.parse(raw) : [];
        return Array.isArray(rows) ? rows : [];
    } catch (_) {
        return [];
    }
}
function saveDraftProducts(customerNo, rows) {
    localStorage.setItem(getDraftStorageKey(customerNo), JSON.stringify(Array.isArray(rows) ? rows : []));
}

function customerPriceGroupText(row) {
    if (!row) return '-';
    return row.CustPrGr == null ? '-' : String(row.CustPrGr);
}

function calcNextProductNo(rows) {
    const numericRows = (Array.isArray(rows) ? rows : []).map(row => {
        const prodNo = String((row && row.ProdNo) || '').trim();
        const match = prodNo.match(/^\d+$/);
        if (!match) return null;
        return { raw: prodNo, value: BigInt(prodNo) };
    }).filter(Boolean);

    if (!numericRows.length) {
        return { last: null, next: null, hasNumeric: false };
    }

    let maxRow = numericRows[0];
    for (let i = 1; i < numericRows.length; i += 1) {
        if (numericRows[i].value > maxRow.value) maxRow = numericRows[i];
    }

    const nextRaw = (maxRow.value + 1n).toString();
    const next = nextRaw.padStart(maxRow.raw.length, '0');
    return { last: maxRow.raw, next, hasNumeric: true };
}

function setCalcCustomerMetaTexts(row, meta) {
    const chosenEl = document.getElementById('calcCustomerChosen');
    const priceEl = document.getElementById('calcCustomerPriceInfo');
    const prodEl = document.getElementById('calcCustomerProductInfo');

    if (!row) {
        chosenEl.textContent = 'Ingen kunde valgt - standard prisliste';
        priceEl.textContent = 'Kundens prisliste: -';
        prodEl.textContent = 'Seneste produktnr: - · Næste forslag: -';
        return;
    }

    chosenEl.textContent = 'Valgt: ' + (row.Nm || row.CustNo) + ' · nr ' + (row.CustNo || '-') + ' · prisliste ' + customerPriceGroupText(row);
    priceEl.textContent = 'Kundens prisliste (CustPrGr): ' + customerPriceGroupText(row) + ' · Pristype kan ændres manuelt.';

    if (!meta) {
        prodEl.textContent = 'Finder seneste produktnr...';
        return;
    }

    if (!meta.hasNumeric) {
        prodEl.textContent = 'Ingen rent numeriske produktnumre fundet for kunden.';
        return;
    }

    prodEl.textContent = 'Seneste produktnr (højeste): ' + meta.last + ' · Næste forslag: ' + meta.next;
}

async function refreshCalcCustomerMeta(row) {
    if (!row) {
        setCalcCustomerMetaTexts(null, null);
        return;
    }

    setCalcCustomerMetaTexts(row, null);

    const customerCode = row.Gr || row['Varenr.'] || '';
    const cacheKey = String(row.CustNo || '') + '|' + String(customerCode || '');
    if (state.calcCustomerMetaCache[cacheKey]) {
        setCalcCustomerMetaTexts(row, state.calcCustomerMetaCache[cacheKey]);
        return;
    }

    try {
        const data = await fetchJson('/bom/products?customerNo=' + encodeURIComponent(row.CustNo) + '&customerCode=' + encodeURIComponent(customerCode));
        const backendRows = data.rows || [];
        const draftRows = loadDraftProducts(row.CustNo);
        const meta = calcNextProductNo(backendRows.concat(draftRows));
        state.calcCustomerMetaCache[cacheKey] = meta;
        if (state.calcCustomer && String(state.calcCustomer.CustNo) === String(row.CustNo)) {
            setCalcCustomerMetaTexts(row, meta);
        }
    } catch (_) {
        const prodEl = document.getElementById('calcCustomerProductInfo');
        prodEl.textContent = 'Kunne ikke hente produktnumre for kunden.';
    }
}
