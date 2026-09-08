// ── BOM Workspace · beregner ─────────────────────────────────────────
// Priskalkulation som i Excel: laser (skæreparametre), buk, svejs,
// flad (typer), stykliste (R8200 komponenter), nesting og prismatrix.
// Alt er læsning — der skrives ALDRIG til Visma.

// Processfamilier følger BgtLn.R7 fra ressourcekataloget:
//   11/12 = laser · 21 = buk · 50/50-1/51/56 = svejs · 60-64 = flad/efterbehandling · 82 = stykliste
const processDefs = [
    { key: 'laser', label: 'Laser (skæring)', isLaser: true, defaultOn: true },
    { key: 'buk', label: 'Buk (kantbukning)', kind: 'buk', r7: ['21'], defaultRes: 'R2100' },
    { key: 'svejs', label: 'Svejs', kind: 'svejs', r7: ['50', '50-1', '51', '56'], defaultRes: 'R5300' },
    { key: 'flad', label: 'Flad / efterbehandling', kind: 'flad', r7: ['60', '61', '62', '63', '64'], defaultRes: 'R6104' },
    { key: 'stykliste', label: 'Stykliste (R8200) — komponentliste', kind: 'stykliste', r7: ['82'] },
    { key: 'andet', label: 'Andet (montage, PEM, valse, save, underleverandør ...)', kind: 'andet', r7: null }
];
let quoteDebounceTimer = null;
const dxfMeasureState = {
    pointA: null,
    pointB: null,
    hoverPoint: null,
    hoverKind: '',
    projection: null,
    eventsBound: false
};
const calcWizardSteps = ['customer', 'drawing', 'material', 'processes', 'result'];
const calcWizardMeta = {
    customer: ['Vælg kunde', 'Søg kunden og kontrollér prisindstillingerne.'],
    drawing: ['Indlæs tegning', 'Upload DXF, STEP eller PDF og kontrollér emnemålene.'],
    material: ['Vælg plade', 'Find den rigtige kvalitet, tykkelse og pladestørrelse.'],
    processes: ['Vælg processer', 'Kontrollér laser, buk, svejs og øvrige operationer.']
};
let activeCalcWizardStep = '';
const LASER_COLUMN_WIDTHS_KEY = 'bomLaserTechnologyColumnWidths';

function renderPieceThumbnail() {
    const canvas = document.getElementById('pieceThumbnailCanvas');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#eef5fc';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    const polygon = state.fileAnalysis && Array.isArray(state.fileAnalysis.polygon) && state.fileAnalysis.polygon.length >= 3
        ? state.fileAnalysis.polygon : null;
    if (!polygon) {
        ctx.strokeStyle = '#9ab2ca';
        ctx.lineWidth = 3;
        ctx.strokeRect(68, 42, 88, 68);
        ctx.beginPath();
        ctx.moveTo(82, 76);
        ctx.lineTo(142, 76);
        ctx.stroke();
        return;
    }
    const xs = polygon.map(point => Number(point[0] || 0));
    const ys = polygon.map(point => Number(point[1] || 0));
    const minX = Math.min(...xs), maxX = Math.max(...xs);
    const minY = Math.min(...ys), maxY = Math.max(...ys);
    const width = Math.max(1, maxX - minX);
    const height = Math.max(1, maxY - minY);
    const padding = 18;
    const scale = Math.min((canvas.width - padding * 2) / width, (canvas.height - padding * 2) / height);
    const offsetX = (canvas.width - width * scale) / 2;
    const offsetY = (canvas.height - height * scale) / 2;
    ctx.beginPath();
    polygon.forEach((point, index) => {
        const x = offsetX + (Number(point[0] || 0) - minX) * scale;
        const y = offsetY + (Number(point[1] || 0) - minY) * scale;
        if (index === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
    });
    ctx.closePath();
    ctx.fillStyle = 'rgba(21, 101, 192, 0.16)';
    ctx.fill();
    ctx.strokeStyle = '#1565c0';
    ctx.lineWidth = 3;
    ctx.stroke();
}

function initLaserTechnologyColumnResize() {
    const table = document.querySelector('.technology-comparison');
    if (!table || table.dataset.resizable === '1') return;
    const columns = [...table.querySelectorAll('col')];
    const headers = [...table.querySelectorAll('thead th')];
    let savedWidths = [];
    try { savedWidths = JSON.parse(localStorage.getItem(LASER_COLUMN_WIDTHS_KEY) || '[]'); } catch (_) {}
    function applyWidths(widths) {
        columns.forEach((column, index) => {
            const width = Math.max(64, Number(widths[index]) || parseFloat(column.style.width) || 100);
            column.style.width = width + 'px';
        });
        table.style.width = columns.reduce((sum, column) => sum + parseFloat(column.style.width), 0) + 'px';
    }
    function saveWidths() {
        const widths = columns.map(column => Math.round(parseFloat(column.style.width)));
        try { localStorage.setItem(LASER_COLUMN_WIDTHS_KEY, JSON.stringify(widths)); } catch (_) {}
    }
    applyWidths(savedWidths);
    headers.forEach((header, index) => {
        const handle = document.createElement('span');
        handle.className = 'column-resizer';
        handle.tabIndex = 0;
        handle.setAttribute('role', 'separator');
        handle.setAttribute('aria-label', 'Tilpas kolonnebredde for ' + header.textContent.trim());
        function resizeBy(delta) {
            const widths = columns.map(column => parseFloat(column.style.width));
            widths[index] = Math.max(64, widths[index] + delta);
            applyWidths(widths);
        }
        handle.addEventListener('pointerdown', event => {
            event.preventDefault();
            const startX = event.clientX;
            const startWidth = parseFloat(columns[index].style.width);
            handle.setPointerCapture(event.pointerId);
            const onPointerMove = moveEvent => {
                const widths = columns.map(column => parseFloat(column.style.width));
                widths[index] = Math.max(64, startWidth + moveEvent.clientX - startX);
                applyWidths(widths);
            };
            const onPointerUp = () => {
                handle.removeEventListener('pointermove', onPointerMove);
                handle.removeEventListener('pointerup', onPointerUp);
                saveWidths();
            };
            handle.addEventListener('pointermove', onPointerMove);
            handle.addEventListener('pointerup', onPointerUp);
        });
        handle.addEventListener('keydown', event => {
            if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
            event.preventDefault();
            resizeBy(event.key === 'ArrowLeft' ? -10 : 10);
            saveWidths();
        });
        header.appendChild(handle);
    });
    table.dataset.resizable = '1';
}

function calcWizardComplete(step) {
    if (step === 'customer') return Boolean(state.calcCustomer);
    if (step === 'drawing') return Boolean(state.fileAnalysis);
    if (step === 'material') return Boolean(state.calcMaterial);
    if (step === 'processes') return Boolean(state.calcWizardProcessesReady || state.lastQuote);
    return Boolean(state.lastQuote);
}

function updateProductRecap() {
    const customer = state.calcCustomer;
    const material = state.calcMaterial;
    const drawing = state.fileAnalysis;
    const quote = state.lastQuote && state.lastQuote.result;
    const qty = Math.max(1, Number((document.getElementById('calcQty') || {}).value || 1));
    const activeProcesses = processCards ? [...processCards.querySelectorAll('.proc-card')]
        .filter(card => card.querySelector('.proc-toggle') && card.querySelector('.proc-toggle').checked)
        .map(card => (processDefs.find(def => def.key === card.dataset.proc) || {}).label || card.dataset.proc) : [];
    const setText = (id, value) => { const el = document.getElementById(id); if (el) el.textContent = value; };
    setText('recapProduct', drawing && drawing.filename ? drawing.filename : 'Ny beregning');
    setText('recapDrawing', drawing ? (formatMoney(drawing.widthMm) + ' × ' + formatMoney(drawing.lengthMm) + ' mm') : 'Ingen tegning indlæst');
    setText('recapCustomer', customer ? (customer.Nm || customer.CustNo || '-') : 'Ikke valgt');
    setText('recapCustomerNo', customer ? ('Kundenr. ' + (customer.CustNo || '-')) : 'Standard prisliste');
    setText('recapMaterial', material ? (material.ProdNo || '-') : 'Ikke valgt');
    setText('recapMaterialMeta', material
        ? ((material.beskrivelse || material.Descr || '') + ' · ' + formatMoney(material.tykklese) + ' mm')
        : 'Materiale og tykkelse');
    setText('recapQuantity', qty + ' stk');
    setText('recapProcesses', activeProcesses.length ? activeProcesses.map(label => label.split(' ')[0]).join(' · ') : 'Ingen processer');
    setText('recapPrice', quote ? (formatMoney(quote.perPiece.unitPrice) + ' DKK/stk') : '-');
    setText('recapTotal', quote ? (formatMoney(quote.total.totalPrice) + ' DKK i alt') : 'Ikke beregnet');
    renderPieceThumbnail();
}

function updateCalcWizard() {
    const statuses = {
        customer: state.calcCustomer ? (state.calcCustomer.Nm || state.calcCustomer.CustNo) : 'Vælg kunde',
        drawing: state.fileAnalysis ? ((state.fileAnalysis.filename || 'Tegning') + ' indlæst') : 'Indlæs fil',
        material: state.calcMaterial ? ((state.calcMaterial.ProdNo || '') + ' · ' + formatMoney(state.calcMaterial.tykklese) + ' mm') : 'Vælg materiale',
        processes: calcWizardComplete('processes') ? 'Operationer kontrolleret' : 'Kontrollér operationer',
        result: state.lastQuote ? (formatMoney(state.lastQuote.result.perPiece.unitPrice) + ' DKK/stk') : 'Afventer beregning'
    };
    document.querySelectorAll('[data-calc-step]').forEach(button => {
        const step = button.dataset.calcStep;
        button.classList.toggle('active', step === activeCalcWizardStep);
        button.classList.toggle('complete', calcWizardComplete(step));
        const status = document.getElementById('calcStep' + step.charAt(0).toUpperCase() + step.slice(1) + 'Status');
        if (status) status.textContent = statuses[step];
    });
    updateProductRecap();
}

function closeCalcStep() {
    const modal = document.getElementById('calcStepModal');
    const body = document.getElementById('calcStepModalBody');
    const store = document.getElementById('calcInputStore');
    if (body && store && body.firstElementChild) store.appendChild(body.firstElementChild);
    if (modal) modal.classList.remove('open');
    activeCalcWizardStep = '';
    updateCalcWizard();
}

function openCalcStep(step) {
    if (step === 'result') {
        closeCalcStep();
        const result = document.getElementById('calcResultAnchor');
        if (result) result.scrollIntoView({ behavior: 'smooth', block: 'start' });
        return;
    }
    const panel = document.getElementById('calcStep' + step.charAt(0).toUpperCase() + step.slice(1) + 'Panel');
    const modal = document.getElementById('calcStepModal');
    const body = document.getElementById('calcStepModalBody');
    if (!panel || !modal || !body || !calcWizardMeta[step]) return;
    closeCalcStep();
    activeCalcWizardStep = step;
    document.getElementById('calcStepModalTitle').textContent = calcWizardMeta[step][0];
    document.getElementById('calcStepModalSubtitle').textContent = calcWizardMeta[step][1];
    body.appendChild(panel);
    const stepIndex = calcWizardSteps.indexOf(step);
    const back = document.getElementById('calcStepBackBtn');
    const next = document.getElementById('calcStepNextBtn');
    back.hidden = stepIndex === 0;
    next.textContent = step === 'processes' ? 'Beregn pris' : 'Næste';
    next.disabled = step === 'processes' ? false : !calcWizardComplete(step);
    modal.classList.add('open');
    updateCalcWizard();
    const firstInput = panel.querySelector('input:not([disabled]), select:not([disabled])');
    if (firstInput) setTimeout(() => firstInput.focus(), 80);
}

function advanceCalcWizard(completedStep) {
    updateCalcWizard();
    if (!calcWizardComplete(completedStep)) return;
    const nextStep = calcWizardSteps[calcWizardSteps.indexOf(completedStep) + 1];
    setTimeout(() => openCalcStep(nextStep), 120);
}

function initCalcWizard() {
    document.querySelectorAll('[data-calc-step]').forEach(button => {
        if (button.dataset.wizardBound) return;
        button.dataset.wizardBound = '1';
        button.addEventListener('click', () => openCalcStep(button.dataset.calcStep));
    });
    if (!state.calcWizardStarted) {
        state.calcWizardStarted = true;
        setTimeout(() => openCalcStep('customer'), 100);
    }
    updateCalcWizard();
}

function rateFromResource(row) {
    // PrDcMat: CstPr = kostpris/min, SalePr = salgspris/min. Følger valgt pristype.
    const basis = document.getElementById('calcPriceBasis') ? document.getElementById('calcPriceBasis').value : 'sale';
    const sale = Number(row.SalePr || 0);
    const cst = Number(row.CstPr || 0);
    const v = basis === 'cost' ? (cst || sale) : (sale || cst);
    return v > 0 ? Math.round(v * 100) / 100 : 0;
}
function resourcesForDef(def) {
    const seen = new Set();
    return state.resources.filter(row => {
        if (seen.has(row.ProdNo)) return false;
        seen.add(row.ProdNo);
        if (!def.r7) return true; // 'andet' må vælge alt
        return def.r7.includes(String(row.R7 || '').trim());
    });
}
function findResource(prodNo) {
    return state.resources.find(r => String(r.ProdNo) === String(prodNo)) || null;
}

// ── Procesvise felter og automatiske tider ──
function svejsAutoMinutes(card) {
    const lgd = Number(card.querySelector('.svejs-lgd').value || 0);       // mm
    const hast = Number(card.querySelector('.svejs-hast').value || 0);     // mm/min
    const efter = Number(card.querySelector('.svejs-efter').value || 0);   // min
    const svejsMin = hast > 0 ? lgd / hast : 0;
    return Math.round((svejsMin + efter) * 100) / 100;
}
const SVEJS_SPEEDS = { mig: 350, tig: 150, punkt: 500 };

function buildProcessCards() {
    if (processCards.dataset.built) return;
    processCards.dataset.built = '1';
    processDefs.forEach(def => {
        const card = document.createElement('div');
        card.className = 'proc-card' + (def.defaultOn ? ' on' : '');
        card.dataset.proc = def.key;
        if (def.isLaser) {
            card.innerHTML = '<label class="proc-head"><input type="checkbox" class="proc-toggle" checked /> ' + escapeHtml(def.label) + ' <span class="proc-res-label">tid beregnes ud fra skæreparametre — kan rettes</span></label>'
                + '<div class="proc-body"><div class="proc-fields">'
                + '<div class="field"><label>Maskine</label><input class="proc-machine" value="R1100" /></div>'
                + '<div class="field"><label>Teknologi</label><select class="proc-technology"><option value="">Automatisk · billigste teknologi</option></select></div>'
                + '<div class="field"><label>Globale gaspriser</label><div class="proc-gas-prices muted">Azot 0,00 · Oxygen 0,00 DKK/Nm³</div></div>'
                + '<div class="field"><label>Minutsats (dkk/min)</label><input class="proc-rate" type="number" step="0.5" value="12" /></div>'
                + '<div class="field"><label>Tid (min/emne) — tom = auto</label><input class="proc-time" type="number" step="0.01" min="0" placeholder="auto" /></div>'
                + '<div class="field"><label>Afstand mellem emner (mm)</label><input class="laser-gap" type="number" step="1" min="0" value="5" /></div>'
                + '<div class="field"><label>Afstand til pladekant (mm)</label><input class="laser-margin" type="number" step="1" min="0" value="10" /></div>'
                + '</div></div>';
        } else if (def.kind === 'stykliste') {
            card.innerHTML = '<label class="proc-head"><input type="checkbox" class="proc-toggle" /> ' + escapeHtml(def.label) + ' <span class="proc-res-label">komponenter pr emne — som R8200 i Excel</span></label>'
                + '<div class="proc-body">'
                + '<div class="picker"><input class="comp-search" placeholder="søg komponent (varenr eller tekst)..." autocomplete="off" /><div class="picker-list"></div></div>'
                + '<div class="proc-fields">'
                + '<div class="field"><label>Antal pr emne</label><input class="comp-qty" type="number" step="1" min="1" value="1" /></div>'
                + '<div class="field"><label>&nbsp;</label><button type="button" class="comp-add alt">+ Tilføj komponent</button></div>'
                + '</div>'
                + '<div class="comp-lines muted">Ingen komponenter tilføjet.</div>'
                + '</div>';
        } else if (def.kind === 'flad') {
            card.innerHTML = '<label class="proc-head"><input type="checkbox" class="proc-toggle" /> ' + escapeHtml(def.label) + ' <span class="proc-res-label"></span></label>'
                + '<div class="proc-body"><div class="proc-fields">'
                + '<div class="field"><label>Type (R05 / R10 / R15 / R20 ...)</label><select class="flad-type"></select></div>'
                + '<div class="field"><label>Hastighed</label><input class="flad-speed" type="number" step="0.1" min="0.01" value="1" /></div>'
                + '<div class="field"><label>Faktor</label><input class="flad-factor" type="number" step="0.1" min="0.01" value="10" /></div>'
                + '<div class="field"><label>Minutter pr emne (tom = auto)</label><input class="proc-min" type="number" step="0.000001" min="0" placeholder="auto" /></div>'
                + '<div class="field"><label>Sats (dkk/min)</label><input class="proc-rate" type="number" step="0.1" value="9.38" /></div>'
                + '<div class="field"><label>Opstart (min/ordre)</label><input class="proc-opstart" type="number" step="1" min="0" value="0" /></div>'
                + '</div></div>';
        } else if (def.kind === 'buk') {
            card.innerHTML = '<label class="proc-head"><input type="checkbox" class="proc-toggle" /> ' + escapeHtml(def.label) + ' <span class="proc-res-label"></span></label>'
                + '<div class="proc-body">'
                + '<div class="picker"><input class="proc-res-search" placeholder="søg buk-ressource..." autocomplete="off" /><div class="picker-list"></div></div>'
                + '<div class="proc-fields">'
                + '<div class="field"><label>Antal buk pr emne</label><input class="buk-antal" type="number" step="1" min="0" value="2" /></div>'
                + '<div class="field"><label>Samlet bukkelængde (mm)</label><input class="buk-laengde" type="number" step="10" min="0" value="1000" /></div>'
                + '<div class="field"><label>Gennemsnitlig vinkel</label><input class="buk-vinkel" type="number" step="1" min="1" max="180" value="90" /></div>'
                + '<div class="field"><label>V-åbning (mm)</label><input class="buk-v-aabning" type="number" step="1" min="0.1" placeholder="auto: 8 × tykkelse" /></div>'
                + '<div class="field"><label>Trækstyrke (MPa)</label><input class="buk-traekstyrke" type="number" step="10" min="1" value="450" /></div>'
                + '<div class="field"><label>90° rotationer</label><input class="buk-rotationer" type="number" step="1" min="0" value="1" /></div>'
                + '<div class="field"><label>Vendinger</label><input class="buk-vendinger" type="number" step="1" min="0" value="0" /></div>'
                + '<div class="field"><label>Minutter pr emne (tom = auto)</label><input class="proc-min" type="number" step="0.01" min="0" placeholder="auto" /></div>'
                + '<div class="field"><label>Sats (dkk/min)</label><input class="proc-rate" type="number" step="0.1" value="10.31" /></div>'
                + '<div class="field"><label>Opstart (tom = maskinparameter)</label><input class="proc-opstart" type="number" step="1" min="0" placeholder="auto" /></div>'
                + '</div></div>';
        } else if (def.kind === 'svejs') {
            card.innerHTML = '<label class="proc-head"><input type="checkbox" class="proc-toggle" /> ' + escapeHtml(def.label) + ' <span class="proc-res-label"></span></label>'
                + '<div class="proc-body">'
                + '<div class="picker"><input class="proc-res-search" placeholder="søg svejse-ressource..." autocomplete="off" /><div class="picker-list"></div></div>'
                + '<div class="proc-fields">'
                + '<div class="field"><label>Metode</label><select class="svejs-metode"><option value="mig">MIG/MAG</option><option value="tig">TIG</option><option value="punkt">Punktsvejs</option></select></div>'
                + '<div class="field"><label>Svejselængde (mm pr emne)</label><input class="svejs-lgd" type="number" step="10" min="0" value="200" /></div>'
                + '<div class="field"><label>Hastighed (mm/min)</label><input class="svejs-hast" type="number" step="10" min="0" value="350" /></div>'
                + '<div class="field"><label>Efterarbejde (min/emne)</label><input class="svejs-efter" type="number" step="0.5" min="0" value="1" /></div>'
                + '<div class="field"><label>Minutter pr emne (auto — kan rettes)</label><input class="proc-min" type="number" step="0.01" value="1.57" /></div>'
                + '<div class="field"><label>Sats (dkk/min)</label><input class="proc-rate" type="number" step="0.1" value="9.38" /></div>'
                + '<div class="field"><label>Opstart (min/ordre)</label><input class="proc-opstart" type="number" step="1" min="0" value="0" /></div>'
                + '</div></div>';
        } else {
            card.innerHTML = '<label class="proc-head"><input type="checkbox" class="proc-toggle" /> ' + escapeHtml(def.label) + ' <span class="proc-res-label"></span></label>'
                + '<div class="proc-body">'
                + '<div class="picker"><input class="proc-res-search" placeholder="søg ressource (fx montage, valse, save)..." autocomplete="off" /><div class="picker-list"></div></div>'
                + '<div class="proc-fields">'
                + '<div class="field"><label>Minutter pr emne</label><input class="proc-min" type="number" step="0.1" value="1" /></div>'
                + '<div class="field"><label>Sats (dkk/min)</label><input class="proc-rate" type="number" step="0.1" value="9.38" /></div>'
                + '<div class="field"><label>Opstart (min/ordre)</label><input class="proc-opstart" type="number" step="1" min="0" value="0" /></div>'
                + '</div></div>';
        }
        const toggle = card.querySelector('.proc-toggle');
        toggle.addEventListener('change', () => card.classList.toggle('on', toggle.checked));
        wireProcessCard(card, def);
        processCards.appendChild(card);
    });
}

function getLaserCard() {
    return processCards.querySelector('.proc-card[data-proc="laser"]');
}

function syncSpacingFromGlobalToLaserCard() {
    const card = getLaserCard();
    if (!card) return;
    const gapInput = card.querySelector('.laser-gap');
    const marginInput = card.querySelector('.laser-margin');
    if (gapInput) gapInput.value = Number(document.getElementById('calcGap').value || 5);
    if (marginInput) marginInput.value = Number(document.getElementById('calcMargin').value || 10);
}

function syncSpacingFromLaserCardToGlobal() {
    const card = getLaserCard();
    if (!card) return;
    const gapInput = card.querySelector('.laser-gap');
    const marginInput = card.querySelector('.laser-margin');
    const gap = Math.max(0, Number(gapInput ? gapInput.value : 5) || 0);
    const margin = Math.max(0, Number(marginInput ? marginInput.value : 10) || 0);
    document.getElementById('calcGap').value = gap;
    document.getElementById('calcMargin').value = margin;
}

function getNestingSpacing() {
    const card = getLaserCard();
    if (!card) {
        return {
            margin: Math.max(0, Number(document.getElementById('calcMargin').value || 10) || 0),
            gap: Math.max(0, Number(document.getElementById('calcGap').value || 5) || 0)
        };
    }
    const gapInput = card.querySelector('.laser-gap');
    const marginInput = card.querySelector('.laser-margin');
    const gap = Math.max(0, Number(gapInput ? gapInput.value : 5) || 0);
    const margin = Math.max(0, Number(marginInput ? marginInput.value : 10) || 0);
    return { margin, gap };
}

function wireProcessCard(card, def) {
    const resLabel = card.querySelector('.proc-res-label');
    function applyResource(row) {
        if (!row) return;
        card.dataset.prodNo = row.ProdNo || '';
        const rate = rateFromResource(row);
        if (resLabel) resLabel.textContent = 'ressource: ' + (row.ProdNo || '') + ' · kost ' + (row.CstPr == null ? '-' : row.CstPr) + ' / salg ' + (row.SalePr == null ? '-' : row.SalePr) + ' dkk/min';
        const rateInput = card.querySelector('.proc-rate');
        if (rateInput && rate > 0) rateInput.value = rate;
    }
    // Søge-picker afgrænset til processens R7-familie
    const searchInput = card.querySelector('.proc-res-search');
    if (searchInput) {
        const searchList = card.querySelector('.picker-list');
        attachPicker({
            input: searchInput,
            list: searchList,
            getRows: () => resourcesForDef(def),
            rowLabel: resourceRowLabel,
            rowSub: resourceRowSub,
            onPick: row => {
                searchInput.value = (row.ProdNo || '') + ' · ' + (row.Descr || '');
                applyResource(row);
            }
        });
        // forvalgt standardressource når kataloget er hentet
        card.dataset.defaultRes = def.defaultRes || '';
    }
    // Flad: typevalg som select
    const fladType = card.querySelector('.flad-type');
    if (fladType) {
        card.dataset.needsFladOptions = '1';
        fladType.addEventListener('change', () => {
            const resource = findResource(fladType.value);
            applyResource(resource);
            const typeMatch = String(resource && resource.Descr || '').toUpperCase().match(/\b(R05|R10|R15|R20|B05)\b/);
            card.dataset.flatType = typeMatch ? typeMatch[1] : '';
        });
    }
    // Svejs: metode sætter hastighed, auto-beregn minutter
    if (def.kind === 'svejs') {
        const metode = card.querySelector('.svejs-metode');
        metode.addEventListener('change', () => {
            card.querySelector('.svejs-hast').value = SVEJS_SPEEDS[metode.value] || 350;
            card.querySelector('.proc-min').value = svejsAutoMinutes(card);
        });
        ['svejs-lgd', 'svejs-hast', 'svejs-efter'].forEach(cls => {
            card.querySelector('.' + cls).addEventListener('input', () => {
                card.querySelector('.proc-min').value = svejsAutoMinutes(card);
            });
        });
    }
    // Stykliste (R8200): komponentlinjer
    if (def.kind === 'stykliste') {
        const compSearch = card.querySelector('.comp-search');
        const compList = card.querySelector('.picker-list');
        let pendingComponent = null;
        attachPicker({
            input: compSearch,
            list: compList,
            getRows: () => state.components,
            rowLabel: componentRowLabel,
            rowSub: componentRowSub,
            onPick: row => {
                pendingComponent = row;
                compSearch.value = (row.ProdNo || '') + ' · ' + (row.Descr || '');
            }
        });
        card.querySelector('.comp-add').addEventListener('click', () => {
            if (!pendingComponent) return;
            const qty = Math.max(1, Number(card.querySelector('.comp-qty').value || 1));
            const basis = document.getElementById('calcPriceBasis').value;
            const basePris = parseDaNumber(pendingComponent.Pris);
            const avance = Number(pendingComponent.Avance || 0);
            const unitPrice = Math.round((basis === 'cost' ? basePris : basePris * (1 + avance / 100)) * 100) / 100;
            state.calcComponents.push({
                prodNo: pendingComponent.ProdNo || '',
                descr: pendingComponent.Descr || '',
                qty,
                unitPrice
            });
            pendingComponent = null;
            compSearch.value = '';
            card.querySelector('.comp-qty').value = 1;
            renderComponentLines(card);
        });
    }
    if (def.isLaser) {
        const gapInput = card.querySelector('.laser-gap');
        const marginInput = card.querySelector('.laser-margin');
        if (gapInput) gapInput.addEventListener('input', syncSpacingFromLaserCardToGlobal);
        if (marginInput) marginInput.addEventListener('input', syncSpacingFromLaserCardToGlobal);
        syncSpacingFromGlobalToLaserCard();
    }
}
function renderComponentLines(cardArg) {
    const card = cardArg || processCards.querySelector('.proc-card[data-proc="stykliste"]');
    if (!card) return;
    const wrap = card.querySelector('.comp-lines');
    if (!state.calcComponents.length) {
        wrap.className = 'comp-lines muted';
        wrap.textContent = 'Ingen komponenter tilføjet.';
        return;
    }
    wrap.className = 'comp-lines';
    const total = state.calcComponents.reduce((s, c) => s + c.qty * c.unitPrice, 0);
    wrap.innerHTML = state.calcComponents.map((c, idx) =>
        '<div class="comp-line"><span><strong>' + escapeHtml(c.prodNo) + '</strong> ' + escapeHtml(c.descr) + '</span>'
        + '<span>' + c.qty + ' stk × <input type="number" step="0.01" min="0" class="comp-price" data-idx="' + idx + '" value="' + c.unitPrice + '" style="width:80px; padding:4px 6px;" /> dkk = ' + formatMoney(c.qty * c.unitPrice) + ' dkk</span>'
        + '<button type="button" class="alt comp-remove" data-idx="' + idx + '" style="padding:4px 8px; font-size:11px;">Fjern</button></div>'
    ).join('') + '<div class="comp-line" style="font-weight:700;"><span>Komponenter i alt pr emne</span><span>' + formatMoney(total) + ' dkk</span><span></span></div>';
    wrap.querySelectorAll('.comp-remove').forEach(btn => btn.addEventListener('click', () => {
        state.calcComponents.splice(Number(btn.getAttribute('data-idx')), 1);
        renderComponentLines(card);
    }));
    wrap.querySelectorAll('.comp-price').forEach(inp => inp.addEventListener('change', () => {
        const c = state.calcComponents[Number(inp.getAttribute('data-idx'))];
        if (c) c.unitPrice = Math.max(0, Number(inp.value || 0));
        renderComponentLines(card);
    }));
}
function populateProcessDefaults() {
    // udfyld flad-typer og standardressourcer når ressourcekataloget er klar
    processCards.querySelectorAll('.proc-card').forEach(card => {
        const def = processDefs.find(d => d.key === card.dataset.proc);
        if (!def) return;
        const fladType = card.querySelector('.flad-type');
        if (fladType && card.dataset.needsFladOptions && state.resources.length) {
            const rows = resourcesForDef(def);
            fladType.innerHTML = rows.map(r => '<option value="' + escapeHtml(r.ProdNo) + '">' + escapeHtml((r.ProdNo || '') + ' · ' + (r.Descr || '')) + '</option>').join('');
            const preferred = rows.find(r => String(r.ProdNo) === def.defaultRes)
                || rows.find(r => /\bR05\b/i.test(String(r.Descr || ''))) || rows[0];
            if (preferred) {
                fladType.value = preferred.ProdNo;
                fladType.dispatchEvent(new Event('change'));
            }
            delete card.dataset.needsFladOptions;
        }
        const searchInput = card.querySelector('.proc-res-search');
        if (searchInput && !card.dataset.prodNo && card.dataset.defaultRes && state.resources.length) {
            const row = findResource(card.dataset.defaultRes);
            if (row) {
                searchInput.value = (row.ProdNo || '') + ' · ' + (row.Descr || '');
                card.dataset.prodNo = row.ProdNo;
                const resLabel = card.querySelector('.proc-res-label');
                if (resLabel) resLabel.textContent = 'ressource: ' + row.ProdNo + ' · kost ' + (row.CstPr == null ? '-' : row.CstPr) + ' / salg ' + (row.SalePr == null ? '-' : row.SalePr) + ' dkk/min';
                const rate = rateFromResource(row);
                if (rate > 0) card.querySelector('.proc-rate').value = rate;
            }
        }
    });
    syncSpacingFromGlobalToLaserCard();
}
function collectOperations() {
    const ops = [];
    processCards.querySelectorAll('.proc-card').forEach(card => {
        const def = processDefs.find(d => d.key === card.dataset.proc);
        if (!def || def.isLaser || def.kind === 'stykliste') return;
        if (!card.querySelector('.proc-toggle').checked) return;
        let prodNo = card.dataset.prodNo || '';
        const fladType = card.querySelector('.flad-type');
        if (fladType) prodNo = fladType.value || prodNo;
        const minutesInput = card.querySelector('.proc-min');
        const opstartInput = card.querySelector('.proc-opstart');
        const operation = {
            key: def.key,
            label: def.label.split('(')[0].trim(),
            prodNo,
            minutes: Number(minutesInput.value || 0),
            rate: Number(card.querySelector('.proc-rate').value || 0),
            opstartMinutes: String(opstartInput.value || '').trim() === '' ? null : Number(opstartInput.value)
        };
        if (def.kind === 'flad') {
            const resource = findResource(prodNo);
            const typeMatch = String(resource && resource.Descr || '').toUpperCase().match(/\b(R05|R10|R15|R20|B05)\b/);
            operation.flatType = card.dataset.flatType || (typeMatch ? typeMatch[1] : '');
            operation.speed = Number(card.querySelector('.flad-speed').value || 0);
            operation.factor = Number(card.querySelector('.flad-factor').value || 0);
            operation.minutesOverride = String(minutesInput.value || '').trim() === '' ? null : Number(minutesInput.value);
        }
        if (def.kind === 'buk') {
            operation.minutesOverride = String(minutesInput.value || '').trim() === '' ? null : Number(minutesInput.value);
            operation.bendCount = Number(card.querySelector('.buk-antal').value || 0);
            operation.totalBendLengthMm = Number(card.querySelector('.buk-laengde').value || 0);
            operation.averageAngleDeg = Number(card.querySelector('.buk-vinkel').value || 90);
            operation.dieOpeningMm = String(card.querySelector('.buk-v-aabning').value || '').trim() === '' ? null : Number(card.querySelector('.buk-v-aabning').value);
            operation.tensileStrengthMpa = Number(card.querySelector('.buk-traekstyrke').value || 450);
            operation.rotate90Count = Number(card.querySelector('.buk-rotationer').value || 0);
            operation.flipCount = Number(card.querySelector('.buk-vendinger').value || 0);
        }
        ops.push(operation);
    });
    return ops;
}
function collectComponents() {
    const card = processCards.querySelector('.proc-card[data-proc="stykliste"]');
    if (!card || !card.querySelector('.proc-toggle').checked) return [];
    return state.calcComponents.slice();
}
function laserCardState() {
    const card = processCards.querySelector('.proc-card[data-proc="laser"]');
    if (!card) return { enabled: true, machine: 'R1100', rate: 12, technology: '', timeOverride: null };
    const timeRaw = String(card.querySelector('.proc-time').value || '').trim();
    return {
        enabled: card.querySelector('.proc-toggle').checked,
        machine: String(card.querySelector('.proc-machine').value || 'R1100').trim(),
        rate: Number(card.querySelector('.proc-rate').value || 12),
        geniusRate: rateFromResource(findResource('R1102') || {}) || Number(card.querySelector('.proc-rate').value || 12),
        technology: String(card.querySelector('.proc-technology').value || '').trim(),
        timeOverride: timeRaw === '' ? null : Number(timeRaw)
    };
}

function detectLaserMaterialFamily(material) {
    const prodNo = String(material && material.ProdNo || '').trim();
    const description = String(material && (material.beskrivelse || material.Descr) || '').toLowerCase();
    if (prodNo.startsWith('301')) return 'SORT';
    if (prodNo.startsWith('311')) return 'RF';
    if (prodNo.startsWith('321')) return 'AL';
    if (prodNo.startsWith('331')) return 'GAL';
    if (prodNo.startsWith('371')) return 'ME';
    if (prodNo.startsWith('381')) return description.includes('kobber') ? 'CO' : description.includes('messing') ? 'ME' : null;
    return null;
}

function refreshLaserTechnologyOptions() {
    const select = processCards.querySelector('.proc-card[data-proc="laser"] .proc-technology');
    if (!select) return;
    const selected = select.value;
    const materialFamily = detectLaserMaterialFamily(state.calcMaterial);
    const thickness = Number(state.calcMaterial && (state.calcMaterial.tykklese != null
        ? state.calcMaterial.tykklese : state.calcMaterial.Thickness));
    const compatibleRows = materialFamily && Number.isFinite(thickness)
        ? state.laserTechnicalParams.filter(row => String(row.Material || '').trim().toUpperCase() === materialFamily
            && Math.abs(Number(row.Thickness) - thickness) < 0.001)
        : [];
    select.innerHTML = '<option value="">Automatisk · billigste teknologi</option>' + compatibleRows.map(row =>
        '<option value="' + escapeHtml(row.Technology) + '">' + escapeHtml(row.Technology + ' · ' + row.Material + ' ' + row.Thickness + ' mm · ' + row.GasPressureBar + ' bar') + '</option>'
    ).join('');
    if (compatibleRows.some(row => row.Technology === selected)) select.value = selected;
    const prices = processCards.querySelector('.proc-card[data-proc="laser"] .proc-gas-prices');
    if (prices) prices.textContent = 'Azot ' + formatMoney(state.laserGasPrices.nitrogenPricePerKg)
        + ' · Oxygen ' + formatMoney(state.laserGasPrices.oxygenPricePerKg) + ' DKK/kg'
        + ' · MixLine ' + formatMoney(state.laserGasPrices.mixLineOxygenPercent) + '% O2';
}

async function primeBeregner() {
    buildProcessCards();
    initLaserTechnologyColumnResize();
    try {
        await Promise.all([ensureMaterials(), ensureResources(), ensureCustomers(), ensureComponents(), (async () => {
            if (state.laserTechnicalParams.length) return;
            const data = await fetchJson('/bom/calculators/laser-params');
            state.laserTechnicalParams = data.technicalRows || [];
            state.laserGasPrices = data.gasPrices || state.laserGasPrices;
        })()]);
        refreshLaserTechnologyOptions();
        populateProcessDefaults();
        initDxfViewerInteractions();
        renderDxfViewer();
        initCalcWizard();
    } catch (err) {
        calcMaterialChosen.textContent = 'Fejl ved hentning: ' + err.message;
    }
}

// ── Fil-analyse (DXF / STEP / PDF) ──
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result).split(',').pop());
        reader.onerror = () => reject(new Error('Kunne ikke læse filen'));
        reader.readAsDataURL(file);
    });
}
async function analyzeDrawingFile(file) {
    fileAnalysisStatus.textContent = 'Analyserer ' + file.name + '...';
    fileAnalysisGrid.innerHTML = '';
    try {
        const data = await fileToBase64(file);
        const result = await fetchJson('/bom/analyze-file', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filename: file.name, data })
        });
        applyDrawingAnalysis(file.name, result);
        return true;
    } catch (err) {
        fileAnalysisStatus.textContent = 'Fejl: ' + err.message;
        resetDxfMeasure(true);
        renderDxfViewer();
        updateCalcWizard();
        return false;
    }
}

function applyDrawingAnalysis(filename, result) {
    state.fileAnalysis = { ...result, filename };
    resetDxfMeasure(true);
    const kvs = [
        ['Format', (result.format || '').toUpperCase()],
        ['Bredde (mm)', result.widthMm == null ? '-' : result.widthMm],
        ['Længde (mm)', result.lengthMm == null ? '-' : result.lengthMm],
        ['Tykkelse (mm)', result.thicknessMm == null ? '-' : result.thicknessMm],
        ['Skærelængde (m)', result.cutLengthM == null ? '-' : result.cutLengthM],
        ['Piercings (estimat)', result.piercingsEstimate == null ? '-' : result.piercingsEstimate],
        ['Form til nesting', (result.polygon && result.polygon.length >= 3) ? 'fundet (' + result.polygon.length + ' punkter)' : 'nej — bruger rektangel']
    ];
    fileAnalysisGrid.innerHTML = kvs.map(([label, value]) => '<div class="kv"><label>' + escapeHtml(label) + '</label><div>' + escapeHtml(value) + '</div></div>').join('');
    fileAnalysisStatus.textContent = result.note || (filename + ' analyseret — felterne er udfyldt nedenfor.');
    if (result.widthMm) document.getElementById('calcPieceW').value = result.widthMm;
    if (result.lengthMm) document.getElementById('calcPieceL').value = result.lengthMm;
    if (result.cutLengthM) document.getElementById('calcCutLength').value = result.cutLengthM;
    if (result.piercingsEstimate) document.getElementById('calcPiercings').value = result.piercingsEstimate;
    renderDxfViewer();
    updateCalcWizard();
    advanceCalcWizard('drawing');
}

function initDxfViewerInteractions() {
    if (dxfMeasureState.eventsBound) return;
    const canvas = document.getElementById('dxfViewerCanvas');
    const resetBtn = document.getElementById('dxfMeasureResetBtn');
    if (!canvas) return;
    canvas.style.cursor = 'crosshair';
    canvas.addEventListener('mousemove', evt => {
        if (!dxfMeasureState.projection) return;
        const rect = canvas.getBoundingClientRect();
        if (!rect.width || !rect.height) return;
        const sx = canvas.width / rect.width;
        const sy = canvas.height / rect.height;
        const px = (evt.clientX - rect.left) * sx;
        const py = (evt.clientY - rect.top) * sy;
        const p = dxfCanvasToMm(px, py);
        if (!p) return;

        const snap = getSmartSnapPoint(p, dxfMeasureState.pointA, 14 / Math.max(1, dxfMeasureState.projection.scale));
        dxfMeasureState.hoverPoint = snap.point;
        dxfMeasureState.hoverKind = snap.kind;
        renderDxfViewer();
    });
    canvas.addEventListener('mouseleave', () => {
        dxfMeasureState.hoverPoint = null;
        dxfMeasureState.hoverKind = '';
        renderDxfViewer();
    });
    canvas.addEventListener('click', evt => {
        if (!dxfMeasureState.projection) return;
        const rect = canvas.getBoundingClientRect();
        if (!rect.width || !rect.height) return;
        const sx = canvas.width / rect.width;
        const sy = canvas.height / rect.height;
        const px = (evt.clientX - rect.left) * sx;
        const py = (evt.clientY - rect.top) * sy;
        const p = dxfCanvasToMm(px, py);
        if (!p) return;

        const snap = getSmartSnapPoint(p, dxfMeasureState.pointA, 14 / Math.max(1, dxfMeasureState.projection.scale));
        const selectedPoint = snap.point;
        if (!dxfMeasureState.pointA || (dxfMeasureState.pointA && dxfMeasureState.pointB)) {
            dxfMeasureState.pointA = selectedPoint;
            dxfMeasureState.pointB = null;
        } else {
            dxfMeasureState.pointB = selectedPoint;
        }
        renderDxfViewer();
    });
    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            resetDxfMeasure();
            renderDxfViewer();
        });
    }
    dxfMeasureState.eventsBound = true;
}

function resetDxfMeasure(skipRender) {
    dxfMeasureState.pointA = null;
    dxfMeasureState.pointB = null;
    dxfMeasureState.hoverPoint = null;
    dxfMeasureState.hoverKind = '';
    if (!skipRender) renderDxfViewer();
}

function dxfCanvasToMm(px, py) {
    const pr = dxfMeasureState.projection;
    if (!pr || !pr.scale) return null;
    const x = ((px - pr.ox) / pr.scale) + pr.minX;
    const y = ((py - pr.oy) / pr.scale) + pr.minY;
    return [x, y];
}

function dxfMmToCanvas(pt) {
    const pr = dxfMeasureState.projection;
    if (!pr || !pr.scale) return null;
    return [
        pr.ox + (pt[0] - pr.minX) * pr.scale,
        pr.oy + (pt[1] - pr.minY) * pr.scale
    ];
}

function nearestPointOnSegment(pt, a, b) {
    const ax = Number(a[0] || 0), ay = Number(a[1] || 0);
    const bx = Number(b[0] || 0), by = Number(b[1] || 0);
    const abx = bx - ax;
    const aby = by - ay;
    const denom = abx * abx + aby * aby;
    if (denom <= 1e-9) return [ax, ay];
    const tRaw = ((pt[0] - ax) * abx + (pt[1] - ay) * aby) / denom;
    const t = Math.max(0, Math.min(1, tRaw));
    return [ax + abx * t, ay + aby * t];
}

function scoreCandidateForSnap(candidate, cursorPt, anchorPt) {
    const dx = candidate[0] - cursorPt[0];
    const dy = candidate[1] - cursorPt[1];
    const dist = Math.hypot(dx, dy);
    if (!anchorPt) return dist;

    // If first point is set, reward candidates that align horizontally/vertically for "normal" measurements.
    const ax = Math.abs(candidate[0] - anchorPt[0]);
    const ay = Math.abs(candidate[1] - anchorPt[1]);
    const orthoPenalty = Math.min(ax, ay) * 0.30;
    return dist + orthoPenalty;
}

function getSmartSnapPoint(pt, anchorPt, radiusMm) {
    const polygon = (state.fileAnalysis && Array.isArray(state.fileAnalysis.polygon)) ? state.fileAnalysis.polygon : [];
    let best = [pt[0], pt[1]];
    let bestKind = 'fri';
    let bestScore = Infinity;

    const vertexRadius = Math.max(radiusMm, 0.01);
    const edgeRadius = Math.max(radiusMm * 0.8, 0.01);

    // 1) Try snapping to vertices (primary preference)
    for (let i = 0; i < polygon.length; i += 1) {
        const p = polygon[i];
        const cand = [Number(p[0] || 0), Number(p[1] || 0)];
        const d = Math.hypot(cand[0] - pt[0], cand[1] - pt[1]);
        if (d <= vertexRadius) {
            const s = scoreCandidateForSnap(cand, pt, anchorPt) - 0.08; // small bias for corners
            if (s < bestScore) {
                bestScore = s;
                best = cand;
                bestKind = 'hjørne';
            }
        }
    }

    // 2) If no good vertex, snap to nearest edge point
    if (polygon.length >= 2) {
        for (let i = 0; i < polygon.length; i += 1) {
            const a = polygon[i];
            const b = polygon[(i + 1) % polygon.length];
            const cand = nearestPointOnSegment(pt, a, b);
            const d = Math.hypot(cand[0] - pt[0], cand[1] - pt[1]);
            if (d <= edgeRadius) {
                const s = scoreCandidateForSnap(cand, pt, anchorPt);
                if (s < bestScore) {
                    bestScore = s;
                    best = cand;
                    bestKind = 'kant';
                }
            }
        }
    }

    return { point: best, kind: bestKind };
}

function renderDxfViewer() {
    const canvas = document.getElementById('dxfViewerCanvas');
    const meta = document.getElementById('dxfViewerMeta');
    const measureInfo = document.getElementById('dxfMeasureInfo');
    if (!canvas || !meta) return;
    canvas.style.cursor = 'crosshair';
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#0e1722';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    const polygon = (state.fileAnalysis && Array.isArray(state.fileAnalysis.polygon) && state.fileAnalysis.polygon.length >= 3)
        ? state.fileAnalysis.polygon
        : null;
    if (!polygon) {
        meta.textContent = 'Ingen DXF-kontur klar';
        if (measureInfo) measureInfo.textContent = 'Klik på to punkter i konturen for at måle afstand (mm)';
        dxfMeasureState.projection = null;
        dxfMeasureState.hoverPoint = null;
        dxfMeasureState.hoverKind = '';
        ctx.fillStyle = '#9fb3c8';
        ctx.font = '13px system-ui, sans-serif';
        ctx.fillText('Upload en DXF-fil for 2D visning og måling.', 18, 24);
        return;
    }

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    polygon.forEach(p => {
        const x = Number(p[0] || 0), y = Number(p[1] || 0);
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
    });
    const w = Math.max(1, maxX - minX);
    const h = Math.max(1, maxY - minY);
    const pad = 24;
    const scale = Math.min((canvas.width - pad * 2) / w, (canvas.height - pad * 2) / h);
    const ox = (canvas.width - w * scale) / 2;
    const oy = (canvas.height - h * scale) / 2;
    dxfMeasureState.projection = { minX, minY, scale, ox, oy };

    const project2D = (pt) => {
        const x = ox + (pt[0] - minX) * scale;
        const y = oy + (pt[1] - minY) * scale;
        return [x, y];
    };

    const pts = polygon.map(project2D);
    ctx.fillStyle = 'rgba(47,129,247,0.20)';
    ctx.beginPath();
    pts.forEach((p, i) => i === 0 ? ctx.moveTo(p[0], p[1]) : ctx.lineTo(p[0], p[1]));
    ctx.closePath();
    ctx.fill();
    ctx.strokeStyle = '#5ba0f7';
    ctx.lineWidth = 2;
    ctx.stroke();

    meta.textContent = '2D visning · konturpunkter ' + polygon.length;
    if (measureInfo) {
        if (dxfMeasureState.pointA && dxfMeasureState.pointB) {
            const dx = dxfMeasureState.pointB[0] - dxfMeasureState.pointA[0];
            const dy = dxfMeasureState.pointB[1] - dxfMeasureState.pointA[1];
            const d = Math.hypot(dx, dy);
            measureInfo.textContent = 'Måling: ' + d.toFixed(1) + ' mm (ΔX ' + dx.toFixed(1) + ' · ΔY ' + dy.toFixed(1) + ')';
        } else if (dxfMeasureState.pointA) {
            const snapTxt = dxfMeasureState.hoverKind ? ' · snap ' + dxfMeasureState.hoverKind : '';
            measureInfo.textContent = 'Punkt A valgt · klik punkt B' + snapTxt;
        } else {
            const snapTxt = dxfMeasureState.hoverKind ? ' · snap ' + dxfMeasureState.hoverKind : '';
            measureInfo.textContent = 'Klik på to punkter i konturen for at måle afstand (mm)' + snapTxt;
        }
    }

    const a = dxfMeasureState.pointA ? dxfMmToCanvas(dxfMeasureState.pointA) : null;
    const b = dxfMeasureState.pointB ? dxfMmToCanvas(dxfMeasureState.pointB) : null;
    const hoverCanvasPt = dxfMeasureState.hoverPoint ? dxfMmToCanvas(dxfMeasureState.hoverPoint) : null;
    if (hoverCanvasPt) {
        ctx.strokeStyle = '#ffe58a';
        ctx.lineWidth = 1.2;
        ctx.setLineDash([3, 3]);
        ctx.beginPath();
        ctx.arc(hoverCanvasPt[0], hoverCanvasPt[1], 8, 0, Math.PI * 2);
        ctx.stroke();
        ctx.setLineDash([]);
    }
    if (a) {
        ctx.fillStyle = '#ffe58a';
        ctx.beginPath();
        ctx.arc(a[0], a[1], 4, 0, Math.PI * 2);
        ctx.fill();
        ctx.fillStyle = '#ffe58a';
        ctx.font = '11px system-ui, sans-serif';
        ctx.fillText('A', a[0] + 7, a[1] - 6);
    }
    if (b) {
        ctx.fillStyle = '#9ff3d0';
        ctx.beginPath();
        ctx.arc(b[0], b[1], 4, 0, Math.PI * 2);
        ctx.fill();
        ctx.fillStyle = '#9ff3d0';
        ctx.font = '11px system-ui, sans-serif';
        ctx.fillText('B', b[0] + 7, b[1] - 6);
    }
    if (a && b) {
        ctx.strokeStyle = '#ffd166';
        ctx.lineWidth = 1.6;
        ctx.setLineDash([6, 4]);
        ctx.beginPath();
        ctx.moveTo(a[0], a[1]);
        ctx.lineTo(b[0], b[1]);
        ctx.stroke();
        ctx.setLineDash([]);
    }
}

function resolveNestingVisualQty(result) {
    const mode = String((document.getElementById('nestingQtyMode') || {}).value || 'order');
    const orderQty = Number(result && result.total ? result.total.qty : (document.getElementById('calcQty').value || 1));
    const perSheet = Number(result && result.nesting && result.nesting.best ? result.nesting.best.total : 1);
    const custom = Math.max(1, Number((document.getElementById('nestingCustomQty') || {}).value || 1));
    if (mode === 'full') return Math.max(1, perSheet || 1);
    if (mode === 'custom') return custom;
    return Math.max(1, orderQty || 1);
}

// ── Nesting-tegning på pladen ──
function drawNesting(n, qtyOverride) {
    const canvas = document.getElementById('nestingCanvas');
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    if (!n || !n.best || n.best.total <= 0) { canvas.style.display = 'none'; return; }
    canvas.style.display = 'block';
    const { sheetW, sheetL, pieceW, pieceL, margin, gap } = n.input;
    const totalQty = Math.max(1, Number(qtyOverride || document.getElementById('calcQty').value || 1));
    const piecesPerSheet = Math.max(1, Number(n.best.total || 1));
    const sheetsNeeded = Math.max(1, Math.ceil(totalQty / piecesPerSheet));
    const previewLimit = Math.max(1, Math.min(6, Number(document.getElementById('calcPreviewSheets').value || 3) || 3));
    const sheetsToDraw = Math.max(1, Math.min(sheetsNeeded, previewLimit));

    const cols = sheetsToDraw <= 2 ? sheetsToDraw : 2;
    const rows = Math.ceil(sheetsToDraw / cols);
    const outerPad = 10;
    const cellGap = 12;
    const cellW = (canvas.width - outerPad * 2 - cellGap * (cols - 1)) / cols;
    const cellH = (canvas.height - outerPad * 2 - cellGap * (rows - 1)) / rows;

    const drawSheet = (sheetIndex, x0, y0, wBox, hBox, piecesOnSheet) => {
        const pad = 8;
        const scale = Math.min((wBox - 2 * pad) / sheetL, (hBox - 2 * pad) / sheetW);
        const ox = x0 + (wBox - sheetL * scale) / 2;
        const oy = y0 + (hBox - sheetW * scale) / 2;
        const X = mm => ox + mm * scale;
        const Y = mm => oy + mm * scale;

        ctx.fillStyle = '#0f1722';
        ctx.fillRect(x0, y0, wBox, hBox);
        ctx.strokeStyle = '#36465a';
        ctx.lineWidth = 1;
        ctx.strokeRect(x0 + 0.5, y0 + 0.5, wBox - 1, hBox - 1);

        ctx.fillStyle = '#16202b';
        ctx.fillRect(X(0), Y(0), sheetL * scale, sheetW * scale);
        ctx.strokeStyle = '#3d4f63';
        ctx.lineWidth = 1.1;
        ctx.strokeRect(X(0), Y(0), sheetL * scale, sheetW * scale);

        let remaining = Math.max(0, piecesOnSheet);
        if (n.mode === 'shape' && Array.isArray(n.placements) && Array.isArray(n.polygon)) {
            const pcw = pieceW;
            const pch = pieceL;
            const rotPt = (p, rot) => {
                if (rot === 90) return [pch - p[1], p[0]];
                if (rot === 180) return [pcw - p[0], pch - p[1]];
                if (rot === 270) return [p[1], pcw - p[0]];
                return [p[0], p[1]];
            };
            const colors = { 0: ['rgba(47,129,247,0.5)', '#5ba0f7'], 90: ['rgba(242,163,60,0.5)', '#f2a33c'], 180: ['rgba(94,201,134,0.5)', '#5ec986'], 270: ['rgba(218,112,214,0.5)', '#da70d6'] };
            for (let i = 0; i < n.placements.length && remaining > 0; i += 1) {
                const plc = n.placements[i];
                const pts = n.polygon.map(p => rotPt(p, plc.rot));
                ctx.beginPath();
                pts.forEach((p, idx) => {
                    const px = X(plc.x + p[0]);
                    const py = Y(plc.y + p[1]);
                    if (idx === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
                });
                ctx.closePath();
                const c = colors[plc.rot] || colors[0];
                ctx.fillStyle = c[0];
                ctx.fill();
                ctx.strokeStyle = c[1];
                ctx.lineWidth = 0.9;
                ctx.stroke();
                remaining -= 1;
            }
        } else {
            const best = n.best;
            const pw = best.rotation ? pieceL : pieceW;
            const pl = best.rotation ? pieceW : pieceL;
            const drawPiece = (xMm, yMm, wMm, hMm, fill, stroke) => {
                ctx.fillStyle = fill;
                ctx.fillRect(X(xMm), Y(yMm), wMm * scale, hMm * scale);
                ctx.strokeStyle = stroke;
                ctx.lineWidth = 0.9;
                ctx.strokeRect(X(xMm), Y(yMm), wMm * scale, hMm * scale);
            };
            for (let r = 0; r < best.rows && remaining > 0; r += 1) {
                for (let c = 0; c < best.cols && remaining > 0; c += 1) {
                    drawPiece(margin + r * (pl + gap), margin + c * (pw + gap), pl, pw, 'rgba(47,129,247,0.55)', '#5ba0f7');
                    remaining -= 1;
                }
            }
            if (remaining > 0 && best.mixedExtra > 0 && best.rotation === 0) {
                const usedL = best.rows * (pieceL + gap) - gap;
                const startL = margin + usedL + gap;
                const usableW = sheetW - 2 * margin;
                const stripCols = Math.floor((usableW + gap) / (pieceL + gap));
                let drawn = 0;
                let sr = 0;
                while (drawn < best.mixedExtra && remaining > 0) {
                    for (let sc = 0; sc < stripCols && drawn < best.mixedExtra && remaining > 0; sc += 1) {
                        drawPiece(startL + sr * (pieceW + gap), margin + sc * (pieceL + gap), pieceW, pieceL, 'rgba(242,163,60,0.55)', '#f2a33c');
                        drawn += 1;
                        remaining -= 1;
                    }
                    sr += 1;
                    if (sr > 200) break;
                }
            }
        }

        ctx.fillStyle = '#9fb3c8';
        ctx.font = '11px system-ui, sans-serif';
        ctx.fillText('Plade ' + (sheetIndex + 1) + ': ' + piecesOnSheet + ' stk', x0 + 8, y0 + 14);
    };

    for (let i = 0; i < sheetsToDraw; i += 1) {
        const col = i % cols;
        const row = Math.floor(i / cols);
        const x = outerPad + col * (cellW + cellGap);
        const y = outerPad + row * (cellH + cellGap);
        const piecesOnSheet = i < sheetsNeeded - 1 ? piecesPerSheet : Math.max(0, totalQty - piecesPerSheet * (sheetsNeeded - 1));
        drawSheet(i, x, y, cellW, cellH, piecesOnSheet);
    }

    if (sheetsNeeded > sheetsToDraw) {
        ctx.fillStyle = '#9fb3c8';
        ctx.font = '12px system-ui, sans-serif';
        ctx.fillText('Viser ' + sheetsToDraw + ' af ' + sheetsNeeded + ' plader', 12, canvas.height - 8);
    }
}

// ── Saml forespørgsel og beregn ──
function buildQuoteBody(qtyOverride) {
    const laser = laserCardState();
    const laserOpstartOn = document.getElementById('calcLaserOpstartChk').checked;
    const spacing = getNestingSpacing();
    const useCustomSheet = !!document.getElementById('calcUseCustomSheet').checked;
    const customSheetW = Math.max(0, Number(document.getElementById('calcCustomSheetW').value || 0));
    const customSheetL = Math.max(0, Number(document.getElementById('calcCustomSheetL').value || 0));
    document.getElementById('calcMargin').value = spacing.margin;
    document.getElementById('calcGap').value = spacing.gap;
    return {
        materialProdNo: state.calcMaterial ? state.calcMaterial.ProdNo : '',
        pieceWidth: Number(document.getElementById('calcPieceW').value || 0),
        pieceLength: Number(document.getElementById('calcPieceL').value || 0),
        qty: qtyOverride != null ? qtyOverride : Number(document.getElementById('calcQty').value || 1),
        cutLengthM: Number(document.getElementById('calcCutLength').value || 0),
        piercings: Number(document.getElementById('calcPiercings').value || 1),
        machine: laser.machine,
        laserTechnology: laser.technology,
        laserEnabled: laser.enabled,
        laserRate: laser.rate,
        laserGeniusRate: laser.geniusRate,
        laserMinutesOverride: laser.timeOverride,
        laserOpstartMinutes: laserOpstartOn ? Number(document.getElementById('calcLaserOpstartMin').value || 0) : 0,
        priceBasis: document.getElementById('calcPriceBasis').value,
        sheetWidth: useCustomSheet ? customSheetW : undefined,
        sheetLength: useCustomSheet ? customSheetL : undefined,
        margin: spacing.margin,
        gap: spacing.gap,
        shapePolygon: (state.fileAnalysis && Array.isArray(state.fileAnalysis.polygon) && state.fileAnalysis.polygon.length >= 3) ? state.fileAnalysis.polygon : null,
        minimumOrderAmount: Number(document.getElementById('calcMinAmount').value || 0),
        minimumQty: Number(document.getElementById('calcMinQty').value || 0),
        operations: collectOperations(),
        components: collectComponents()
    };
}
async function postQuote(body) {
    return fetchJson('/bom/calc/quote', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
    });
}
async function runQuote() {
    if (!state.calcMaterial) {
        quoteStatus.textContent = 'Søg og vælg en plade/materiale først (trin 2).';
        return;
    }
    quoteStatus.textContent = 'Beregner...';
    try {
        const body = buildQuoteBody();
        const result = await postQuote(body);
        runQuote._lastErrToast = null;
        renderQuoteResult(result, body);
        if (!body.__skipMatrix) {
            renderPriceMatrix(body).catch(() => {});
        }
    } catch (err) {
        quoteStatus.textContent = 'Fejl: ' + err.message;
        quotePriceBig.textContent = '-';
        if (quotePriceTotal) quotePriceTotal.textContent = '-';
        const bd = document.getElementById('quoteBreakdown');
        if (bd) bd.hidden = true;
        const comparison = document.getElementById('laserTechnologyComparison');
        if (comparison) comparison.hidden = true;
        if (copyQuoteBtn) copyQuoteBtn.disabled = true;
        if (runQuote._lastErrToast !== err.message) {
            runQuote._lastErrToast = err.message;
            showToast('Beregning fejlede: ' + err.message, 'err');
        }
    }
}

function scheduleQuoteRecalc(delayMs) {
    const liveChk = document.getElementById('calcLiveUpdateChk');
    if (liveChk && !liveChk.checked) return;
    clearTimeout(quoteDebounceTimer);
    const wait = Math.max(120, Number(delayMs || 280));
    quoteDebounceTimer = setTimeout(() => {
        runQuote().catch(() => {});
    }, wait);
}
function renderCostBreakdown(result) {
    const wrap = document.getElementById('quoteBreakdown');
    const bar = document.getElementById('breakdownBar');
    const legend = document.getElementById('breakdownLegend');
    if (!wrap || !bar || !legend) return;
    const p = result.perPiece;
    const segs = [
        { label: 'Materiale', value: Number(p.materialPrice || 0), color: '#4fc3f7' },
        { label: 'Laser', value: Number(p.laserCost || 0), color: '#ffb74d' },
        { label: 'Processer', value: Number(p.resourceCost || 0), color: '#9575cd' },
        { label: 'Komponenter', value: Number(p.componentsCost || 0), color: '#4db6ac' },
        { label: 'Opstart', value: Number(p.opstartShare || 0), color: '#f06292' }
    ].filter(s => s.value > 0);
    const sum = segs.reduce((acc, s) => acc + s.value, 0);
    if (sum <= 0) { wrap.hidden = true; return; }
    wrap.hidden = false;
    bar.innerHTML = segs.map(s =>
        '<div class="breakdown-seg" style="width:' + ((s.value / sum) * 100).toFixed(2) + '%; background:' + s.color + ';" title="'
        + escapeHtml(s.label + ': ' + formatMoney(s.value) + ' dkk/stk (' + ((s.value / sum) * 100).toFixed(1) + ' %)') + '"></div>'
    ).join('');
    legend.innerHTML = segs.map(s =>
        '<span class="legend-item"><span class="legend-dot" style="background:' + s.color + ';"></span>'
        + escapeHtml(s.label) + ' ' + ((s.value / sum) * 100).toFixed(0) + ' % · ' + formatMoney(s.value) + ' dkk</span>'
    ).join('');
}

function renderLaserTechnologyComparison(result) {
    const wrap = document.getElementById('laserTechnologyComparison');
    const body = document.getElementById('laserTechnologyComparisonBody');
    const meta = document.getElementById('laserTechnologyComparisonMeta');
    if (!wrap || !body || !meta) return;
    const rows = result.laserTechnologyAlternatives || [];
    wrap.hidden = rows.length === 0;
    if (!rows.length) { body.innerHTML = ''; return; }
    const material = rows[0].material || '-';
    const thickness = rows[0].thickness;
    const eligibleCount = rows.filter(row => row.eligibleForAutomatic).length;
    meta.textContent = rows.length + ' kompatible · ' + eligibleCount + ' kan prisberegnes · '
        + material + ' · ' + formatMoney(thickness) + ' mm';
    body.innerHTML = rows.map(row => {
        const gas = row.gasType === 'nitrogen' ? 'N2' : row.gasType === 'oxygen' ? 'O2'
            : row.gasType === 'mixline' ? 'MixLine ' + formatMoney(row.mixLineOxygenPercent) + '% O2' : '-';
        const process = (row.machineName || '-') + (row.technologyLine && row.technologyLine !== '-' ? ' ' + row.technologyLine : '') + ' · ' + gas;
        const choice = row.selected ? '<span class="technology-choice"><span class="tag">Valgt</span></span>'
            : row.eligibleForAutomatic ? 'Alternativ' : escapeHtml(row.unavailableReason || 'Ikke prissat');
        return '<tr class="' + (row.selected ? 'selected' : '') + (!row.eligibleForAutomatic ? ' unavailable' : '') + '">'
            + '<td>' + choice + '</td><td><strong>' + escapeHtml(row.technology) + '</strong></td>'
            + '<td>' + escapeHtml(row.material) + '</td><td>' + formatMoney(row.thickness) + ' mm</td>'
            + '<td>' + escapeHtml(process) + '</td><td>' + escapeHtml(row.lens || '-') + '</td>'
            + '<td>' + formatNumber(row.feedrateMmMin) + ' mm/min</td><td>' + formatNumber(row.piercingMilliseconds) + ' ms</td>'
            + '<td>' + formatMoney(row.gasPressureBar) + ' bar / ' + formatMoney(row.nozzleSizeMm) + ' mm</td>'
            + '<td>' + (row.minutes == null ? '-' : formatMoney(row.minutes) + ' min') + '</td>'
            + '<td>' + (row.cost == null ? '-' : formatMoney(row.cost) + ' dkk/stk') + '</td></tr>';
    }).join('');
}

function buildQuoteClipboardText(result, body) {
    const fmt = formatMoney;
    const p = result.perPiece;
    const lines = [];
    lines.push('Tilbud — Gantech BOM-beregner');
    if (state.calcCustomer) lines.push('Kunde: ' + (state.calcCustomer.Nm || '') + ' (' + state.calcCustomer.CustNo + ')');
    lines.push('Materiale: ' + (result.material.prodNo || '') + ' · ' + (result.material.descr || ''));
    lines.push('Emne: ' + body.pieceWidth + ' x ' + body.pieceLength + ' mm · ' + result.total.qty + ' stk');
    lines.push('Pristype: ' + (result.material.priceBasis === 'cost' ? 'kostpris' : 'salgspris'));
    lines.push('');
    lines.push('Materiale: ' + fmt(p.materialPrice) + ' dkk/stk');
    if (Number(p.laserCost || 0) > 0) lines.push('Laser (' + p.laserMinutes + ' min): ' + fmt(p.laserCost) + ' dkk/stk');
    (result.operations || []).forEach(op => {
        lines.push(op.label + (op.prodNo ? ' (' + op.prodNo + ')' : '') + ': ' + op.minutes + ' min · ' + fmt(op.cost) + ' dkk/stk');
    });
    (result.components || []).forEach(c => {
        lines.push('Komponent ' + c.prodNo + ': ' + c.qty + ' stk × ' + fmt(c.unitPrice) + ' = ' + fmt(c.lineCost) + ' dkk/stk');
    });
    const perOrder = result.perOrder || {};
    if (Number(perOrder.opstartCost || 0) > 0) lines.push('Opstart pr ordre: ' + fmt(perOrder.opstartCost) + ' dkk (' + fmt(p.opstartShare) + ' dkk/stk)');
    if (result.total.minimumApplied) lines.push('Minimumsbeløb anvendt: ' + fmt(result.total.totalPrice) + ' dkk');
    lines.push('----------------------------------------');
    lines.push('Pris pr stk: ' + fmt(p.unitPrice) + ' dkk');
    lines.push('I alt (' + result.total.qty + ' stk): ' + fmt(result.total.totalPrice) + ' dkk');
    if (result.total.sheetsNeeded != null) lines.push('Plader til ordren: ' + result.total.sheetsNeeded);
    return lines.join('\n');
}

async function copyQuoteToClipboard() {
    if (!state.lastQuote) return;
    const text = buildQuoteClipboardText(state.lastQuote.result, state.lastQuote.body);
    try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(text);
        } else {
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            ta.remove();
        }
        showToast('Tilbud kopieret til udklipsholderen', 'ok');
    } catch (err) {
        showToast('Kunne ikke kopiere: ' + err.message, 'err');
    }
}

function vismaSqlString(value) {
    return "'" + String(value == null ? '' : value).replace(/'/g, "''") + "'";
}

function vismaSqlComment(value) {
    return String(value == null ? '' : value).replace(/[\r\n]+/g, ' ');
}

function vismaStructInsert(row) {
    return [
        'INSERT INTO Struct',
        '    (ProdNo, LnNo, SubProd, Descr, NoPerStr, Srt, ProdTp4,',
        '     PrM1, PrM2, PrM3, PrM4, TrInf4, Inf, Inf2, R7)',
        'VALUES',
        '    (' + vismaSqlString(row.prodNo) + ', ' + row.lineNo + ', ' + vismaSqlString(row.subProd) + ', ' + vismaSqlString(row.descr || '') + ', '
            + Number(row.quantity || 0) + ', 0, ' + row.prodType4 + ', '
            + row.prM1 + ', ' + row.prM2 + ', ' + row.prM3 + ', ' + row.prM4 + ', '
            + vismaSqlString(row.route || '0') + ', ' + vismaSqlString(row.inf || '0') + ', '
            + vismaSqlString(row.inf2 || '') + ', ' + vismaSqlString(row.r7 || '') + ');'
    ].join('\n');
}

function vismaOperationNo(resource) {
    const prodNo = String(resource.ProdNo || '').trim().toUpperCase();
    const exact = {
        R1100: 10, R1201: 10, R2100: 30, R5300: 50, R5200: 80,
        R5100: 10, R6104: 15, R6101: 70, R6106: 15, R6120: 14,
        R6200: 80, R8100: 80, R8200: 10
    };
    if (exact[prodNo] != null) return exact[prodNo];
    const family = String(resource.R7 || '').trim();
    if (family === '11' || family === '12') return 10;
    if (family === '21') return 30;
    if (family === '61' || family === '63') return 15;
    if (family === '82') return 10;
    return 0;
}

function buildVismaQueryPreview() {
    if (!state.lastQuote) return '';
    const body = state.lastQuote.body || {};
    const result = state.lastQuote.result || {};
    const customer = state.calcCustomer || {};
    const sameCustomer = state.selectedCustomer && customer.CustNo
        && String(state.selectedCustomer.CustNo) === String(customer.CustNo);
    const selectedProduct = sameCustomer ? state.selectedProduct : null;
    const drawing = state.fileAnalysis || {};
    const material = state.calcMaterial || {};
    const perPiece = result.perPiece || {};
    const total = result.total || {};
    const technology = result.laserTechnology || {};
    const prodNo = selectedProduct && selectedProduct.ProdNo ? selectedProduct.ProdNo : '<PRODUKTNR_MANGLER>';
    const descr = selectedProduct && selectedProduct.Descr
        ? selectedProduct.Descr : String(drawing.filename || 'BOMe+ beregning').replace(/\.[^.]+$/, '');
    const tgNo = selectedProduct && selectedProduct.TgNo
        ? selectedProduct.TgNo : String(drawing.filename || '').replace(/\.[^.]+$/, '');
    const revision = selectedProduct && selectedProduct.RevNo ? selectedProduct.RevNo : '';
    const customerCode = String(customer.Gr || customer.CustNo || '');
    const thickness = Number(material.tykklese || material.thickness || (result.material || {}).thickness || 0);
    const density = Number(material.density || material.DensU || (result.material || {}).density || 7.85);
    const routeProdNo = 'V' + prodNo;
    const laserProdNo = prodNo + 'L';
    const laserResource = String(technology.machineCode || technology.resourceNo || body.machine || 'R1100');
    const resourceInfo = resourceNo => state.resources.find(row => String(row.ProdNo || '') === String(resourceNo || '')) || {};
    const operations = (result.operations || []).map(op =>
        String(op.prodNo || op.label || op.key || '') + '=' + Number(op.minutes || 0).toFixed(6) + ' min'
    ).join(', ') || 'ingen';
    const structRows = [];
    let mainLine = 1;
    if (body.laserEnabled) {
        structRows.push({ prodNo, lineNo: mainLine++, subProd: laserProdNo, quantity: 1, prodType4: 2,
            prM1: 1073741824, prM2: 0, prM3: 256, prM4: 2097152 });
    }
    structRows.push({ prodNo, lineNo: mainLine++, subProd: routeProdNo, quantity: 1, prodType4: 5,
        prM1: 0, prM2: 16384, prM3: 0, prM4: 1310720 });
    (result.components || []).forEach(component => {
        structRows.push({ prodNo, lineNo: mainLine++, subProd: component.prodNo, quantity: component.qty,
            prodType4: 4, prM1: 1073741824, prM2: 0, prM3: 256, prM4: 2097152 });
    });

    let routeLine = 1;
    const addRouteResource = (resourceNo, label, minutes, prodType4, infoText) => {
        if (!resourceNo || Number(minutes || 0) <= 0) return;
        const resource = resourceInfo(resourceNo);
        structRows.push({ prodNo: routeProdNo, lineNo: routeLine++, subProd: resourceNo, descr: label,
            quantity: minutes, prodType4, prM1: prodType4 === 3 ? 1107296256 : 1073741824,
            prM2: 16384, prM3: 0, prM4: 2097153, route: vismaOperationNo(resource),
            inf: resource.BasePrice || 0, inf2: infoText, r7: resource.R7 || '' });
    };
    if (body.laserEnabled) {
        addRouteResource(laserResource, 'Opstart, Laserskæring', body.laserOpstartMinutes, 3,
            'min:' + Number(body.laserOpstartMinutes || 0));
    }
    (result.operations || []).forEach(operation => {
        const resource = resourceInfo(operation.prodNo);
        addRouteResource(operation.prodNo, 'Opstart, ' + String(operation.label || resource.Descr || ''),
            operation.opstartMinutes, 3, 'min:' + Number(operation.opstartMinutes || 0));
        addRouteResource(operation.prodNo, String(operation.label || resource.Descr || ''),
            operation.minutes, 1, 'min:' + Number(operation.minutes || 0));
    });
    if ((result.components || []).length) {
        addRouteResource('R8200', 'Opstart, Stykliste', 2, 3, 'min:2');
        addRouteResource('R8200', 'Stykliste', 1, 1, 'min:1');
    }
    if (body.laserEnabled) {
        structRows.push({ prodNo: laserProdNo, lineNo: 1, subProd: body.materialProdNo, quantity: perPiece.weightKg,
            prodType4: 2, prM1: 1140850688, prM2: 0, prM3: 768, prM4: 2097152, inf: 18 });
        const laserResourceRow = resourceInfo(laserResource);
        structRows.push({ prodNo: laserProdNo, lineNo: 2, subProd: laserResource, descr: 'Laserskæring',
            quantity: perPiece.laserMinutes, prodType4: 1, prM1: 1073741824, prM2: 16384, prM3: 0, prM4: 2097153,
            route: vismaOperationNo(laserResourceRow), inf: laserResourceRow.BasePrice || 0,
            inf2: 'Min:' + Number(perPiece.laserMinutes || 0), r7: laserResourceRow.R7 || '' });
    }
    return [
        '-- PREVIEW FRA BOMe+ - SENDES IKKE TIL VISMA',
        '-- Produktnummer skal udfyldes, hvis det står som <PRODUKTNR_MANGLER>.',
        '-- Materiale: ' + vismaSqlComment(body.materialProdNo || ''),
        '-- Antal: ' + Number(body.qty || 0),
        '-- Pris pr. stk: ' + Number(perPiece.unitPrice || 0).toFixed(2) + ' DKK',
        '-- Pris i alt: ' + Number(total.totalPrice || 0).toFixed(2) + ' DKK',
        '-- Laserteknologi: ' + vismaSqlComment(technology.technology || 'fravalgt'),
        '-- Operationer: ' + vismaSqlComment(operations),
        '',
        'SET XACT_ABORT ON;',
        'BEGIN TRANSACTION;',
        '',
        'DECLARE @ProdNo varchar(50) = ' + vismaSqlString(prodNo) + ';',
        'DECLARE @Descr varchar(60) = ' + vismaSqlString(String(descr).slice(0, 60)) + ';',
        'DECLARE @TgNo varchar(60) = ' + vismaSqlString(tgNo) + ';',
        'DECLARE @CustomerCode varchar(60) = ' + vismaSqlString(customerCode) + ';',
        'DECLARE @Revision varchar(20) = ' + vismaSqlString(revision) + ';',
        '',
        '-- 1) PRODUKTANAGRAFIK: hovedprodukt, rute og laserprodukt',
        'INSERT INTO Prod',
        '    (ProdNo, Descr, ProdGr, Inf2, Inf3, Inf7, Inf8,',
        '     HgtU, LgtU, WdtU, DensU, Inf, Free2, StSaleUn,',
        '     ProdPrGr, PrCatNo, Gr8, NWgtU, CreDt, Rsp)',
        'VALUES',
        '    (@ProdNo, @Descr, 1, @TgNo, @CustomerCode, @Revision, \'A4\',',
        '     ' + thickness + ', ' + (Number(body.pieceLength || 0) / 1000) + ', ' + (Number(body.pieceWidth || 0) / 1000) + ', ' + density + ', 0, ' + Number(body.margin || 0) + ', 1,',
        '     ' + Number(customer.CustPrGr || 0) + ', 1, 1, ' + Number(perPiece.weightKg || 0) + ', CONVERT(varchar(8), GETDATE(), 112), 0),',
        '    (' + vismaSqlString(routeProdNo) + ', ' + vismaSqlString(('Rute for ' + descr).slice(0, 60)) + ', 1, \'0\', @CustomerCode, \'\', \'\',',
        '     0, 0, 0, 0, 0, 0, 1, ' + Number(customer.CustPrGr || 0) + ', 1, 0, 0, CONVERT(varchar(8), GETDATE(), 112), 0)' + (body.laserEnabled ? ',' : ';'),
        ...(body.laserEnabled ? [
            '    (' + vismaSqlString(laserProdNo) + ', ' + vismaSqlString(('Laser ' + descr).slice(0, 60)) + ', 2, @TgNo, @CustomerCode, @Revision, @Revision,',
            '     ' + thickness + ', ' + Number(body.pieceLength || 0) + ', ' + Number(body.pieceWidth || 0) + ', ' + density + ', 0, 0, 1,',
            '     ' + Number(customer.CustPrGr || 0) + ', 1, 1, ' + Number(perPiece.weightKg || 0) + ', CONVERT(varchar(8), GETDATE(), 112), 0);'
        ] : []),
        '',
        '-- 2) STRUKTUR: hovedvare -> laser/rute/komponenter; laser -> materiale/tid; rute -> opstart/processer',
        ...structRows.flatMap((row, index) => [vismaStructInsert(row), index === structRows.length - 1 ? '' : '']),
        '',
        '-- SIKKERHED: previewen gemmer aldrig ændringer og kaldes ikke af et write-endpoint.',
        'ROLLBACK TRANSACTION;'
    ].join('\n');
}

function openVismaQueryPreview() {
    if (!state.permissions.bomVismaPreview) return;
    if (!state.lastQuote) {
        showToast('Beregn en pris først.', 'err');
        return;
    }
    document.getElementById('vismaQueryText').textContent = buildVismaQueryPreview();
    document.getElementById('vismaQueryModal').classList.add('open');
}

function closeVismaQueryPreview() {
    document.getElementById('vismaQueryModal').classList.remove('open');
}

function renderQuoteResult(result, body) {
    const fmt = formatMoney;
    state.lastQuote = { result, body };
    updateCalcWizard();
    animateMoney(quotePriceBig, result.perPiece.unitPrice, ' dkk');
    if (quotePriceTotal) animateMoney(quotePriceTotal, result.total.totalPrice, ' dkk');
    [quotePriceBig, quotePriceTotal].forEach(el => {
        if (!el) return;
        el.classList.remove('price-flash');
        void el.offsetWidth;
        el.classList.add('price-flash');
    });
    renderCostBreakdown(result);
    renderLaserTechnologyComparison(result);
    if (copyQuoteBtn) copyQuoteBtn.disabled = false;
    const vismaPreviewBtn = document.getElementById('vismaQueryPreviewBtn');
    if (vismaPreviewBtn && state.permissions.bomVismaPreview) vismaPreviewBtn.disabled = false;
    const laserOn = body.laserEnabled;
    const technology = result.laserTechnology;
    const eligibleTechnologyCount = (result.laserTechnologyAlternatives || []).filter(row => row.eligibleForAutomatic).length;
    const automaticChoiceLabel = eligibleTechnologyCount > 1 ? 'billigst af ' + eligibleTechnologyCount : 'eneste beregnelige';
    quoteStatus.textContent = laserOn
        ? (technology ? ('Skæredata: ' + technology.technology
            + (technology.selectionMode === 'automatic' ? ' (' + automaticChoiceLabel + ')' : ' (manuel)')
            + ' · ' + (technology.machineName || '') + (technology.technologyLine && technology.technologyLine !== '-' ? ' ' + technology.technologyLine : '')
            + ' · ' + (technology.gasType === 'nitrogen' ? 'N2' : technology.gasType === 'oxygen' ? 'O2'
                : 'MixLine ' + formatMoney(technology.mixLineOxygenPercent) + '% O2')
            + ' · ' + formatNumber(technology.feedrateMmMin) + ' mm/min'
            + ' · ' + technology.gasPressureBar + ' bar · ' + formatMoney(technology.gasFlowNm3Hour) + ' m³/h'
            + ' · ' + formatMoney(technology.gasConsumptionKgHour) + ' kg/h')
            : result.cutParam ? ('Skæredata: ' + (result.cutParam.maskine || '') + ' · ' + result.cutParam.skaerehast + ' m/min')
                : 'OBS: ingen skæreparametre fundet for materialet — laser-tid er 0.')
        : 'Laser fravalgt — kun materiale + processer.';
    if (result.nesting) {
        const n = result.nesting;
        const visualQty = resolveNestingVisualQty(result);
        const visualSheets = n.best.total > 0 ? Math.ceil(visualQty / n.best.total) : 0;
        nestingMeta.textContent = 'plade ' + n.input.sheetW + ' x ' + n.input.sheetL + ' mm' + (n.mode === 'shape' ? ' · ægte form-nesting fra DXF' : '') + ' · visning ' + visualQty + ' stk';
        const layoutTxt = n.mode === 'shape'
            ? 'form-nesting (rotationer: ' + (n.best.rotationsUsed || []).join('°, ') + '°)'
            : n.best.cols + ' x ' + n.best.rows + (n.best.rotation ? ' (roteret 90°)' : '') + (n.best.mixedExtra ? ' + ' + n.best.mixedExtra + ' blandet' : '');
        const piecesPerSheet = Math.max(1, Number(n.best.total || 1));
        const sheetKg = Number(result.material.sheetWeightKg || 0);
        const pieceKg = Number(result.perPiece.weightKg || 0);
        const qty = visualQty;
        const sheetsNeeded = visualSheets;
        const fullUsedKg = piecesPerSheet * pieceKg;
        const fullWasteKg = Math.max(0, sheetKg - fullUsedKg);
        const fullWastePct = sheetKg > 0 ? (fullWasteKg / sheetKg) * 100 : 0;
        const orderSheetKg = sheetsNeeded > 0 ? sheetKg * sheetsNeeded : 0;
        const orderPartsKg = pieceKg * qty;
        const orderWasteKg = Math.max(0, orderSheetKg - orderPartsKg);
        const orderWastePct = orderSheetKg > 0 ? (orderWasteKg / orderSheetKg) * 100 : 0;

        nestingGrid.innerHTML = [
            ['Pladekilde', result.material && result.material.sheetSource === 'custom' ? 'Tilpasset plade' : 'Materialets standardplade'],
            ['Emner pr plade (bedst)', n.best.total],
            ['Layout', layoutTxt],
            ['Nesting-strategi', n.best.strategy || (n.best.fragmented ? 'blandet' : 'kompakt')],
            ['Nesting afstande (laser)', 'afstand til kant ' + n.input.margin + ' mm · afstand mellem emner ' + n.input.gap + ' mm'],
            ['Bevarbart reststykke', n.best.reusableRemnant ? (Math.round(Number(n.best.reusableRemnant.width || 0)) + ' x ' + Math.round(Number(n.best.reusableRemnant.length || 0)) + ' mm') : '-'],
            ['Rest-bevaring', n.best.restPreservationPct == null ? '-' : (n.best.restPreservationPct + ' % af brugbar plade')],
            ['Udnyttelse', n.utilizationPct + ' %'],
            ['Spild pr fuld plade', formatMoney(fullWasteKg) + ' kg · ' + fullWastePct.toFixed(1) + ' %'],
            ['Plader til vist antal', sheetsNeeded || '-'],
            ['Spild for vist antal', formatMoney(orderWasteKg) + ' kg · ' + orderWastePct.toFixed(1) + ' %'],
            ['Plader til ordren', result.total.sheetsNeeded == null ? '-' : result.total.sheetsNeeded]
        ].map(([label, value]) => '<div class="kv"><label>' + escapeHtml(label) + '</label><div>' + escapeHtml(value) + '</div></div>').join('');
        drawNesting(n, visualQty);
    } else {
        nestingMeta.textContent = 'plademål mangler på materialet';
        drawNesting(null);
        nestingGrid.innerHTML = '<div class="empty" style="grid-column:1 / -1;">Ingen nesting — materialet har ikke plade-dimensioner.</div>';
    }
    const p = result.perPiece;
    const laserTimeInput = processCards.querySelector('.proc-card[data-proc="laser"] .proc-time');
    if (laserTimeInput) laserTimeInput.placeholder = 'auto: ' + p.autoLaserMinutes + ' min';
    const bendingLine = (result.operations || []).find(op => op.key === 'buk' && op.bending);
    const bendingTimeInput = processCards.querySelector('.proc-card[data-proc="buk"] .proc-min');
    if (bendingTimeInput && bendingLine) bendingTimeInput.placeholder = 'auto: ' + bendingLine.minutes + ' min';
    quoteMeta.textContent = (result.material.descr || result.material.prodNo) + (state.calcCustomer ? ' · ' + (state.calcCustomer.Nm || '') : '');
    const opLines = (result.operations || []).map(op => {
        const bendingText = op.bending ? ' · kraft ' + op.bending.requiredForceKn + '/' + op.bending.availableForceKn + ' kN'
            + ' · cyklus ' + op.bending.cycleSeconds + ' sek. · håndtering ' + op.bending.handlingSeconds + ' sek. (' + op.bending.handlingBand + ')' : '';
        return [op.label + (op.prodNo ? ' (' + op.prodNo + ')' : ''), op.minutes + ' min · ' + fmt(op.cost) + ' dkk'
            + (op.opstartMinutes ? ' · opstart ' + op.opstartMinutes + ' min' : '') + bendingText];
    });
    const compLines = (result.components || []).map(c => ['Komponent ' + c.prodNo, c.qty + ' stk × ' + fmt(c.unitPrice) + ' = ' + fmt(c.lineCost) + ' dkk']);
    const isCost = result.material.priceBasis === 'cost';
    const perOrder = result.perOrder || {};
    const opstartLines = (perOrder.opstartCost > 0)
        ? [['Opstart i alt (pr ordre)', fmt(perOrder.opstartCost) + ' dkk · ' + fmt(p.opstartShare) + ' dkk/stk'
            + (perOrder.laserOpstartCost > 0 ? ' (laser ' + perOrder.laserOpstartMinutes + ' min)' : '')]]
        : [];
    const minLines = result.total.minimumApplied
        ? [['Minimum anvendt', 'beregnet ' + fmt(result.total.rawTotal) + ' dkk → faktureres ' + fmt(result.total.totalPrice) + ' dkk']]
        : [];
    quoteGrid.innerHTML = [
        ['Emnevægt', p.weightKg + ' kg'],
        ['Materiale (kost)', fmt(p.materialCost) + ' dkk'],
        [isCost ? 'Materiale (kostpris — pristype)' : 'Materiale (m. avance ' + result.material.avancePct + '%)', fmt(p.materialPrice) + ' dkk'],
        ['Laser-tid' + (p.laserMinutesOverridden ? ' (rettet manuelt)' : ''), p.laserMinutes + ' min'],
        ['Laser-pris', fmt(p.laserCost) + ' dkk']
    ].concat(opLines).concat(compLines).concat(opstartLines).concat([
        ['Processer i alt', p.resourceMinutes + ' min · ' + fmt(p.resourceCost) + ' dkk']
    ]).concat(p.componentsCost > 0 ? [['Komponenter i alt', fmt(p.componentsCost) + ' dkk']] : []).concat(minLines).concat([
        ['Pris pr stk', fmt(p.unitPrice) + ' dkk'],
        ['I alt (' + result.total.qty + ' stk)', fmt(result.total.totalPrice) + ' dkk']
    ]).map(([label, value]) => '<div class="kv"><label>' + escapeHtml(label) + '</label><div>' + escapeHtml(value) + '</div></div>').join('');
}

// ── Prismatrix: pris pr stk ved forskellige antal ──
async function renderPriceMatrix(baseBody) {
    const matrixBody = document.getElementById('matrixBody');
    const matrixMeta = document.getElementById('matrixMeta');
    const tiersRaw = String(document.getElementById('calcMatrixTiers').value || '1, 5, 10, 25, 50, 100');
    const currentQty = Number(document.getElementById('calcQty').value || 1);
    const tiers = Array.from(new Set(
        tiersRaw.split(/[,;\s]+/).map(v => Number(v)).filter(v => v > 0).concat([currentQty])
    )).sort((a, b) => a - b).slice(0, 10);
    matrixMeta.textContent = 'beregner...';
    matrixBody.innerHTML = '';
    const results = await Promise.all(tiers.map(qty => postQuote({ ...baseBody, qty }).catch(() => null)));
    const rows = tiers.map((qty, i) => ({ qty, r: results[i] })).filter(x => x.r);
    matrixBody.innerHTML = rows.map(({ qty, r }) => {
        const effUnit = r.total.totalPrice / qty;
        const isCurrent = qty === currentQty;
        return '<tr data-qty="' + qty + '"' + (isCurrent ? ' style="background:#eaf4ff; font-weight:700;"' : '') + '>'
            + '<td>' + qty + ' stk</td>'
            + '<td>' + formatMoney(effUnit) + ' dkk</td>'
            + '<td>' + formatMoney(r.total.totalPrice) + ' dkk</td>'
            + '<td>' + (r.total.minimumApplied ? 'minimum' : (r.total.sheetsNeeded == null ? '-' : r.total.sheetsNeeded + ' plader')) + '</td>'
            + '</tr>';
    }).join('');
    matrixBody.querySelectorAll('tr[data-qty]').forEach(tr => {
        tr.style.cursor = 'pointer';
        tr.title = 'Klik for at bruge antal i beregningen';
        tr.addEventListener('click', () => {
            const q = Math.max(1, Number(tr.getAttribute('data-qty') || 1));
            document.getElementById('calcQty').value = q;
            const body = buildQuoteBody();
            body.__skipMatrix = true;
            postQuote(body).then(result => {
                renderQuoteResult(result, body);
                matrixBody.querySelectorAll('tr[data-qty]').forEach(x => {
                    x.style.background = '';
                    x.style.fontWeight = '';
                });
                tr.style.background = '#eaf4ff';
                tr.style.fontWeight = '700';
            }).catch(() => {});
        });
    });
    matrixMeta.textContent = rows.length + ' trin' + (Number(document.getElementById('calcMinAmount').value || 0) > 0 ? ' · minimumsbeløb ' + formatMoney(document.getElementById('calcMinAmount').value) + ' dkk' : '');
}

// ── Ny beregning: nulstil hele beregneren ──
function resetBeregner() {
    state.calcMaterial = null;
    state.calcCustomer = null;
    state.fileAnalysis = null;
    state.calcComponents = [];
    state.calcWizardProcessesReady = false;
    document.getElementById('calcCustomerSearch').value = '';
    document.getElementById('calcCustomerChosen').textContent = 'Ingen kunde valgt — standard prisliste';
    document.getElementById('calcCustomerPriceInfo').textContent = 'Kundens prisliste: -';
    document.getElementById('calcCustomerProductInfo').textContent = 'Seneste produktnr: - · Næste forslag: -';
    document.getElementById('calcPriceBasis').value = 'sale';
    document.getElementById('calcLaserOpstartChk').checked = true;
    document.getElementById('calcLaserOpstartMin').value = 15;
    document.getElementById('calcMinAmount').value = 0;
    document.getElementById('calcMinQty').value = 0;
    document.getElementById('calcThicknessFilter').value = '';
    document.getElementById('calcOnlyStock').checked = false;
    document.getElementById('calcUseCustomSheet').checked = false;
    document.getElementById('calcCustomSheetW').value = 1500;
    document.getElementById('calcCustomSheetL').value = 3000;
    document.getElementById('calcCustomSheetW').disabled = true;
    document.getElementById('calcCustomSheetL').disabled = true;
    calcMaterialSearch.value = '';
    calcMaterialChosen.textContent = 'Ingen plade valgt';
    document.getElementById('drawingFileInput').value = '';
    fileAnalysisStatus.textContent = 'Ingen fil — udfyld felterne manuelt.';
    fileAnalysisGrid.innerHTML = '';
    document.getElementById('calcPieceW').value = 100;
    document.getElementById('calcPieceL').value = 200;
    document.getElementById('calcQty').value = 1;
    document.getElementById('calcCutLength').value = 0.6;
    document.getElementById('calcPiercings').value = 1;
    document.getElementById('calcMargin').value = 10;
    document.getElementById('calcGap').value = 5;
    document.getElementById('calcPreviewSheets').value = 3;
    document.getElementById('calcLiveUpdateChk').checked = true;
    document.getElementById('nestingQtyMode').value = 'order';
    document.getElementById('nestingCustomQty').value = 1;
    document.getElementById('calcMatrixTiers').value = '1, 5, 10, 25, 50, 100';
    // genbyg proceskort fra bunden
    processCards.innerHTML = '';
    delete processCards.dataset.built;
    buildProcessCards();
    populateProcessDefaults();
    refreshLaserTechnologyOptions();
    syncSpacingFromGlobalToLaserCard();
    // ryd resultater
    state.lastQuote = null;
    quotePriceBig.textContent = '-';
    delete quotePriceBig.dataset.animVal;
    if (quotePriceTotal) {
        quotePriceTotal.textContent = '-';
        delete quotePriceTotal.dataset.animVal;
    }
    const bd = document.getElementById('quoteBreakdown');
    if (bd) bd.hidden = true;
    const comparison = document.getElementById('laserTechnologyComparison');
    if (comparison) comparison.hidden = true;
    if (copyQuoteBtn) copyQuoteBtn.disabled = true;
    const vismaPreviewBtn = document.getElementById('vismaQueryPreviewBtn');
    if (vismaPreviewBtn) vismaPreviewBtn.disabled = true;
    quoteStatus.textContent = 'Følg trin 1-4 og tryk Beregn pris.';
    quoteGrid.innerHTML = '';
    quoteMeta.textContent = '-';
    nestingGrid.innerHTML = '';
    nestingMeta.textContent = '-';
    drawNesting(null);
    renderDxfViewer();
    updateCalcWizard();
    setTimeout(() => openCalcStep('customer'), 100);
    document.getElementById('matrixBody').innerHTML = '';
    document.getElementById('matrixMeta').textContent = '-';
}
