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
function bukAutoMinutes(card) {
    const antal = Number(card.querySelector('.buk-antal').value || 0);
    const sek = Number(card.querySelector('.buk-sek').value || 0);
    const haandt = Number(card.querySelector('.buk-haandt').value || 0);
    return Math.round(((antal * sek + haandt) / 60) * 100) / 100;
}
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
                + '<div class="field"><label>Minutter pr emne</label><input class="proc-min" type="number" step="0.1" value="2" /></div>'
                + '<div class="field"><label>Sats (dkk/min)</label><input class="proc-rate" type="number" step="0.1" value="9.38" /></div>'
                + '<div class="field"><label>Opstart (min/ordre)</label><input class="proc-opstart" type="number" step="1" min="0" value="0" /></div>'
                + '</div></div>';
        } else if (def.kind === 'buk') {
            card.innerHTML = '<label class="proc-head"><input type="checkbox" class="proc-toggle" /> ' + escapeHtml(def.label) + ' <span class="proc-res-label"></span></label>'
                + '<div class="proc-body">'
                + '<div class="picker"><input class="proc-res-search" placeholder="søg buk-ressource..." autocomplete="off" /><div class="picker-list"></div></div>'
                + '<div class="proc-fields">'
                + '<div class="field"><label>Antal buk pr emne</label><input class="buk-antal" type="number" step="1" min="0" value="2" /></div>'
                + '<div class="field"><label>Sekunder pr buk</label><input class="buk-sek" type="number" step="5" min="0" value="30" /></div>'
                + '<div class="field"><label>Håndtering (sek/emne)</label><input class="buk-haandt" type="number" step="5" min="0" value="15" /></div>'
                + '<div class="field"><label>Minutter pr emne (auto — kan rettes)</label><input class="proc-min" type="number" step="0.01" value="1.25" /></div>'
                + '<div class="field"><label>Sats (dkk/min)</label><input class="proc-rate" type="number" step="0.1" value="10.31" /></div>'
                + '<div class="field"><label>Opstart (min/ordre)</label><input class="proc-opstart" type="number" step="1" min="0" value="10" /></div>'
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
        fladType.addEventListener('change', () => applyResource(findResource(fladType.value)));
    }
    // Buk: auto-beregn minutter
    if (def.kind === 'buk') {
        ['buk-antal', 'buk-sek', 'buk-haandt'].forEach(cls => {
            card.querySelector('.' + cls).addEventListener('input', () => {
                card.querySelector('.proc-min').value = bukAutoMinutes(card);
            });
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
            const preferred = rows.find(r => String(r.ProdNo) === def.defaultRes) || rows[0];
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
        ops.push({
            key: def.key,
            label: def.label.split('(')[0].trim(),
            prodNo,
            minutes: Number(card.querySelector('.proc-min').value || 0),
            rate: Number(card.querySelector('.proc-rate').value || 0),
            opstartMinutes: Number(card.querySelector('.proc-opstart').value || 0)
        });
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
    if (!card) return { enabled: true, machine: 'R1100', rate: 12, timeOverride: null };
    const timeRaw = String(card.querySelector('.proc-time').value || '').trim();
    return {
        enabled: card.querySelector('.proc-toggle').checked,
        machine: String(card.querySelector('.proc-machine').value || 'R1100').trim(),
        rate: Number(card.querySelector('.proc-rate').value || 12),
        timeOverride: timeRaw === '' ? null : Number(timeRaw)
    };
}
async function primeBeregner() {
    buildProcessCards();
    try {
        await Promise.all([ensureMaterials(), ensureResources(), ensureCustomers(), ensureComponents()]);
        populateProcessDefaults();
        initDxfViewerInteractions();
        renderDxfViewer();
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
        state.fileAnalysis = result;
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
        fileAnalysisStatus.textContent = result.note || (file.name + ' analyseret — felterne er udfyldt nedenfor.');
        if (result.widthMm) document.getElementById('calcPieceW').value = result.widthMm;
        if (result.lengthMm) document.getElementById('calcPieceL').value = result.lengthMm;
        if (result.cutLengthM) document.getElementById('calcCutLength').value = result.cutLengthM;
        if (result.piercingsEstimate) document.getElementById('calcPiercings').value = result.piercingsEstimate;
        renderDxfViewer();
    } catch (err) {
        fileAnalysisStatus.textContent = 'Fejl: ' + err.message;
        resetDxfMeasure(true);
        renderDxfViewer();
    }
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
        laserEnabled: laser.enabled,
        laserRate: laser.rate,
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
        renderQuoteResult(result, body);
        if (!body.__skipMatrix) {
            renderPriceMatrix(body).catch(() => {});
        }
    } catch (err) {
        quoteStatus.textContent = 'Fejl: ' + err.message;
        quotePriceBig.textContent = '-';
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
function renderQuoteResult(result, body) {
    const fmt = formatMoney;
    quotePriceBig.textContent = fmt(result.perPiece.unitPrice) + ' dkk/stk · ' + fmt(result.total.totalPrice) + ' dkk i alt';
    const laserOn = body.laserEnabled;
    quoteStatus.textContent = laserOn
        ? (result.cutParam ? ('Skæredata: ' + (result.cutParam.maskine || '') + ' · ' + result.cutParam.skaerehast + ' m/min') : 'OBS: ingen skæreparametre fundet for materialet — laser-tid er 0.')
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
    quoteMeta.textContent = (result.material.descr || result.material.prodNo) + (state.calcCustomer ? ' · ' + (state.calcCustomer.Nm || '') : '');
    const opLines = (result.operations || []).map(op => [op.label + (op.prodNo ? ' (' + op.prodNo + ')' : ''), op.minutes + ' min · ' + fmt(op.cost) + ' dkk' + (op.opstartMinutes ? ' · opstart ' + op.opstartMinutes + ' min' : '')]);
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
    syncSpacingFromGlobalToLaserCard();
    // ryd resultater
    quotePriceBig.textContent = '-';
    quoteStatus.textContent = 'Følg trin 1-4 og tryk Beregn pris.';
    quoteGrid.innerHTML = '';
    quoteMeta.textContent = '-';
    nestingGrid.innerHTML = '';
    nestingMeta.textContent = '-';
    drawNesting(null);
    renderDxfViewer();
    document.getElementById('matrixBody').innerHTML = '';
    document.getElementById('matrixMeta').textContent = '-';
}
