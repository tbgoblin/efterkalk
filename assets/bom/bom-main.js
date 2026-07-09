// ── BOM Workspace · main ─────────────────────────────────────────────
// Event-bindinger og opstart. Indlæses som sidste modul.

document.getElementById('loadStyklisteBtn').addEventListener('click', async () => { await loadCustomers(); await loadProducts(); await loadRevisions(); });
document.getElementById('loadResourcesBtn').addEventListener('click', loadResources);
document.getElementById('loadMaterialsBtn').addEventListener('click', loadMaterials);
document.getElementById('loadCalculatorsBtn').addEventListener('click', loadCalculators);
document.getElementById('invalidateBtn').addEventListener('click', async () => { try { await invalidateActiveCache(); } catch (err) { setStatus('Kunne ikke rydde cache: ' + err.message); } });
document.getElementById('openDraftProductBtn').addEventListener('click', openDraftModal);
document.getElementById('closeDraftProductBtn').addEventListener('click', closeDraftModal);
document.getElementById('cancelDraftProductBtn').addEventListener('click', closeDraftModal);
draftProductModal.addEventListener('click', evt => {
    if (evt.target === draftProductModal) closeDraftModal();
});

// ── Auto-preview produktnr. mens bruger skriver ──
if (draftProdNoSuffix) {
    draftProdNoSuffix.addEventListener('input', () => {
        const custCode = state.selectedCustomer ? String(state.selectedCustomer.Gr || state.selectedCustomer.CustNo || '') : '';
        const suffix = String(draftProdNoSuffix.value || '').trim();
        const preview = custCode + suffix;
        if (draftProdNoPreview) draftProdNoPreview.textContent = preview || '—';
        // Skjul preview-panel ved ændring
        if (vismaPreviewPanel) vismaPreviewPanel.style.display = 'none';
        if (createVismaBtn) { createVismaBtn.style.display = 'none'; createVismaBtn.disabled = true; }
    });
}

// ── "Tjek Visma" — preview hvad der vil blive oprettet ──
document.getElementById('previewVismaBtn').addEventListener('click', async () => {
    if (!state.selectedCustomer) { setStatus('Vælg en kunde først.'); return; }
    const custCode = String(state.selectedCustomer.Gr || state.selectedCustomer.CustNo || '');
    const suffix   = String(draftProdNoSuffix ? draftProdNoSuffix.value || '' : '').trim();
    const descr    = String(draftDescr.value || '').trim();
    if (!suffix || !descr) {
        setStatus('Udfyld produktnr.-suffiks og beskrivelse.');
        return;
    }
    const btn = document.getElementById('previewVismaBtn');
    btn.disabled = true;
    btn.textContent = '⏳ Tjekker…';
    try {
        const body = {
            customerCode:    custCode,
            customerNo:      String(state.selectedCustomer.CustNo || ''),
            prodNoSuffix:    suffix,
            descr,
            tgNo:            String(draftTgNo ? draftTgNo.value || '' : '').trim(),
            revNo:           String(draftRevNo ? draftRevNo.value || '' : '').trim(),
            tgForm:          draftTgForm ? draftTgForm.value : 'A4',
            customerNoAlt:   draftCustomerNoAlt ? String(draftCustomerNoAlt.value || '').trim() : '',
            createRoute:     !!(draftCreateRoute && draftCreateRoute.checked),
            createLaserPart: !!(draftCreateLaser && draftCreateLaser.checked),
            prodPrGr:        state.selectedCustomer.CustPrGr || 0
        };
        const resp = await fetchJson('/bom/create-products/preview', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        if (vismaPreviewPanel) vismaPreviewPanel.style.display = 'block';
        if (vismaPreviewBody) {
            vismaPreviewBody.innerHTML = (resp.records || []).map(r =>
                '<div style="padding:3px 0;border-bottom:1px solid #dde;"><strong>' + escapeHtml(r.ProdNo) + '</strong> — ' +
                escapeHtml(r.Descr) + ' <span style="color:#57718f;">(ProdGr=' + r.ProdGr + ')</span></div>'
            ).join('');
        }
        const conflicts = resp.conflicts || [];
        if (vismaPreviewConflict) {
            if (conflicts.length > 0) {
                vismaPreviewConflict.style.display = 'block';
                vismaPreviewConflict.textContent = '⚠️ Eksisterer allerede: ' + conflicts.join(', ');
            } else {
                vismaPreviewConflict.style.display = 'none';
            }
        }
        if (createVismaBtn) {
            createVismaBtn.style.display = 'inline-block';
            createVismaBtn.disabled = conflicts.length > 0;
            createVismaBtn.dataset.payload = JSON.stringify(body);
        }
    } catch (err) {
        setStatus('Fejl ved tjek: ' + err.message);
    } finally {
        btn.disabled = false;
        btn.textContent = '👁 Tjek Visma';
    }
});

// ── "Opret i Visma" — udfør INSERT ──
if (createVismaBtn) {
    createVismaBtn.addEventListener('click', async () => {
        if (!confirm('Er du sikker på, at du vil oprette disse produkter direkte i Visma?\nDenne handling kan ikke fortrydes automatisk.')) return;
        createVismaBtn.disabled = true;
        createVismaBtn.textContent = '⏳ Opretter…';
        try {
            const payload = JSON.parse(createVismaBtn.dataset.payload || '{}');
            const resp = await fetchJson('/bom/create-products/execute', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const created = (resp.created || []).map(r => r.ProdNo).join(', ');
            setStatus('✅ Oprettet i Visma: ' + created);
            closeDraftModal();
            await loadProducts();
        } catch (err) {
            setStatus('❌ Fejl ved oprettelse i Visma: ' + err.message);
            createVismaBtn.disabled = false;
            createVismaBtn.textContent = '✅ Opret i Visma';
        }
    });
}

draftProductForm.addEventListener('submit', evt => {
    evt.preventDefault();
    if (!state.selectedCustomer) return;
    const custCode = String(state.selectedCustomer.Gr || state.selectedCustomer.CustNo || '');
    const suffix   = String(draftProdNoSuffix ? draftProdNoSuffix.value || '' : '').trim();
    const prodNo   = custCode + suffix;
    const descr    = String(draftDescr.value || '').trim();
    if (!suffix || !descr) {
        setStatus('Produktnr.-suffiks og beskrivelse er påkrævet for lokal kladde.');
        return;
    }
    const rows = loadDraftProducts(state.selectedCustomer.CustNo);
    if (rows.some(row => String(row.ProdNo) === prodNo)) {
        setStatus('Der findes allerede en lokal kladde med dette produktnr.');
        return;
    }
    const selectedResources = state.draftResources.map(row => ({
        ProdNo: row.ProdNo || '',
        Descr: row.Descr || ''
    }));
    const draftRow = {
        ProdNo: prodNo,
        Descr: descr,
        TgNo: String(draftTgNo ? draftTgNo.value || '' : '').trim(),
        RevNo: String(draftRevNo ? draftRevNo.value || '' : '').trim() || 'KLADDE',
        PosNo: '',
        Inf3: state.selectedCustomer ? state.selectedCustomer.CustNo : '',
        Inf4: draftCustomerNoAlt ? String(draftCustomerNoAlt.value || '').trim() : '',
        chck: 'LOCAL_DRAFT',
        IsLocalDraft: true,
        MaterialProdNo: state.draftMaterial ? (state.draftMaterial.ProdNo || '') : '',
        MaterialDescr: state.draftMaterial ? (state.draftMaterial.beskrivelse || state.draftMaterial.Descr || '') : '',
        Resources: selectedResources,
        Sublevels: collectSublevels(prodNo)
    };
    rows.unshift(draftRow);
    saveDraftProducts(state.selectedCustomer.CustNo, rows);
    state.draftProducts = rows;
    state.products = state.products.filter(row => String(row.ProdNo) !== prodNo).concat([draftRow]);
    state.selectedProduct = draftRow;
    if (draftRow.TgNo) tgnInput.value = draftRow.TgNo;
    closeDraftModal();
    renderProducts();
    renderProductDetail();
    loadProductTree();
    updateContext();
    setMetric('metricProducts', formatNumber(state.products.length));
    setStatus('Lokal produktkladde gemt (sendes ikke til Visma).', state.products.length);
});
customerSearchInput.addEventListener('input', renderCustomers);
productSearchInput.addEventListener('input', renderProducts);
resourcesSearchInput.addEventListener('input', loadResources);
document.getElementById('loadComponentsBtn').addEventListener('click', loadComponents);
componentsSearchInput.addEventListener('input', () => { loadComponents().catch(() => {}); });
document.getElementById('loadSuppliersBtn').addEventListener('click', loadSuppliers);
suppliersSearchInput.addEventListener('input', () => { loadSuppliers().catch(() => {}); });
document.getElementById('addSublevelBtn').addEventListener('click', () => addSublevelRow());
document.getElementById('runQuoteBtn').addEventListener('click', runQuote);
document.getElementById('resetQuoteBtn').addEventListener('click', resetBeregner);
if (copyQuoteBtn) copyQuoteBtn.addEventListener('click', copyQuoteToClipboard);

// ── Tastaturgenveje: Alt+1..8 skifter område, "/" fokuserer søgning ──
const viewSearchFocus = {
    stykliste: 'customerSearchInput',
    komponenter: 'componentsSearchInput',
    resources: 'resourcesSearchInput',
    materials: 'materialsSearchInput',
    calculators: 'laserMachineInput',
    beregner: 'calcMaterialSearch',
    leverandorer: 'suppliersSearchInput'
};
document.addEventListener('keydown', evt => {
    if (evt.altKey && !evt.ctrlKey && !evt.shiftKey && !evt.metaKey) {
        const idx = Number(evt.key);
        if (idx >= 1 && idx <= navItems.length) {
            evt.preventDefault();
            switchView(navItems[idx - 1].key);
        }
        return;
    }
    if (evt.key === '/' && !evt.ctrlKey && !evt.metaKey) {
        const tag = String((evt.target && evt.target.tagName) || '').toLowerCase();
        if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
        const id = viewSearchFocus[state.view];
        const el = id ? document.getElementById(id) : null;
        if (el) {
            evt.preventDefault();
            el.focus();
            if (el.select) el.select();
        }
    }
});
// Enter/Space aktiverer kunde-/produktkort (tastatur-navigation)
[customersList, productsList].forEach(listEl => {
    if (!listEl) return;
    listEl.addEventListener('keydown', evt => {
        if (evt.key !== 'Enter' && evt.key !== ' ') return;
        const item = evt.target && evt.target.closest ? evt.target.closest('.list-item') : null;
        if (item) {
            evt.preventDefault();
            item.click();
        }
    });
});
document.getElementById('drawingFileInput').addEventListener('change', evt => {
    const file = evt.target.files && evt.target.files[0];
    if (file) analyzeDrawingFile(file);
    scheduleQuoteRecalc(220);
});
const useCustomSheetEl = document.getElementById('calcUseCustomSheet');
const customSheetWEl = document.getElementById('calcCustomSheetW');
const customSheetLEl = document.getElementById('calcCustomSheetL');
function updateCustomSheetState() {
    if (!useCustomSheetEl || !customSheetWEl || !customSheetLEl) return;
    const on = !!useCustomSheetEl.checked;
    customSheetWEl.disabled = !on;
    customSheetLEl.disabled = !on;
}
if (useCustomSheetEl) {
    useCustomSheetEl.addEventListener('change', () => {
        updateCustomSheetState();
        scheduleQuoteRecalc(180);
    });
    updateCustomSheetState();
}
if (customSheetWEl) {
    customSheetWEl.addEventListener('input', () => scheduleQuoteRecalc(180));
    customSheetWEl.addEventListener('change', () => scheduleQuoteRecalc(180));
}
if (customSheetLEl) {
    customSheetLEl.addEventListener('input', () => scheduleQuoteRecalc(180));
    customSheetLEl.addEventListener('change', () => scheduleQuoteRecalc(180));
}
['nestingQtyMode', 'nestingCustomQty', 'calcPreviewSheets'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('input', () => scheduleQuoteRecalc(180));
    el.addEventListener('change', () => scheduleQuoteRecalc(180));
});
const nestingModeEl = document.getElementById('nestingQtyMode');
const nestingCustomEl = document.getElementById('nestingCustomQty');
function updateNestingCustomState() {
    if (!nestingModeEl || !nestingCustomEl) return;
    const isCustom = nestingModeEl.value === 'custom';
    nestingCustomEl.disabled = !isCustom;
}
if (nestingModeEl && nestingCustomEl) {
    nestingModeEl.addEventListener('change', updateNestingCustomState);
    updateNestingCustomState();
}
const liveChk = document.getElementById('calcLiveUpdateChk');
if (liveChk) {
    liveChk.addEventListener('change', () => {
        if (liveChk.checked) scheduleQuoteRecalc(120);
    });
}

// ── Søge-pickers (Excel-agtig typeahead) ──
attachPicker({
    input: calcMaterialSearch,
    list: calcMaterialList,
    getRows: filteredCalcMaterials,
    rowLabel: materialRowLabel,
    rowSub: materialRowSub,
    onPick: row => {
        state.calcMaterial = row;
        calcMaterialSearch.value = row.ProdNo || '';
        if (customSheetWEl && customSheetLEl) {
            const wMm = Math.round(Number(row.Bredde || 0) * 1000);
            const lMm = Math.round(Number(row['Længde'] || 0) * 1000);
            if (wMm > 0) customSheetWEl.value = wMm;
            if (lMm > 0) customSheetLEl.value = lMm;
        }
        const lagerTxt = row.Bal == null ? 'lager ?' : 'lager ' + Number(row.Bal);
        calcMaterialChosen.textContent = 'Valgt: ' + materialOptionLabel(row) + ' · plade ' + (row.Bredde || '-') + 'x' + (row['Længde'] || '-') + ' m · tyk ' + (row.tykklese == null ? '-' : row.tykklese) + ' · ' + lagerTxt;
        renderDxfViewer();
        scheduleQuoteRecalc(180);
    }
});
['calcThicknessFilter', 'calcOnlyStock'].forEach(id => {
    document.getElementById(id).addEventListener('input', () => {
        if (document.activeElement === calcMaterialSearch || calcMaterialList.classList.contains('open')) {
            calcMaterialSearch.dispatchEvent(new Event('input'));
        }
    });
});
['calcMargin', 'calcGap'].forEach(id => {
    document.getElementById(id).addEventListener('input', () => {
        syncSpacingFromGlobalToLaserCard();
        scheduleQuoteRecalc(220);
    });
});
attachPicker({
    input: document.getElementById('calcCustomerSearch'),
    list: document.getElementById('calcCustomerList'),
    getRows: () => state.customers,
    rowLabel: customerRowLabel,
    rowSub: customerRowSub,
    onPick: async row => {
        state.calcCustomer = row;
        document.getElementById('calcCustomerSearch').value = row.Nm || String(row.CustNo || '');
        applyCalcCustomerPref(row);
        await refreshCalcCustomerMeta(row);
        scheduleQuoteRecalc(180);
    }
});
['calcPriceBasis', 'calcLaserOpstartChk', 'calcLaserOpstartMin', 'calcMinAmount', 'calcMinQty'].forEach(id => {
    document.getElementById(id).addEventListener('change', () => {
        saveCalcCustomerPref();
        scheduleQuoteRecalc(180);
    });
});
['calcPieceW', 'calcPieceL', 'calcQty', 'calcCutLength', 'calcPiercings', 'calcMatrixTiers'].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('input', () => scheduleQuoteRecalc(250));
    el.addEventListener('change', () => scheduleQuoteRecalc(250));
});
if (processCards) {
    processCards.addEventListener('input', evt => {
        const target = evt.target;
        if (!target || !target.matches) return;
        if (target.matches('input,select,textarea')) scheduleQuoteRecalc(250);
    });
    processCards.addEventListener('change', evt => {
        const target = evt.target;
        if (!target || !target.matches) return;
        if (target.matches('input,select,textarea')) scheduleQuoteRecalc(180);
    });
}
attachPicker({
    input: draftMaterialSearch,
    list: draftMaterialList,
    getRows: () => state.materials,
    rowLabel: materialRowLabel,
    rowSub: materialRowSub,
    onPick: row => {
        state.draftMaterial = row;
        draftMaterialSearch.value = row.ProdNo || '';
        draftMaterialChosen.textContent = 'Valgt: ' + materialOptionLabel(row);
    }
});
attachPicker({
    input: draftResourceSearch,
    list: draftResourceList,
    getRows: () => state.resources,
    rowLabel: resourceRowLabel,
    rowSub: resourceRowSub,
    onPick: row => {
        if (!state.draftResources.some(r => String(r.ProdNo) === String(row.ProdNo))) {
            state.draftResources.push(row);
            renderDraftResourceChips();
        }
        draftResourceSearch.value = '';
    }
});
customerSelect.addEventListener('change', () => {
    const row = state.customers.find(item => String(item.CustNo) === String(customerSelect.value));
    state.selectedCustomer = row || null;
    state.selectedProduct = null;
    state.selectedRevision = null;
    productSearchInput.value = '';
    updateContext();
    renderCustomers();
    loadProducts();
    loadCustomerNotes();
});

// ── Opstart ──
renderNav();
switchView('overview');
(async function boot() {
    try {
        await loadCustomers();
        await loadProducts();
        loadCustomerNotes();
        updateContext();
        await primeOverviewCounts();
        setStatus('BOM-arbejdsområdet er klar');
    } catch (err) {
        setStatus('Fejl ved opstart: ' + err.message);
    }
})();
