// ── goToDashboard + inizializzazione Omsaetning (orari lavoro, budget giornaliero,
// conti, riepilogo mensile) · client ───────────────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: closeSideMenu (assets/js/side-menu.js), escapeHtmlFE
// (assets/js/order-notes.js), renderDashboardWidgets (dichiarato nell'inline
// script) e numerose funzioni di assets/js/omsaetning.js /
// assets/js/ordreindgang.js (initializeOrdreindgangHolidaySettings).
function goToDashboard() {
    orderDetailReturnModule = null;
    const dashboard = document.getElementById('mainDashboard');
    const workspace = document.getElementById('mainWorkspace');
    const ordreoversigt = document.getElementById('mainOrdreoversigt');
    const salgordreVia = document.getElementById('mainSalgordreVia');
    const administration = document.getElementById('mainAdministration');
    const omsaetning = document.getElementById('mainOmsaetning');
    const ordreindgang = document.getElementById('mainOrdreindgang');
    const belastning = document.getElementById('mainBelastning');
    const lagerliste = document.getElementById('mainLagerliste');
    closeSideMenu();
    if (workspace) workspace.style.display = 'none';
    if (ordreoversigt) ordreoversigt.style.display = 'none';
    if (salgordreVia) salgordreVia.style.display = 'none';
    if (administration) administration.style.display = 'none';
    if (omsaetning) omsaetning.style.display = 'none';
    if (ordreindgang) ordreindgang.style.display = 'none';
    if (belastning) belastning.style.display = 'none';
    if (lagerliste) lagerliste.style.display = 'none';
    if (dashboard) dashboard.style.display = 'block';
    renderDashboardWidgets();
    const detailModal = document.getElementById('orderDetailModal');
    const detailBody = document.getElementById('orderDetailModalBody');
    if (detailModal) detailModal.style.display = 'none';
    if (detailBody) detailBody.innerHTML = '';
    document.body.classList.remove('report-modal-open');
    setTimeout(syncStickyOffsets, 0);
}

async function initializeOmsaetningIfNeeded() {
    if (omsaetningInitialized) {
        await Promise.all([loadOmsaetningWorkingDays(), loadOmsaetningDailyBudgetSettings()]);
        if (getOmsaetningThresholdInputs().useDailyBudget) scheduleOmsaetningAutoReload();
        return;
    }
    omsaetningInitialized = true;

    const currentFiscalYear = getCurrentFiscalYearStart();
    omsaetningSelectedFiscalYears = new Set([currentFiscalYear]);
    applySelectedFiscalYearsToInputs();
    renderOmsaetningYearChips(currentFiscalYear);
    applyOmsaetningThresholdInputs(loadOmsaetningThresholdUiConfig());
    initializeOrdreindgangHolidaySettings();
    renderOmsaetningCustomerMode();
    renderOmsaetningCustomerResults();

    await Promise.all([loadOmsaetningWorkingDays(), loadOmsaetningDailyBudgetSettings()]);
    await loadOmsaetningAccounts();
    await loadOmsaetningSummary();
}

async function loadOmsaetningWorkingDays() {
    try {
        const response = await fetch('/omsaetning/working-days');
        const data = await response.json();
        if (!response.ok || !data.ok) return;
        omsaetningWorkingDaysByMonth = typeof OmsaetningDailyThresholds !== 'undefined'
            ? OmsaetningDailyThresholds.normalizeWorkingDaysByMonth(data.months)
            : (data.months || {});
    } catch {}
}

async function loadOmsaetningDailyBudgetSettings() {
    try {
        const response = await fetch('/omsaetning/daily-budget-settings');
        const data = await response.json();
        if (!response.ok || !data.ok) return;
        const current = getOmsaetningThresholdInputs();
        const merged = saveOmsaetningThresholdUiConfig({
            ...current,
            useDailyBudget: data.useDailyBudget === true,
            dailyBreakEvenDkk: data.dailyBreakEvenDkk,
            dailyBudgetDkk: data.dailyBudgetDkk
        });
        applyOmsaetningThresholdInputs(merged);
    } catch {}
}

async function loadOmsaetningAccounts() {
    const list = document.getElementById('omsaetningAccountsList');
    if (!list) return;
    list.innerHTML = '<div class="omsaetning-account-item"><span>Indlæser konti...</span></div>';
    try {
        const response = await fetch('/omsaetning/accounts');
        if (!response.ok) throw new Error('HTTP ' + response.status);
        const payload = await response.json();
        const accounts = Array.isArray(payload.accounts) ? payload.accounts : [];
        omsaetningAccounts = accounts;
        const allAccounts = accounts.map(acc => String(acc.acNo || '').trim()).filter(Boolean);
        const ssrsPreset = allAccounts.filter(acNo => OMSAETNING_SSRS_DEFAULT_ACCOUNTS.has(acNo));
        if (ssrsPreset.length === 0) {
            omsaetningSelectedAccounts = new Set(allAccounts);
        } else {
            omsaetningSelectedAccounts = new Set(ssrsPreset);
        }
        renderOmsaetningAccountsList();
    } catch (err) {
        list.innerHTML = '<div class="omsaetning-account-item"><span>Fejl ved konti</span></div>';
        console.error('loadOmsaetningAccounts failed:', err);
    }
}

async function loadOmsaetningSummary(options) {
    const safeOptions = options && typeof options === 'object' ? options : {};
    const silentValidation = safeOptions.silentValidation === true;
    const loadBtn = document.getElementById('omsaetningLoadBtn');
    const empty = document.getElementById('omsaetningEmpty');
    const tableWrap = document.getElementById('omsaetningTableWrap');
    const detailsWrap = document.getElementById('omsaetningDetailsWrap');
    const thresholdWrap = document.getElementById('omsaetningThresholdWrap');
    const thresholdTable = document.getElementById('omsaetningThresholdTable');
    const chartsWrap = document.getElementById('omsaetningChartsWrap');
    const totalEl = document.getElementById('omsaetningTotalMio');
    const rowsEl = document.getElementById('omsaetningRowsCount');
    const periodsEl = document.getElementById('omsaetningPeriodsCount');

    const periodRange = buildOmsaetningPeriodRange();
    if (!periodRange) {
        if (!silentValidation) {
            alert('Vælg gyldig periode (Fra måned skal være før eller lig Til måned).');
        }
        return;
    }

    const fra = periodRange.fra;
    const til = periodRange.til;
    const monthKeysForPeriod = [];

    const selected = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean);
    if (selected.length === 0) {
        if (!silentValidation) {
            alert('Vælg mindst én konto.');
        }
        return;
    }
    const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys()).filter(Boolean);

    const thresholdConfig = getOmsaetningThresholdInputs();
    applyOmsaetningThresholdInputs(thresholdConfig);

    if (loadBtn) {
        loadBtn.disabled = true;
        loadBtn.textContent = 'Indlæser...';
    }

    try {
        const payload = await fetchOmsaetningSummaryCached(fra, til, selected, selectedCustomers, safeOptions);

        const persistTargets = await resolveOmsaetningThresholdPersistTargets(
            selectedCustomers,
            thresholdConfig,
            safeOptions
        );

        if (persistTargets.length > 0) {
            await persistOmsaetningThresholdsForCustomers(persistTargets, thresholdConfig);

            for (const custNo of persistTargets) {
                omsaetningThresholdsByCustomer.set(custNo, { ...thresholdConfig });
            }
            renderOmsaetningCustomerThresholds();
        }

        const rows = Array.isArray(payload.rows) ? payload.rows : [];
        const uniquePeriods = new Set(monthKeysForPeriod.length > 0 ? monthKeysForPeriod : rows.map(r => normalizeOmsaetningMonthKey(r.date)));

        const monthTotals = new Map();
        for (const monthKey of monthKeysForPeriod) {
            monthTotals.set(String(monthKey), 0);
        }
        for (const row of rows) {
            const monthKey = normalizeOmsaetningMonthKey(row.date);
            const prev = monthTotals.get(monthKey) || 0;
            monthTotals.set(monthKey, prev + Number(row.revenueMio || 0));
        }

        if (totalEl) totalEl.textContent = formatMio(payload.totalRevenueMio || 0);
        if (rowsEl) rowsEl.textContent = formatCount(rows.length);
        if (periodsEl) periodsEl.textContent = formatCount(uniquePeriods.size);

        const sortedMonths = (monthKeysForPeriod.length > 0
            ? monthKeysForPeriod.map(monthKey => [monthKey, Number(monthTotals.get(String(monthKey)) || 0)])
            : Array.from(monthTotals.entries()).sort((a, b) => String(a[0]).localeCompare(String(b[0]))));
        let periodRevenueMio = 0;
        let periodZeroPointMio = 0;
        let periodBudgetMio = 0;
        let thresholdHtml = '<table class="omsaetning-table omsaetning-threshold-table"><thead><tr>' +
            '<th>Måned</th><th class="omsaetning-cell-right">Omsætning (Mio)</th><th>Tærskel</th>' +
            '</tr></thead><tbody>';
        for (const [monthKey, amountMio] of sortedMonths) {
            const selectableMonth = String(monthKey || '').slice(0, 7);
            const monthThresholds = getOmsaetningThresholdsForMonth(selectableMonth, thresholdConfig);
            const warnThreshold = monthThresholds.warnThreshold;
            const goodThreshold = monthThresholds.goodThreshold;
            periodRevenueMio += Number(amountMio || 0);
            periodZeroPointMio += Number(warnThreshold || 0);
            periodBudgetMio += Number(goodThreshold || 0);
            const statusClass = getOmsaetningStatusClass(amountMio, warnThreshold, goodThreshold);
            const monthMioLabel = formatMio(amountMio);
            const monthDkkLabel = formatDkkFromMio(amountMio);
            const gauge = buildOmsaetningGaugeData(amountMio, warnThreshold, goodThreshold);
            const marginPctLabel = formatSigned(gauge.marginPct, 1) + '%';
            const deltaWarnLabel = formatSigned(gauge.deltaWarn, 3) + ' Mio';
            const deltaGoodLabel = formatSigned(gauge.deltaGood, 3) + ' Mio';
            const workdayLabel = monthThresholds.useDailyBudget ? (' · ' + String(monthThresholds.workingDays) + ' arb.dage') : '';
            const zeroGaugeLabel = '0: ' + formatMio(warnThreshold);
            const budgetGaugeLabel = 'Budget: ' + formatMio(goodThreshold);
            thresholdHtml += '<tr class="omsaetning-threshold-row' + (omsaetningSelectedMonthKey === selectableMonth ? ' selected' : '') + '" data-month="' + escapeHtmlFE(selectableMonth) + '" onclick="loadOmsaetningMonthDetail(\'' + escapeHtmlFE(selectableMonth) + '\')" style="cursor:pointer;">' +
                '<td><button type="button" class="omsaetning-month-pick" title="Vis bogførte ordrer og Ordreindgang for måneden">' + escapeHtmlFE(formatMonthDa(monthKey)) + '</button></td>' +
                '<td class="omsaetning-cell-right" title="' + escapeHtmlFE(monthDkkLabel + ' DKK') + '">' + escapeHtmlFE(monthMioLabel) + '</td>' +
                '<td>' +
                    '<div class="omsaetning-gauge-wrap">' +
                        '<div class="omsaetning-gauge-meta">' +
                            '<span class="omsaetning-status ' + statusClass + '">' + escapeHtmlFE(getOmsaetningStatusLabel(statusClass) + workdayLabel) + '</span>' +
                            '<strong>' + escapeHtmlFE(marginPctLabel) + '</strong>' +
                        '</div>' +
                        '<div class="omsaetning-gauge-track">' +
                            '<div class="omsaetning-gauge-fill ' + gauge.fillClass + '" style="left:' + gauge.fillLeft.toFixed(2) + '%;width:' + gauge.fillWidth.toFixed(2) + '%;"></div>' +
                            '<span class="omsaetning-gauge-marker" style="left:' + gauge.zeroLeft.toFixed(2) + '%;"></span>' +
                            '<span class="omsaetning-gauge-marker" style="left:' + gauge.targetLeft.toFixed(2) + '%;"></span>' +
                            '<span class="omsaetning-gauge-point" style="left:' + gauge.pointLeft.toFixed(2) + '%;"></span>' +
                        '</div>' +
                        '<div class="omsaetning-gauge-legend"><span>-30%</span><span>' + escapeHtmlFE(zeroGaugeLabel) + '</span><span>' + escapeHtmlFE(budgetGaugeLabel) + '</span><span>60%</span></div>' +
                        '<div class="omsaetning-gauge-delta">vs 0: <strong>' + escapeHtmlFE(deltaWarnLabel) + '</strong> · vs budget: <strong>' + escapeHtmlFE(deltaGoodLabel) + '</strong></div>' +
                    '</div>' +
                '</td>' +
                '</tr>';
        }
        thresholdHtml += '</tbody></table>';
        const periodZeroDeltaMio = periodRevenueMio - periodZeroPointMio;
        const periodZeroDeltaPct = periodZeroPointMio > 0 ? (periodZeroDeltaMio / periodZeroPointMio) * 100 : 0;
        const periodDeltaMio = periodRevenueMio - periodBudgetMio;
        const periodDeltaPct = periodBudgetMio > 0 ? (periodDeltaMio / periodBudgetMio) * 100 : 0;
        const periodIsOverBudget = periodDeltaMio >= 0;
        const periodStatusLabel = periodIsOverBudget ? 'Over budget for perioden' : 'Under budget for perioden';
        thresholdHtml = '<div class="omsaetning-period-budget">' +
            '<span class="omsaetning-status ' + (periodIsOverBudget ? 'good' : 'low') + '">' + escapeHtmlFE(periodStatusLabel) + '</span>' +
            '<div class="omsaetning-period-budget-values">' +
                '<span>Omsætning: <strong>' + escapeHtmlFE(formatMio(periodRevenueMio)) + ' Mio</strong></span>' +
                '<span>0-punkt: <strong>' + escapeHtmlFE(formatMio(periodZeroPointMio)) + ' Mio</strong></span>' +
                '<span>Budget: <strong>' + escapeHtmlFE(formatMio(periodBudgetMio)) + ' Mio</strong></span>' +
            '</div>' +
            '<div class="omsaetning-period-budget-deltas">' +
                '<span class="' + (periodZeroDeltaMio >= 0 ? 'positive' : 'negative') + '">Forskel til 0-punkt: <strong>' + escapeHtmlFE(formatSigned(periodZeroDeltaMio, 3)) + ' Mio (' + escapeHtmlFE(formatSigned(periodZeroDeltaPct, 1)) + '%)</strong></span>' +
                '<span class="' + (periodDeltaMio >= 0 ? 'positive' : 'negative') + '">Forskel til budget: <strong>' + escapeHtmlFE(formatSigned(periodDeltaMio, 3)) + ' Mio (' + escapeHtmlFE(formatSigned(periodDeltaPct, 1)) + '%)</strong></span>' +
            '</div>' +
        '</div>' + thresholdHtml;

        let html = '<table class="omsaetning-table"><thead><tr>' +
            '<th>Måned</th><th>Konto</th><th>Navn</th><th>Kunde</th><th>Kundenavn</th><th style="text-align:right;">Omsætning (Mio)</th>' +
            '</tr></thead><tbody>';

        if (rows.length === 0) {
            html += '<tr><td colspan="6" style="color:#5f7892;">Ingen bevægelser i perioden (0 for alle måneder).</td></tr>';
        } else {
            for (const row of rows) {
                const rowMioLabel = formatMio(row.revenueMio || 0);
                const rowDkkLabel = formatDkkFromMio(row.revenueMio || 0);
                html += '<tr>' +
                    '<td>' + escapeHtmlFE(formatMonthDa(row.date)) + '</td>' +
                    '<td>' + escapeHtmlFE(String(row.acNo || '')) + '</td>' +
                    '<td>' + escapeHtmlFE(String(row.name || '')) + '</td>' +
                    '<td>' + escapeHtmlFE(row.custNo === null || row.custNo === undefined ? '' : String(row.custNo)) + '</td>' +
                    '<td>' + escapeHtmlFE(String(row.customerName || '')) + '</td>' +
                    '<td style="text-align:right;" title="' + escapeHtmlFE(rowDkkLabel + ' DKK') + '">' + escapeHtmlFE(rowMioLabel) + '</td>' +
                    '</tr>';
            }
        }

        html += '</tbody></table>';
        if (OMSAETNING_SHOW_THRESHOLD_SECTION) {
            if (thresholdTable) thresholdTable.innerHTML = thresholdHtml;
            if (thresholdWrap) thresholdWrap.style.display = 'block';
        } else {
            if (thresholdTable) thresholdTable.innerHTML = '';
            if (thresholdWrap) thresholdWrap.style.display = 'none';
        }
        if (tableWrap) {
            tableWrap.innerHTML = html;
        }
        if (detailsWrap) {
            detailsWrap.style.display = 'block';
            applyOmsaetningDetailsCollapsedState();
        }
        renderOmsaetningCharts(rows, monthKeysForPeriod);
        if (omsaetningSelectedMonthKey) {
            const selectedMonthStillVisible = sortedMonths.some(entry => String(entry[0] || '').slice(0, 7) === omsaetningSelectedMonthKey);
            if (selectedMonthStillVisible) {
                loadOmsaetningMonthDetail(omsaetningSelectedMonthKey, { forceRefresh: safeOptions.forceRefresh === true });
            } else {
                omsaetningSelectedMonthKey = '';
                omsaetningMonthDetailToken += 1;
                const monthDetailWrap = document.getElementById('omsaetningMonthDetailWrap');
                if (monthDetailWrap) monthDetailWrap.style.display = 'none';
            }
        }
        if (empty) empty.style.display = 'none';
    } catch (err) {
        if (tableWrap) {
            tableWrap.style.display = 'none';
            tableWrap.innerHTML = '';
        }
        if (detailsWrap) detailsWrap.style.display = 'none';
        if (thresholdWrap) thresholdWrap.style.display = 'none';
        if (thresholdTable) thresholdTable.innerHTML = '';
        omsaetningMonthDetailToken += 1;
        const monthDetailWrap = document.getElementById('omsaetningMonthDetailWrap');
        if (monthDetailWrap) monthDetailWrap.style.display = 'none';
        if (chartsWrap) chartsWrap.style.display = 'none';
        if (empty) {
            empty.style.display = 'block';
            empty.textContent = 'Fejl ved indlæsning: ' + (err && err.message ? err.message : 'ukendt fejl');
        }
        console.error('loadOmsaetningSummary failed:', err);
    } finally {
        if (loadBtn) {
            loadBtn.disabled = false;
            loadBtn.textContent = 'Opdater';
        }
    }
}

