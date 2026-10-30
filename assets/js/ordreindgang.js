// ── Ordreindgang (incoming-order trend) + modulo di stampa condiviso · client ──
// Estratto verbatim dall'inline script di server.js. Contiene anche i generici
// helper di stampa (buildModulePrintCssText, openModulePrintWindow,
// openConfigurableModulePrintPreview, ecc.) e printOmsaetningReport, perché
// erano fisicamente intrecciati col codice Ordreindgang e condivisi da
// entrambi i moduli nello script originale. Usa i globali condivisi della
// pagina: formatNumber, escapeHtml, escapeHtmlFE, canAccessModule,
// pushModalStack/removeModalStack (dichiarati nell'inline script) e
// getOmsaetningCacheEntry/setOmsaetningCacheEntry/getOmsaetningThresholdInputs/
// scheduleOmsaetningAutoReload (dichiarati in assets/js/omsaetning.js,
// caricato prima di questo file).
function normalizeWeekKeyInput(value) {
    return String(value || '').replace(/[^0-9]/g, '').slice(0, 6);
}

function parseWeekKeyMeta(value) {
    const raw = normalizeWeekKeyInput(value);
    if (!/^[0-9]{6}$/.test(raw)) return null;
    const year = Number(raw.slice(0, 4));
    const week = Number(raw.slice(4, 6));
    if (!Number.isFinite(year) || !Number.isFinite(week) || week < 1 || week > 53) return null;
    return { raw, year, week };
}

function getIsoWeekMeta(dateValue) {
    const d = new Date(Date.UTC(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    const isoYear = d.getUTCFullYear();
    const yearStart = new Date(Date.UTC(isoYear, 0, 1));
    const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return {
        isoYear,
        week,
        weekKey: String(isoYear) + String(week).padStart(2, '0')
    };
}

function getIsoWeekStartDate(isoYear, week) {
    const jan4 = new Date(Date.UTC(isoYear, 0, 4));
    const jan4Day = jan4.getUTCDay() || 7;
    const mondayWeek1 = new Date(jan4);
    mondayWeek1.setUTCDate(jan4.getUTCDate() - jan4Day + 1);
    const weekStart = new Date(mondayWeek1);
    weekStart.setUTCDate(mondayWeek1.getUTCDate() + ((week - 1) * 7));
    return weekStart;
}

function formatWeekLabel(weekKey) {
    const meta = parseWeekKeyMeta(weekKey);
    if (!meta) return String(weekKey || '-');
    return String(meta.year) + '-W' + String(meta.week).padStart(2, '0');
}

function formatDkkDa(value) {
    return Number(value || 0).toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatPctDa(value) {
    if (value === null || value === undefined || !Number.isFinite(Number(value))) return '-';
    return Number(value).toLocaleString('da-DK', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + '%';
}

function shouldShowOrdreindgangTilbudLine() {
    const seriesEl = document.getElementById('ordreindgangSeriesTilbud');
    if (seriesEl) return seriesEl.checked === true;
    const el = document.getElementById('ordreindgangShowTilbud');
    return !!el && el.checked === true;
}

function shouldShowOrdreindgangTrendLine() {
    const seriesEl = document.getElementById('ordreindgangSeriesTrend');
    if (seriesEl) return seriesEl.checked === true;
    const el = document.getElementById('ordreindgangShowTrend');
    return !!el && el.checked === true;
}

function shouldShowOrdreindgangOrdreLine() {
    const el = document.getElementById('ordreindgangSeriesOrdreLine');
    return !(el && el.checked === false);
}

function shouldShowOrdreindgangBudgetLine() {
    const el = document.getElementById('ordreindgangSeriesBudget');
    return !(el && el.checked === false);
}

function sanitizeOrdreindgangSeriesConfig(rawConfig) {
    const cfg = rawConfig && typeof rawConfig === 'object' ? rawConfig : {};
    return {
        showTilbud: !(cfg.showTilbud === false),
        showTrend: !(cfg.showTrend === false),
        showOrdreLine: !(cfg.showOrdreLine === false),
        showBudgetLine: !(cfg.showBudgetLine === false)
    };
}

function loadOrdreindgangSeriesConfig() {
    try {
        const raw = localStorage.getItem(ORDREINDGANG_SERIES_STORAGE_KEY);
        if (!raw) return sanitizeOrdreindgangSeriesConfig(null);
        return sanitizeOrdreindgangSeriesConfig(JSON.parse(raw));
    } catch {
        return sanitizeOrdreindgangSeriesConfig(null);
    }
}

function saveOrdreindgangSeriesConfig(config) {
    const safe = sanitizeOrdreindgangSeriesConfig(config);
    try {
        localStorage.setItem(ORDREINDGANG_SERIES_STORAGE_KEY, JSON.stringify(safe));
    } catch {}
    return safe;
}

function applyOrdreindgangSeriesConfigToInputs(config) {
    const safe = sanitizeOrdreindgangSeriesConfig(config);
    const tilbudTopEl = document.getElementById('ordreindgangShowTilbud');
    const trendTopEl = document.getElementById('ordreindgangShowTrend');
    const tilbudLegendEl = document.getElementById('ordreindgangSeriesTilbud');
    const trendLegendEl = document.getElementById('ordreindgangSeriesTrend');
    const ordreLegendEl = document.getElementById('ordreindgangSeriesOrdreLine');
    const budgetLegendEl = document.getElementById('ordreindgangSeriesBudget');
    if (tilbudTopEl) tilbudTopEl.checked = safe.showTilbud;
    if (trendTopEl) trendTopEl.checked = safe.showTrend;
    if (tilbudLegendEl) tilbudLegendEl.checked = safe.showTilbud;
    if (trendLegendEl) trendLegendEl.checked = safe.showTrend;
    if (ordreLegendEl) ordreLegendEl.checked = safe.showOrdreLine;
    if (budgetLegendEl) budgetLegendEl.checked = safe.showBudgetLine;
}

function getOrdreindgangSeriesConfigFromInputs() {
    const tilbudLegendEl = document.getElementById('ordreindgangSeriesTilbud');
    const trendLegendEl = document.getElementById('ordreindgangSeriesTrend');
    const tilbudTopEl = document.getElementById('ordreindgangShowTilbud');
    const trendTopEl = document.getElementById('ordreindgangShowTrend');
    const ordreLegendEl = document.getElementById('ordreindgangSeriesOrdreLine');
    const budgetLegendEl = document.getElementById('ordreindgangSeriesBudget');
    return sanitizeOrdreindgangSeriesConfig({
        showTilbud: tilbudLegendEl ? tilbudLegendEl.checked : (tilbudTopEl ? tilbudTopEl.checked : true),
        showTrend: trendLegendEl ? trendLegendEl.checked : (trendTopEl ? trendTopEl.checked : true),
        showOrdreLine: !(ordreLegendEl && ordreLegendEl.checked === false),
        showBudgetLine: !(budgetLegendEl && budgetLegendEl.checked === false)
    });
}

function onOrdreindgangSeriesChanged(source) {
    const safe = saveOrdreindgangSeriesConfig(getOrdreindgangSeriesConfigFromInputs());
    if (source === 'top') {
        const tilbudLegendEl = document.getElementById('ordreindgangSeriesTilbud');
        const trendLegendEl = document.getElementById('ordreindgangSeriesTrend');
        const tilbudTopEl = document.getElementById('ordreindgangShowTilbud');
        const trendTopEl = document.getElementById('ordreindgangShowTrend');
        if (tilbudLegendEl && tilbudTopEl) tilbudLegendEl.checked = tilbudTopEl.checked;
        if (trendLegendEl && trendTopEl) trendLegendEl.checked = trendTopEl.checked;
    } else {
        const tilbudLegendEl = document.getElementById('ordreindgangSeriesTilbud');
        const trendLegendEl = document.getElementById('ordreindgangSeriesTrend');
        const tilbudTopEl = document.getElementById('ordreindgangShowTilbud');
        const trendTopEl = document.getElementById('ordreindgangShowTrend');
        if (tilbudLegendEl && tilbudTopEl) tilbudTopEl.checked = tilbudLegendEl.checked;
        if (trendLegendEl && trendTopEl) trendTopEl.checked = trendLegendEl.checked;
    }
    applyOrdreindgangSeriesConfigToInputs(safe);
    renderOrdreindgangFromLastPayload();
}

function getOrdreindgangAnomalyThresholdPct() {
    const el = document.getElementById('ordreindgangAnomalyThreshold');
    const raw = Number(el && el.value);
    if (!Number.isFinite(raw) || raw < 0) return 20;
    return raw;
}

function normalizeOrdreindgangHolidayWeeksText(value) {
    return String(value || '')
        .split(',')
        .map(part => {
            const cleaned = String(part || '').trim().replace(/[^0-9\-]/g, '');
            const rangeMatch = /^([0-9]{6})-([0-9]{6})$/.exec(cleaned);
            if (rangeMatch && rangeMatch[2] >= rangeMatch[1]) return rangeMatch[1] + '-' + rangeMatch[2];
            const single = cleaned.replace(/-/g, '').slice(0, 6);
            return /^[0-9]{6}$/.test(single) ? single : '';
        })
        .filter(Boolean)
        .join(',');
}

function parseOrdreindgangHolidayWeeksSet(value) {
    return window.GohReportCharts.orderHolidayWeeks(value);
}

function getOrdreindgangHolidaySettingsFromInputs() {
    const weeksEl = document.getElementById('ordreindgangHolidayWeeks');
    const ignoreEl = document.getElementById('ordreindgangIgnoreHolidayWeeks');
    return {
        holidayWeeksText: normalizeOrdreindgangHolidayWeeksText(weeksEl ? weeksEl.value : ''),
        ignoreHolidayWeeks: !(ignoreEl && ignoreEl.checked === false)
    };
}

function applyOrdreindgangHolidaySettingsToInputs(settings) {
    const safe = settings && typeof settings === 'object' ? settings : {};
    const weeksEl = document.getElementById('ordreindgangHolidayWeeks');
    const ignoreEl = document.getElementById('ordreindgangIgnoreHolidayWeeks');
    const normalized = normalizeOrdreindgangHolidayWeeksText(safe.holidayWeeksText || '');
    if (weeksEl) weeksEl.value = normalized;
    if (ignoreEl) ignoreEl.checked = safe.ignoreHolidayWeeks !== false;
}

function loadOrdreindgangHolidaySettings() {
    try {
        const raw = localStorage.getItem(ORDREINDGANG_HOLIDAY_SETTINGS_STORAGE_KEY);
        if (!raw) return { holidayWeeksText: '', ignoreHolidayWeeks: true };
        const parsed = JSON.parse(raw);
        return {
            holidayWeeksText: normalizeOrdreindgangHolidayWeeksText(parsed && parsed.holidayWeeksText),
            ignoreHolidayWeeks: !(parsed && parsed.ignoreHolidayWeeks === false)
        };
    } catch {
        return { holidayWeeksText: '', ignoreHolidayWeeks: true };
    }
}

function saveOrdreindgangHolidaySettings(settings) {
    const safe = {
        holidayWeeksText: normalizeOrdreindgangHolidayWeeksText(settings && settings.holidayWeeksText),
        ignoreHolidayWeeks: !(settings && settings.ignoreHolidayWeeks === false)
    };
    try {
        localStorage.setItem(ORDREINDGANG_HOLIDAY_SETTINGS_STORAGE_KEY, JSON.stringify(safe));
    } catch {}
    // Delt lagring på GOH: gælder alle maskiner (fail-soft)
    try {
        fetch('/ordreindgang/holiday-settings', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(safe)
        }).catch(() => {});
    } catch {}
    return safe;
}

function getOrdreindgangHolidayWeeksSet() {
    const settings = getOrdreindgangHolidaySettingsFromInputs();
    return parseOrdreindgangHolidayWeeksSet(settings.holidayWeeksText);
}

function shouldIgnoreOrdreindgangHolidayWeeks() {
    const settings = getOrdreindgangHolidaySettingsFromInputs();
    return settings.ignoreHolidayWeeks !== false;
}

function onOrdreindgangHolidaySettingsChanged() {
    const safe = saveOrdreindgangHolidaySettings(getOrdreindgangHolidaySettingsFromInputs());
    applyOrdreindgangHolidaySettingsToInputs(safe);
    renderOrdreindgangFromLastPayload();
    if (omsaetningInitialized && getOmsaetningThresholdInputs().useDailyBudget) {
        scheduleOmsaetningAutoReload();
    }
    if (ordreindgangLastPayload && Array.isArray(ordreindgangLastPayload.weeklyRows)) {
        const range = buildOrdreindgangRange();
        const budgetCfg = getOrdreindgangBudgetConfigFromInputs();
        const budgetTargets = computeOrdreindgangBudgetTargets(budgetCfg);
        const holidayCount = getOrdreindgangRowsForView(ordreindgangLastPayload.weeklyRows).filter(row => Number(row.totalBudget || 0) === 0 && Number(row.totalOrd || 0) === 0).length;
        const budgetText = budgetCfg.useManualBudget
            ? ('Budget manuel: ' + formatDkkDa(budgetTargets.daily) + ' /1000 pr. dag')
            : 'Budget fra SSRS';
        if (range) {
            setOrdreindgangStatus('Periode: ' + formatWeekLabel(range.fraWeek) + ' til ' + formatWeekLabel(range.tilWeek) + ' · rækker: ' + String(ordreindgangLastPayload.weeklyRows.length) + ' · ferieuger: ' + String(holidayCount) + ' · ' + budgetText);
        }
    }
}

function saveOrdreindgangHolidaySettingsManual() {
    onOrdreindgangHolidaySettingsChanged();
    const infoEl = document.getElementById('ordreindgangHolidaySaveInfo');
    if (!infoEl) return;
    infoEl.textContent = 'Gemt';
    infoEl.style.color = '#1b5e20';
    setTimeout(() => {
        if (infoEl.textContent === 'Gemt') infoEl.textContent = '';
    }, 1500);
}

function initializeOrdreindgangHolidaySettings() {
    if (ordreindgangHolidaySettingsInitialized) return;
    ordreindgangHolidaySettingsInitialized = true;
    const saved = loadOrdreindgangHolidaySettings();
    applyOrdreindgangHolidaySettingsToInputs(saved);
    // Delte ferieuger fra GOH vinder over lokale (fail-soft hvis GOH er nede)
    fetch('/ordreindgang/holiday-settings')
        .then(r => r.json())
        .then(d => {
            if (!d || !d.ok || !d.settings || typeof d.settings !== 'object') return;
            const shared = {
                holidayWeeksText: normalizeOrdreindgangHolidayWeeksText(d.settings.holidayWeeksText),
                ignoreHolidayWeeks: !(d.settings.ignoreHolidayWeeks === false)
            };
            try { localStorage.setItem(ORDREINDGANG_HOLIDAY_SETTINGS_STORAGE_KEY, JSON.stringify(shared)); } catch {}
            applyOrdreindgangHolidaySettingsToInputs(shared);
            renderOrdreindgangFromLastPayload();
            if (omsaetningInitialized && getOmsaetningThresholdInputs().useDailyBudget) {
                scheduleOmsaetningAutoReload();
            }
        })
        .catch(() => {});
}

function sanitizeOrdreindgangBudgetConfig(rawConfig) {
    const cfg = rawConfig && typeof rawConfig === 'object' ? rawConfig : {};
    const workDaysRaw = Math.round(Number(cfg.workDaysPerYear));
    const dailyRaw = Number(cfg.dailyBudget);
    const useManual = cfg.useManualBudget !== false;
    const workDaysPerYear = Number.isFinite(workDaysRaw)
        ? Math.max(1, Math.min(366, workDaysRaw))
        : ORDREINDGANG_DEFAULT_WORK_DAYS;
    const dailyBudget = Number.isFinite(dailyRaw)
        ? Math.max(0, dailyRaw)
        : ORDREINDGANG_DEFAULT_DAILY_BUDGET;
    return {
        workDaysPerYear,
        dailyBudget,
        useManualBudget: useManual
    };
}

function loadOrdreindgangBudgetPanelCollapsed() {
    try {
        const stored = localStorage.getItem(ORDREINDGANG_BUDGET_PANEL_STORAGE_KEY);
        if (stored === null || stored === undefined) return true;
        return stored === '1';
    } catch {
        return true;
    }
}

function saveOrdreindgangBudgetPanelCollapsed(collapsed) {
    try {
        localStorage.setItem(ORDREINDGANG_BUDGET_PANEL_STORAGE_KEY, collapsed ? '1' : '0');
    } catch {}
}

function applyOrdreindgangBudgetPanelState() {
    const bodyEl = document.getElementById('ordreindgangBudgetBody');
    const btnEl = document.getElementById('ordreindgangBudgetToggleBtn');
    if (bodyEl) {
        bodyEl.classList.toggle('hidden', ordreindgangBudgetPanelCollapsed);
    }
    if (btnEl) {
        btnEl.textContent = ordreindgangBudgetPanelCollapsed ? 'Vis budget' : 'Skjul budget';
    }
}

function toggleOrdreindgangBudgetPanel() {
    ordreindgangBudgetPanelCollapsed = !ordreindgangBudgetPanelCollapsed;
    saveOrdreindgangBudgetPanelCollapsed(ordreindgangBudgetPanelCollapsed);
    applyOrdreindgangBudgetPanelState();
}

function computeOrdreindgangBudgetTargets(config) {
    return window.GohReportCharts.orderBudgetTargets(sanitizeOrdreindgangBudgetConfig(config));
}

function getOrdreindgangBudgetConfigFromInputs() {
    const workDaysEl = document.getElementById('ordreindgangWorkDays');
    const dailyEl = document.getElementById('ordreindgangDailyBudget');
    const useManualEl = document.getElementById('ordreindgangUseManualBudget');
    return sanitizeOrdreindgangBudgetConfig({
        workDaysPerYear: workDaysEl ? workDaysEl.value : ORDREINDGANG_DEFAULT_WORK_DAYS,
        dailyBudget: dailyEl ? dailyEl.value : ORDREINDGANG_DEFAULT_DAILY_BUDGET,
        useManualBudget: useManualEl ? useManualEl.checked : true
    });
}

function applyOrdreindgangBudgetConfigToInputs(config) {
    const safe = sanitizeOrdreindgangBudgetConfig(config);
    const workDaysEl = document.getElementById('ordreindgangWorkDays');
    const dailyEl = document.getElementById('ordreindgangDailyBudget');
    const useManualEl = document.getElementById('ordreindgangUseManualBudget');
    if (workDaysEl) workDaysEl.value = String(safe.workDaysPerYear);
    if (dailyEl) dailyEl.value = String(Math.round(safe.dailyBudget));
    if (useManualEl) useManualEl.checked = safe.useManualBudget;
    const disabled = !safe.useManualBudget;
    if (workDaysEl) workDaysEl.disabled = disabled;
    if (dailyEl) dailyEl.disabled = disabled;
}

function loadOrdreindgangBudgetConfig() {
    try {
        const raw = localStorage.getItem(ORDREINDGANG_BUDGET_STORAGE_KEY);
        if (!raw) return sanitizeOrdreindgangBudgetConfig(null);
        return sanitizeOrdreindgangBudgetConfig(JSON.parse(raw));
    } catch {
        return sanitizeOrdreindgangBudgetConfig(null);
    }
}

function saveOrdreindgangBudgetConfig(config) {
    const safe = sanitizeOrdreindgangBudgetConfig(config);
    try {
        localStorage.setItem(ORDREINDGANG_BUDGET_STORAGE_KEY, JSON.stringify(safe));
    } catch {}
    return safe;
}

function renderOrdreindgangBudgetMetrics() {
    const metricsEl = document.getElementById('ordreindgangBudgetMetrics');
    if (!metricsEl) return;
    const targets = computeOrdreindgangBudgetTargets(getOrdreindgangBudgetConfigFromInputs());
    const weeklyRows = Array.isArray(ordreindgangLastPayload && ordreindgangLastPayload.weeklyRows)
        ? getOrdreindgangRowsForView(ordreindgangLastPayload.weeklyRows)
        : [];
    const actualPeriod = weeklyRows.reduce((sum, row) => sum + Number(row.totalOrd || 0), 0);
    const expectedPeriod = weeklyRows.reduce((sum, row) => sum + Number(row.totalBudget || 0), 0);
    const deltaPeriod = actualPeriod - expectedPeriod;
    const deltaClass = deltaPeriod > 0.00001
        ? 'delta-up'
        : (deltaPeriod < -0.00001 ? 'delta-down' : 'delta-flat');
    const deltaLabel = deltaPeriod > 0.00001
        ? 'Over budget i periode'
        : (deltaPeriod < -0.00001 ? 'Under budget i periode' : 'På budget i periode');
    metricsEl.innerHTML = '' +
        '<div class="line"><span>Budget / dag (÷1000)</span><strong>' + escapeHtmlFE(formatDkkDa(targets.daily)) + '</strong></div>' +
        '<div class="line"><span>Budget / uge (÷1000)</span><strong>' + escapeHtmlFE(formatDkkDa(targets.weekly)) + '</strong></div>' +
        '<div class="line"><span>Budget / måned (÷1000)</span><strong>' + escapeHtmlFE(formatDkkDa(targets.monthly)) + '</strong></div>' +
        '<div class="line"><span>Budget / år (÷1000)</span><strong>' + escapeHtmlFE(formatDkkDa(targets.annual)) + '</strong></div>' +
        '<div class="line"><span>Forventet i periode</span><strong>' + escapeHtmlFE(formatDkkDa(expectedPeriod)) + '</strong></div>' +
        '<div class="line"><span>Aktuelt i periode</span><strong>' + escapeHtmlFE(formatDkkDa(actualPeriod)) + '</strong></div>' +
        '<div class="line ' + deltaClass + '"><span>Afvigelse</span><strong>' + escapeHtmlFE(formatDkkDa(deltaPeriod)) + '</strong></div>' +
        '<div class="ordreindgang-budget-summary ' + deltaClass + '">' + escapeHtmlFE(deltaLabel) + '</div>';
}

function getOrdreindgangRowsForView(rows) {
    return window.GohReportCharts.orderRowsForView(rows, getOrdreindgangBudgetConfigFromInputs(),
        getOrdreindgangHolidayWeeksSet(), shouldIgnoreOrdreindgangHolidayWeeks());
}

function onOrdreindgangBudgetConfigChanged() {
    const safe = saveOrdreindgangBudgetConfig(getOrdreindgangBudgetConfigFromInputs());
    applyOrdreindgangBudgetConfigToInputs(safe);
    renderOrdreindgangBudgetMetrics();
    renderOrdreindgangFromLastPayload();
    if (ordreindgangLastPayload && Array.isArray(ordreindgangLastPayload.weeklyRows)) {
        const range = buildOrdreindgangRange();
        const targets = computeOrdreindgangBudgetTargets(safe);
        const holidayCount = getOrdreindgangRowsForView(ordreindgangLastPayload.weeklyRows).filter(row => Number(row.totalBudget || 0) === 0 && Number(row.totalOrd || 0) === 0).length;
        const budgetText = safe.useManualBudget
            ? ('Budget manuel: ' + formatDkkDa(targets.daily) + ' /1000 pr. dag')
            : 'Budget fra SSRS';
        if (range) {
            setOrdreindgangStatus('Periode: ' + formatWeekLabel(range.fraWeek) + ' til ' + formatWeekLabel(range.tilWeek) + ' · rækker: ' + String(ordreindgangLastPayload.weeklyRows.length) + ' · ferieuger: ' + String(holidayCount) + ' · ' + budgetText);
        } else {
            setOrdreindgangStatus(budgetText);
        }
    }
}

function initializeOrdreindgangBudgetConfig() {
    const saved = loadOrdreindgangBudgetConfig();
    applyOrdreindgangBudgetConfigToInputs(saved);
    initializeOrdreindgangHolidaySettings();
    const seriesSaved = loadOrdreindgangSeriesConfig();
    applyOrdreindgangSeriesConfigToInputs(seriesSaved);
    ordreindgangBudgetPanelCollapsed = loadOrdreindgangBudgetPanelCollapsed();
    applyOrdreindgangBudgetPanelState();
    renderOrdreindgangBudgetMetrics();
}

function setOrdreindgangStatus(message) {
    const el = document.getElementById('ordreindgangStatus');
    if (el) el.textContent = String(message || '');
}

function buildModulePrintCssText(options) {
    const safeOptions = options && typeof options === 'object' ? options : {};
    const orientation = safeOptions.orientation === 'landscape' ? 'landscape' : 'portrait';
    const reportMaxWidth = orientation === 'landscape' ? '277mm' : '190mm';
    return '' +
        '@page { size: A4 ' + orientation + '; margin: 12mm; }' +
        'body { font-family: Segoe UI, Arial, sans-serif; margin:0; color:#172b3c; background:#fff; }' +
        '.report { max-width: ' + reportMaxWidth + '; margin:0 auto; }' +
        '.report-head { border-bottom:2px solid #d9e6f5; padding:0 0 8px 0; margin:0 0 10px 0; }' +
        '.report-title { font-size:22px; font-weight:800; color:#0f3560; margin:0; }' +
        '.report-sub { margin:3px 0 0 0; font-size:12px; color:#4b6783; }' +
        '.report-meta { display:flex; gap:8px 16px; flex-wrap:wrap; margin-top:8px; font-size:12px; color:#355675; }' +
        '.report-meta strong { color:#0f3560; }' +
        '.kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:8px; margin:10px 0 12px 0; }' +
        '.kpi { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#f8fbff; }' +
        '.kpi .lbl { font-size:11px; color:#4f6d8c; text-transform:uppercase; font-weight:700; }' +
        '.kpi .val { margin-top:3px; font-size:16px; font-weight:800; color:#0f3560; }' +
        '.section { margin-top:10px; page-break-inside:avoid; }' +
        '.section h3 { margin:0 0 6px 0; font-size:14px; color:#214867; border-bottom:1px solid #e2ebf7; padding-bottom:4px; }' +
        '.chart-box { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#fff; }' +
        '.chart-box svg { width:100%; height:auto; display:block; }' +
        '.legend-line { margin:0 0 6px 0; font-size:12px; color:#355675; font-weight:700; }' +
        '.omsaetning-legend-item,.ordreindgang-legend-item { display:inline-flex; align-items:center; gap:6px; margin-right:10px; font-size:12px; font-weight:700; color:#355675; }' +
        '.omsaetning-legend-swatch,.ordreindgang-legend-swatch { width:12px; height:12px; border-radius:3px; display:inline-block; }' +
        '.context-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; }' +
        '.context-card { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#f8fbff; }' +
        '.context-card h4 { margin:0 0 6px 0; font-size:12px; color:#214867; text-transform:uppercase; }' +
        '.context-card p { margin:0; font-size:12px; color:#355675; }' +
        '.context-line { margin-top:4px; font-size:12px; color:#355675; }' +
        '.pill-list { display:flex; flex-wrap:wrap; gap:4px; margin-top:6px; }' +
        '.pill { display:inline-block; padding:2px 7px; border-radius:999px; background:#e8f1fc; color:#1f476a; font-size:11px; font-weight:700; }' +
        '.report-landscape .kpis { grid-template-columns:repeat(4,minmax(0,1fr)); }' +
        '.report-landscape table { font-size:11px; }' +
        '.table-wrap { border:1px solid #dbe8f9; border-radius:8px; overflow:hidden; }' +
        'table { width:100%; border-collapse:collapse; font-size:12px; }' +
        'th { background:#1565c0; color:#fff; text-align:left; padding:7px 8px; }' +
        'td { border-bottom:1px solid #e7eef8; padding:6px 8px; }' +
        'td[style*="text-align:right"], th[style*="text-align:right"] { text-align:right !important; }' +
        '.muted { color:#6a829b; font-size:11px; }';
}

function buildModulePrintStyles(options) {
    return '<style>' + buildModulePrintCssText(options) + '</style>';
}

function openModulePrintWindow(title, subtitle, metaHtml, kpiHtml, sectionsHtml, options) {
    const safeOptions = options && typeof options === 'object' ? options : {};
    const orientation = safeOptions.orientation === 'landscape' ? 'landscape' : 'portrait';
    const html = '<!doctype html><html><head><meta charset="utf-8" />' +
        '<title>' + escapeHtmlFE(title) + '</title>' +
        buildModulePrintStyles({ orientation }) +
        '</head><body>' +
        '<div class="report report-' + orientation + '">' +
        '<header class="report-head">' +
        '<h1 class="report-title">' + escapeHtmlFE(title) + '</h1>' +
        '<p class="report-sub">' + escapeHtmlFE(subtitle) + '</p>' +
        '<div class="report-meta">' + metaHtml + '</div>' +
        '</header>' +
        '<section class="kpis">' + kpiHtml + '</section>' +
        sectionsHtml +
        '</div>' +
        '</body></html>';

    const existing = document.getElementById('modulePrintFrame');
    if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

    const frame = document.createElement('iframe');
    frame.id = 'modulePrintFrame';
    frame.setAttribute('aria-hidden', 'true');
    frame.style.position = 'fixed';
    frame.style.right = '0';
    frame.style.bottom = '0';
    frame.style.width = '1px';
    frame.style.height = '1px';
    frame.style.border = '0';
    frame.style.opacity = '0';
    document.body.appendChild(frame);

    const cleanup = () => {
        const current = document.getElementById('modulePrintFrame');
        if (current && current.parentNode) current.parentNode.removeChild(current);
    };

    const runPrint = () => {
        const w = frame.contentWindow;
        if (!w) {
            alert('Kunne ikke oprette print-visning. Prøv igen.');
            cleanup();
            return;
        }

        w.onafterprint = cleanup;
        w.focus();
        w.print();
        setTimeout(cleanup, 2000);
    };

    frame.onload = () => {
        setTimeout(runPrint, 120);
    };

    // srcdoc avoids popup blockers and is more reliable in Electron than window.open.
    frame.srcdoc = html;
}

function getOrientationLabelDa(orientation) {
    return orientation === 'landscape' ? 'Liggende' : 'Stående';
}

function getPrintOrientationPreference(moduleKey) {
    const id = moduleKey === 'ordreindgang' ? 'ordreindgangPrintOrientation' : 'omsaetningPrintOrientation';
    const raw = String((document.getElementById(id) || {}).value || 'auto').toLowerCase();
    if (raw === 'portrait' || raw === 'landscape') return raw;
    return 'auto';
}

function resolvePrintOrientation(preference, autoOrientation) {
    if (preference === 'portrait' || preference === 'landscape') return preference;
    return autoOrientation === 'landscape' ? 'landscape' : 'portrait';
}

function getOrientationSourceLabelDa(preference) {
    return preference === 'auto' ? 'Auto' : 'Manuel';
}

function chooseOmsaetningPrintOrientation(context) {
    const c = context && typeof context === 'object' ? context : {};
    if (c.detailsOpen || c.thresholdOpen) return 'landscape';
    if (Number(c.selectedCustomers || 0) > 1) return 'landscape';
    if (Number(c.selectedAccounts || 0) > 6) return 'landscape';
    if (Number(c.periods || 0) > 12) return 'landscape';
    return 'portrait';
}

function chooseOrdreindgangPrintOrientation(context) {
    const c = context && typeof context === 'object' ? context : {};
    if (c.weeklyOpen || c.customersOpen) return 'landscape';
    if (Number(c.weeks || 0) > 20) return 'landscape';
    return 'portrait';
}

function buildModulePreviewReportHtml(state, selectedKeys) {
    const safeState = state && typeof state === 'object' ? state : {};
    const safeSections = Array.isArray(safeState.sections) ? safeState.sections : [];
    const selected = selectedKeys instanceof Set ? selectedKeys : new Set();
    const sectionsHtml = safeSections
        .filter(section => section && section.key && selected.has(section.key))
        .map(section => String(section.html || ''))
        .join('');

    return '<div class="report report-' + (safeState.orientation || 'portrait') + '">' +
        '<header class="report-head">' +
        '<h1 class="report-title">' + escapeHtmlFE(safeState.title || 'Rapport') + '</h1>' +
        '<p class="report-sub">' + escapeHtmlFE(safeState.subtitle || '') + '</p>' +
        '<div class="report-meta">' + String(safeState.metaHtml || '') + '</div>' +
        '</header>' +
        '<section class="kpis">' + String(safeState.kpiHtml || '') + '</section>' +
        sectionsHtml +
        '</div>';
}

function updateModulePrintPreviewContent() {
    if (!modulePrintPreviewState) return;
    const previewEl = document.getElementById('modulePrintPreviewContent');
    if (!previewEl) return;
    const checkboxes = Array.from(document.querySelectorAll('.module-print-section'));
    const selected = new Set(
        checkboxes
            .filter(cb => cb && cb.checked)
            .map(cb => String(cb.getAttribute('data-key') || '').trim())
            .filter(Boolean)
    );
    modulePrintPreviewState.selected = selected;
    previewEl.innerHTML = buildModulePreviewReportHtml(modulePrintPreviewState, selected);
}

function toggleModulePrintPreviewSection() {
    updateModulePrintPreviewContent();
}

function openConfigurableModulePrintPreview(config) {
    const safeConfig = config && typeof config === 'object' ? config : {};
    const sections = Array.isArray(safeConfig.sections)
        ? safeConfig.sections.filter(section => section && section.key && section.label && section.html)
        : [];
    if (sections.length === 0) {
        alert('Ingen sektioner at forhåndsvise for denne rapport.');
        return;
    }

    const selected = new Set(
        sections
            .filter(section => section.defaultChecked !== false)
            .map(section => section.key)
    );

    modulePrintPreviewState = {
        title: String(safeConfig.title || 'Rapport'),
        subtitle: String(safeConfig.subtitle || ''),
        metaHtml: String(safeConfig.metaHtml || ''),
        kpiHtml: String(safeConfig.kpiHtml || ''),
        orientation: safeConfig.orientation === 'landscape' ? 'landscape' : 'portrait',
        sections,
        selected
    };

    const controlsHtml = '<div class="context-card" style="margin-bottom:10px;">' +
        '<h4 style="margin:0 0 6px 0;">Vælg indhold til print</h4>' +
        '<div style="display:flex; flex-wrap:wrap; gap:10px 14px;">' +
        sections.map(section => {
            const checked = selected.has(section.key) ? ' checked' : '';
            return '<label style="display:inline-flex; align-items:center; gap:6px; font-size:12px; color:#244766;">' +
                '<input class="module-print-section" type="checkbox" data-key="' + escapeHtmlFE(section.key) + '" onchange="toggleModulePrintPreviewSection()"' + checked + ' />' +
                '<span>' + escapeHtmlFE(section.label) + '</span>' +
                '</label>';
        }).join('') +
        '</div>' +
        '</div>' +
        '<div id="modulePrintPreviewContent"></div>';

    renderPrintPreview(String(safeConfig.previewTitle || 'Forhåndsvisning - rapport'), controlsHtml, 'module');
    updateModulePrintPreviewContent();
}

function printOmsaetningReport() {
    const chartsWrap = document.getElementById('omsaetningChartsWrap');
    if (!chartsWrap || chartsWrap.style.display === 'none') {
        alert('Ingen Omsætning-data at printe. Tryk Opdater først.');
        return;
    }

    const detailsWrapEl = document.getElementById('omsaetningTableWrap');
    const thresholdWrapEl = document.getElementById('omsaetningThresholdWrap');
    const detailsOpen = !!detailsWrapEl && detailsWrapEl.style.display !== 'none';
    const thresholdOpen = !!thresholdWrapEl && thresholdWrapEl.style.display !== 'none';

    const fra = String((document.getElementById('omsaetningFraMonth') || {}).value || '-');
    const til = String((document.getElementById('omsaetningTilMonth') || {}).value || '-');
    const total = String((document.getElementById('omsaetningTotalMio') || {}).textContent || '-').trim();
    const rows = String((document.getElementById('omsaetningRowsCount') || {}).textContent || '-').trim();
    const periods = String((document.getElementById('omsaetningPeriodsCount') || {}).textContent || '-').trim();
    const stackedSvg = (document.getElementById('omsaetningStackedChart') || {}).outerHTML || '';
    const trendSvg = (document.getElementById('omsaetningTrendChart') || {}).outerHTML || '';
    const legend = (document.getElementById('omsaetningLegend') || {}).innerHTML || '';
    const thresholdTable = (document.getElementById('omsaetningThresholdTable') || {}).innerHTML || '';
    const detailsTable = (document.getElementById('omsaetningTableWrap') || {}).innerHTML || '<div class="muted">Ingen detaljetabel tilgængelig.</div>';
    const customerModeText = String((document.getElementById('omsaetningCustomerMode') || {}).textContent || '').trim();
    const customerThresholdsHtml = (document.getElementById('omsaetningCustomerThresholds') || {}).innerHTML || '';
    const selectedAccounts = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean).map(v => String(v));
    const selectedCustomerEntries = Array.from(omsaetningSelectedCustomers.entries());
    const numericPeriods = Number(periods.replace(/\./g, '').replace(',', '.')) || 0;
    const orientationPreference = getPrintOrientationPreference('omsaetning');
    const autoOrientation = chooseOmsaetningPrintOrientation({
        detailsOpen,
        thresholdOpen,
        selectedCustomers: selectedCustomerEntries.length,
        selectedAccounts: selectedAccounts.length,
        periods: numericPeriods
    });
    const orientation = resolvePrintOrientation(orientationPreference, autoOrientation);

    const accountPills = selectedAccounts.slice(0, 18).map(v => '<span class="pill">' + escapeHtmlFE(v) + '</span>').join('');
    const accountRest = selectedAccounts.length > 18
        ? '<span class="pill">+' + escapeHtmlFE(String(selectedAccounts.length - 18)) + '</span>'
        : '';
    const customerPills = selectedCustomerEntries.slice(0, 12).map(([custNo, custName]) => {
        const label = String(custName || '').trim() || String(custNo || '').trim();
        return '<span class="pill">' + escapeHtmlFE(label + ' (' + String(custNo || '') + ')') + '</span>';
    }).join('');
    const customerRest = selectedCustomerEntries.length > 12
        ? '<span class="pill">+' + escapeHtmlFE(String(selectedCustomerEntries.length - 12)) + '</span>'
        : '';

    const metaHtml =
        '<div><strong>Periode:</strong> ' + escapeHtmlFE(fra + ' → ' + til) + '</div>' +
        '<div><strong>Layoutvalg:</strong> ' + escapeHtmlFE(getOrientationSourceLabelDa(orientationPreference)) + '</div>' +
        '<div><strong>Layout:</strong> ' + escapeHtmlFE(getOrientationLabelDa(orientation)) + '</div>' +
        '<div><strong>Udskrevet:</strong> ' + escapeHtmlFE(new Date().toLocaleString('da-DK')) + '</div>';

    const kpiHtml =
        '<div class="kpi"><div class="lbl">Omsætning (Mio)</div><div class="val">' + escapeHtmlFE(total) + '</div></div>' +
        '<div class="kpi"><div class="lbl">Rækker</div><div class="val">' + escapeHtmlFE(rows) + '</div></div>' +
        '<div class="kpi"><div class="lbl">Perioder</div><div class="val">' + escapeHtmlFE(periods) + '</div></div>' +
        '<div class="kpi"><div class="lbl">Modul</div><div class="val">Omsætning</div></div>';

    const contextSection =
        '<section class="section"><h3>Aktive filtre og visning</h3>' +
            '<div class="context-grid">' +
                '<div class="context-card">' +
                    '<h4>Konti</h4>' +
                    '<p>' + escapeHtmlFE(String(selectedAccounts.length)) + ' aktiv</p>' +
                    '<div class="pill-list">' + accountPills + accountRest + '</div>' +
                '</div>' +
                '<div class="context-card">' +
                    '<h4>Kunder</h4>' +
                    '<p>' + escapeHtmlFE(String(selectedCustomerEntries.length)) + ' valgt</p>' +
                    '<div class="pill-list">' + (selectedCustomerEntries.length > 0 ? (customerPills + customerRest) : '<span class="pill">Ingen kunde</span>') + '</div>' +
                    '<div class="context-line"><strong>Visning:</strong> ' + escapeHtmlFE(customerModeText || 'Standardvisning') + '</div>' +
                '</div>' +
            '</div>' +
        '</section>';

    const customerCompareSection = selectedCustomerEntries.length > 1
        ? '<section class="section"><h3>Kunde-sammenligning (aktiv)</h3>' +
            '<div class="context-card"><p>Flere kunder er valgt. Graf og tabeller er baseret på sammenligning pr. måned.</p></div>' +
            '<div class="context-line"></div>' +
            '<div class="table-wrap">' + (customerThresholdsHtml || '<div class="muted" style="padding:8px;">Ingen kundetærskler tilgængelig.</div>') + '</div>' +
        '</section>'
        : '';

    const thresholdSection = thresholdTable
        ? '<section class="section"><h3>Tærskel-tabel</h3><div class="table-wrap">' + thresholdTable + '</div></section>'
        : '';

    const customerOverviewSection = '<section class="section"><h3>Kundeoversigt</h3>' +
        '<div class="context-card"><p>' + (selectedCustomerEntries.length > 0
            ? ('Valgte kunder: ' + escapeHtmlFE(String(selectedCustomerEntries.length)))
            : 'Ingen kunder valgt i filteret.') + '</p></div>' +
        '<div class="table-wrap" style="margin-top:8px;">' + (customerThresholdsHtml || '<div class="muted" style="padding:8px;">Ingen kundedata at vise.</div>') + '</div>' +
        '</section>';

    const sections = [
        { key: 'context', label: 'Aktive filtre', html: contextSection, defaultChecked: true },
        { key: 'customer-overview', label: 'Kundeoversigt', html: customerOverviewSection, defaultChecked: true },
        { key: 'customer-compare', label: 'Kunde-sammenligning', html: customerCompareSection, defaultChecked: selectedCustomerEntries.length > 1 },
        { key: 'stacked-chart', label: 'Stacked graf', html: '<section class="section"><h3>Stacked graf</h3><div class="chart-box"><div class="legend-line">' + legend + '</div>' + stackedSvg + '</div></section>', defaultChecked: true },
        { key: 'trend-chart', label: 'Trend graf', html: '<section class="section"><h3>Trend graf</h3><div class="chart-box">' + trendSvg + '</div></section>', defaultChecked: true },
        { key: 'thresholds', label: 'Tærskel-tabel', html: thresholdSection, defaultChecked: thresholdOpen },
        { key: 'details', label: 'Detaljer', html: '<section class="section"><h3>Detaljer</h3><div class="table-wrap">' + detailsTable + '</div></section>', defaultChecked: detailsOpen }
    ];

    openConfigurableModulePrintPreview({
        previewTitle: 'Forhåndsvisning - Omsætning',
        title: 'Gantech Operations Hub - Omsætning',
        subtitle: 'Rapportudskrift (' + getOrientationLabelDa(orientation).toLowerCase() + ')',
        metaHtml,
        kpiHtml,
        orientation,
        sections
    });
}

function printOrdreindgangReport() {
    const chartsWrap = document.getElementById('ordreindgangChartsWrap');
    if (!chartsWrap || chartsWrap.style.display === 'none') {
        alert('Ingen Ordreindgang-data at printe. Tryk Opdater først.');
        return;
    }

    const weeklyWrapEl = document.getElementById('ordreindgangWeeklyTable');
    const customersWrapEl = document.getElementById('ordreindgangCustomersTable');
    const weeklyOpen = !!weeklyWrapEl && weeklyWrapEl.style.display !== 'none';
    const customersOpen = !!customersWrapEl && customersWrapEl.style.display !== 'none';

    const fraWeek = weekInputValueToKey(String((document.getElementById('ordreindgangFraWeek') || {}).value || '-'));
    const tilWeek = weekInputValueToKey(String((document.getElementById('ordreindgangTilWeek') || {}).value || '-'));
    const totalOrd = String((document.getElementById('ordreindgangTotalOrd') || {}).textContent || '-').trim();
    const totalTilbud = String((document.getElementById('ordreindgangTotalTilbud') || {}).textContent || '-').trim();
    const avgOrd = String((document.getElementById('ordreindgangAvgOrd') || {}).textContent || '-').trim();
    const conv = String((document.getElementById('ordreindgangConv') || {}).textContent || '-').trim();
    const trendSvg = (document.getElementById('ordreindgangTrendChart') || {}).outerHTML || '';
    const legend = (document.getElementById('ordreindgangLegend') || {}).innerHTML || '';
    const weeklyTable = (document.getElementById('ordreindgangWeeklyTable') || {}).innerHTML || '<div class="muted">Ingen ugetabel tilgængelig.</div>';
    const customerTable = (document.getElementById('ordreindgangCustomersTable') || {}).innerHTML || '<div class="muted">Ingen kundetabel tilgængelig.</div>';
    const statusText = String((document.getElementById('ordreindgangStatus') || {}).textContent || '').trim();
    const tilbudEnabled = !!((document.getElementById('ordreindgangShowTilbud') || {}).checked);
    const trendEnabled = !!((document.getElementById('ordreindgangShowTrend') || {}).checked);
    const thresholdPct = getOrdreindgangAnomalyThresholdPct();
    const budgetCfg = getOrdreindgangBudgetConfigFromInputs();
    const budgetTargets = computeOrdreindgangBudgetTargets(budgetCfg);
    const weekCount = Array.isArray(ordreindgangLastPayload && ordreindgangLastPayload.weeklyRows)
        ? ordreindgangLastPayload.weeklyRows.length
        : 0;
    const customerRowsRaw = Array.isArray(ordreindgangLastPayload && ordreindgangLastPayload.customerRows)
        ? ordreindgangLastPayload.customerRows
        : [];
    const orientationPreference = getPrintOrientationPreference('ordreindgang');
    const autoOrientation = chooseOrdreindgangPrintOrientation({
        weeklyOpen,
        customersOpen,
        weeks: weekCount
    });
    const orientation = resolvePrintOrientation(orientationPreference, autoOrientation);

    const metaHtml =
        '<div><strong>Periode:</strong> ' + escapeHtmlFE(fraWeek + ' → ' + tilWeek) + '</div>' +
        '<div><strong>Tilbud-linje:</strong> ' + (tilbudEnabled ? 'Aktiv' : 'Skjult') + '</div>' +
        '<div><strong>Trendlinje:</strong> ' + (trendEnabled ? 'Aktiv (3 uger)' : 'Skjult') + '</div>' +
        '<div><strong>Anomali-soglia:</strong> ±' + escapeHtmlFE(formatPctDa(thresholdPct)) + '</div>' +
        '<div><strong>Budgetkilde:</strong> ' + (budgetCfg.useManualBudget ? 'Manuel' : 'SSRS') + '</div>' +
        '<div><strong>Budget / dag:</strong> ' + escapeHtmlFE(formatDkkDa(budgetTargets.daily)) + ' DKK</div>' +
        '<div><strong>Budget / uge:</strong> ' + escapeHtmlFE(formatDkkDa(budgetTargets.weekly)) + ' DKK</div>' +
        '<div><strong>Layoutvalg:</strong> ' + escapeHtmlFE(getOrientationSourceLabelDa(orientationPreference)) + '</div>' +
        '<div><strong>Layout:</strong> ' + escapeHtmlFE(getOrientationLabelDa(orientation)) + '</div>' +
        '<div><strong>Udskrevet:</strong> ' + escapeHtmlFE(new Date().toLocaleString('da-DK')) + '</div>';

    const kpiHtml =
        '<div class="kpi"><div class="lbl">Total Ordre</div><div class="val">' + escapeHtmlFE(totalOrd) + '</div></div>' +
        '<div class="kpi"><div class="lbl">Total Tilbud</div><div class="val">' + escapeHtmlFE(totalTilbud) + '</div></div>' +
        '<div class="kpi"><div class="lbl">Gns. Ordre</div><div class="val">' + escapeHtmlFE(avgOrd) + '</div></div>' +
        '<div class="kpi"><div class="lbl">Tilbud → Ordre</div><div class="val">' + escapeHtmlFE(conv) + '</div></div>';

    const contextSection =
        '<section class="section"><h3>Aktive filtre og visning</h3>' +
            '<div class="context-grid">' +
                '<div class="context-card">' +
                    '<h4>Periode</h4>' +
                    '<p>' + escapeHtmlFE(fraWeek + ' → ' + tilWeek) + '</p>' +
                '</div>' +
                '<div class="context-card">' +
                    '<h4>Trend/anomali</h4>' +
                    '<p>' + (tilbudEnabled ? 'Aktiv (vises i graf)' : 'Skjult (kun ordre)') + '</p>' +
                    '<div class="context-line"><strong>Soglia:</strong> ±' + escapeHtmlFE(formatPctDa(thresholdPct)) + '</div>' +
                    '<div class="context-line"><strong>Budget:</strong> ' + (budgetCfg.useManualBudget ? 'Manuel' : 'SSRS') + ' (' + escapeHtmlFE(formatDkkDa(budgetTargets.weekly)) + ' DKK/uge)</div>' +
                    '<div class="context-line"><strong>Status:</strong> ' + escapeHtmlFE(statusText || 'Ingen status') + '</div>' +
                '</div>' +
            '</div>' +
        '</section>';

    const topCustomersRows = customerRowsRaw
        .slice()
        .sort((a, b) => Number(b.ordSum || 0) - Number(a.ordSum || 0))
        .slice(0, 10)
        .map(row => {
            const name = String(row.customerName || '').trim() || String(row.custNo || '-');
            return '<tr>' +
                '<td>' + escapeHtmlFE(name) + '</td>' +
                '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.ordSum || 0)) + '</td>' +
                '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.tilbudSum || 0)) + '</td>' +
                '<td style="text-align:right;">' + escapeHtmlFE(formatPctDa(row.conversionPct)) + '</td>' +
                '</tr>';
        }).join('');

    const customerGeneralSection = '<section class="section"><h3>Kunderapport (generel)</h3>' +
        '<div class="context-card"><p>Top 10 kunder i perioden baseret på ordrebeløb.</p></div>' +
        '<div class="table-wrap" style="margin-top:8px;">' +
        '<table><thead><tr><th>Kunde</th><th style="text-align:right;">Ordre</th><th style="text-align:right;">Tilbud</th><th style="text-align:right;">Tilbud → Ordre</th></tr></thead>' +
        '<tbody>' + (topCustomersRows || '<tr><td colspan="4" class="muted">Ingen kundedata.</td></tr>') + '</tbody></table>' +
        '</div></section>';

    const sections = [
        { key: 'context', label: 'Aktive filtre', html: contextSection, defaultChecked: true },
        { key: 'trend-chart', label: 'Ugeudvikling graf', html: '<section class="section"><h3>Ugeudvikling</h3><div class="chart-box"><div class="legend-line">' + legend + '</div>' + trendSvg + '</div></section>', defaultChecked: true },
        { key: 'weekly-table', label: 'Ugetabel', html: '<section class="section"><h3>Ugetabel</h3><div class="table-wrap">' + weeklyTable + '</div></section>', defaultChecked: weeklyOpen },
        { key: 'top-customers', label: 'Topkunder', html: '<section class="section"><h3>Topkunder</h3><div class="table-wrap">' + customerTable + '</div></section>', defaultChecked: customersOpen },
        { key: 'customer-general', label: 'Kunderapport (generel)', html: customerGeneralSection, defaultChecked: true }
    ];

    openConfigurableModulePrintPreview({
        previewTitle: 'Forhåndsvisning - Ordreindgang',
        title: 'Gantech Operations Hub - Ordreindgang',
        subtitle: 'Rapportudskrift (' + getOrientationLabelDa(orientation).toLowerCase() + ')',
        metaHtml,
        kpiHtml,
        orientation,
        sections
    });
}

// Standardperiode: uge 27 i indeværende år → nuværende uge (året før hvis uge < 27)
function getOrdreindgangDefaultRangeKeys() {
    const nowMeta = getIsoWeekMeta(new Date());
    const currentYear = Number(String(nowMeta.weekKey).slice(0, 4));
    const currentWeek = Number(String(nowMeta.weekKey).slice(4, 6));
    const fraYear = currentWeek >= 27 ? currentYear : (currentYear - 1);
    return { fraKey: String(fraYear) + '27', tilKey: String(nowMeta.weekKey) };
}

function getIsoWeeksInYear(year) {
    // ISO: 28. december ligger altid i årets sidste uge
    const meta = getIsoWeekMeta(new Date(year, 11, 28));
    return Number(String(meta.weekKey).slice(4, 6)) || 52;
}

// Dropdown med uger (mandag-baseret, dansk etiket) — uafhængig af browser-locale
function populateOrdreindgangWeekSelects() {
    const fraEl = document.getElementById('ordreindgangFraWeek');
    const tilEl = document.getElementById('ordreindgangTilWeek');
    if (!fraEl || !tilEl || fraEl.options.length > 0) return;
    const nowMeta = getIsoWeekMeta(new Date());
    const currentYear = Number(String(nowMeta.weekKey).slice(0, 4));
    const currentWeek = Number(String(nowMeta.weekKey).slice(4, 6));
    const fmt = d => String(d.getDate()).padStart(2, '0') + '.' + String(d.getMonth() + 1).padStart(2, '0');
    const options = [];
    for (let year = currentYear - 2; year <= currentYear; year++) {
        const lastWeek = year === currentYear ? currentWeek : getIsoWeeksInYear(year);
        for (let week = 1; week <= lastWeek; week++) {
            const start = getIsoWeekStartDate(year, week);
            const end = new Date(start.getTime());
            end.setDate(end.getDate() + 6);
            const key = String(year) + String(week).padStart(2, '0');
            options.push('<option value="' + key + '">Uge ' + week + ' · ' + year + ' (' + fmt(start) + '–' + fmt(end) + ')</option>');
        }
    }
    options.reverse(); // nyeste øverst
    fraEl.innerHTML = options.join('');
    tilEl.innerHTML = options.join('');
}

function applyOrdreindgangDefaultWeeks() {
    const fraEl = document.getElementById('ordreindgangFraWeek');
    const tilEl = document.getElementById('ordreindgangTilWeek');
    if (!fraEl || !tilEl) return;

    populateOrdreindgangWeekSelects();
    const def = getOrdreindgangDefaultRangeKeys();
    fraEl.value = def.fraKey;
    tilEl.value = def.tilKey;
}

// Konvertering mellem uge-input ("2026-W27") og weekKey ("202627")
function weekInputValueToKey(value) {
    const match = /^(\d{4})-W(\d{2})$/.exec(String(value || '').trim());
    return match ? (match[1] + match[2]) : String(value || '').trim();
}

function weekKeyToInputValue(weekKey) {
    const s = String(weekKey || '').trim();
    return /^\d{6}$/.test(s) ? (s.slice(0, 4) + '-W' + s.slice(4)) : '';
}

function buildOrdreindgangRange() {
    const fraEl = document.getElementById('ordreindgangFraWeek');
    const tilEl = document.getElementById('ordreindgangTilWeek');
    populateOrdreindgangWeekSelects();
    let fraMeta = parseWeekKeyMeta(weekInputValueToKey(fraEl ? fraEl.value : ''));
    let tilMeta = parseWeekKeyMeta(weekInputValueToKey(tilEl ? tilEl.value : ''));
    if (!fraMeta || !tilMeta) {
        // Selvhelbredende: tomme/ugyldige felter får automatisk standardperioden
        const def = getOrdreindgangDefaultRangeKeys();
        if (fraEl) fraEl.value = def.fraKey;
        if (tilEl) tilEl.value = def.tilKey;
        fraMeta = parseWeekKeyMeta(def.fraKey);
        tilMeta = parseWeekKeyMeta(def.tilKey);
    }
    if (!fraMeta || !tilMeta) return null;

    const fromStart = getIsoWeekStartDate(fraMeta.year, fraMeta.week);
    const toStart = getIsoWeekStartDate(tilMeta.year, tilMeta.week);
    if (toStart.getTime() < fromStart.getTime()) return null;

    return {
        fraWeek: fraMeta.raw,
        tilWeek: tilMeta.raw
    };
}

function buildOrdreindgangSummaryCacheKey(range) {
    return JSON.stringify({
        fraWeek: String(range.fraWeek || ''),
        tilWeek: String(range.tilWeek || '')
    });
}

async function fetchOrdreindgangSummaryCached(range, options) {
    const safeOptions = options && typeof options === 'object' ? options : {};
    const forceRefresh = safeOptions.forceRefresh === true;
    const cacheKey = buildOrdreindgangSummaryCacheKey(range);

    if (!forceRefresh) {
        const cached = getOmsaetningCacheEntry(ordreindgangSummaryCache, cacheKey, ORDREINDGANG_SUMMARY_CACHE_TTL_MS);
        if (cached) return cached;
    }

    if (!forceRefresh) {
        const inFlight = ordreindgangSummaryInFlight.get(cacheKey);
        if (inFlight) return await inFlight;
    }

    const reqPromise = (async () => {
        const query = new URLSearchParams({
            fraWeek: String(range.fraWeek),
            tilWeek: String(range.tilWeek)
        });
        const response = await fetch('/ordreindgang/summary?' + query.toString());
        if (!response.ok) throw new Error('HTTP ' + response.status);
        const payload = await response.json();
        setOmsaetningCacheEntry(ordreindgangSummaryCache, cacheKey, payload);
        return payload;
    })();

    ordreindgangSummaryInFlight.set(cacheKey, reqPromise);
    try {
        return await reqPromise;
    } finally {
        ordreindgangSummaryInFlight.delete(cacheKey);
    }
}

function computeOrdreindgangMovingAvg(values, windowSize, skipMask) {
    return window.GohReportCharts.orderMovingAverage(values, windowSize, skipMask);
}

function enrichOrdreindgangTrendRows(rows) {
    const safeRows = Array.isArray(rows) ? rows : [];
    const ordValues = safeRows.map(row => Number(row.totalOrd || 0));
    const holidaySet = getOrdreindgangHolidayWeeksSet();
    const ignoreHolidays = shouldIgnoreOrdreindgangHolidayWeeks();
    const holidayMask = safeRows.map(row => {
        const weekKey = String(row && row.weekKey || '');
        const isHolidayTagged = holidaySet.has(weekKey);
        const isZeroWeek = Number(row && row.totalOrd || 0) === 0;
        return isHolidayTagged && isZeroWeek;
    });
    const zeroWeekMask = safeRows.map(row => Number(row && row.totalOrd || 0) === 0);
    const ma3Values = computeOrdreindgangMovingAvg(ordValues, 3, zeroWeekMask);
    const thresholdPct = getOrdreindgangAnomalyThresholdPct();

    return safeRows.map((row, idx) => {
        const isHoliday = Boolean(holidayMask[idx]);
        const ord = Number(row.totalOrd || 0);
        let prevOrd = null;
        if (idx > 0) {
            let prevIdx = idx - 1;
            while (prevIdx >= 0 && ignoreHolidays && holidayMask[prevIdx]) {
                prevIdx -= 1;
            }
            if (prevIdx >= 0) {
                prevOrd = Number(safeRows[prevIdx].totalOrd || 0);
            }
        }
        const ma3Raw = ma3Values[idx];
        const ma3 = Number.isFinite(Number(ma3Raw)) ? Number(ma3Raw) : null;
        const wowPct = (prevOrd !== null && prevOrd > 0)
            ? ((ord - prevOrd) / prevOrd) * 100
            : null;
        const devFromMa3Pct = ma3 !== null && ma3 > 0
            ? ((ord - ma3) / ma3) * 100
            : null;
        const absDev = Number.isFinite(devFromMa3Pct) ? Math.abs(devFromMa3Pct) : 0;
        const isAnomaly = !zeroWeekMask[idx] && Number.isFinite(devFromMa3Pct) && absDev >= thresholdPct;
        const anomalyDir = isAnomaly
            ? (devFromMa3Pct >= 0 ? 'high' : 'low')
            : 'ok';
        return {
            ...row,
            isHoliday,
            ma3,
            wowPct: isHoliday ? null : wowPct,
            devFromMa3Pct: isHoliday ? null : devFromMa3Pct,
            isAnomaly,
            anomalyDir
        };
    });
}

function buildSmoothPath(points) {
    const safePoints = Array.isArray(points) ? points : [];
    if (safePoints.length === 0) return '';
    if (safePoints.length === 1) return 'M ' + safePoints[0].x + ' ' + safePoints[0].y;

    let path = 'M ' + safePoints[0].x + ' ' + safePoints[0].y;
    for (let i = 0; i < safePoints.length - 1; i += 1) {
        const p0 = i > 0 ? safePoints[i - 1] : safePoints[i];
        const p1 = safePoints[i];
        const p2 = safePoints[i + 1];
        const p3 = (i + 2 < safePoints.length) ? safePoints[i + 2] : p2;

        const cp1x = p1.x + (p2.x - p0.x) / 6;
        const cp1y = p1.y + (p2.y - p0.y) / 6;
        const cp2x = p2.x - (p3.x - p1.x) / 6;
        const cp2y = p2.y - (p3.y - p1.y) / 6;

        path += ' C ' + cp1x + ' ' + cp1y + ', ' + cp2x + ' ' + cp2y + ', ' + p2.x + ' ' + p2.y;
    }
    return path;
}

function renderOrdreindgangTrendChart(rows) {
    const wrap = document.getElementById('ordreindgangChartsWrap');
    const svg = document.getElementById('ordreindgangTrendChart');
    const legendEl = document.getElementById('ordreindgangLegend');
    const safeRows = getOrdreindgangRowsForView(rows);
    const showTilbudLine = shouldShowOrdreindgangTilbudLine();
    const showTrendLine = shouldShowOrdreindgangTrendLine();
    const showOrdreLine = shouldShowOrdreindgangOrdreLine();
    const showBudgetLine = shouldShowOrdreindgangBudgetLine();
    if (!wrap || !svg) return;

    if (safeRows.length === 0) {
        wrap.style.display = 'none';
        svg.innerHTML = '';
        if (legendEl) legendEl.innerHTML = '';
        return;
    }

    const trendRows = enrichOrdreindgangTrendRows(safeRows);
    const labels = trendRows.map(r => {
        const key = String(r.weekKey || '');
        if (/^\d{6}$/.test(key)) return key.slice(0, 4) + '-' + key.slice(4, 6);
        return formatWeekLabel(key);
    });
    const ordValues = trendRows.map(r => Number(r.totalOrd || 0));
    const tilbudValues = trendRows.map(r => Number(r.totalTilbud || 0));
    const budgetValues = trendRows.map(r => Number(r.totalBudget || 0));
    const movingAvgValues = trendRows.map(r => (Number.isFinite(Number(r.ma3)) ? Number(r.ma3) : null));
    const rawPeriodAverage = Number(ordreindgangLastPayload && ordreindgangLastPayload.kpis
        ? ordreindgangLastPayload.kpis.avgSumOrd
        : NaN);
    const periodAverageValue = Number.isFinite(rawPeriodAverage) ? rawPeriodAverage : null;
    const ignoreHolidayWeeks = shouldIgnoreOrdreindgangHolidayWeeks();
    let allValues = showTilbudLine
        ? ordValues.concat(tilbudValues).concat(budgetValues).concat(showTrendLine ? movingAvgValues.filter(v => Number.isFinite(v)) : [])
        : ordValues.concat(budgetValues).concat(showTrendLine ? movingAvgValues.filter(v => Number.isFinite(v)) : []);
    if (periodAverageValue !== null) {
        allValues = allValues.concat([periodAverageValue]);
    }
    const rawMax = Math.max(...allValues, 0);
    const chartMax = rawMax <= 0 ? 1000 : (Math.ceil(rawMax / 1000) * 1000);

    const leftPad = 72;
    const topPad = 36;
    const bottomPad = 108;
    const viewportWidth = Math.max(760, (wrap.clientWidth || 0) - 24);
    const visibleWeeksTarget = Math.max(12, Math.min(24, safeRows.length || 12));
    const slot = Math.max(18, Math.min(36, Math.floor(viewportWidth / visibleWeeksTarget)));
    const width = Math.max(viewportWidth, safeRows.length * slot);
    const height = Math.max(320, Math.min(440, Math.round(window.innerHeight * 0.46)));
    const innerHeight = height - topPad - bottomPad;
    const viewWidth = leftPad + width + 20;
    const toY = val => topPad + ((chartMax - Math.max(0, Number(val || 0))) / chartMax) * innerHeight;
    const toXCenter = idx => leftPad + (idx * slot) + (slot / 2);
    const yZero = toY(0);
    const buildPathWithGaps = (points) => {
        let path = '';
        let open = false;
        for (const point of points) {
            if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y)) {
                open = false;
                continue;
            }
            if (!open) {
                path += (path ? ' ' : '') + 'M' + point.x + ' ' + point.y;
                open = true;
            } else {
                path += ' L' + point.x + ' ' + point.y;
            }
        }
        return path;
    };

    let html = '<g>';
    for (let i = 0; i <= 6; i++) {
        const ratio = i / 6;
        const y = topPad + (innerHeight * ratio);
        const tickVal = chartMax * (1 - ratio);
        html += '<line x1="' + leftPad + '" y1="' + y + '" x2="' + (leftPad + width) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
        html += '<text x="' + (leftPad - 8) + '" y="' + (y + 5) + '" text-anchor="end" font-size="12" fill="#5f7892">' + escapeHtmlFE(Math.round(tickVal).toLocaleString('da-DK')) + '</text>';
    }

    const ordBarWidth = showTilbudLine ? 8 : 12;
    const tilbudBarWidth = 7;

    const labelStep = safeRows.length > 70 ? 4 : (safeRows.length > 52 ? 3 : (safeRows.length > 36 ? 2 : 1));
    trendRows.forEach((row, idx) => {
        const centerX = toXCenter(idx);
        const ordY = toY(ordValues[idx]);
        const ordH = Math.max(1, yZero - ordY);

        if (row.isHoliday) {
            html += '<rect x="' + (centerX - (slot / 2)) + '" y="' + topPad + '" width="' + slot + '" height="' + innerHeight + '" fill="rgba(207,216,220,0.20)"><title>' +
                escapeHtmlFE(labels[idx] + ' markeret som ferieuge') + '</title></rect>';
        }

        if (showTilbudLine) {
            const tilbudY = toY(tilbudValues[idx]);
            const tilbudH = Math.max(1, yZero - tilbudY);
            const tilbudX = centerX - tilbudBarWidth - 1;
            html += '<rect x="' + tilbudX + '" y="' + tilbudY + '" width="' + tilbudBarWidth + '" height="' + tilbudH + '" fill="#8ec3f7" rx="1"><title>' +
                escapeHtmlFE(labels[idx] + ' Tilbud: ' + formatDkkDa(tilbudValues[idx])) + '</title></rect>';
        }

        const ordX = showTilbudLine ? (centerX + 1) : (centerX - (ordBarWidth / 2));
        html += '<rect x="' + ordX + '" y="' + ordY + '" width="' + ordBarWidth + '" height="' + ordH + '" fill="#2f5ea5" rx="1"><title>' +
            escapeHtmlFE(labels[idx] + ' Ordre: ' + formatDkkDa(ordValues[idx])) + '</title></rect>';

        if (row.isAnomaly) {
            const markerColor = row.anomalyDir === 'high' ? '#e65100' : '#1565c0';
            html += '<circle cx="' + centerX + '" cy="' + (ordY - 7) + '" r="3.8" fill="#fff" stroke="' + markerColor + '" stroke-width="2">' +
                '<title>' + escapeHtmlFE(labels[idx] + ' Anomali: dev fra MA3 ' + formatPctDa(row.devFromMa3Pct) + ' (soglia ' + formatPctDa(getOrdreindgangAnomalyThresholdPct()) + ')') + '</title>' +
                '</circle>';
        }

        if ((idx % labelStep) === 0 || idx === trendRows.length - 1) {
            html += '<text x="' + centerX + '" y="' + (topPad + innerHeight + 50) + '" text-anchor="middle" transform="rotate(-90 ' + centerX + ' ' + (topPad + innerHeight + 50) + ')" font-size="12" fill="#5f7892">' + escapeHtmlFE(labels[idx]) + '</text>';
        }
    });

    if (showTrendLine) {
        const movingPoints = movingAvgValues.map((value, idx) => ({
            x: toXCenter(idx),
            y: Number.isFinite(value) ? toY(value) : null
        }));
        const smoothPath = buildPathWithGaps(movingPoints);
        html += '<path d="' + smoothPath + '" fill="none" stroke="#ff6f00" stroke-width="4.5" stroke-linejoin="round" stroke-linecap="round" />';
        movingAvgValues.forEach((value, idx) => {
            if (!Number.isFinite(value)) return;
            const cx = toXCenter(idx);
            const cy = toY(value);
            html += '<circle cx="' + cx + '" cy="' + cy + '" r="3.4" fill="#ff6f00">' +
                '<title>' + escapeHtmlFE(labels[idx] + ' Gns. ordre (3 uger): ' + formatDkkDa(value)) + '</title>' +
                '</circle>';
        });
    }

    if (showOrdreLine) {
        const ordPoints = ordValues.map((value, idx) => ({
            x: toXCenter(idx),
            y: (ignoreHolidayWeeks && trendRows[idx] && trendRows[idx].isHoliday) ? null : toY(value)
        }));
        const ordPath = buildPathWithGaps(ordPoints);
        html += '<path d="' + ordPath + '" fill="none" stroke="#0b3f88" stroke-width="2.8" stroke-linecap="round" />';
        ordPoints.forEach((point, idx) => {
            if (!Number.isFinite(point.y)) return;
            html += '<circle cx="' + point.x + '" cy="' + point.y + '" r="2.6" fill="#0b3f88">' +
                '<title>' + escapeHtmlFE(labels[idx] + ' Ordre (linje): ' + formatDkkDa(ordValues[idx])) + '</title>' +
                '</circle>';
        });
    }

    if (showBudgetLine) {
        const budgetPoints = budgetValues.map((value, idx) => ({
            x: toXCenter(idx),
            y: toY(value)
        }));
        const budgetPath = budgetPoints.map((p, idx) => (idx === 0 ? 'M' : 'L') + p.x + ' ' + p.y).join(' ');
        html += '<path d="' + budgetPath + '" fill="none" stroke="#1b8f3b" stroke-width="2.4" stroke-linecap="round" stroke-dasharray="8 5" />';
        budgetPoints.forEach((point, idx) => {
            html += '<circle cx="' + point.x + '" cy="' + point.y + '" r="2.8" fill="#1b8f3b">' +
                '<title>' + escapeHtmlFE(labels[idx] + ' Budget: ' + formatDkkDa(budgetValues[idx])) + '</title>' +
                '</circle>';
        });
    }
    if (periodAverageValue !== null) {
        const periodAverageY = toY(periodAverageValue);
        const periodAverageLabelY = Math.max(topPad + 14, periodAverageY - 7);
        const periodAverageText = 'Gns. ordre i perioden: ' + formatDkkDa(periodAverageValue);
        html += '<line x1="' + leftPad + '" y1="' + periodAverageY + '" x2="' + (leftPad + width) + '" y2="' + periodAverageY + '" stroke="#7b1fa2" stroke-width="2.6" vector-effect="non-scaling-stroke">' +
            '<title>' + escapeHtmlFE(periodAverageText) + '</title></line>';
        html += '<text x="' + (leftPad + width - 8) + '" y="' + periodAverageLabelY + '" text-anchor="end" font-size="12" font-weight="700" fill="#7b1fa2" paint-order="stroke" stroke="#ffffff" stroke-width="4" stroke-linejoin="round">' +
            escapeHtmlFE(periodAverageText) + '</text>';
    }
    html += '</g>';

    svg.setAttribute('viewBox', '0 0 ' + viewWidth + ' ' + height);
    svg.innerHTML = html;
    if (legendEl) {
        let legendHtml = '';
        legendHtml += '<span class="omsaetning-legend-item" style="margin-right:14px"><label class="ordreindgang-toggle" style="display:inline-flex;align-items:center;gap:6px;"><input id="ordreindgangSeriesOrdreLine" type="checkbox" ' + (showOrdreLine ? 'checked' : '') + ' onchange="onOrdreindgangSeriesChanged()" /><span>Ordre-linje</span></label></span>';
        legendHtml += '<span class="omsaetning-legend-item" style="margin-right:14px"><label class="ordreindgang-toggle" style="display:inline-flex;align-items:center;gap:6px;"><input id="ordreindgangSeriesTrend" type="checkbox" ' + (showTrendLine ? 'checked' : '') + ' onchange="onOrdreindgangSeriesChanged()" /><span>Trendlinje</span></label></span>';
        legendHtml += '<span class="omsaetning-legend-item" style="margin-right:14px"><label class="ordreindgang-toggle" style="display:inline-flex;align-items:center;gap:6px;"><input id="ordreindgangSeriesBudget" type="checkbox" ' + (showBudgetLine ? 'checked' : '') + ' onchange="onOrdreindgangSeriesChanged()" /><span>Budgetlinje</span></label></span>';
        legendHtml += '<span class="omsaetning-legend-item" style="margin-right:14px"><label class="ordreindgang-toggle" style="display:inline-flex;align-items:center;gap:6px;"><input id="ordreindgangSeriesTilbud" type="checkbox" ' + (showTilbudLine ? 'checked' : '') + ' onchange="onOrdreindgangSeriesChanged()" /><span>Tilbud</span></label></span>';
        legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#2f5ea5"></span>Ordre (søjle)</span>';
        if (showOrdreLine) {
            legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#0b3f88"></span>Ordre (linje)</span>';
        }
        if (showTilbudLine) {
            legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#8ec3f7"></span>Tilbud</span>';
        }
        if (showTrendLine) {
            legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#ff6f00"></span>Gns. ordre (3 uger)</span>';
        }
        if (showBudgetLine) {
            legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#1b8f3b"></span>Budget</span>';
        }
        if (periodAverageValue !== null) {
            legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#7b1fa2"></span>Gns. ordre i perioden: ' + escapeHtmlFE(formatDkkDa(periodAverageValue)) + '</span>';
        }
        if (trendRows.some(row => row.isHoliday)) {
            legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#cfd8dc"></span>Ferieuge (0,00)</span>';
        }
        if (trendRows.some(row => row.isAnomaly)) {
            legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#fff;border:2px solid #e65100"></span>Anomali (dev fra MA3)</span>';
        }
        legendEl.innerHTML = legendHtml;
        applyOrdreindgangSeriesConfigToInputs(loadOrdreindgangSeriesConfig());
    }
    wrap.style.display = 'grid';
}

function applyOrdreindgangWeeklyCollapsedState() {
    const tableWrap = document.getElementById('ordreindgangWeeklyTable');
    const toggleBtn = document.getElementById('ordreindgangWeeklyToggleBtn');
    if (!tableWrap || !toggleBtn) return;
    tableWrap.style.display = ordreindgangWeeklyCollapsed ? 'none' : 'block';
    toggleBtn.textContent = ordreindgangWeeklyCollapsed ? 'Vis tabel' : 'Skjul tabel';
}

function applyOrdreindgangCustomersCollapsedState() {
    const tableWrap = document.getElementById('ordreindgangCustomersTable');
    const toggleBtn = document.getElementById('ordreindgangCustomersToggleBtn');
    if (!tableWrap || !toggleBtn) return;
    tableWrap.style.display = ordreindgangCustomersCollapsed ? 'none' : 'block';
    toggleBtn.textContent = ordreindgangCustomersCollapsed ? 'Vis tabel' : 'Skjul tabel';
}

function toggleOrdreindgangWeeklyTable() {
    ordreindgangWeeklyCollapsed = !ordreindgangWeeklyCollapsed;
    applyOrdreindgangWeeklyCollapsedState();
}

function toggleOrdreindgangCustomersTable() {
    ordreindgangCustomersCollapsed = !ordreindgangCustomersCollapsed;
    applyOrdreindgangCustomersCollapsedState();
}

function renderOrdreindgangWeeklyTable(rows) {
    const wrapCard = document.getElementById('ordreindgangWeeklyWrap');
    const wrap = document.getElementById('ordreindgangWeeklyTable');
    const safeRows = getOrdreindgangRowsForView(rows);
    const showTilbudColumn = shouldShowOrdreindgangTilbudLine();
    if (!wrapCard || !wrap) return;
    if (safeRows.length === 0) {
        wrapCard.style.display = 'none';
        wrap.innerHTML = '';
        return;
    }

    const trendRows = enrichOrdreindgangTrendRows(safeRows);

    const body = trendRows.map((row) => {
        const rowClass = row.isAnomaly ? (' is-anomaly-' + row.anomalyDir) : '';
        const cells = [
            '<td>' + escapeHtmlFE(formatWeekLabel(row.weekKey)) + '</td>',
            '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalOrd)) + '</td>'
        ];
        if (showTilbudColumn) {
            cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalTilbud)) + '</td>');
        }
        cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalBudget)) + '</td>');
        cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.ma3)) + '</td>');
        cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatPctDa(row.wowPct)) + '</td>');
        cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatPctDa(row.devFromMa3Pct)) + '</td>');
        const anomalyLabel = row.isHoliday
            ? '<span class="ordreindgang-trend-chip holiday">Ferie</span>'
            : (row.isAnomaly
            ? ('<span class="ordreindgang-trend-chip ' + row.anomalyDir + '">' + (row.anomalyDir === 'high' ? 'Over' : 'Under') + '</span>')
            : '<span class="ordreindgang-trend-chip ok">OK</span>');
        cells.push('<td style="text-align:center;">' + anomalyLabel + '</td>');
        return '<tr class="' + rowClass.trim() + '">' +
            cells.join('') +
            '</tr>';
    }).join('');

    const headers = [
        '<th>Uge</th>',
        '<th class="omsaetning-cell-right">Ordre</th>'
    ];
    if (showTilbudColumn) {
        headers.push('<th class="omsaetning-cell-right">Tilbud</th>');
    }
    headers.push('<th class="omsaetning-cell-right">Budget</th>');
    headers.push('<th class="omsaetning-cell-right">Gns. ordre (3 uger)</th>');
    headers.push('<th class="omsaetning-cell-right">WoW</th>');
    headers.push('<th class="omsaetning-cell-right">Dev fra MA3</th>');
    headers.push('<th class="omsaetning-cell-right">Anomali</th>');

    const colgroup = showTilbudColumn
        ? '<colgroup><col style="width:12%;" /><col style="width:14%;" /><col style="width:14%;" /><col style="width:14%;" /><col style="width:14%;" /><col style="width:11%;" /><col style="width:11%;" /><col style="width:10%;" /></colgroup>'
        : '<colgroup><col style="width:14%;" /><col style="width:16%;" /><col style="width:16%;" /><col style="width:16%;" /><col style="width:13%;" /><col style="width:13%;" /><col style="width:12%;" /></colgroup>';

    wrap.innerHTML = '<table class="omsaetning-table ordreindgang-weekly-table">' +
        colgroup +
        '<thead><tr>' + headers.join('') + '</tr></thead>' +
        '<tbody>' + body + '</tbody></table>';
    wrapCard.style.display = 'block';
    applyOrdreindgangWeeklyCollapsedState();
}

function renderOrdreindgangCustomersTable(rows) {
    const wrapCard = document.getElementById('ordreindgangCustomersWrap');
    const wrap = document.getElementById('ordreindgangCustomersTable');
    const safeRows = Array.isArray(rows) ? rows : [];
    if (!wrapCard || !wrap) return;
    if (safeRows.length === 0) {
        wrapCard.style.display = 'none';
        wrap.innerHTML = '';
        return;
    }

    const body = safeRows.map(row => {
        const label = String(row.customerName || '').trim() || String(row.custNo || '-');
        return '<tr>' +
            '<td>' + escapeHtmlFE(label) + '</td>' +
            '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.ordSum)) + '</td>' +
            '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.tilbudSum)) + '</td>' +
            '<td style="text-align:right;">' + escapeHtmlFE(formatPctDa(row.conversionPct)) + '</td>' +
            '</tr>';
    }).join('');

    wrap.innerHTML = '<table class="omsaetning-table">' +
        '<thead><tr><th>Kunde</th><th>Ordre</th><th>Tilbud</th><th>Tilbud → Ordre</th></tr></thead>' +
        '<tbody>' + body + '</tbody></table>';
    wrapCard.style.display = 'block';
    applyOrdreindgangCustomersCollapsedState();
}

function scheduleOrdreindgangAutoReload() {
    if (ordreindgangAutoReloadTimer) {
        clearTimeout(ordreindgangAutoReloadTimer);
    }
    ordreindgangAutoReloadTimer = setTimeout(() => {
        ordreindgangAutoReloadTimer = null;
        loadOrdreindgangSummary({ silentValidation: true });
    }, ORDREINDGANG_AUTO_RELOAD_DELAY_MS);
}

async function initializeOrdreindgangIfNeeded() {
    if (ordreindgangInitialized) return;
    ordreindgangInitialized = true;
    window.addEventListener('resize', () => {
        if (ordreindgangResizeTimer) clearTimeout(ordreindgangResizeTimer);
        ordreindgangResizeTimer = setTimeout(() => {
            ordreindgangResizeTimer = null;
            renderOrdreindgangFromLastPayload();
        }, 120);
    });
    applyOrdreindgangDefaultWeeks();
    initializeOrdreindgangBudgetConfig();
    await loadOrdreindgangSummary({ forceRefresh: true });
}

function renderOrdreindgangFromLastPayload() {
    if (!ordreindgangLastPayload || !Array.isArray(ordreindgangLastPayload.weeklyRows)) return;
    const rowsForView = getOrdreindgangRowsForView(ordreindgangLastPayload.weeklyRows);
    renderOrdreindgangBudgetMetrics();
    renderOrdreindgangTrendChart(rowsForView);
    renderOrdreindgangWeeklyTable(rowsForView);
}

async function loadOrdreindgangSummary(options) {
    const safeOptions = options && typeof options === 'object' ? options : {};
    const emptyEl = document.getElementById('ordreindgangEmpty');
    const loadBtn = document.getElementById('ordreindgangLoadBtn');
    const range = buildOrdreindgangRange();

    if (!range) {
        setOrdreindgangStatus('Ugyldig ugeperiode. Brug format YYYYWW.');
        if (safeOptions.silentValidation !== true) {
            alert('Ugyldig ugeperiode. Brug format YYYYWW.');
        }
        return;
    }

    const persistedHolidaySettings = saveOrdreindgangHolidaySettings(getOrdreindgangHolidaySettingsFromInputs());
    applyOrdreindgangHolidaySettingsToInputs(persistedHolidaySettings);

    if (loadBtn) loadBtn.disabled = true;
    if (emptyEl) {
        emptyEl.style.display = 'block';
        emptyEl.textContent = 'Henter ordreindgang...';
    }
    setOrdreindgangStatus('Henter data...');

    try {
        const payload = await fetchOrdreindgangSummaryCached(range, { forceRefresh: safeOptions.forceRefresh === true });
        ordreindgangLastPayload = payload;
        const kpis = payload && payload.kpis ? payload.kpis : {};
        const weeklyRows = Array.isArray(payload && payload.weeklyRows) ? payload.weeklyRows : [];
        const customerRows = Array.isArray(payload && payload.customerRows) ? payload.customerRows : [];

        const ordEl = document.getElementById('ordreindgangTotalOrd');
        const tilbudEl = document.getElementById('ordreindgangTotalTilbud');
        const avgEl = document.getElementById('ordreindgangAvgOrd');
        const convEl = document.getElementById('ordreindgangConv');

        if (ordEl) ordEl.textContent = formatDkkDa(kpis.totalOrdSum || 0);
        if (tilbudEl) tilbudEl.textContent = formatDkkDa(kpis.totalTilbudSum || 0);
        if (avgEl) avgEl.textContent = formatDkkDa(kpis.avgSumOrd || 0);
        if (convEl) convEl.textContent = formatPctDa(kpis.conversionPct);

        renderOrdreindgangFromLastPayload();
        renderOrdreindgangCustomersTable(customerRows);

        if (emptyEl) {
            if (weeklyRows.length === 0) {
                emptyEl.style.display = 'block';
                emptyEl.textContent = 'Ingen data i valgt ugeperiode.';
            } else {
                emptyEl.style.display = 'none';
            }
        }

        const budgetCfg = getOrdreindgangBudgetConfigFromInputs();
        const budgetTargets = computeOrdreindgangBudgetTargets(budgetCfg);
        const holidayCount = getOrdreindgangRowsForView(weeklyRows).filter(row => Number(row.totalBudget || 0) === 0 && Number(row.totalOrd || 0) === 0).length;
        const budgetText = budgetCfg.useManualBudget
            ? ('Budget manuel: ' + formatDkkDa(budgetTargets.daily) + ' /1000 pr. dag')
            : 'Budget fra SSRS';
        setOrdreindgangStatus('Periode: ' + formatWeekLabel(range.fraWeek) + ' til ' + formatWeekLabel(range.tilWeek) + ' · rækker: ' + String(weeklyRows.length) + ' · ferieuger: ' + String(holidayCount) + ' · ' + budgetText);
    } catch (err) {
        setOrdreindgangStatus('Fejl ved hentning af ordreindgang.');
        if (emptyEl) {
            emptyEl.style.display = 'block';
            emptyEl.textContent = 'Fejl: ' + (err && err.message ? err.message : 'ukendt fejl');
        }
    } finally {
        if (loadBtn) loadBtn.disabled = false;
    }
}
