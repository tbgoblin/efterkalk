const { defaults, compileDefinition, evaluateRows } = require('./bilancioDefinition');
const { createSourceResolver } = require('./bilancioSources');
// Some lines aggregate several leaf Kontogrupper (e.g. 19_Lønomk_faste = 15+16+17),
// mirroring the "Aggr. kontogruppe" rollups defined in the chart of accounts.
const LINES = [
    { type: 'group', codes: ['10_Omsætning'], name: 'Omsætning' },
    { type: 'group', codes: ['12_Vareforbrug'], name: 'Vareforbrug' },
    { type: 'subtotal', name: 'Dækningsbidrag I' },
    { type: 'group', codes: ['15_Lønninger', '16_Socialeudgifter', '17_Personaleudgifter'], name: 'Lønninger, produktion inkl.vikar' },
    { type: 'group', codes: ['20_Produktionsomk'], name: 'Produktionsomkostninger' },
    { type: 'group', codes: ['22_Salgsomk'], name: 'Salgsomkostninger' },
    { type: 'subtotal', name: 'Dækningsbidrag II' },
    { type: 'group', codes: ['23\u001f_Driftsmidler ialt'], name: 'Leje af driftsmidler' },
    { type: 'group', codes: ['24_Lokaleomk'], name: 'Lokaleomkostninger' },
    { type: 'group', codes: ['25_Funktionærløn'], name: 'Funktionærløn' },
    { type: 'group', codes: ['26_Administrationomk'], name: 'Administration' },
    { type: 'group', codes: ['28_Transportomk'], name: 'Transportomkostninger' },
    { type: 'subtotal', name: 'Resultat før afskrivninger og renter' },
    { type: 'group', codes: ['30_Afskrivninger'], name: 'Afskrivninger' },
    { type: 'subtotal', name: 'Resultat før afskrivninger' },
    { type: 'group', codes: ['35_Renteindtægter', '36_Renteudgifter'], name: 'Renter Netto' },
    { type: 'subtotal', name: 'Resultat før SKAT' },
    { type: 'computed', name: 'Skat (beregnet 22%)', rate: -0.22 },
    { type: 'subtotal', name: 'Årets resultat' }
];
// Fiscal periods run Juli(1)..Juni(12); period 1 of fiscal year Y is calendar Juli Y.
function fiscalPeriodToCalendar(year, period) {
    return { calendarYear: year + (period > 6 ? 1 : 0), calendarMonth: ((period - 1 + 6) % 12) + 1 };
}
function todayMonthKeyCopenhagen() {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
        timeZone: 'Europe/Copenhagen', year: 'numeric', month: '2-digit'
    }).formatToParts(new Date()).map(part => [part.type, part.value]));
    return `${parts.year}-${parts.month}`;
}
// Same-month whole-order sales value minus the existing finished-order cost.
// Older snapshots without sales valuation must not be treated as zero sales.
function computeFinishedOrderMargin(payload) {
    const totals = payload && payload.totals;
    if (!totals || totals.finishedNotInvoicedSales == null || totals.finishedNotInvoiced == null) return null;
    const sales = Number(totals.finishedNotInvoicedSales);
    const cost = Number(totals.finishedNotInvoiced);
    return Number.isFinite(sales) && Number.isFinite(cost) ? Math.round((sales - cost) * 100) / 100 : null;
}
function validatePeriod(year, period) {
    if (!Number.isInteger(year) || year < 2000 || year > 2100 || !Number.isInteger(period) || period < 1 || period > 12) {
        throw new Error('Vælg et regnskabsår (2000–2100) og en periode (1–12).');
    }
}
function buildReport(records, year, period) {
    validatePeriod(year, period);
    const rows = [];
    let running = [0, 0, 0, 0];
    let revenueAmounts = null;
    for (const line of LINES) {
        if (line.type === 'subtotal') {
            const subtotalAmounts = running.slice();
            rows.push({
                type: 'subtotal', name: line.name, amounts: subtotalAmounts,
                percentages: subtotalAmounts.map((v, i) => revenueAmounts[i] === 0 ? null : v / revenueAmounts[i] * 100)
            });
            continue;
        }
        if (line.type === 'computed') {
            const amounts = running.map(v => v * line.rate);
            const percentages = amounts.map((v, i) => revenueAmounts[i] === 0 ? null : v / revenueAmounts[i] * 100);
            running = running.map((v, i) => v + amounts[i]);
            rows.push({ type: 'computed', name: line.name, amounts, percentages });
            continue;
        }
        const codes = new Set(line.codes);
        const accounts = records.filter(r => codes.has(String(r.AcGr).trim())).map(r => ({
            account: Number(r.AcNo), name: String(r.Nm || '').trim(),
            amounts: ['Month', 'PriorMonth', 'Ytd', 'PriorYtd'].map(key => {
                const value = Number(r[key]);
                if (!Number.isFinite(value)) throw new Error('Ugyldigt kontobeløb');
                return value === 0 ? 0 : -value;
            })
        }));
        const amounts = [0, 1, 2, 3].map(i => accounts.reduce((sum, a) => sum + a.amounts[i], 0));
        if (!revenueAmounts) revenueAmounts = amounts;
        const percentages = amounts.map((v, i) => revenueAmounts[i] === 0 ? null : v / revenueAmounts[i] * 100);
        running = running.map((v, i) => v + amounts[i]);
        rows.push({ type: 'group', name: line.name, accounts, amounts, percentages });
    }
    return { year, period, currency: 'DKK', generatedAt: new Date().toISOString(), rows };
}
function resolveAssetGroups(chart) {
    const groups = new Set(['47_Driftsmidler_ialt']);
    let changed = true;
    while (changed) {
        changed = false;
        for (const row of chart) {
            const code = String(row.AcGr || '').trim();
            if (code && groups.has(String(row.AgAcGr || '').trim()) && !groups.has(code)) {
                groups.add(code); changed = true;
            }
        }
    }
    return groups;
}
function buildAssetRows(records, year, operatingGroups) {
    const specs = [
        { name: 'Depusitum husleje + driftsmidler', matches: row => [66980, 66981].includes(Number(row.AcNo)) },
        { name: 'Grunde og bygninger', matches: row => String(row.AcGr).trim() === '46_Grunde_og_bygninger' },
        { name: 'Driftsmidler i alt', matches: row => operatingGroups.has(String(row.AcGr).trim()) }
    ];
    const seen = new Set();
    for (const row of records) {
        const key = row.AcNo + '/' + row.Yr;
        if (seen.has(key) || row.Balance == null || !Number.isFinite(Number(row.Balance))) throw new Error('Ugyldig eller dubleret kontosaldo: ' + key);
        seen.add(key);
    }
    for (const yr of [year, year - 1]) for (const ac of [66980, 66981, 61100, 61120, 61850, 61900, 66100]) {
        if (!seen.has(ac + '/' + yr)) throw new Error('Saldo mangler for konto ' + ac);
    }
    const rows = specs.map(spec => {
        const selected = records.filter(spec.matches);
        return { name: spec.name, accounts: [...new Set(selected.map(row => Number(row.AcNo)))],
            amounts: [year, year - 1].map(yr => selected.filter(row => Number(row.Yr) === yr).reduce((sum, row) => sum + Number(row.Balance), 0)) };
    });
    const allAccounts = rows.flatMap(row => row.accounts);
    if (new Set(allAccounts).size !== allAccounts.length) throw new Error('En konto indgår i flere aktivposter.');
    rows.push({ type: 'subtotal', name: 'Anlægsaktiver i alt', amounts: [0, 1].map(i => rows.reduce((sum, row) => sum + row.amounts[i], 0)) });
    const inventoryAccounts = [61100, 61120, 61850, 61900];
    rows.push({ name: 'Varebeholdninger', accounts: inventoryAccounts,
        amounts: [year, year - 1].map(yr => records.filter(row => Number(row.Yr) === yr && inventoryAccounts.includes(Number(row.AcNo))).reduce((sum, row) => sum + Number(row.Balance), 0)) });
    // Trade receivables use the explicitly selected accounts, not the whole group.
    const receivables = records.filter(row => Number(row.AcNo) === 66100);
    rows.push({ name: 'Tilgodehavender fra salg', accounts: [...new Set(receivables.map(row => Number(row.AcNo)))],
        amounts: [year, year - 1].map(yr => receivables.filter(row => Number(row.Yr) === yr).reduce((sum, row) => sum + Number(row.Balance), 0)) });
    const detailedAccounts = rows.filter(row => row.type !== 'subtotal').flatMap(row => row.accounts);
    if (new Set(detailedAccounts).size !== detailedAccounts.length) throw new Error('En konto indgår i flere aktivposter.');
    return rows;
}
function createBilancioService({ getConnection, sql, lagerlisteService, fs, definitionStore }) {
    async function catalog() {
        const pool = await getConnection();
        const result = await pool.request().query('SELECT AcNo,Nm,AcGr FROM Ac ORDER BY AcNo; SELECT AcGr,AgAcGr FROM AcGr ORDER BY AcGr;');
        return { accounts: result.recordsets[0], groups: result.recordsets[1] };
    }
    async function assetBalances(pool, year, period, definition, resolveSources, pnlRows) {
        const request = pool.request().input('year', sql.Int, year).input('period', sql.Int, period);
        const ids = [...new Set(definition.balance.flatMap(row => row.selectedAccounts || []))];
        request.input('accounts', sql.NVarChar(sql.MAX), ids.join(','));
        const result = await request
            .query(`SELECT A.AcNo, A.Nm, A.AcGr, Y.Yr,
                COALESCE(B.DbIB,0)+COALESCE(B.DbCh,0)-COALESCE(B.CrIB,0)-COALESCE(B.CrCh,0) AS Balance
                FROM Ac A
                CROSS JOIN (VALUES (@year),(@year-1)) Y(Yr)
                OUTER APPLY (
                    SELECT TOP (1) B.DbIB,B.DbCh,B.CrIB,B.CrCh
                    FROM AcBal B WHERE B.AcNo=A.AcNo
                        AND (B.Yr<Y.Yr OR (B.Yr=Y.Yr AND B.Pr<=@period))
                    ORDER BY B.Yr DESC,B.Pr DESC
                ) B WHERE A.AcNo IN (SELECT TRY_CONVERT(int,value) FROM STRING_SPLIT(@accounts,','))
                ORDER BY Y.Yr DESC,A.AcNo`);
        const records = result.recordset || [];
        const byAccount = new Map();
        const seen = new Set();
        for (const row of records) {
            const key = row.AcNo + '/' + row.Yr;
            if (seen.has(key)) throw new Error('Dubleret kontosaldo: ' + key);
            seen.add(key);
            const record = byAccount.get(row.AcNo) || { AcNo: row.AcNo, Nm: row.Nm };
            record[Number(row.Yr) === year ? 'Current' : 'Previous'] = row.Balance;
            byAccount.set(row.AcNo, record);
        }
        // Balance formulas may pull in a P&L row: its year-to-date columns line up with this point-in-time snapshot.
        const crossSection = new Map(pnlRows.map(row => [row.id, [row.amounts[2], row.amounts[3]]]));
        return { title: definition.balanceTitle, currency: 'DKK', rows: evaluateRows(definition.balance, [...byAccount.values()], 2, ['Current', 'Previous'], await resolveSources(definition.balance, 2), crossSection) };
    }
    // Closed months use their snapshot; only the current open month uses live values.
    async function finishedOrderMarginForMonth(year, period, todayKey) {
        if (!lagerlisteService || !fs) return null;
        const { calendarYear, calendarMonth } = fiscalPeriodToCalendar(year, period);
        const monthKey = calendarYear + '-' + String(calendarMonth).padStart(2, '0');
        const snapshot = await lagerlisteService.loadMonthlySnapshot({ fs, month: monthKey, includeCurrentSales: true });
        if (snapshot) return computeFinishedOrderMargin(snapshot.current);
        return monthKey === todayKey ? computeFinishedOrderMargin(await lagerlisteService.getCurrent()) : null;
    }
    return { catalog, async report(year, period, override, reportId = 'default') {
        validatePeriod(year, period);
        const config = override || (definitionStore ? await definitionStore.load(reportId) : defaults(LINES));
        const definition = compileDefinition(config, await catalog());
        const todayKey = todayMonthKeyCopenhagen();
        const resolveSources = createSourceResolver({ lagerlisteService, fs, year, period, today: todayKey });
        const pool = await getConnection();
        const request = pool.request().input('year', sql.Int, year).input('period', sql.Int, period);
        const ids = [...new Set(definition.pnl.flatMap(row => row.selectedAccounts || []))];
        request.input('accounts', sql.NVarChar(sql.MAX), ids.join(','));
        const result = await request.query(`SELECT A.AcNo, A.Nm, A.AcGr,
                SUM(CASE WHEN T.AcYr=@year AND T.AcPr=@period THEN COALESCE(T.AcAm,0) ELSE 0 END) AS Month,
                SUM(CASE WHEN T.AcYr=@year-1 AND T.AcPr=@period THEN COALESCE(T.AcAm,0) ELSE 0 END) AS PriorMonth,
                SUM(CASE WHEN T.AcYr=@year THEN COALESCE(T.AcAm,0) ELSE 0 END) AS Ytd,
                SUM(CASE WHEN T.AcYr=@year-1 THEN COALESCE(T.AcAm,0) ELSE 0 END) AS PriorYtd
                FROM Ac A LEFT JOIN AcTr T ON T.AcNo=A.AcNo
                    AND T.AcYr IN (@year,@year-1) AND T.AcPr BETWEEN 1 AND @period
                WHERE A.AcNo IN (SELECT TRY_CONVERT(int,value) FROM STRING_SPLIT(@accounts,','))
                GROUP BY A.AcNo,A.Nm,A.AcGr ORDER BY A.AcNo`);
        const rows = evaluateRows(definition.pnl, result.recordset || [], 4, ['Month', 'PriorMonth', 'Ytd', 'PriorYtd'], await resolveSources(definition.pnl, 4));
        const revenueAmounts = rows.find(row => row.id === definition.revenueRow).amounts;
        for (const row of rows) row.percentages = row.amounts.map((v, i) => v == null || !revenueAmounts[i] ? null : v / revenueAmounts[i] * 100);
        const report = { year, period, reportId: definition.reportId, reportName: definition.reportName, showPnl: definition.showPnl, currency: 'DKK', generatedAt: new Date().toISOString(), rows, revenueAmounts, definitionVersion: definition.version };
        const resultFørSkat = rows.find(row => row.id === definition.beforeTaxRow);
        const finishedMargin = definition.legacyPeriodRows === false ? [null, null] : await Promise.all([
            finishedOrderMarginForMonth(year, period, todayKey),
            finishedOrderMarginForMonth(year - 1, period, todayKey)
        ]);
        const totalAmounts = finishedMargin.map((margin, i) => margin === null || resultFørSkat.amounts[i] == null ? null : resultFørSkat.amounts[i] + margin);
        const pctOfRevenue = (v, i) => v === null || revenueAmounts[i] === 0 ? null : v / revenueAmounts[i] * 100;
        report.periodOnly = [
            {
                type: 'subtotal', name: 'Periodens resultat før skat',
                amounts: [resultFørSkat.amounts[0], resultFørSkat.amounts[1], null, null],
                percentages: [pctOfRevenue(resultFørSkat.amounts[0], 0), pctOfRevenue(resultFørSkat.amounts[1], 1), null, null]
            },
            {
                type: 'computed', name: 'Regulering fortjeneste ej realiseret (VIA)',
                description: 'Færdige SO salgpris − Færdige SO kostpris. Manglende salgspriser hentes fra den aktuelle ordreværdi i Visma.',
                amounts: [finishedMargin[0], finishedMargin[1], null, null],
                percentages: [pctOfRevenue(finishedMargin[0], 0), pctOfRevenue(finishedMargin[1], 1), null, null]
            },
            {
                type: 'subtotal', name: 'Periodens resultat i alt',
                amounts: [totalAmounts[0], totalAmounts[1], null, null],
                percentages: [pctOfRevenue(totalAmounts[0], 0), pctOfRevenue(totalAmounts[1], 1), null, null]
            }
        ];
        if (definition.legacyPeriodRows === false) report.periodOnly = [];
        report.assets = await assetBalances(pool, year, period, definition, resolveSources, rows);
        return report;
    } };
}
module.exports = { LINES, validatePeriod, buildReport, createBilancioService, resolveAssetGroups, buildAssetRows, computeFinishedOrderMargin };
