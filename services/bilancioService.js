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
function previousFiscalPeriod(year, period) {
    return period === 1 ? { year: year - 1, period: 12 } : { year, period: period - 1 };
}
function todayMonthKeyCopenhagen() {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
        timeZone: 'Europe/Copenhagen', year: 'numeric', month: '2-digit'
    }).formatToParts(new Date()).map(part => [part.type, part.value]));
    return `${parts.year}-${parts.month}`;
}
// Matches the "Vare i arbejde" formula in assets/js/lagerliste.js: Færdige SO kostpris + VIA Tid + VIA Laser + VIA Stang + indkøbte dele + VIA Plader.
function computeWorkInProgress(payload) {
    if (!payload) return null;
    const totals = payload.totals || {};
    const categories = payload.categories || {};
    const viaRows = Array.isArray(categories.salgordreVia) ? categories.salgordreVia : [];
    const viaTid = viaRows.reduce((sum, row) => sum + Number(row.TimeCost || 0), 0);
    const viaLaser = viaRows.reduce((sum, row) => sum + Number(row.MaterialCost || 0), 0);
    const viaStang = viaRows.reduce((sum, row) => sum + Number(row.StangCost || 0), 0);
    const viaIndkobt = viaRows.reduce((sum, row) => sum + Number(row.PurchasedPartCost || 0), 0);
    const nestingRows = Array.isArray(categories.nestingCutting) ? categories.nestingCutting : [];
    const nestingCountedValue = row => {
        if (row && row.CountedValue !== undefined && row.CountedValue !== null) return Number(row.CountedValue || 0);
        const value = Number(row && row.Value || 0);
        return value < 0 ? 0 : value;
    };
    const viaPlader = nestingRows.reduce((sum, row) => sum + nestingCountedValue(row), 0);
    return Number(totals.finishedNotInvoiced || 0) + viaTid + viaLaser + viaStang + viaIndkobt + viaPlader;
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
function createBilancioService({ getConnection, sql, lagerlisteService, fs }) {
    // The selected month vs the month before it, mirroring the Lagerliste module: a closed month uses its
    // snapshot, the still-open current month falls back to the live figure (assets/js/lagerliste.js pattern).
    async function workInProgressForMonth(calendarYear, calendarMonth, todayKey) {
        if (!lagerlisteService || !fs) return null;
        const monthKey = calendarYear + '-' + String(calendarMonth).padStart(2, '0');
        const snapshot = await lagerlisteService.loadMonthlySnapshot({ fs, month: monthKey });
        if (snapshot) return computeWorkInProgress(snapshot.current);
        return monthKey === todayKey ? computeWorkInProgress(await lagerlisteService.getCurrent()) : null;
    }
    async function workInProgressChange(year, period, todayKey) {
        const current = fiscalPeriodToCalendar(year, period);
        const previous = previousFiscalPeriod(year, period);
        const previousCalendar = fiscalPeriodToCalendar(previous.year, previous.period);
        const [currentWip, previousWip] = await Promise.all([
            workInProgressForMonth(current.calendarYear, current.calendarMonth, todayKey),
            workInProgressForMonth(previousCalendar.calendarYear, previousCalendar.calendarMonth, todayKey)
        ]);
        return currentWip === null || previousWip === null ? null : currentWip - previousWip;
    }
    return { async report(year, period) {
        validatePeriod(year, period);
        const pool = await getConnection();
        const request = pool.request().input('year', sql.Int, year).input('period', sql.Int, period);
        const leafCodes = [...new Set(LINES.filter(l => l.type === 'group').flatMap(l => l.codes))];
        const groupParams = leafCodes.map((code, i) => {
            request.input('g' + i, sql.NVarChar, code);
            return '@g' + i;
        });
        const result = await request.query(`SELECT A.AcNo, A.Nm, A.AcGr,
                SUM(CASE WHEN T.AcYr=@year AND T.AcPr=@period THEN COALESCE(T.AcAm,0) ELSE 0 END) AS Month,
                SUM(CASE WHEN T.AcYr=@year-1 AND T.AcPr=@period THEN COALESCE(T.AcAm,0) ELSE 0 END) AS PriorMonth,
                SUM(CASE WHEN T.AcYr=@year THEN COALESCE(T.AcAm,0) ELSE 0 END) AS Ytd,
                SUM(CASE WHEN T.AcYr=@year-1 THEN COALESCE(T.AcAm,0) ELSE 0 END) AS PriorYtd
                FROM Ac A LEFT JOIN AcTr T ON T.AcNo=A.AcNo
                    AND T.AcYr IN (@year,@year-1) AND T.AcPr BETWEEN 1 AND @period
                WHERE A.AcGr IN (${groupParams.join(',')})
                GROUP BY A.AcNo,A.Nm,A.AcGr ORDER BY A.AcNo`);
        const report = buildReport(result.recordset || [], year, period);
        const revenueAmounts = report.rows[0].amounts;
        const resultFørSkat = report.rows.find(row => row.name === 'Resultat før SKAT');
        const todayKey = todayMonthKeyCopenhagen();
        const wipChange = await Promise.all([
            workInProgressChange(year, period, todayKey),
            workInProgressChange(year - 1, period, todayKey)
        ]);
        const totalAmounts = wipChange.map((change, i) => change === null ? null : resultFørSkat.amounts[i] + change);
        const pctOfRevenue = (v, i) => v === null || revenueAmounts[i] === 0 ? null : v / revenueAmounts[i] * 100;
        report.periodOnly = [
            {
                type: 'subtotal', name: 'Periodens resultat før skat',
                amounts: [resultFørSkat.amounts[0], resultFørSkat.amounts[1], null, null],
                percentages: [pctOfRevenue(resultFørSkat.amounts[0], 0), pctOfRevenue(resultFørSkat.amounts[1], 1), null, null]
            },
            {
                type: 'computed', name: 'Regulering fortjeneste ej realiseret (VIA)',
                amounts: [wipChange[0], wipChange[1], null, null],
                percentages: [pctOfRevenue(wipChange[0], 0), pctOfRevenue(wipChange[1], 1), null, null]
            },
            {
                type: 'subtotal', name: 'Periodens resultat i alt',
                amounts: [totalAmounts[0], totalAmounts[1], null, null],
                percentages: [pctOfRevenue(totalAmounts[0], 0), pctOfRevenue(totalAmounts[1], 1), null, null]
            }
        ];
        return report;
    } };
}
module.exports = { LINES, validatePeriod, buildReport, createBilancioService };
