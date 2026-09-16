const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.join(__dirname, '..');

test('management report endpoints require the dedicated permission', () => {
    const source = fs.readFileSync(path.join(root, 'routes/apiRoutes.js'), 'utf8');
    assert.match(source, /router\.get\('\/ledelsesrapport\/config', requireModulePermission\('ledelsesrapport'\)/);
    assert.match(source, /router\.get\('\/ledelsesrapport\/data', requireModulePermission\('ledelsesrapport'\)/);
    assert.match(source, /\[5, 10, 20, 50\]\.includes\(topCustomers\)/);
    assert.match(source, /monthCount < 1 \|\| monthCount > 36/);
    assert.match(source, /loadDays < 1 \|\| loadDays > 180/);
    assert.match(source, /!allowedAccounts\.has\(account\)/);
});

test('management report is exposed through a separately administered module permission', () => {
    const source = fs.readFileSync(path.join(root, 'server.js'), 'utf8');
    assert.match(source, /ledelsesrapport: 'ledelsesrapport'/);
    assert.match(source, /\['ledelsesrapport', 'Ledelsesrapport'\]/);
    assert.match(source, /data-module-key="ledelsesrapport"/);
});

test('management report page is landscape printable and its client script parses', () => {
    const html = fs.readFileSync(path.join(root, 'assets/ledelsesrapport.html'), 'utf8');
    const script = fs.readFileSync(path.join(root, 'assets/js/ledelsesrapport.js'), 'utf8');
    assert.match(html, /@page \{ size:A4 landscape;/);
    assert.match(html, /id="configDialog"/);
    assert.match(html, /id="accounts"/);
    assert.match(html, /Print liggende/);
    assert.doesNotThrow(() => new Function(script));
    assert.match(script, /Omsætning pr\. måned/);
    assert.match(script, /Største kunder/);
    assert.match(script, /Ordre og budget pr\. uge/);
    assert.doesNotMatch(script, /totalTilbud|label:'Tilbud'/);
    assert.match(script, /data-week-value="ordre"/);
    assert.match(html, /id="orderFromWeek" type="week"/);
    assert.match(html, /id="orderToWeek" type="week"/);
    assert.match(script, /GohReportCharts\.revenueStacked/);
    assert.match(script, /GohReportCharts\.belastningCluster/);
    assert.match(script, /Forventet salg \(åbne ordrer\)/);
    assert.match(script, /Samlet kost/);
    assert.match(script, /Belastning pr\. ressource/);
    assert.match(script, /VIA kostfordeling/);
});

test('shared revenue keeps credits, empty months and Danish precision', () => {
    const charts = require('../assets/js/report-charts');
    const chart = charts.revenueStacked([
        { acNo: 11012, name: '<Salg>', date: '2026-07-01', revenueMio: 2 },
        { acNo: 11015, name: 'Kredit', date: '2026-07-01', revenueMio: -0.5 }
    ], ['2026-07-01', '2026-08-01']);
    assert.match(chart.html, /-0,500 Mio DKK/);
    assert.match(chart.html, /1,500 Mio/);
    assert.match(chart.html, /0,000 Mio/);
    assert.match(chart.html, /#1565c0/);
    assert.match(chart.legend, /&lt;Salg&gt;/);
    assert.doesNotMatch(chart.html, /NaN|Infinity/);
    const shell = fs.readFileSync(path.join(root, 'server.js'), 'utf8');
    assert.match(shell, /GohReportCharts\.revenueStacked\(safeRows, monthKeys\)/);
    assert.match(shell, /GohReportCharts\.belastningCluster/);
});

test('daily resource charts preserve original minute values, backlog, evening and clicks', () => {
    const charts = require('../assets/js/report-charts');
    const rows = [
        { Dato: null, Kap: 100, Resv: 200, Aften: 20 },
        { Dato: '2026-09-15', Kap: 300, Resv: 60, Aften: 10 },
        { Dato: '2026-09-16', Kap: 1200, Resv: 900, Aften: 300 }
    ];
    const days = charts.loadDays(rows, { today: '2026-09-16' });
    assert.equal(days.length, 2);
    assert.equal(days[0].__dayKey, 'before');
    assert.equal(days[0].Kap, 0);
    assert.equal(days[0].Resv, 260);
    assert.equal(days[0].Aften, 30);
    const svg = charts.belastningCluster(rows, { today: '2026-09-16', fit: true });
    assert.match(svg, /Reservationer: 900/);
    assert.match(svg, /Kapacitet: 1\.200/);
    assert.match(svg, /Rest Aften: 300/);
    assert.doesNotMatch(svg, /min-width|onclick|NaN/);
    const interactive = charts.belastningCluster(rows, { today: '2026-09-16', clickable: true, resGr: '11', activeDayKey: 'before' });
    assert.match(interactive, /onBelastningDayColumnClick/);
    assert.match(interactive, /belastning-day-band active/);
});

function reportClientFixture() {
    const source = fs.readFileSync(path.join(root, 'assets/js/ledelsesrapport.js'), 'utf8');
    const body = source.slice(source.indexOf('{') + 1, source.indexOf('    async function loadConfig()'));
    const charts = require('../assets/js/report-charts');
    return new Function('document', 'window', body + ';return { stackedRevenue, orderReportView, lineChart, resourceLoadPanels, render, report };')(
        { getElementById: () => ({}) }, { GohReportCharts: charts });
}

test('revenue net totals sit above the whole column including when credits reduce the net', () => {
    const charts = require('../assets/js/report-charts');
    for (const amounts of [[6.5, -0.604], [2, -3], [-3], [0]]) {
        const chart = charts.revenueStacked(amounts.map((revenueMio, index) => ({
            date: '2026-08-01', acNo: 11012 + index, revenueMio
        })), ['2026-08-01']);
        const labelY = Number(/class="omsaetning-month-total"[^>]* y="([\d.]+)"/.exec(chart.html)[1]);
        for (const bar of chart.html.matchAll(/<rect[^>]* y="([\d.]+)"/g)) assert.ok(labelY + 4 < Number(bar[1]));
        assert.match(chart.html, /paint-order="stroke" stroke="#fff"/);
        assert.doesNotMatch(chart.html, /NaN|Infinity/);
    }
    const chart = charts.revenueStacked([{ date: '2026-08-01', acNo: 11012, revenueMio: 6.5 },
        { date: '2026-08-01', acNo: 11015, revenueMio: -0.604 }], ['2026-08-01']);
    assert.match(chart.html, />5,896 Mio<\/text>/);
});

test('report revenue keeps all selected months and one legend in a single chart', () => {
    const client = reportClientFixture();
    for (const count of [1, 12, 13, 15, 24, 36]) {
        const dates = Array.from({ length: count }, (_, index) => new Date(Date.UTC(2025, 6 + index, 1)).toISOString().slice(0, 10));
        const rows = dates.flatMap(date => [{ date, acNo: 11012, name: 'Salg', revenueMio: 6.5 },
            { date, acNo: 11050, name: 'Regulering', revenueMio: -0.604 }]);
        const html = client.stackedRevenue({ filters: { from: '2025-07', to: dates[count - 1].slice(0, 7) }, revenue: { rows } });
        assert.equal((html.match(/revenue-panel/g) || []).length, 1);
        assert.equal((html.match(/class="omsaetning-chart-svg"/g) || []).length, 1);
        assert.equal((html.match(/class="omsaetning-legend"/g) || []).length, 1);
        assert.equal((html.match(/class="omsaetning-month-total"/g) || []).length, count);
        assert.equal((html.match(/>5,896 Mio<\/text>/g) || []).length, count);
        assert.doesNotMatch(html, /NaN|Infinity/);
    }
});

test('order report fits all weeks in one chart without dropping weekly numbers', () => {
    const client = reportClientFixture();
    for (const count of [1, 17, 20, 21, 52, 157]) {
        const weeklyRows = Array.from({ length: count }, (_, index) => ({ weekKey: String(202601 + index), totalOrd: index * 100, totalBudget: 0 }));
        client.render({ generatedAt: '2026-09-16', filters: { from: '2026-07', to: '2026-09', orderFrom: '2026-W01', orderTo: '2026-W38', loadDays: 30 },
            orders: { weeklyRows, kpis: { avgSumOrd: 1453.22 } }, revenue: { rows: [], topCustomers: [] }, load: { rows: [] }, via: { asOf: '2026-09-16', costs: {} }
        }, { budget: {}, holidays: {} });
        const html = client.report.innerHTML;
        assert.equal((html.match(/class="weekly-chart"/g) || []).length, 1);
        assert.equal((html.match(/data-series="ordreBars"/g) || []).length, count);
        assert.equal((html.match(/data-week-value="ordre"/g) || []).length, count);
        assert.equal((html.match(/data-week-label=/g) || []).length, count);
        assert.equal(html.includes('transform="rotate(-90'), count > 20);
        assert.doesNotMatch(html, /NaN|Infinity/);
    }
});

test('order budget matches manual settings, preserves SSRS mode and skips only empty holiday weeks', () => {
    const charts = require('../assets/js/report-charts');
    const rows = [
        { weekKey: '202632', totalOrd: 1500, totalBudget: 0 },
        { weekKey: '202633', totalOrd: 0, totalBudget: 300 },
        { weekKey: '202634', totalOrd: 100, totalBudget: 400 }
    ];
    const holidays = new Set(['202633', '202634']);
    const defaults = charts.orderRowsForView(rows, {}, holidays);
    assert.equal(defaults[0].totalBudget, (280851 / 1000) * 235 / 52);
    assert.equal(defaults[1].totalBudget, 0);
    assert.equal(defaults[2].totalBudget, defaults[0].totalBudget);
    const config = { dailyBudget: 300000, workDaysPerYear: 220 };
    assert.equal(charts.orderRowsForView(rows, config, holidays, false)[1].totalBudget, 300 * 220 / 52);
    assert.strictEqual(charts.orderRowsForView(rows, { useManualBudget: false }, holidays), rows);
    assert.equal(rows[0].totalBudget, 0);
    assert.deepEqual([...charts.orderHolidayWeeks('202652-202702')], ['202652', '202653', '202701', '202702']);
    assert.deepEqual(charts.orderMovingAverage([100, 0, 200, 300], 3, [false, true, false, false]), [100, 100, 150, 200]);
    assert.deepEqual(charts.orderMovingAverage([100, 0, 200, 300], 3), [100, 50, 100, 500 / 3]);
});

test('report order chart keeps the full trend and source period average without offer or budget numbers', () => {
    const client = reportClientFixture();
    const weeklyRows = Array.from({ length: 18 }, (_, index) => ({ weekKey: String(202601 + index), totalOrd: (index + 1) * 100, totalBudget: 0 }));
    const view = client.orderReportView({ weeklyRows, kpis: { avgSumOrd: 1453.22 } }, { budget: {}, holidays: {} });
    assert.equal(view.rows[16].ma3, 1600);
    assert.equal(view.average, 1453.22);
    const svg = client.lineChart(view.rows, view.average);
    assert.equal((svg.match(/data-series="ordreBars"/g) || []).length, 18);
    assert.equal((svg.match(/data-week-value="ordre"/g) || []).length, 18);
    for (const series of ['totalOrd', 'totalBudget', 'ma3', 'periodAverage']) assert.match(svg, new RegExp('data-series="' + series + '"'));
    assert.match(svg, /Gns\. ordre i perioden: 1\.453,22/);
    assert.doesNotMatch(svg, /Tilbud|data-week-value="budget"|Budget:|NaN|Infinity/);
    const holiday = client.orderReportView({ weeklyRows: [{ weekKey: '202601', totalOrd: 0 }] }, { budget: {}, holidays: { holidayWeeksText: '202601' } });
    assert.equal(holiday.rows[0].skipOrdLine, true);
    assert.equal(holiday.rows[0].isAnomaly, false);
    assert.equal(holiday.average, null);
});

test('resource horizons always stay in one card per resource with all bars inside the SVG', () => {
    const client = reportClientFixture();
    const charts = require('../assets/js/report-charts');
    for (const count of [25, 32, 33, 180]) {
        const rows = Array.from({ length: count }, (_, index) => ({ ResGr: '42', Nm: 'Bore',
            Dato: new Date(Date.UTC(2026, 8, 16 + index)).toISOString().slice(0, 10), Kap: 1110, Resv: 900, Aften: 100 }));
        const html = client.resourceLoadPanels({ generatedAt: '2026-09-16T12:00:00Z', filters: { loadDays: count }, load: { rows } });
        assert.equal((html.match(/data-resource="42"/g) || []).length, 1);
        assert.equal((html.match(/Kapacitet: 1\.110/g) || []).length, count);
        assert.doesNotMatch(html, /1\/2|2\/2|NaN|Infinity/);
        assert.equal(html.includes('resource-grid single-column'), count > 32);
        const svg = charts.belastningCluster(rows, { today: '2026-09-16', fit: true, viewportWidth: count > 32 ? 2400 : 1200 });
        const width = Number(/viewBox="0 0 ([\d.]+)/.exec(svg)[1]);
        assert.ok(width <= (count > 32 ? 1060 : 556));
    }
});

function reportFixture(options = {}) {
    const source = fs.readFileSync(path.join(root, 'routes/apiRoutes.js'), 'utf8');
    const start = source.indexOf('    function parseLedelsesrapportMonth(');
    const end = source.indexOf("    router.get('/aftercalc/:ordno',", start);
    const handlers = new Map();
    const calls = [];
    const dailyRow = { ResGr: '11', Nm: 'Laser', Dato: null, DatoX: '', Kap: 0, Resv: 300, Aften: 10 };
    new Function('router', 'requireModulePermission', 'omsaetningService', 'ordreindgangService', 'fetchBelastningRows',
        'getViaPayload', 'viaContext', 'getConnection', 'sql', 'logEvent', source.slice(start, end))(
        { get(route, guard, handler) { handlers.set(route, { guard, handler }); } },
        permission => (req, res, next) => req.user?.permissions?.[permission] ? next() : res.status(403).json({ error: 'Denied' }),
        { getAccounts: async () => { calls.push('accounts'); return [{ acNo: 11012 }]; }, getSummary: async params => { calls.push({ revenue: params }); return options.revenue ? options.revenue(params) : { rows: [], totalRevenueMio: 0 }; } },
        { getSummary: async params => { calls.push({ orders: params }); return { weeklyRows: [] }; } },
        async params => { calls.push({ load: params }); return params.parity === 1 ? (options.loadRows || [dailyRow]) : []; },
        async () => ({ rows: options.viaRows || [] }), () => ({}), () => {}, {}, () => {}
    );
    return { calls, dailyRow, async request(query, allowed = true) {
        const req = { query, user: { permissions: { ledelsesrapport: allowed } } };
        const result = { status: 200, payload: null };
        const res = { status(code) { result.status = code; return this; }, json(payload) { result.payload = payload; return this; } };
        const endpoint = handlers.get('/ledelsesrapport/data');
        await endpoint.guard(req, res, () => endpoint.handler(req, res));
        return result;
    } };
}

test('report forwards independent weeks without changing revenue months or daily units', async () => {
    const fixture = reportFixture();
    const result = await fixture.request({ from: '2026-07', to: '2026-09', orderFrom: '2025-W52', orderTo: '2026-W02' });
    assert.equal(result.status, 200);
    assert.deepEqual(fixture.calls.find(call => call.orders).orders, { fraWeek: '202552', tilWeek: '202602' });
    const revenue = fixture.calls.find(call => call.revenue).revenue;
    assert.equal(revenue.fra, '202601');
    assert.equal(revenue.til, '202604');
    assert.equal(result.payload.filters.orderFrom, '2025-W52');
    assert.deepEqual(result.payload.load.rows, [fixture.dailyRow]);
});

test('invalid week ranges and unauthorized requests stop before database reads', async () => {
    for (const range of [
        { orderFrom: '2025-W53', orderTo: '2026-W02' },
        { orderFrom: '2026-W20', orderTo: '2026-W10' },
        { orderFrom: '2026-W00', orderTo: '2026-W10' },
        { orderFrom: '2026-W20' },
        { orderFrom: '2020-W01', orderTo: '2026-W01' }
    ]) {
        const fixture = reportFixture();
        assert.equal((await fixture.request({ from: '2026-07', to: '2026-09', ...range })).status, 400);
        assert.deepEqual(fixture.calls, []);
    }
    const fixture = reportFixture();
    assert.equal((await fixture.request({ from: '2026-07', to: '2026-09' }, false)).status, 403);
    assert.deepEqual(fixture.calls, []);
});

test('customer month range has its own ranking and identical ranges reuse the revenue query', async () => {
    const fixture = reportFixture({ revenue: params => ({ totalRevenueMio: 99, rows: [
        { custNo: params.fra === '202507' ? 1 : 2, customerName: params.fra === '202507' ? 'Year customer' : 'Month customer', revenueMio: 3 }
    ] }) });
    const result = await fixture.request({ from: '2026-01', to: '2026-12', customerFrom: '2026-09', customerTo: '2026-09' });
    assert.equal(result.status, 200);
    assert.deepEqual(fixture.calls.filter(call => call.revenue).map(call => [call.revenue.fra, call.revenue.til]), [['202507', '202607'], ['202603', '202604']]);
    assert.equal(result.payload.revenue.rows[0].custNo, 1);
    assert.equal(result.payload.revenue.topCustomers[0].custNo, 2);
    assert.equal(result.payload.filters.customerFrom, '2026-09');
    for (const query of [{}, { customerFrom: '2026-01', customerTo: '2026-12' }]) {
        const same = reportFixture();
        assert.equal((await same.request({ from: '2026-01', to: '2026-12', ...query })).status, 200);
        assert.equal(same.calls.filter(call => call.revenue).length, 1);
    }
});

test('load planning period is inclusive and preserves backlog while VIA filters current open orders by order date', async () => {
    const fixture = reportFixture({ loadRows: [
        { ResGr: '11', Dato: null, Resv: 10 }, { ResGr: '11', Dato: '2027-01-10', Resv: 20 },
        { ResGr: '11', Dato: '2027-01-12', Resv: 30 }, { ResGr: '11', Dato: '2027-01-13', Resv: 40 }
    ], viaRows: [
        { OrderDate: '2026-09-01', CostDataAvailable: true, MaterialCost: 10, SalesValue: 20 },
        { OrderDate: 20260930, CostDataAvailable: false, SalesValue: 30 },
        { OrderDate: '2026-10-01', CostDataAvailable: true, MaterialCost: 900 },
        { OrderDate: null, CostDataAvailable: true, MaterialCost: 100 }
    ] });
    const result = await fixture.request({ from: '2026-01', to: '2026-12', loadFrom: '2027-01-10', loadTo: '2027-01-12', viaFrom: '2026-09-01', viaTo: '2026-09-30' });
    assert.equal(result.status, 200);
    assert.deepEqual(fixture.calls.filter(call => call.load).map(call => [call.load.toDay, call.load.dage]), [['2027-01-10', 2], ['2027-01-10', 2]]);
    assert.equal(result.payload.filters.loadDays, 3);
    assert.deepEqual(result.payload.load.rows.map(row => row.Dato), [null, '2027-01-10', '2027-01-12']);
    assert.equal(result.payload.via.orderCount, 2);
    assert.equal(result.payload.via.unknownCostCount, 1);
    assert.equal(result.payload.via.totalCost, 10);
    assert.equal(result.payload.via.totalSales, 50);
    assert.equal(result.payload.via.excludedUndatedCount, 1);
    const oneDay = reportFixture();
    assert.equal((await oneDay.request({ from: '2026-01', to: '2026-01', loadFrom: '2026-01-01', loadTo: '2026-01-01' })).status, 200);
    assert.equal(oneDay.calls.find(call => call.load).load.dage, 0);
});

test('invalid independent report dates fail before reading accounts or report data', async () => {
    for (const range of [
        { customerFrom: '2026-01' }, { customerFrom: '2026-13', customerTo: '2026-13' },
        { customerFrom: '2026-02', customerTo: '2026-01' }, { customerFrom: '2020-01', customerTo: '2026-01' },
        { loadFrom: '2026-02-29', loadTo: '2026-03-01' }, { loadFrom: '2026-03-01' },
        { loadFrom: '2026-03-02', loadTo: '2026-03-01' }, { loadFrom: '2026-01-01', loadTo: '2026-12-31' },
        { viaFrom: '2026-01-01' }, { viaFrom: '2026-02-30', viaTo: '2026-03-01' },
        { viaFrom: '2026-03-02', viaTo: '2026-03-01' }
    ]) {
        const fixture = reportFixture();
        assert.equal((await fixture.request({ from: '2026-01', to: '2026-12', ...range })).status, 400);
        assert.deepEqual(fixture.calls, []);
    }
});