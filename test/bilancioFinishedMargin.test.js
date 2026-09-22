const test = require('node:test');
const assert = require('node:assert/strict');
const { createBilancioService, computeFinishedOrderMargin } = require('../services/bilancioService');
const { allocateSharedOrders } = require('../services/lagerlisteAllocation');

const snapshot = (sales, cost) => ({ current: { totals: { finishedNotInvoicedSales: sales, finishedNotInvoiced: cost } } });
function fixture(snapshots, enrichedSnapshots = {}) {
    const requested = [];
    let liveCalls = 0;
    const definition = { schema: 1, version: 0, revenueRow: 'sales', beforeTaxRow: 'sales', balanceTitle: 'Aktiver',
        pnl: [{ id: 'sales', name: 'Sales', type: 'accounts', accounts: [100], groups: [], exclude: [], children: false, sign: -1 }],
        balance: [{ id: 'heading', name: 'Aktiver', type: 'heading' }] };
    const pool = { request: () => ({ input() { return this; }, async query(query) {
        if (query.startsWith('SELECT AcNo')) return { recordsets: [[{ AcNo: 100, Nm: 'Sales' }], []] };
        if (query.includes('OUTER APPLY')) return { recordset: [] };
        return { recordset: [{ AcNo: 100, Nm: 'Sales', Month: -1000, PriorMonth: -800, Ytd: -2000, PriorYtd: -1600 }] };
    } }) };
    const service = createBilancioService({ getConnection: async () => pool,
        sql: { Int: 'Int', NVarChar: () => 'NVarChar' }, fs: {},
        lagerlisteService: {
            loadMonthlySnapshot: async ({ month, includeCurrentSales }) => {
                requested.push(month);
                return (includeCurrentSales && enrichedSnapshots[month]) || snapshots[month] || null;
            },
            getCurrent: async () => { liveCalls++; return snapshot(900, 500).current; }
        } });
    return { report: (year, period, override) => service.report(year, period, override || definition), definition, requested, liveCalls: () => liveCalls };
}

test('finished margin uses sales minus cost, including losses and genuine zero', () => {
    assert.equal(computeFinishedOrderMargin(snapshot(1200, 800).current), 400);
    assert.equal(computeFinishedOrderMargin(snapshot(400, 650).current), -250);
    assert.equal(computeFinishedOrderMargin(snapshot(0, 0).current), 0);
    assert.equal(computeFinishedOrderMargin(snapshot(0.3, 0.1).current), 0.2);
    for (const sales of [undefined, null, NaN, Infinity]) assert.equal(computeFinishedOrderMargin(snapshot(sales, 100).current), null);
    assert.equal(computeFinishedOrderMargin(null), null);
});

test('report requests the same missing-sales enrichment as the Lagerliste screen', async () => {
    const f = fixture({ '2026-08': snapshot(undefined, 800), '2025-08': snapshot(undefined, 600) },
        { '2026-08': snapshot(1200, 800), '2025-08': snapshot(500, 600) });
    const report = await f.report(2026, 2);
    assert.deepEqual(report.periodOnly[1].amounts, [400, -100, null, null]);
    assert.deepEqual(report.periodOnly[2].amounts, [1400, 700, null, null]);
    assert.equal(f.liveCalls(), 0);
});

test('custom Lagerliste and formula rows flow through the report without duplicate fixed VIA rows', async () => {
    const f=fixture({'2026-08':snapshot(1200,800),'2025-08':snapshot(500,600)});
    const definition=structuredClone(f.definition);
    definition.legacyPeriodRows=false;
    definition.pnl.push(
        {id:'wipSales',name:'SO salgpris',type:'lager',metric:'sales',period:'selected',visibility:'hidden'},
        {id:'wipCost',name:'SO kostpris',type:'lager',metric:'cost',period:'selected',visibility:'hidden'},
        {id:'margin',name:'Margin',type:'formula',formula:'[wipSales] - [wipCost]',visibility:'period'},
        {id:'total',name:'Total',type:'formula',formula:'[sales] + [margin]',visibility:'period',bold:true}
    );
    definition.balance.push({id:'inventory',name:'Inventory',type:'manual',period:'selected',monthValues:{'2026-08':50,'2025-08':40}});
    const report=await f.report(2026,2,definition);
    assert.deepEqual(report.periodOnly,[]);
    assert.deepEqual(report.rows.find(r=>r.id==='total').amounts,[1400,700,2400,1500]);
    assert.equal(report.rows.find(r=>r.id==='wipSales').visibility,'hidden');
    assert.deepEqual(report.assets.rows[1].amounts,[50,40]);
    assert.equal(f.requested.length,2);
});

test('report uses only the selected month in each year, and adds the margin to the result', async () => {
    const f = fixture({ '2026-08': snapshot(1200, 800), '2025-08': snapshot(500, 600), '2026-07': snapshot(99999, 0) });
    const report = await f.report(2026, 2);
    assert.deepEqual(f.requested.sort(), ['2025-08', '2026-08']);
    assert.deepEqual(report.periodOnly[1].amounts, [400, -100, null, null]);
    assert.deepEqual(report.periodOnly[2].amounts, [1400, 700, null, null]);
    assert.deepEqual(report.periodOnly[1].percentages, [40, -12.5, null, null]);
    assert.equal(f.liveCalls(), 0);
});

test('old or missing closed snapshots stay unavailable, without replacing history with live data', async () => {
    const f = fixture({ '2020-08': snapshot(undefined, 100) });
    const report = await f.report(2020, 2);
    assert.deepEqual(report.periodOnly[1].amounts, [null, null, null, null]);
    assert.deepEqual(report.periodOnly[2].amounts, [null, null, null, null]);
    assert.equal(f.liveCalls(), 0);
});

test('fiscal January selects January of the next calendar year', async () => {
    const f = fixture({ '2021-01': snapshot(30, 10), '2020-01': snapshot(20, 10) });
    assert.deepEqual((await f.report(2020, 7)).periodOnly[1].amounts, [20, 10, null, null]);
    assert.deepEqual(f.requested.sort(), ['2020-01', '2021-01']);
});

test('current open month can use live values, including same-month prior-year comparison', async () => {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/Copenhagen', year: 'numeric', month: 'numeric' }).formatToParts(new Date()).map(p => [p.type, p.value]));
    const year = Number(parts.year), month = Number(parts.month);
    const f = fixture({ [(year - 1) + '-' + String(month).padStart(2, '0')]: snapshot(500, 400) });
    const report = await f.report(year - (month < 7 ? 1 : 0), ((month + 5) % 12) + 1);
    assert.deepEqual(report.periodOnly[1].amounts, [400, 100, null, null]);
    assert.equal(f.liveCalls(), 1);
});

test('shared-order cost allocation preserves the whole-order sales value requested by the user', () => {
    const result = allocateSharedOrders([{ OrdNo: 1, Value: 800, SalesValue: 1200 }],
        [{ OrdNo: 1, Value: 800, MaterialCost: 800 }], [{ orderNo: 1, packedRatio: 0.25 }]);
    assert.equal(result.finished[0].Value, 200);
    assert.equal(result.finished[0].SalesValue, 1200);
});
