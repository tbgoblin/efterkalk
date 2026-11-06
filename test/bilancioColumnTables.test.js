const test = require('node:test');
const assert = require('node:assert/strict');
const { createBilancioService } = require('../services/bilancioService');
const { createColumnSourceResolver } = require('../services/bilancioSources');

const accounts = [{ AcNo: 1000, Nm: 'Salg', AcGr: '10' }, { AcNo: 2000, Nm: 'Køb', AcGr: '12' }];
const definition = {
    schema: 1, version: 3, reportId: 'default', reportName: 'Økonomirapport', balanceTitle: 'Aktiver',
    revenueRow: 'rev', beforeTaxRow: 'res',
    pnl: [
        { id: 'rev', name: 'Omsætning', type: 'accounts', accounts: [1000], groups: [], exclude: [], children: false, sign: -1 },
        { id: 'cost', name: 'Vareforbrug', type: 'accounts', accounts: [2000], groups: [], exclude: [], children: false, sign: -1 },
        { id: 'res', name: 'Resultat', type: 'sum', sources: ['rev', 'cost'] }
    ],
    balance: [{ id: 'h', name: 'Aktiver', type: 'heading' }]
};
function fixture(dataRows) {
    const calls = [];
    const pool = { request: () => ({
        inputs: {},
        input(name, _type, value) { this.inputs[name] = value; return this; },
        async query(text) {
            if (text.includes('FROM Ac ORDER BY AcNo')) return { recordsets: [accounts, []] };
            calls.push({ inputs: this.inputs, text });
            return { recordset: dataRows(this.inputs, text) };
        }
    }) };
    const service = createBilancioService({
        getConnection: async () => pool,
        sql: { Int: 'Int', MAX: 'MAX', NVarChar: () => 'NVarChar' },
        definitionStore: { load: async () => structuredClone(definition) }
    });
    return { service, calls };
}

test('monthly: 12 rolling fiscal months across the year boundary, oldest first, with sign flip and percentages', async () => {
    const { service, calls } = fixture(() => [
        { AcNo: 1000, Nm: 'Salg', AcGr: '10', AcYr: 2025, AcPr: 5, Amount: -100 },
        { AcNo: 1000, Nm: 'Salg', AcGr: '10', AcYr: 2026, AcPr: 4, Amount: -250 },
        { AcNo: 2000, Nm: 'Køb', AcGr: '12', AcYr: 2026, AcPr: 4, Amount: 50 },
        { AcNo: 2000, Nm: 'Køb', AcGr: '12', AcYr: null, AcPr: null, Amount: null }
    ]);
    const report = await service.monthly(2026, 4);
    assert.deepEqual(report.columns[0], { year: 2025, period: 5 });
    assert.deepEqual(report.columns[11], { year: 2026, period: 4 });
    assert.equal(report.columns.length, 12);
    assert.deepEqual(calls[0].inputs, { year: 2026, period: 4, accounts: '1000,2000' });
    const rev = report.rows.find(row => row.id === 'rev');
    const res = report.rows.find(row => row.id === 'res');
    assert.deepEqual(rev.amounts, [100, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 250]);
    assert.deepEqual(res.amounts, [100, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 200]);
    assert.equal(rev.percentages[0], 100);
    assert.equal(rev.percentages[1], null); // no revenue -> no percentage
    assert.equal(res.percentages[11], 80);
});

test('yearly: 12 fiscal years ending at the selected one, closed years are full and the current stops at the period', async () => {
    const { service, calls } = fixture(() => [
        { AcNo: 1000, Nm: 'Salg', AcGr: '10', AcYr: 2015, Amount: -50 },
        { AcNo: 1000, Nm: 'Salg', AcGr: '10', AcYr: 2026, Amount: -75 },
        { AcNo: 2000, Nm: 'Køb', AcGr: '12', AcYr: null, Amount: null }
    ]);
    const report = await service.yearly(2026, 4);
    assert.equal(report.columns.length, 12);
    assert.deepEqual(report.columns[0], { year: 2015, period: 12 });
    assert.deepEqual(report.columns[11], { year: 2026, period: 4 });
    assert.deepEqual(calls[0].inputs, { year: 2026, period: 4, accounts: '1000,2000' });
    const rev = report.rows.find(row => row.id === 'rev');
    assert.deepEqual(rev.amounts, [50, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 75]);
});

test('a posting outside the requested window is rejected instead of silently dropped', async () => {
    const { service } = fixture(() => [{ AcNo: 1000, Nm: 'Salg', AcGr: '10', AcYr: 2020, AcPr: 9, Amount: -1 }]);
    await assert.rejects(service.monthly(2026, 4), /Uventet regnskabsperiode/);
});

test('column source resolver maps each fiscal column to its own calendar month', async () => {
    const resolve = createColumnSourceResolver({ today: '2026-11' });
    // Fiscal 2025 period 12 = juni 2026; fiscal 2026 period 10 = april 2027.
    const rows = [{ id: 'x', type: 'manual', period: 'selected', monthValues: { '2026-06': 3, '2027-04': 7 } }];
    const values = await resolve(rows, [{ year: 2025, period: 12 }, { year: 2026, period: 10 }]);
    assert.deepEqual(values.get('x'), [3, 7]);
});
