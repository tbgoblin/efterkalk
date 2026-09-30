const test = require('node:test');
const assert = require('node:assert/strict');
const { createBilancioService } = require('../services/bilancioService');

function fixture(recordset) {
    const calls = [];
    const queries = [];
    const pool = { request: () => ({
        inputs: {},
        input(name, _type, value) { this.inputs[name] = value; return this; },
        async query(text) { calls.push(this.inputs); queries.push(text); return { recordset }; }
    }) };
    const service = createBilancioService({ getConnection: async () => pool, sql: { Int: 'Int', NVarChar: () => 'NVarChar' } });
    return { service, calls, queries };
}

test('missing and partial invoice values are split by status and never summed together', async () => {
    const { service } = fixture([
        { Status: 'missing', OrderCount: 3, UninvoicedValue: 200000 },
        { Status: 'partial', OrderCount: 2, UninvoicedValue: 50000 }
    ]);
    const result = await service.purchaseInvoiceGapForMonth('2026-09');
    assert.equal(result.missingValue, 200000);
    assert.equal(result.missingOrderCount, 3);
    assert.equal(result.partialValue, 50000);
    assert.equal(result.partialOrderCount, 2);
    assert.equal(result.totalValue, 250000);
});

test('a status with no rows falls back to zero instead of throwing', async () => {
    const { service } = fixture([{ Status: 'missing', OrderCount: 1, UninvoicedValue: 100 }]);
    const result = await service.purchaseInvoiceGapForMonth('2026-09');
    assert.equal(result.missingValue, 100);
    assert.equal(result.partialValue, 0);
    assert.equal(result.partialOrderCount, 0);
    assert.equal(result.totalValue, 100);
});

test('an empty result set returns all-zero, not an error', async () => {
    const { service } = fixture([]);
    const result = await service.purchaseInvoiceGapForMonth('2026-09');
    assert.deepEqual(result, { missingValue: 0, missingOrderCount: 0, partialValue: 0, partialOrderCount: 0, totalValue: 0 });
});

test('the month is translated into calendar day-boundary integers, not fiscal periods', async () => {
    const { service, calls } = fixture([]);
    await service.purchaseInvoiceGapForMonth('2026-09');
    assert.deepEqual(calls[0], { from: 20260901, to: 20260930 });
    await service.purchaseInvoiceGapForMonth('2026-02'); // 2026 is not a leap year
    assert.deepEqual(calls[1], { from: 20260201, to: 20260228 });
    await service.purchaseInvoiceGapForMonth('2028-02'); // 2028 is a leap year
    assert.deepEqual(calls[2], { from: 20280201, to: 20280229 });
});

test('purchase order lines whose ProdNo starts with U (underleverandør work) are excluded', async () => {
    // Regression: ProdNo starting with 'U' (e.g. varmgalvanisering, pulverlak) is subcontractor
    // work, not material purchased into stock — it must not count toward "ikke fuldt faktureret".
    const { service, queries } = fixture([]);
    await service.purchaseInvoiceGapForMonth('2026-09');
    assert.match(queries[0], /L\.ProdNo NOT LIKE 'U%'/);
});

test('an invalid or missing month is rejected', async () => {
    const { service } = fixture([]);
    await assert.rejects(service.purchaseInvoiceGapForMonth('2026-13'));
    await assert.rejects(service.purchaseInvoiceGapForMonth(''));
    await assert.rejects(service.purchaseInvoiceGapForMonth(undefined));
});
