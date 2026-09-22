const test = require('node:test');
const assert = require('node:assert/strict');
const { createLagerlisteService } = require('../services/lagerlisteService');

function fixture(records) {
    let calls = 0;
    const service = createLagerlisteService({
        sql: { NVarChar: () => 'nvarchar' },
        getConnection: async () => ({ request: () => ({
            input(name, type, value) { assert.equal(name, 'orderNos'); assert.equal(value, '41880,42024'); return this; },
            async query() { calls++; return { recordset: records }; }
        }) })
    });
    return { service, calls: () => calls };
}
const snapshot = () => ({ current: { totals: { finishedNotInvoiced: 300, total: 900 }, categories: {
    finishedNotInvoiced: [{ OrdNo: 41880, Value: 100 }, { OrdNo: 42024, Value: 200 }, { OrdNo: 411421, Value: 0, SalesValue: 50 }]
} } });

test('missing sales are filled by order number without altering saved costs or the source snapshot', async () => {
    const f = fixture([{ OrdNo: 42024, SalesValue: 450 }, { OrdNo: 41880, SalesValue: 250 }]);
    const original = snapshot(), before = JSON.stringify(original);
    const result = await f.service.withCurrentFinishedSales(original);
    assert.equal(JSON.stringify(original), before);
    assert.equal(result.current.totals.finishedNotInvoicedSales, 750);
    assert.equal(result.current.totals.finishedNotInvoiced, 300);
    assert.equal(result.current.totals.total, 900);
    assert.deepEqual(result.current.categories.finishedNotInvoiced.map(r => r.SalesValue), [250, 450, 50]);
    assert.equal(result.current.categories.finishedNotInvoiced[0].SalesValueSource, 'current-order');
    await f.service.withCurrentFinishedSales(result);
    assert.equal(f.calls(), 1);
});

test('unavailable orders do not become zero sales or a misleading partial total', async () => {
    const f = fixture([{ OrdNo: 41880, SalesValue: 0 }]);
    const result = await f.service.withCurrentFinishedSales(snapshot());
    assert.equal(result.current.categories.finishedNotInvoiced[0].SalesValue, 0);
    assert.equal(result.current.categories.finishedNotInvoiced[1].SalesValue, undefined);
    assert.equal(result.current.totals.finishedNotInvoicedSales, null);
});

test('duplicate sales records are rejected', async () => {
    const f = fixture([{ OrdNo: 41880, SalesValue: 20 }, { OrdNo: 41880, SalesValue: 30 }]);
    await assert.rejects(f.service.withCurrentFinishedSales(snapshot()), /Dubleret/);
});
