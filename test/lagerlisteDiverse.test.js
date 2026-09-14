const test = require('node:test');
const assert = require('node:assert/strict');
const { calculateRows, defaultRows, createLagerlisteDiverseService, monthNow } = require('../services/lagerlisteDiverseService');
const { createLagerlisteService } = require('../services/lagerlisteService');

test('Skrot uses pallets times kg per pallet times price; blank is not confirmed zero', () => {
    const rows = calculateRows([{ category: 'Skrot Alu', quantity: 8, kg: 1200, price: 12 }]);
    assert.equal(rows[0].Value, 115200);
    assert.equal(rows[0].complete, true);
    assert.ok(defaultRows().every(row => !row.complete));
    assert.equal(calculateRows([{ category: 'Gasser', mode: 'amount', amount: 0 }])[0].complete, true);
});

test('Diverse disallows invalid amounts, stang and direct amount plus details double counting', () => {
    assert.throws(() => calculateRows([{ category: 'Stang', mode: 'amount', amount: 2 }]), /kategori/);
    assert.throws(() => calculateRows([{ category: 'Gasser', mode: 'amount', amount: -1 }]), /tal/);
    assert.throws(() => calculateRows([{ category: 'Gasser', mode: 'amount', amount: 'oops' }]), /tal/);
    assert.throws(() => calculateRows([{ category: 'Gasser', mode: 'amount', amount: 20 }, { category: 'Gasser', mode: 'quantity', quantity: 1, price: 2 }]), /ikke begge/);
});

test('monthly save preserves older revisions and does not write closure keys', async () => {
    const states = new Map();
    const gohData = { getAppState: async key => states.has(key) ? { payload: states.get(key) } : null,
        setAppState: async (key, payload) => { states.set(key, payload); return true; } };
    const service = createLagerlisteDiverseService({ gohData });
    const august = await service.save('2026-08', defaultRows(), 'admin');
    await service.save('2026-09', defaultRows(), 'admin');
    assert.equal((await service.load('2026-08')).revision, august.revision);
    await service.save('2026-08', defaultRows(), 'admin');
    assert.equal(states.get('lagerliste_diverse_rev_' + august.revision).updatedBy, 'admin');
    assert.ok([...states.keys()].every(key => !key.startsWith('lagerliste_month_')));
    await assert.rejects(service.load('2026-99'), /måned/);
});

test('failed GOH save is an error, not a success', async () => {
    const service = createLagerlisteDiverseService({ gohData: { setAppState: async () => false } });
    await assert.rejects(service.save('2026-08', defaultRows(), 'admin'), /GOH/);
});

test('template contains the photograph detail lines without inventing monthly quantities or prices', () => {
    const rows = defaultRows();
    assert.equal(rows.filter(row => row.category === 'Paller').length, 7);
    assert.equal(rows.filter(row => row.category === 'Forbrugsmatl. Pakkeri').length, 15);
    assert.equal(rows.filter(row => row.category === 'Forbrugsmatl. Svejseafd.').length, 12);
    assert.equal(rows.filter(row => row.category === 'Gasser').length, 11);
    assert.ok(rows.every(row => !row.complete && row.Value === 0));
});

test('selected historical month gets its saved manual values without changing snapshot or querying live Visma', async () => {
    const states = new Map();
    const service = createLagerlisteDiverseService({
        gohData: { getAppState: async key => states.has(key) ? { payload: states.get(key) } : null, setAppState: async (key, value) => { states.set(key, value); return true; } },
        getConnection: async () => { throw new Error('Must not read current stock for old months'); }
    });
    const rows = defaultRows().map(row => ({ ...row, amount: 0, quantity: 0, price: 0, kg: 0 }));
    rows.find(row => row.category === 'Div. bolte').amount = 12000;
    await service.save('2026-08', rows, 'admin');
    const original = { current: { generatedAt: '2026-08-31T20:00:00Z', categories: {}, totals: { total: 100000, plates: 10 } } };
    const before = JSON.stringify(original);
    const report = await service.applyToSnapshot(original, '2026-08');
    assert.equal(report.current.totals.diverse, 12000);
    assert.equal(report.current.totals.total, 112000);
    assert.equal(report.current.diverseStatus.complete, false); // historical automatic values are absent
    assert.equal(report.current.categories.diverse.filter(row => row.mode === 'visma' && !row.complete).length, 4);
    assert.equal(JSON.stringify(original), before);
    assert.equal((await service.applyToSnapshot(report, '2026-08')).current.totals.total, 112000);
    assert.equal(await service.applyToSnapshot(original, '2026-07'), original);
});

test('historical overlay retains saved Visma values, replacing only manual Diverse', async () => {
    const service = createLagerlisteDiverseService({ gohData: { getAppState: async () => ({ payload: { saved: true, rows: [{ category: 'Div. bolte', mode: 'amount', amount: 30 }] } }) } });
    const snapshot = { current: { diverseStatus: { complete: true }, totals: { total: 120, diverse: 20 }, categories: { diverse: [
        { category: 'PEM (44)', mode: 'visma', Value: 15, complete: true },
        { category: 'Div. bolte', mode: 'amount', Value: 5, complete: true }
    ] } } };
    const result = await service.applyToSnapshot(snapshot, '2026-08');
    assert.equal(result.current.totals.diverse, 45);
    assert.equal(result.current.totals.total, 145);
});

test('unavailable GOH does not appear as an empty new month', async () => {
    const service = createLagerlisteDiverseService({ gohData: { getAppState: async () => null, isEnabled: () => false } });
    await assert.rejects(service.load('2026-08'), /GOH/);
});

test('cached reports refresh Diverse once; incomplete monthly closure is blocked', async () => {
    const base = { valuationVersion: 31, generatedAt: new Date().toISOString(),
        categories: { plates: [], gr5Items: [], opfolgningvare: [], finishedNotInvoiced: [], salgordreVia: [] },
        totals: { total: 100, diverse: 10, finishedNotInvoiced: 0, salgordreVia: 0 } };
    const service = createLagerlisteService({ diskCache: { get: () => base },
        getDiverse: async () => ({ month: monthNow(), rows: [], complete: false, total: 20 }) });
    const updated = await service.getCurrent();
    assert.equal(updated.totals.total, 110);
    assert.equal(base.totals.total, 100);
    assert.equal((await service.getCurrent()).totals.total, 110);
    await assert.rejects(service.saveMonthlySnapshot({ fs: {}, month: monthNow(), currentOverride: updated }), /Diverse/);
});

test('existing stock-category overlap refuses a duplicated Diverse valuation', async () => {
    const base = { valuationVersion: 31, categories: { plates: [{ ProdNo: '44001' }], gr5Items: [], opfolgningvare: [], finishedNotInvoiced: [], salgordreVia: [] }, totals: { finishedNotInvoiced: 0, salgordreVia: 0 } };
    const service = createLagerlisteService({ diskCache: { get: () => base },
        getDiverse: async () => ({ rows: [{ ProdNo: '44001', Value: 20 }], total: 20 }) });
    await assert.rejects(service.getCurrent(), /overlapper/);
});

test('Visma Diverse uses warehouse 1, user availability filter and physical quantity times standard price', async () => {
    let query;
    const service = createLagerlisteDiverseService({
        gohData: { getAppState: async () => null },
        getConnection: async () => ({ request: () => ({ query: async sql => { query = sql; return { recordset: [
            { ProdNo: '44001', Quantity: 10, Price: 2.5 }, { ProdNo: '45001', Quantity: 3, Price: null }
        ] }; } }) })
    });
    const result = await service.current();
    assert.match(query, /B.StcNo = 1/);
    assert.match(query, /P.Gr5 IN \(2,3\)/);
    assert.match(query, /B.PicNotR/);
    assert.match(query, /P.Inf2 AS TegnNr/);
    assert.equal(result.month, monthNow());
    assert.equal(result.total, 25);
    assert.equal(result.complete, false);
    assert.equal(result.rows.find(row => row.ProdNo === '44001').category, 'PEM (44)');
});
