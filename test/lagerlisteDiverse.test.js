const test = require('node:test');
const assert = require('node:assert/strict');
const { calculateRows, defaultRows, createLagerlisteDiverseService, monthNow } = require('../services/lagerlisteDiverseService');

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
