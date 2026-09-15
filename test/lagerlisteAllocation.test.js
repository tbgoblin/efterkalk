const test = require('node:test');
const assert = require('node:assert/strict');
const { allocateSharedOrders, allocateComponentStock, allocatePurchasedPartsFromFollowUpStock, validateValuation, validateClosure } = require('../services/lagerlisteAllocation');
const { createLagerlisteService } = require('../services/lagerlisteService');
const { canonicalValueSummary } = require('../assets/js/lagerliste2-engine');

const finished = [{ OrdNo: 1, Value: 1000 }];
const via = [{ OrdNo: 1, MaterialCost: 500, TimeCost: 200, StangCost: 100, PurchasedPartCost: 0, Value: 800 }];

test('identical physical quantities in both stock categories are counted once', () => {
    const result = allocateComponentStock([{ ProdNo: 'A', Quantity: 10, UnitCost: 12, FifoValue: 120 }],
        [{ ProdNo: 'A', Beholdning: 10, Value: 120 }]);
    assert.deepEqual(result.rows, []);
    assert.equal(result.audit[0].removedValue, 120);
});

test('partial availability removes only overlap; reserved physical stock is not silently lost', () => {
    const result = allocateComponentStock([{ ProdNo: 'A', Quantity: 10, UnitCost: 12, StandardPrice: 15, FifoValue: 120 }],
        [{ ProdNo: 'A', Beholdning: 6, Value: 72 }]);
    assert.equal(result.rows[0].Quantity, 4);
    assert.equal(result.rows[0].FifoValue + 72, 120);
});

test('physical follow-up stock takes precedence over available quantity for valuation overlap', () => {
    const result = allocateComponentStock([{ ProdNo: 'A', Quantity: 10, UnitCost: 1, FifoValue: 10 }],
        [{ ProdNo: 'A', Beholdning: 4, PoPhStB: 10, Value: 10 }]);
    assert.equal(result.rows.length, 0);
    assert.equal(result.audit[0].overlapQty, 10);
});

test('absent or zero available follow-up quantity does not prove physical stock has been consumed', () => {
    const source = [{ ProdNo: 'A', Quantity: 10, UnitCost: 12, FifoValue: 120 }];
    assert.equal(allocateComponentStock(source, []).rows[0].FifoValue, 120);
    assert.equal(allocateComponentStock(source, [{ ProdNo: 'A', Beholdning: 0 }]).rows[0].FifoValue, 120);
    assert.equal(source[0].Quantity, 10);
});

test('received purchased parts move only reserved physical quantity from follow-up stock to VIA', () => {
    const followUp = [{
        ProdNo: 'A', PoPhStB: 10, Beholdning: 4, ShpRsv: 6, PhCstPr: 12,
        PoPhStBValue: 120, ReservedValue: 72, RemainingReservedValue: 72,
        BalanceDifferenceValue: 0, AvailableValue: 48, Value: 120, _preciseValue: 120, Diff: 72
    }];
    const viaRows = [{ PurchasedPartDetails: [{ prodNo: 'A', stockTransferQty: 8 }] }];
    const result = allocatePurchasedPartsFromFollowUpStock(followUp, viaRows);

    assert.equal(result.rows[0].PoPhStB, 10);
    assert.equal(result.rows[0].TransferredToViaQty, 6);
    assert.equal(result.rows[0].TransferredToViaValue, 72);
    assert.equal(result.rows[0].RemainingReservedValue, 0);
    assert.equal(result.rows[0].Value, 48);
    assert.equal(result.rows[0].Diff, 0);
    assert.equal(result.audit[0].transferredValue, 72);
    assert.deepEqual(allocatePurchasedPartsFromFollowUpStock(result.rows, viaRows).rows, result.rows);
});

test('fully packed shared order keeps finished valuation and removes VIA overlap', () => {
    const result = allocateSharedOrders(finished, via, [{ orderNo: 1, packedRatio: 1 }]);
    assert.equal(result.finished[0].Value, 1000);
    assert.equal(result.via[0].Value, 0);
    assert.equal(via[0].Value, 800);
});

test('unpacked shared order stays in VIA without adding full finished cost', () => {
    const result = allocateSharedOrders(finished, via, [{ orderNo: 1, packedRatio: 0 }]);
    assert.equal(result.finished[0].Value, 0);
    assert.equal(result.via[0].Value, 800);
});

test('partial packing splits costs once, including when opened in Lagerliste2', () => {
    const states = [{ orderNo: 1, packedRatio: 0.25 }];
    const result = allocateSharedOrders(finished, via, states);
    assert.equal(result.finished[0].Value, 250);
    assert.equal(result.via[0].Value, 600);
    const payload = { categories: { finishedNotInvoiced: result.finished, salgordreVia: result.via }, totals: {} };
    assert.equal(canonicalValueSummary(payload, { evidence: { orderStates: states } }).total, 850);
    // Later packing evidence must not reallocate a previously captured snapshot.
    assert.equal(canonicalValueSummary(payload, { evidence: { orderStates: [{ orderNo: 1, packedRatio: 1 }] } }).total, 850);
});

test('partial packing scales purchased part details with their VIA total', () => {
    const purchasedVia = [{
        OrdNo: 1, MaterialCost: 0, StangCost: 0, TimeCost: 0,
        PurchasedPartCost: 100, Value: 100,
        PurchasedPartDetails: [{ prodNo: 'A', countedValue: 100 }]
    }];
    const result = allocateSharedOrders(finished, purchasedVia, [{ orderNo: 1, packedRatio: 0.25 }]);

    assert.equal(result.via[0].PurchasedPartCost, 75);
    assert.equal(result.via[0].PurchasedPartDetails[0].countedValue, 75);
    assert.equal(result.via[0].PurchasedPartDetails[0].originalCountedValue, 100);
});

test('orders in only one category retain their value; absent overlap evidence fails explicitly', () => {
    assert.equal(allocateSharedOrders(finished, [], []).finished[0].Value, 1000);
    assert.equal(allocateSharedOrders([], via, []).via[0].Value, 800);
    assert.throws(() => allocateSharedOrders(finished, via, []), /Missing packing evidence/);
});

function validPayload(now = new Date()) {
    const allocation = allocateSharedOrders(finished, via, [{ orderNo: 1, packedRatio: 0.25 }]);
    return { valuationVersion: 35, generatedAt: now.toISOString(),
        categories: { gr5Items: [], opfolgningvare: [], finishedNotInvoiced: allocation.finished, salgordreVia: allocation.via },
        totals: { finishedNotInvoiced: 250, salgordreVia: 600 } };
}

test('reprocessing partial allocations cannot reduce the same stock or cost twice', () => {
    const follow = [{ ProdNo: 'A', Beholdning: 6 }];
    const first = allocateComponentStock([{ ProdNo: 'A', Quantity: 10, UnitCost: 12, FifoValue: 120 }], follow);
    assert.deepEqual(allocateComponentStock(first.rows, follow).rows, first.rows);
    const allocated = validPayload().categories;
    const again = allocateSharedOrders(allocated.finishedNotInvoiced, allocated.salgordreVia, []);
    assert.deepEqual(again.finished, allocated.finishedNotInvoiced);
    assert.deepEqual(again.via, allocated.salgordreVia);
});

test('closing rejects legacy data, unmatched totals and raw duplicate categories', () => {
    assert.equal(validateValuation(validPayload()), true);
    const payload = validPayload();
    payload.totals.salgordreVia = 800;
    assert.throws(() => validateValuation(payload), /Total stemmer/);
    const legacy = validPayload(); legacy.valuationVersion = 34;
    assert.throws(() => validateValuation(legacy), /ældre version/);
    const stock = validPayload();
    stock.categories.gr5Items = [{ ProdNo: 'A', Quantity: 10 }];
    stock.categories.opfolgningvare = [{ ProdNo: 'A', Beholdning: 10 }];
    assert.throws(() => validateValuation(stock), /Dobbelt lagerbeholdning/);
    const orders = validPayload(); orders.categories.finishedNotInvoiced[0].AllocationApplied = false;
    assert.throws(() => validateValuation(orders), /Dobbelt ordreværdi/);
});

test('closure requires fresh data for the selected Danish calendar month', () => {
    const now = new Date('2026-08-31T22:02:00Z'); // September 1 in Denmark
    validateClosure(validPayload(now), '2026-09', now);
    assert.throws(() => validateClosure(validPayload(now), '2026-08', now), /Måneden passer/);
    assert.throws(() => validateClosure(validPayload(new Date('2026-09-01T10:00:00Z')), '2026-09', new Date('2026-09-01T11:00:00Z')), /for gammel/);
});

test('failed or conflicting GOH closure save never writes a local replacement', async () => {
    let localWrites = 0;
    const fs = { existsSync: () => false, mkdirSync: () => {}, promises: { writeFile: async () => { localWrites++; } } };
    const service = createLagerlisteService({ fs, gohData: { setAppState: async (key, value, options) => {
        assert.equal(options.createOnly, true); return false;
    } } });
    const now = new Date();
    const month = new Intl.DateTimeFormat('sv-SE', { timeZone: 'Europe/Copenhagen', year: 'numeric', month: '2-digit' }).format(now);
    await assert.rejects(service.saveMonthlySnapshot({ fs, month, currentOverride: validPayload(now) }), /Intet er overskrevet/);
    assert.equal(localWrites, 0);
});

test('successful monthly closure waits for persistence and uses exclusive local creation', async () => {
    const events = [];
    const fs = { existsSync: () => false, mkdirSync: () => {}, promises: { writeFile: async (file, data, options) => {
        assert.equal(options.flag, 'wx'); events.push('local');
    } } };
    const service = createLagerlisteService({ fs, gohData: { setAppState: async () => { events.push('shared'); return true; } } });
    const now = new Date();
    const month = new Intl.DateTimeFormat('sv-SE', { timeZone: 'Europe/Copenhagen', year: 'numeric', month: '2-digit' }).format(now);
    await service.saveMonthlySnapshot({ fs, month, currentOverride: validPayload(now) });
    assert.deepEqual(events, ['shared', 'local']);
});
