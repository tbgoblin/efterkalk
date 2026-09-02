const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const {
    getLineKey,
    getLineSalesPrice,
    getLineCostContribution,
    calculateAdjustedCost
} = require('../assets/js/aftercalc-cost-exclusions');
const {
    createAftercalcCostExclusionsService,
    STATE_KEY_PREFIX
} = require('../services/aftercalcCostExclusionsService');

test('excluding a direct sales line subtracts cost but never changes sales price', () => {
    const line = { LnNo: 1, DPrice: 150000, NoFin: 2, EffectiveLineCost: 154007.81 };
    const result = calculateAdjustedCost(327648.75, [line], new Set(['line:1']));

    assert.equal(getLineSalesPrice(line), 300000);
    assert.equal(getLineCostContribution(line), 154007.81);
    assert.equal(result.excludedCost, 154007.81);
    assert.equal(result.adjustedCost, 173640.94);
    assert.equal(result.excludedLineCount, 1);
});

test('a shared production-order cost is removed once and only after every linked line is excluded', () => {
    const lines = [
        { LnNo: 2, PurcNo: 500001, LinkedOrderType: 1, ProductionOrderTotalCost: 1200 },
        { LnNo: 3, PurcNo: 500001, LinkedOrderType: 1, ProductionOrderTotalCost: 1200 }
    ];

    const partial = calculateAdjustedCost(1200, lines, new Set(['line:2']));
    assert.equal(partial.excludedCost, 0);
    assert.equal(partial.adjustedCost, 1200);
    assert.equal(partial.deferredSharedLineCount, 1);
    assert.deepEqual(partial.deferredKeys, ['line:2']);

    const complete = calculateAdjustedCost(1200, lines, new Set(['line:2', 'line:3']));
    assert.equal(complete.excludedCost, 1200);
    assert.equal(complete.adjustedCost, 0);
    assert.equal(complete.deferredSharedLineCount, 0);
});

test('purchase-linked and direct costs remain independent line contributions', () => {
    const lines = [
        { LnNo: 1, EffectiveLineCost: 250 },
        { LnNo: 2, PurcNo: 600001, LinkedOrderType: 6, EffectiveLineCost: 400 }
    ];
    const result = calculateAdjustedCost(650, lines, new Set(['line:1', 'line:2']));

    assert.equal(result.excludedCost, 650);
    assert.equal(result.adjustedCost, 0);
});

test('unknown keys do not alter cost and calculations do not mutate inputs', () => {
    const lines = [{ LnNo: 1, EffectiveLineCost: 100 }];
    const before = JSON.stringify(lines);
    const exclusions = new Set(['line:999']);
    const result = calculateAdjustedCost(100, lines, exclusions);

    assert.equal(result.adjustedCost, 100);
    assert.equal(result.excludedLineCount, 0);
    assert.equal(JSON.stringify(lines), before);
    assert.deepEqual(Array.from(exclusions), ['line:999']);
});

test('line keys use stable positive line numbers and fall back only for display', () => {
    assert.equal(getLineKey({ LnNo: 7 }, 3), 'line:7');
    assert.equal(getLineKey({}, 3), 'index:3');
});

function createFakeGoh() {
    const values = new Map();
    return {
        values,
        async setAppState(key, payload) {
            values.set(key, { payload, updatedAt: payload.updatedAt });
            return true;
        },
        async getAppStatesByPrefix(prefix) {
            return Array.from(values.entries())
                .filter(([key]) => key.startsWith(prefix))
                .map(([key, value]) => ({ key, ...value }));
        }
    };
}

test('GOH persistence is stored per order and line until explicitly changed', async t => {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'aftercalc-cost-'));
    t.after(() => fs.rmSync(tempDir, { recursive: true, force: true }));
    const fakeGoh = createFakeGoh();
    const stateFile = path.join(tempDir, 'state.json');
    const service = createAftercalcCostExclusionsService({ gohData: fakeGoh, stateFile });

    await service.setLine(398383, 1, true, 'Test User');
    assert.equal(fakeGoh.values.get(STATE_KEY_PREFIX + '398383:1').payload.excluded, true);
    assert.deepEqual(Array.from(service.getExcludedLineKeysSync(398383)), ['line:1']);
    assert.match(service.getFingerprintSync(398383), /^1@/);

    const secondInstance = createAftercalcCostExclusionsService({ gohData: fakeGoh, stateFile: path.join(tempDir, 'second.json') });
    const loaded = await secondInstance.listForOrder(398383);
    assert.equal(loaded.shared, true);
    assert.deepEqual(loaded.exclusions.map(item => item.lineNo), [1]);

    await secondInstance.setLine(398383, 1, false, 'Test User');
    const cleared = await service.listForOrder(398383);
    assert.deepEqual(cleared.exclusions, []);
    assert.equal(service.getFingerprintSync(398383), '');
});

test('a failed GOH write is rejected and does not pretend to be permanent', async t => {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'aftercalc-cost-fail-'));
    t.after(() => fs.rmSync(tempDir, { recursive: true, force: true }));
    const service = createAftercalcCostExclusionsService({
        stateFile: path.join(tempDir, 'state.json'),
        gohData: {
            async setAppState() { return false; },
            async getAppStatesByPrefix() { return null; }
        }
    });

    await assert.rejects(
        service.setLine(398383, 1, true, 'Test User'),
        error => error.statusCode === 503
    );
    assert.deepEqual(Array.from(service.getExcludedLineKeysSync(398383)), []);
});
