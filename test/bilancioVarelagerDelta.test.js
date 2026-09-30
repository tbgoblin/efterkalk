const test = require('node:test');
const assert = require('node:assert/strict');
const { createBilancioService } = require('../services/bilancioService');

function todayMonthKeyCopenhagen() {
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
        timeZone: 'Europe/Copenhagen', year: 'numeric', month: '2-digit'
    }).formatToParts(new Date()).map(part => [part.type, part.value]));
    return `${parts.year}-${parts.month}`;
}

function fixture({ snapshots = {}, current = null, diverseOverlay = {} } = {}) {
    const calls = [];
    const lagerlisteService = {
        async loadMonthlySnapshot({ month }) {
            calls.push(['snapshot', month]);
            return snapshots[month] || null;
        },
        async getCurrent() {
            calls.push(['live']);
            return current;
        }
    };
    // Mirrors services/lagerlisteDiverseService.js's applyToSnapshot: overrides the snapshot's
    // frozen "diverse" total with the latest administrative revision for that month.
    const diverseService = {
        async applyToSnapshot(snapshot, month) {
            calls.push(['diverse-overlay', month]);
            if (!(month in diverseOverlay) || !snapshot) return snapshot;
            return { ...snapshot, current: { ...snapshot.current, totals: { ...snapshot.current.totals, diverse: diverseOverlay[month] } } };
        }
    };
    const service = createBilancioService({
        getConnection: async () => ({ request: () => ({ input() { return this; }, async query() { return { recordset: [] }; } }) }),
        sql: { Int: 'Int', NVarChar: () => 'NVarChar' },
        lagerlisteService,
        diverseService,
        fs: {}
    });
    return { service, calls };
}

function payload({ platesFifo = 0, platesStandard = null, opfolgningvare = 0, stang = 0, diverse = 0, restPlates = 0, gr5Items = [] }) {
    // totals.plates defaults to standard price (Prod.Inf) unless the client toggles to FIFO
    // (assets/js/lagerliste.js:14-24) — passing a different platesStandard than platesFifo lets a
    // test prove the server-side calc ignores totals.plates and always uses categories.plates FifoValue.
    return {
        totals: { plates: platesStandard === null ? platesFifo : platesStandard, opfolgningvare, stang, diverse, restPlates },
        categories: { plates: [{ FifoValue: platesFifo }], gr5Items }
    };
}

test('Varelager total sums plates(FIFO) + opfolgningvare + stang + gr5Items FifoValue + diverse + restPlates', async () => {
    const { service } = fixture({
        snapshots: {
            '2026-08': { current: payload({ platesFifo: 100, opfolgningvare: 200, stang: 300, diverse: 40, restPlates: 50, gr5Items: [{ FifoValue: 10 }, { FifoValue: 5 }] }) },
            '2026-07': { current: payload({}) }
        }
    });
    const result = await service.varelagerDeltaForMonth('2026-08');
    // 100+200+300+40+50+(10+5) = 705
    assert.equal(result.currentValue, 705);
    assert.equal(result.previousValue, 0);
    assert.equal(result.delta, 705);
});

test('plates default to FIFO value, never the standard-price totals.plates', async () => {
    // Regression: totals.plates is standard price (Prod.Inf) by default server-side; only the client's
    // FIFO toggle (lagerlistePlateTotals) substitutes categories.plates[].FifoValue. Without an explicit
    // plateMode, the server-side Varelager calc must default to FIFO, matching Materiale (CCstPr).
    const { service } = fixture({
        snapshots: {
            '2026-08': { current: payload({ platesFifo: 2795195.29, platesStandard: 3181854.95, opfolgningvare: 0, stang: 0, diverse: 0, restPlates: 0 }) },
            '2026-07': { current: payload({}) }
        }
    });
    const result = await service.varelagerDeltaForMonth('2026-08');
    assert.equal(result.currentValue, 2795195.29);
    assert.equal(result.plateMode, 'fifo');
});

test('plateMode="standard" lets the user opt into the standard-price totals.plates instead', async () => {
    const { service } = fixture({
        snapshots: {
            '2026-08': { current: payload({ platesFifo: 2795195.29, platesStandard: 3181854.95, opfolgningvare: 0, stang: 0, diverse: 0, restPlates: 0 }) },
            '2026-07': { current: payload({}) }
        }
    });
    const result = await service.varelagerDeltaForMonth('2026-08', 'standard');
    assert.equal(result.currentValue, 3181854.95);
    assert.equal(result.plateMode, 'standard');
});

test('a snapshot\'s frozen diverse total is replaced by the latest diverse overlay, not left stale', async () => {
    // Regression test: the raw snapshot has diverse=0 (frozen at closing time), but Diverse is an
    // administrative value that can be entered/corrected after the month was closed. The overlay
    // must win, exactly like GET /lagerliste/snapshot/:month already does via applyToSnapshot.
    const { service } = fixture({
        snapshots: {
            '2026-08': { current: payload({ platesFifo: 3181854.95, opfolgningvare: 482221.70, stang: 232052.77, restPlates: 277049.01, diverse: 0, gr5Items: [{ FifoValue: 84179 }] }) },
            '2026-07': { current: payload({}) }
        },
        diverseOverlay: { '2026-08': 1285715.85 }
    });
    const result = await service.varelagerDeltaForMonth('2026-08');
    assert.equal(result.currentValue, 5543073.28);
});

test('the previous month is computed correctly across a year boundary', async () => {
    const { service, calls } = fixture({ snapshots: {} });
    await service.varelagerDeltaForMonth('2026-01');
    const snapshotMonths = calls.filter(c => c[0] === 'snapshot').map(c => c[1]);
    assert.ok(snapshotMonths.includes('2025-12'), 'previous month of 2026-01 must be 2025-12');
});

test('delta is null when the previous month has no saved snapshot', async () => {
    const { service } = fixture({
        snapshots: { '2026-08': { current: payload({ platesFifo: 100 }) } }
    });
    const result = await service.varelagerDeltaForMonth('2026-08');
    assert.equal(result.currentValue, 100);
    assert.equal(result.previousValue, null);
    assert.equal(result.delta, null);
});

test('a month with no snapshot falls back to the live total only when it is the current calendar month', async () => {
    const today = todayMonthKeyCopenhagen();
    const { service } = fixture({
        snapshots: {},
        current: payload({ platesFifo: 999 })
    });
    const result = await service.varelagerDeltaForMonth(today);
    assert.equal(result.currentValue, 999);
});

test('a past month with no snapshot and no live fallback resolves to null, not an error', async () => {
    const { service } = fixture({ snapshots: {} });
    const result = await service.varelagerDeltaForMonth('2020-01');
    assert.equal(result.currentValue, null);
    assert.equal(result.previousValue, null);
    assert.equal(result.delta, null);
});

test('an invalid month is rejected', async () => {
    const { service } = fixture({});
    await assert.rejects(service.varelagerDeltaForMonth('2026-13'));
    await assert.rejects(service.varelagerDeltaForMonth(''));
});
