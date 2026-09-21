const test = require('node:test');
const assert = require('node:assert/strict');
const { closingSlot, runMonthlyClose } = require('../services/lagerlisteMonthlyJob');
const { startClientMonthlyScheduler } = require('../services/lagerlisteMonthlyJob');

test('open client waits until 23:59 and triggers once despite repeated ticks', async () => {
    let time = new Date('2026-09-30T21:58:59Z'), tick, runs = 0, finish;
    startClientMonthlyScheduler({ now: () => time, setTimer: fn => { tick = fn; },
        run: () => { runs++; return new Promise(resolve => { finish = resolve; }); } });
    assert.equal(runs, 0);
    time = new Date('2026-09-30T21:59:00Z');
    const pending = tick(); await tick(); assert.equal(runs, 1);
    finish({ status: 'saved' }); await pending; await tick(); assert.equal(runs, 1);
    time = new Date('2026-10-01T00:00:00Z'); await tick(); assert.equal(runs, 1);
});

test('client opened after midnight does not fabricate a previous-month closure; errors are reported once', async () => {
    let tick, errors = 0, runs = 0, time = new Date('2026-09-30T22:00:00Z');
    startClientMonthlyScheduler({ now: () => time, setTimer: fn => { tick = fn; },
        run: async () => { runs++; throw Error('offline'); }, onError: () => { errors++; } });
    await tick(); assert.equal(runs, 0);
    time = new Date('2026-10-31T22:59:00Z'); await tick(); await tick();
    assert.equal(runs, 1); assert.equal(errors, 1);
});

test('month-end trigger uses Danish time, including leap years, winter and summer', () => {
    for (const date of ['2026-09-30T21:59:00Z', '2026-01-31T22:59:59Z', '2028-02-29T22:59:00Z', '2026-03-31T21:59:00Z', '2026-10-31T22:59:00Z']) {
        assert.equal(closingSlot(new Date(date)).due, true, date);
    }
    for (const date of ['2026-09-30T21:58:59Z', '2026-09-30T22:00:00Z', '2026-09-29T21:59:00Z', '2028-02-28T22:59:00Z']) {
        assert.equal(closingSlot(new Date(date)).due, false, date);
    }
});

function fixture(overrides = {}) {
    let state = null, calls = 0, calculated = 0;
    const events = [];
    const started = new Date('2026-09-30T21:59:05Z');
    const completed = new Date('2026-09-30T22:03:00Z');
    const payload = { valuationVersion: 35, generatedAt: completed.toISOString(),
        categories: { salgordreVia: [], finishedNotInvoiced: [] },
        totals: { salgordreVia: 0, finishedNotInvoiced: 0 },
        diverseStatus: { complete: true, month: '2026-09' } };
    const dependencies = {
        now: () => calls++ ? completed : started,
        gohData: {
            getAppState: async (key, options) => { assert.equal(options.strict, true); return state; },
            setAppState: async (key, value, options) => {
                assert.equal(key, 'lagerliste_month_2026-09'); assert.equal(options.createOnly, true);
                events.push('goh'); state = { payload: value }; return true;
            }
        },
        createService: month => {
            assert.equal(month, '2026-09');
            return { getCurrent: async options => { calculated++; assert.deepEqual(options, { forceRefresh: true, forceAftercalc: true, valuationDate: started }); return payload; } };
        },
        writeBackup: async () => { events.push('backup'); },
        ...overrides
    };
    return { dependencies, payload, events, state: () => state, calculated: () => calculated, setState: value => { state = value; } };
}

test('completion after midnight keeps September, pins Diverse and verifies GOH before backup', async () => {
    const f = fixture();
    assert.equal((await runMonthlyClose(f.dependencies)).status, 'saved');
    assert.equal(f.state().payload.month, '2026-09');
    assert.equal(f.state().payload.automaticClose.completedAt, '2026-09-30T22:03:00.000Z');
    assert.equal(f.state().payload.automaticClose.valuationMode, 'live-during-calculation');
    assert.deepEqual(f.events, ['goh', 'backup']);
});

test('existing closure is never recalculated or overwritten', async () => {
    const f = fixture(); f.setState({ payload: { month: '2026-09' } });
    assert.equal((await runMonthlyClose(f.dependencies)).status, 'already-closed');
    assert.equal(f.calculated(), 0); assert.deepEqual(f.events, []);
});

test('non month-end does not contact databases', async () => {
    const f = fixture({ now: () => new Date('2026-09-29T21:59:00Z') });
    assert.equal((await runMonthlyClose(f.dependencies)).status, 'not-due');
    assert.equal(f.calculated(), 0);
});

test('early month-end execution is rejected', async () => {
    const f = fixture({ now: () => new Date('2026-09-30T06:00:00Z') });
    await assert.rejects(runMonthlyClose(f.dependencies), /missed 23:59/);
    assert.equal(f.calculated(), 0);
});

test('incomplete or next-month Diverse never gets saved', async () => {
    for (const status of [{ complete: false, month: '2026-09' }, { complete: true, month: '2026-10' }]) {
        const f = fixture(); f.payload.diverseStatus = status;
        await assert.rejects(runMonthlyClose(f.dependencies), /Diverse/); assert.deepEqual(f.events, []);
    }
});

test('stale calculations and unavailable GOH fail without local success', async () => {
    const f = fixture(); f.payload.generatedAt = '2026-09-30T20:00:00Z';
    await assert.rejects(runMonthlyClose(f.dependencies), /fresh/); assert.deepEqual(f.events, []);
    const g = fixture(); g.dependencies.gohData.getAppState = async () => { throw Error('offline'); };
    await assert.rejects(runMonthlyClose(g.dependencies), /offline/); assert.equal(g.calculated(), 0);
});

test('rejected GOH write cannot produce a local snapshot', async () => {
    const f = fixture(); f.dependencies.gohData.setAppState = async () => false;
    await assert.rejects(runMonthlyClose(f.dependencies), /did not confirm/); assert.deepEqual(f.events, []);
});

test('concurrent winner is preserved', async () => {
    const f = fixture(); f.dependencies.gohData.setAppState = async () => { f.setState({ payload: { winner: true } }); return false; };
    assert.equal((await runMonthlyClose(f.dependencies)).status, 'already-closed');
    assert.deepEqual(f.state(), { payload: { winner: true } }); assert.deepEqual(f.events, []);
});

test('failed verification or backup does not report success', async () => {
    const f = fixture(); f.dependencies.gohData.setAppState = async () => true;
    await assert.rejects(runMonthlyClose(f.dependencies), /verification/);
    const g = fixture(); g.dependencies.writeBackup = async () => { throw Error('disk full'); };
    await assert.rejects(runMonthlyClose(g.dependencies), /disk full/);
    assert.ok(g.state()); // GOH remains authoritative even when the optional local copy fails.
});
