const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createLagerlisteService } = require('../services/lagerlisteService');

function makeService(dataDir, states) {
    const gohData = {
        async getAppState(key) {
            return states.has(key) ? { payload: states.get(key), updatedAt: new Date() } : null;
        },
        async setAppState(key, payload) {
            states.set(key, payload);
            return true;
        },
        async saveRawImport() { return true; }
    };
    return createLagerlisteService({
        getConnection: async () => { throw new Error('not used'); },
        sql: {}, diskCache: { get() {}, set() {} }, fs,
        getSalgordreViaRows: async () => [], getOrComputeAftercalc: async () => ({}),
        getProductionSummary: async () => ({}), getRestPrices: () => ({}), dataDir, gohData
    });
}

test('superadmin migration preserves local months and refuses silent GOH overwrite', async t => {
    const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'lagerliste-goh-'));
    t.after(() => fs.rmSync(dataDir, { recursive: true, force: true }));
    const local = { month:'2026-08', current:{ totals:{ total:123 } } };
    fs.writeFileSync(path.join(dataDir, '2026-08.json'), JSON.stringify(local), 'utf8');
    const states = new Map([['lagerliste_month_2026-08', { month:'2026-08', current:{ totals:{ total:99 } } }]]);
    const service = makeService(dataDir, states);

    const safeResult = await service.migrateLocalMonthlySnapshotsToGoh(fs, { overwrite:false, migratedBy:'admin' });
    assert.deepEqual(safeResult.conflicts, ['2026-08']);
    assert.equal(states.get('lagerliste_month_2026-08').current.totals.total, 99);

    const overwriteResult = await service.migrateLocalMonthlySnapshotsToGoh(fs, { overwrite:true, migratedBy:'admin' });
    assert.equal(overwriteResult.copied, 1);
    assert.deepEqual(states.get('lagerliste_month_2026-08'), local);
    assert.equal(fs.existsSync(path.join(dataDir, '2026-08.json')), true);
    assert.equal(Array.from(states.keys()).some(key => key.startsWith('lagerliste_backup_2026-08_')), true);
});
