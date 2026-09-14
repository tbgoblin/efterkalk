// One-off, guarded revision of the August closure. Default mode is read-only.
// --apply atomically preserves the original, archives revision 1, and activates it.
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const sql = require('mssql/msnodesqlv8');
const { allocateComponentStock } = require('../services/lagerlisteAllocation');

const key = 'lagerliste_month_2026-08';
const originalKey = 'lagerliste_revision_2026-08_original';
const revisionKey = 'lagerliste_revision_2026-08_r1';
const round = n => Math.round(n * 100) / 100;
const hash = text => crypto.createHash('sha256').update(text).digest('hex');
const sum = (rows, field) => rows.reduce((n, r) => n + Number(r[field] || 0), 0);
function figures(snapshot) {
    const c = snapshot.current.categories, t = snapshot.current.totals;
    const components = sum(c.gr5Items, 'FifoValue');
    const warehouse = Number(t.plates) + Number(t.restPlates) + Number(t.stang) + Number(t.opfolgningvare) + components;
    const work = Number(t.finishedNotInvoiced) + ['MaterialCost', 'TimeCost', 'StangCost', 'PurchasedPartCost']
        .reduce((n, f) => n + sum(c.salgordreVia, f), 0)
        + c.nestingCutting.reduce((n, r) => n + Number(r.CountedValue ?? (r.IsEstimatedRest ? 0 : r.Value)), 0);
    return { components: round(components), warehouse: round(warehouse), workInProgress: round(work), total: round(warehouse + work) };
}
function buildRevision(raw, updatedAt) {
    const original = JSON.parse(raw);
    assert.equal(original.month, '2026-08');
    assert.equal(original.current.generatedAt, '2026-08-31T20:49:53.773Z');
    assert.equal(original.revision, undefined, 'Closure already has a revision');
    const before = figures(original);
    assert.equal(before.total, 5955743.76);
    assert.equal(before.components, 297149.49);
    const c = original.current.categories;
    const via = new Set(c.salgordreVia.map(r => Number(r.OrdNo)));
    assert.equal(c.finishedNotInvoiced.filter(r => via.has(Number(r.OrdNo))).length, 0);
    const allocation = allocateComponentStock(c.gr5Items, c.opfolgningvare);
    assert.equal(allocation.audit.length, 11);
    assert.equal(round(sum(allocation.audit, 'removedValue')), 212970.49);
    const revised = structuredClone(original);
    revised.current.categories.gr5Items = allocation.rows;
    const after = figures(revised);
    assert.equal(after.total, 5742773.27);
    assert.equal(after.components, 84179);
    assert.equal(after.workInProgress, before.workInProgress);
    for (const category of Object.keys(c)) {
        if (category !== 'gr5Items') assert.deepEqual(revised.current.categories[category], c[category]);
    }
    // totals.total in legacy snapshots excludes components and cutting; do not
    // subtract this correction from that subtotal as well as from component rows.
    assert.deepEqual(revised.current.totals, original.current.totals);
    revised.revision = { number: 1, revisedAt: new Date().toISOString(),
        reason: 'Remove component quantities already counted in Opfølgningsvarer, using only August snapshot data.',
        originalKey, revisionKey, originalUpdatedAt: updatedAt,
        originalSha256: hash(raw), before, after, adjustmentDkk: -212970.49,
        affectedProducts: allocation.audit };
    return revised;
}
async function read(pool, stateKey) {
    const result = await pool.request().input('key', sql.NVarChar(100), stateKey)
        .query('SELECT Payload, UpdatedAt FROM dbo.AppState WHERE StateKey = @key');
    return result.recordset[0];
}
async function verify(pool) {
    const active = await read(pool, key), saved = await read(pool, originalKey), archive = await read(pool, revisionKey);
    assert.ok(active && saved && archive, 'Missing closure, original or revision archive');
    const payload = JSON.parse(active.Payload);
    assert.equal(payload.revision.number, 1);
    assert.equal(hash(saved.Payload), payload.revision.originalSha256);
    assert.equal(active.Payload, archive.Payload);
    assert.equal(figures(payload).total, 5742773.27);
    assert.equal(figures(JSON.parse(saved.Payload)).total, 5955743.76);
    const rebuilt = buildRevision(saved.Payload, saved.UpdatedAt);
    assert.deepEqual(payload.current, rebuilt.current);
    console.log(JSON.stringify({ verified: true, key, originalKey, revisionKey, revision: payload.revision }, null, 2));
}
async function main() {
    const pool = await new sql.ConnectionPool({
        server: process.env.GOH_CACHE_SERVER || '192.168.17.2\\GOH',
        database: process.env.GOH_DATA_DB || 'GantechOperationHub', driver: 'msnodesqlv8',
        connectionTimeout: 8000, requestTimeout: 30000,
        options: { trustedConnection: true, trustServerCertificate: true }
    }).connect();
    try {
        if (process.argv.includes('--verify')) return await verify(pool);
        const source = await read(pool, key);
        assert.ok(source, 'August closure is missing');
        const revised = buildRevision(source.Payload, source.UpdatedAt);
        if (!process.argv.includes('--apply')) {
            console.log(JSON.stringify({ preview: true, ...revised.revision }, null, 2));
            return;
        }
        const transaction = new sql.Transaction(pool);
        await transaction.begin(sql.ISOLATION_LEVEL.SERIALIZABLE);
        let committed = false;
        try {
            const locked = await new sql.Request(transaction).input('key', sql.NVarChar(100), key)
                .query('SELECT Payload FROM dbo.AppState WITH (UPDLOCK, HOLDLOCK) WHERE StateKey = @key');
            assert.equal(locked.recordset[0]?.Payload, source.Payload, 'Closure changed during preparation');
            await new sql.Request(transaction)
                .input('key', sql.NVarChar(100), key)
                .input('originalKey', sql.NVarChar(100), originalKey)
                .input('revisionKey', sql.NVarChar(100), revisionKey)
                .input('original', sql.NVarChar(sql.MAX), source.Payload)
                .input('revised', sql.NVarChar(sql.MAX), JSON.stringify(revised))
                .query(`
                    IF EXISTS (SELECT 1 FROM dbo.AppState WHERE StateKey IN (@originalKey, @revisionKey))
                        THROW 50001, 'Revision archive already exists; nothing overwritten.', 1;
                    INSERT INTO dbo.AppState (StateKey, Payload) VALUES (@originalKey, @original);
                    INSERT INTO dbo.AppState (StateKey, Payload) VALUES (@revisionKey, @revised);
                    UPDATE dbo.AppState SET Payload = @revised, UpdatedAt = SYSUTCDATETIME() WHERE StateKey = @key;
                    IF @@ROWCOUNT <> 1 THROW 50002, 'Unexpected update count', 1;
                `);
            await transaction.commit();
            committed = true;
        } catch (error) {
            if (!committed) await transaction.rollback().catch(() => {});
            throw error;
        }
        await verify(pool);
    } finally { await pool.close(); }
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
