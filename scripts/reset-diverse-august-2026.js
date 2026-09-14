// Preview by default. --apply backs up and resets ONLY August's Diverse input.
const assert = require('node:assert/strict');
const { createHash, randomUUID } = require('node:crypto');
const sql = require('mssql/msnodesqlv8');
const { defaultRows, calculateRows } = require('../services/lagerlisteDiverseService');
const key = 'lagerliste_diverse_2026-08';
const backupKey = 'lagerliste_diverse_2026-08_before_manual_reset';
const hash = value => createHash('sha256').update(value).digest('hex');
const read = async (connection, stateKey, lock = false) => (await connection.request().input('key', sql.NVarChar(100), stateKey)
    .query('SELECT Payload FROM dbo.AppState' + (lock ? ' WITH (UPDLOCK,HOLDLOCK)' : '') + ' WHERE StateKey=@key')).recordset[0]?.Payload;
async function main() {
    const pool = await new sql.ConnectionPool({ server: process.env.GOH_CACHE_SERVER || '192.168.17.2\\GOH', database: process.env.GOH_DATA_DB || 'GantechOperationHub', connectionTimeout: 8000, requestTimeout: 30000, options: { trustedConnection: true, trustServerCertificate: true } }).connect();
    try {
        const source = await read(pool, key);
        assert.ok(source, 'No August Diverse input found');
        const old = JSON.parse(source);
        assert.equal(old.month, '2026-08');
        const categories = ['PEM (44)', 'Sv. bolte (45)', 'POP nitter (46)', 'Muffer (63)', ...new Set(defaultRows().map(row => row.category))];
        const rows = calculateRows(categories.map(category => ({ category, mode: 'amount' })), { manualAutomatic: true });
        const payload = { month: '2026-08', saved: true, automaticSource: 'manual', rows, revision: randomUUID(), updatedAt: new Date().toISOString(), updatedBy: 'user-authorized August reset', resetBackupKey: backupKey };
        if (!process.argv.includes('--apply')) {
            console.log(JSON.stringify({ preview: true, key, oldRows: old.rows.length, oldTotal: old.rows.reduce((sum, row) => sum + Number(row.Value || 0), 0), sourceHash: hash(source), newRows: rows.length, backupExists: !!await read(pool, backupKey) }));
            return;
        }
        assert.equal(hash(source), process.argv.find(arg => arg.startsWith('--expected='))?.slice(11), 'Source changed or preview hash missing');
        const tx = new sql.Transaction(pool);
        await tx.begin(sql.ISOLATION_LEVEL.SERIALIZABLE);
        try {
            assert.equal(await read(tx, key, true), source, 'Concurrent modification');
            assert.equal(await read(tx, backupKey, true), undefined, 'Reset already performed');
            const closureBefore = await read(tx, 'lagerliste_month_2026-08', true);
            const septemberBefore = await read(tx, 'lagerliste_diverse_2026-09', true);
            const insert = async (stateKey, content) => tx.request().input('key', sql.NVarChar(100), stateKey).input('payload', sql.NVarChar(sql.MAX), content)
                .query('INSERT INTO dbo.AppState(StateKey,Payload,UpdatedAt) VALUES(@key,@payload,SYSUTCDATETIME())');
            await insert(backupKey, source);
            await insert('lagerliste_diverse_rev_' + payload.revision, JSON.stringify(payload));
            await tx.request().input('key', sql.NVarChar(100), key).input('payload', sql.NVarChar(sql.MAX), JSON.stringify(payload))
                .query('UPDATE dbo.AppState SET Payload=@payload,UpdatedAt=SYSUTCDATETIME() WHERE StateKey=@key');
            assert.equal(await read(tx, 'lagerliste_month_2026-08'), closureBefore);
            assert.equal(await read(tx, 'lagerliste_diverse_2026-09'), septemberBefore);
            await tx.commit();
        } catch (err) { await tx.rollback(); throw err; }
        assert.equal(await read(pool, backupKey), source);
        assert.equal(await read(pool, key), JSON.stringify(payload));
        console.log(JSON.stringify({ verified: true, key, backupKey, rows: rows.length, allBlank: rows.every(row => row.amount === null && !row.complete), closureUnchanged: true, septemberUnchanged: true }));
    } finally { await pool.close(); }
}
main().catch(err => { console.error(err.message); process.exitCode = 1; });
