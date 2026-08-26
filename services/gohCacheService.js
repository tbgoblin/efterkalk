// ── GOH shared cache ────────────────────────────────────────────────────────
// Cache condivisa su SQL Server (istanza GOH) come livello intermedio tra la
// diskCache locale e il DB Visma. Fail-soft: se GOH non risponde, il servizio
// si disabilita temporaneamente e l'app continua con la sola cache locale.
const sql = require('mssql/msnodesqlv8');

const GOH_SERVER = process.env.GOH_CACHE_SERVER || '192.168.17.2\\GOH';
const GOH_DATABASE = process.env.GOH_CACHE_DB || 'GOHCache';
const DISABLE_AFTER_ERROR_MS = 5 * 60 * 1000;

let poolPromise = null;
let disabledUntil = 0;
let logEvent = () => {};

function configure({ logEvent: logger } = {}) {
    if (typeof logger === 'function') logEvent = logger;
}

function isEnabled() {
    return Date.now() >= disabledUntil;
}

function markUnavailable(err) {
    disabledUntil = Date.now() + DISABLE_AFTER_ERROR_MS;
    poolPromise = null;
    logEvent('GOH-CACHE UNAVAILABLE (retry in 5 min): ' + (err && err.message ? err.message : err));
}

async function getPool() {
    if (!isEnabled()) return null;
    if (!poolPromise) {
        poolPromise = new sql.ConnectionPool({
            server: GOH_SERVER,
            database: GOH_DATABASE,
            driver: 'msnodesqlv8',
            connectionTimeout: 8000,
            requestTimeout: 30000,
            pool: { max: 4, min: 0, idleTimeoutMillis: 30000 },
            options: { trustedConnection: true, trustServerCertificate: true }
        }).connect()
            .then(pool => {
                pool.on('error', err => markUnavailable(err));
                logEvent('GOH-CACHE CONNECTED: ' + GOH_SERVER + '/' + GOH_DATABASE);
                return pool;
            })
            .catch(err => {
                markUnavailable(err);
                throw err;
            });
    }
    try {
        return await poolPromise;
    } catch {
        return null;
    }
}

async function ping() {
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request().query('SELECT TOP 1 1 AS ok FROM dbo.AppCache WITH(NOLOCK)');
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

async function get(key) {
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request()
            .input('key', sql.NVarChar(180), String(key))
            .query('SELECT Payload, CachedAtMs, TtlMs FROM dbo.AppCache WITH(NOLOCK) WHERE CacheKey = @key');
        const row = result.recordset && result.recordset[0];
        if (!row) return null;
        if (Date.now() - Number(row.CachedAtMs) > Number(row.TtlMs)) return null;
        return JSON.parse(row.Payload);
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

// Lettura bulk per il warmup: una sola query per molte chiavi.
async function getMany(keys) {
    const found = new Map();
    if (!Array.isArray(keys) || keys.length === 0) return found;
    const pool = await getPool();
    if (!pool) return found;
    try {
        const request = pool.request();
        const placeholders = keys.map((k, i) => {
            request.input('k' + i, sql.NVarChar(180), String(k));
            return '@k' + i;
        });
        const result = await request.query(
            'SELECT CacheKey, Payload, CachedAtMs, TtlMs FROM dbo.AppCache WITH(NOLOCK) WHERE CacheKey IN (' + placeholders.join(',') + ')'
        );
        const now = Date.now();
        for (const row of (result.recordset || [])) {
            if (now - Number(row.CachedAtMs) > Number(row.TtlMs)) continue;
            try {
                found.set(String(row.CacheKey), {
                    data: JSON.parse(row.Payload),
                    cachedAtMs: Number(row.CachedAtMs),
                    ttlMs: Number(row.TtlMs)
                });
            } catch { /* payload corrotto: ignora la riga */ }
        }
        return found;
    } catch (err) {
        markUnavailable(err);
        return found;
    }
}

// Scarica tutte le entry ancora fresche (hydrate di avvio / sync periodica).
async function getAllFresh() {
    const found = new Map();
    const pool = await getPool();
    if (!pool) return found;
    try {
        const result = await pool.request()
            .input('now', sql.BigInt, Date.now())
            .query('SELECT CacheKey, Payload, CachedAtMs, TtlMs FROM dbo.AppCache WITH(NOLOCK) WHERE CachedAtMs + TtlMs > @now');
        for (const row of (result.recordset || [])) {
            try {
                found.set(String(row.CacheKey), {
                    data: JSON.parse(row.Payload),
                    cachedAtMs: Number(row.CachedAtMs),
                    ttlMs: Number(row.TtlMs)
                });
            } catch { /* payload corrotto: ignora la riga */ }
        }
        return found;
    } catch (err) {
        markUnavailable(err);
        return found;
    }
}

async function set(key, data, ttlMs) {
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('key', sql.NVarChar(180), String(key))
            .input('payload', sql.NVarChar(sql.MAX), JSON.stringify(data))
            .input('cachedAtMs', sql.BigInt, Date.now())
            .input('ttlMs', sql.BigInt, Number(ttlMs) || 0)
            .query(`MERGE dbo.AppCache AS t
                USING (SELECT @key AS CacheKey) AS s ON t.CacheKey = s.CacheKey
                WHEN MATCHED THEN UPDATE SET Payload = @payload, CachedAtMs = @cachedAtMs, TtlMs = @ttlMs, UpdatedAt = SYSUTCDATETIME()
                WHEN NOT MATCHED THEN INSERT (CacheKey, Payload, CachedAtMs, TtlMs) VALUES (@key, @payload, @cachedAtMs, @ttlMs);`);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

async function del(key) {
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('key', sql.NVarChar(180), String(key))
            .query('DELETE FROM dbo.AppCache WHERE CacheKey = @key');
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

module.exports = { configure, isEnabled, ping, get, getMany, getAllFresh, set, del, serverLabel: GOH_SERVER + '/' + GOH_DATABASE };
