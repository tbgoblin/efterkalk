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

// Svuota tutta la cache condivisa (usato da "Ryd cache": vale per tutte le macchine).
async function clearAll() {
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request().query('DELETE FROM dbo.AppCache');
        logEvent('GOH-CACHE CLEARED (alle maskiner)');
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

// Rimuove le entry scadute (chiamata periodica dalla sync).
async function purgeExpired() {
    const pool = await getPool();
    if (!pool) return 0;
    try {
        const result = await pool.request()
            .input('now', sql.BigInt, Date.now())
            .query('DELETE FROM dbo.AppCache WHERE CachedAtMs + TtlMs < @now');
        return (result.rowsAffected && result.rowsAffected[0]) || 0;
    } catch (err) {
        markUnavailable(err);
        return 0;
    }
}

async function saveEfterkalkMonth({ periodStart, periodEnd, rows, updatedBy, calculationVersion }) {
    const pool = await getPool();
    if (!pool) return { ok: false, unavailable: true };
    const safeRows = Array.isArray(rows) ? rows : [];
    const transaction = new sql.Transaction(pool);
    try {
        await transaction.begin(sql.ISOLATION_LEVEL.READ_COMMITTED);
        const runResult = await new sql.Request(transaction)
            .input('periodStart', sql.Date, periodStart)
            .input('periodEnd', sql.Date, periodEnd)
            .input('version', sql.NVarChar(50), String(calculationVersion || '').slice(0, 50))
            .input('updatedBy', sql.NVarChar(100), String(updatedBy || '').slice(0, 100))
            .query(`
                DECLARE @snapshotId bigint;
                SELECT @snapshotId = SnapshotId
                FROM dbo.EfterkalkSnapshotRun WITH (UPDLOCK, HOLDLOCK)
                WHERE PeriodStart = @periodStart AND IsCurrent = 1;

                IF @snapshotId IS NULL
                BEGIN
                    INSERT dbo.EfterkalkSnapshotRun
                        (PeriodStart, PeriodEnd, RevisionNo, SnapshotStatus, IsCurrent, CalculationVersion, UpdatedBy)
                    SELECT @periodStart, @periodEnd,
                           ISNULL(MAX(RevisionNo), 0) + 1, 'OPEN', 1, @version, @updatedBy
                    FROM dbo.EfterkalkSnapshotRun
                    WHERE PeriodStart = @periodStart;
                    SET @snapshotId = SCOPE_IDENTITY();
                END
                ELSE
                BEGIN
                    UPDATE dbo.EfterkalkSnapshotRun
                    SET PeriodEnd=@periodEnd, CalculationVersion=@version, UpdatedAt=SYSUTCDATETIME(), UpdatedBy=@updatedBy
                    WHERE SnapshotId=@snapshotId;
                END;

                SELECT @snapshotId AS SnapshotId,
                       (SELECT SnapshotStatus FROM dbo.EfterkalkSnapshotRun WHERE SnapshotId=@snapshotId) AS SnapshotStatus;
            `);
        const run = runResult.recordset && runResult.recordset[0];
        if (!run || !run.SnapshotId) throw new Error('Snapshot run kunne ikke oprettes');
        const snapshotId = Number(run.SnapshotId);

        await new sql.Request(transaction)
            .input('snapshotId', sql.BigInt, snapshotId)
            .query('UPDATE dbo.EfterkalkOrderSnapshot SET IsActive=0, UpdatedAt=SYSUTCDATETIME() WHERE SnapshotId=@snapshotId');

        for (const row of safeRows) {
            await new sql.Request(transaction)
                .input('snapshotId', sql.BigInt, snapshotId)
                .input('ordNo', sql.BigInt, Number(row.OrdNo))
                .input('orderDate', sql.Date, row.OrderDate || null)
                .input('invoiceNo', sql.NVarChar(50), String(row.InvoNo || '').slice(0, 50))
                .input('invoiceDate', sql.Date, row.InvoiceDate)
                .input('custNo', sql.NVarChar(50), String(row.CustNo || '').slice(0, 50))
                .input('customerName', sql.NVarChar(250), String(row.CustomerName || row.CustomerShrt || '').slice(0, 250))
                .input('seller', sql.NVarChar(100), String(row.SellerUsr || '').slice(0, 100))
                .input('revenue', sql.Decimal(19, 4), Number(row.InvoAm || 0))
                .input('cost', sql.Decimal(19, 4), row.CostComplete ? Number(row.Cost || 0) : null)
                .input('costComplete', sql.Bit, row.CostComplete ? 1 : 0)
                .query(`MERGE dbo.EfterkalkOrderSnapshot WITH (HOLDLOCK) AS target
                    USING (SELECT @snapshotId SnapshotId, @ordNo OrdNo, @invoiceNo InvoiceNo, @invoiceDate InvoiceDate) source
                    ON target.SnapshotId=source.SnapshotId AND target.OrdNo=source.OrdNo
                       AND target.InvoiceNo=source.InvoiceNo AND target.InvoiceDate=source.InvoiceDate
                    WHEN MATCHED THEN UPDATE SET OrderDate=@orderDate, CustNo=@custNo, CustomerName=@customerName,
                        Seller=@seller, Revenue=@revenue, Cost=@cost, CostComplete=@costComplete,
                        IsActive=1, CalculatedAt=SYSUTCDATETIME(), UpdatedAt=SYSUTCDATETIME()
                    WHEN NOT MATCHED THEN INSERT
                        (SnapshotId, OrdNo, OrderDate, InvoiceNo, InvoiceDate, CustNo, CustomerName, Seller,
                         Revenue, Cost, CostComplete, IsActive)
                    VALUES (@snapshotId, @ordNo, @orderDate, @invoiceNo, @invoiceDate, @custNo, @customerName,
                            @seller, @revenue, @cost, @costComplete, 1);`);
        }

        await new sql.Request(transaction)
            .input('snapshotId', sql.BigInt, snapshotId)
            .query(`UPDATE dbo.EfterkalkSnapshotRun
                    SET CompletedAt=SYSUTCDATETIME(), UpdatedAt=SYSUTCDATETIME()
                    WHERE SnapshotId=@snapshotId`);
        await transaction.commit();
        return { ok: true, snapshotId, rows: safeRows.length };
    } catch (err) {
        try { await transaction.rollback(); } catch { /* transaction not active */ }
        markUnavailable(err);
        return { ok: false, error: err.message };
    }
}

async function getEfterkalkCustomerTrend(custNo, fromDate, toDate) {
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request()
            .input('custNo', sql.NVarChar(50), String(custNo || ''))
            .input('fromDate', sql.Date, fromDate)
            .input('toDate', sql.Date, toDate)
            .query(`SELECT PeriodStart, PeriodEnd, RevisionNo, InvoiceOrderCount, Revenue, Cost,
                           ContributionMargin, MarginPct, CostComplete
                    FROM dbo.vw_EfterkalkCustomerCurrent WITH (NOLOCK)
                    WHERE CustNo=@custNo AND PeriodStart>=@fromDate AND PeriodStart<=@toDate
                    ORDER BY PeriodStart`);
        return result.recordset || [];
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function getEfterkalkMonth(periodStart) {
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request()
            .input('periodStart', sql.Date, periodStart)
            .query(`SELECT r.SnapshotId, r.SnapshotStatus, r.RevisionNo, r.CompletedAt,
                           o.OrdNo, o.OrderDate, o.InvoiceNo, o.InvoiceDate, o.CustNo,
                           o.CustomerName, o.Seller, o.Revenue, o.Cost, o.CostComplete
                    FROM dbo.EfterkalkSnapshotRun r WITH (NOLOCK)
                    LEFT JOIN dbo.EfterkalkOrderSnapshot o WITH (NOLOCK)
                      ON o.SnapshotId=r.SnapshotId AND o.IsActive=1
                    WHERE r.PeriodStart=@periodStart AND r.IsCurrent=1
                    ORDER BY o.InvoiceDate DESC, o.OrdNo DESC`);
        const records = result.recordset || [];
        if (!records.length) return { found:false, rows:[] };
        const head = records[0];
        return {
            found:true,
            snapshotId:Number(head.SnapshotId),
            status:String(head.SnapshotStatus || ''),
            revision:Number(head.RevisionNo || 1),
            completedAt:head.CompletedAt,
            rows:records.filter(row => row.OrdNo !== null && row.OrdNo !== undefined)
        };
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

module.exports = { configure, isEnabled, ping, get, getMany, getAllFresh, set, del, clearAll, purgeExpired, saveEfterkalkMonth, getEfterkalkCustomerTrend, getEfterkalkMonth, serverLabel: GOH_SERVER + '/' + GOH_DATABASE };
