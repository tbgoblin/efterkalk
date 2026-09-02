// ── GOH data hub ─────────────────────────────────────────────────────────────
// Persistenza storica su SQL Server GantechOperationHub: snapshot efterkalk
// (Orders + AfterkalkReadings) e import grezzi (RawImports). Fail-soft come
// gohCacheService: se il DB non risponde si disabilita 5 min e l'app prosegue.
const sql = require('mssql/msnodesqlv8');

const GOH_SERVER = process.env.GOH_CACHE_SERVER || '192.168.17.2\\GOH';
const GOH_DATABASE = process.env.GOH_DATA_DB || 'GantechOperationHub';
const DISABLE_AFTER_ERROR_MS = 5 * 60 * 1000;

const COMPANY_NAME = 'Gantech';
const PLANT_NAME = 'Hovedfabrik';
const MEASUREMENT_TYPES = [
    { name: 'Omsaetning', unit: 'DKK', description: 'Efterkalk: samlet omsætning for ordren' },
    { name: 'Kostpris', unit: 'DKK', description: 'Efterkalk: samlet kostpris for ordren' }
];

let poolPromise = null;
let disabledUntil = 0;
let seedPromise = null;
let logEvent = () => {};
const lastSnapshotByOrder = new Map(); // OrdNo → 'revenue|cost' per dedup in-process

function configure({ logEvent: logger } = {}) {
    if (typeof logger === 'function') logEvent = logger;
}

function isEnabled() {
    return Date.now() >= disabledUntil;
}

function markUnavailable(err) {
    disabledUntil = Date.now() + DISABLE_AFTER_ERROR_MS;
    poolPromise = null;
    seedPromise = null;
    logEvent('GOH-DATA UNAVAILABLE (retry in 5 min): ' + (err && err.message ? err.message : err));
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
                logEvent('GOH-DATA CONNECTED: ' + GOH_SERVER + '/' + GOH_DATABASE);
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

// Seed idempotente di Company/Plant/MeasurementTypes; ritorna gli id (cache in-process).
async function ensureSeed() {
    if (seedPromise) return seedPromise;
    seedPromise = (async () => {
        const pool = await getPool();
        if (!pool) throw new Error('GOH data pool unavailable');

        await pool.request()
            .input('name', sql.NVarChar(200), COMPANY_NAME)
            .query(`IF NOT EXISTS (SELECT 1 FROM dbo.Companies WHERE CompanyName = @name)
                    INSERT INTO dbo.Companies (CompanyName, CompanyCode) VALUES (@name, 'GT')`);
        const companyId = (await pool.request()
            .input('name', sql.NVarChar(200), COMPANY_NAME)
            .query('SELECT TOP 1 CompanyId FROM dbo.Companies WHERE CompanyName = @name')).recordset[0].CompanyId;

        await pool.request()
            .input('companyId', sql.Int, companyId)
            .input('name', sql.NVarChar(200), PLANT_NAME)
            .query(`IF NOT EXISTS (SELECT 1 FROM dbo.Plants WHERE PlantName = @name AND CompanyId = @companyId)
                    INSERT INTO dbo.Plants (CompanyId, PlantName) VALUES (@companyId, @name)`);
        const plantId = (await pool.request()
            .input('companyId', sql.Int, companyId)
            .input('name', sql.NVarChar(200), PLANT_NAME)
            .query('SELECT TOP 1 PlantId FROM dbo.Plants WHERE PlantName = @name AND CompanyId = @companyId')).recordset[0].PlantId;

        const typeIds = {};
        for (const mt of MEASUREMENT_TYPES) {
            await pool.request()
                .input('name', sql.NVarChar(200), mt.name)
                .input('unit', sql.NVarChar(50), mt.unit)
                .input('descr', sql.NVarChar(500), mt.description)
                .query(`IF NOT EXISTS (SELECT 1 FROM dbo.MeasurementTypes WHERE TypeName = @name)
                        INSERT INTO dbo.MeasurementTypes (TypeName, Unit, Description) VALUES (@name, @unit, @descr)`);
            typeIds[mt.name] = (await pool.request()
                .input('name', sql.NVarChar(200), mt.name)
                .query('SELECT TOP 1 MeasurementTypeId FROM dbo.MeasurementTypes WHERE TypeName = @name')).recordset[0].MeasurementTypeId;
        }

        return { companyId, plantId, typeIds };
    })().catch(err => {
        seedPromise = null;
        markUnavailable(err);
        return null;
    });
    return seedPromise;
}

// Snapshot efterkalk: upsert Orders per OrderNumber + 2 righe AfterkalkReadings
// (Omsaetning, Kostpris). Dedup: salta se identico all'ultimo snapshot salvato.
async function recordAftercalcSnapshot(ordNo, summary) {
    if (!isEnabled() || !summary) return false;
    const orderNumber = String(ordNo);
    const revenue = Math.round(Number(summary.totalRevenue || 0) * 100) / 100;
    const cost = Math.round(Number(summary.totalCost || 0) * 100) / 100;
    const fingerprint = revenue + '|' + cost;
    if (lastSnapshotByOrder.get(orderNumber) === fingerprint) return false;

    const seed = await ensureSeed();
    if (!seed) return false;
    const pool = await getPool();
    if (!pool) return false;

    try {
        await pool.request()
            .input('companyId', sql.Int, seed.companyId)
            .input('orderNumber', sql.NVarChar(100), orderNumber)
            .query(`IF NOT EXISTS (SELECT 1 FROM dbo.Orders WHERE OrderNumber = @orderNumber)
                    INSERT INTO dbo.Orders (CompanyId, OrderNumber, Status) VALUES (@companyId, @orderNumber, 'efterkalk')`);
        const orderId = (await pool.request()
            .input('orderNumber', sql.NVarChar(100), orderNumber)
            .query('SELECT TOP 1 OrderId FROM dbo.Orders WHERE OrderNumber = @orderNumber')).recordset[0].OrderId;

        // Dedup persistente: confronta con l'ultimo snapshot su DB
        const lastResult = await pool.request()
            .input('orderId', sql.Int, orderId)
            .query(`SELECT mt.TypeName, r.Value
                    FROM dbo.AfterkalkReadings r
                    JOIN dbo.MeasurementTypes mt ON mt.MeasurementTypeId = r.MeasurementTypeId
                    WHERE r.OrderId = @orderId
                      AND r.ReadingTime = (SELECT MAX(ReadingTime) FROM dbo.AfterkalkReadings WHERE OrderId = @orderId)`);
        const lastValues = {};
        for (const row of (lastResult.recordset || [])) {
            lastValues[row.TypeName] = Math.round(Number(row.Value || 0) * 100) / 100;
        }
        if (lastValues.Omsaetning === revenue && lastValues.Kostpris === cost) {
            lastSnapshotByOrder.set(orderNumber, fingerprint);
            return false;
        }

        const readingTime = new Date();
        const resultStatus = summary.hasInvoiceWarning ? 'NoInvoWarning' : 'OK';
        for (const [typeName, value] of [['Omsaetning', revenue], ['Kostpris', cost]]) {
            await pool.request()
                .input('plantId', sql.Int, seed.plantId)
                .input('orderId', sql.Int, orderId)
                .input('readingTime', sql.DateTime2, readingTime)
                .input('typeId', sql.Int, seed.typeIds[typeName])
                .input('value', sql.Decimal(18, 6), value)
                .input('status', sql.NVarChar(50), resultStatus)
                .query(`INSERT INTO dbo.AfterkalkReadings (PlantId, OrderId, ReadingTime, MeasurementTypeId, Value, ResultStatus)
                        VALUES (@plantId, @orderId, @readingTime, @typeId, @value, @status)`);
        }
        lastSnapshotByOrder.set(orderNumber, fingerprint);
        logEvent('GOH-DATA SNAPSHOT: ordNo=' + orderNumber + ' revenue=' + revenue + ' cost=' + cost);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

// Storico efterkalk di un ordine: righe { readingTime, typeName, value, resultStatus }.
async function getOrderTrend(ordNo) {
    if (!isEnabled()) return [];
    const pool = await getPool();
    if (!pool) return [];
    try {
        const result = await pool.request()
            .input('orderNumber', sql.NVarChar(100), String(ordNo))
            .query(`SELECT r.ReadingTime, mt.TypeName, r.Value, r.ResultStatus
                    FROM dbo.AfterkalkReadings r
                    JOIN dbo.Orders o ON o.OrderId = r.OrderId
                    JOIN dbo.MeasurementTypes mt ON mt.MeasurementTypeId = r.MeasurementTypeId
                    WHERE o.OrderNumber = @orderNumber
                    ORDER BY r.ReadingTime ASC`);
        return (result.recordset || []).map(row => ({
            readingTime: row.ReadingTime,
            typeName: row.TypeName,
            value: Number(row.Value || 0),
            resultStatus: row.ResultStatus || null
        }));
    } catch (err) {
        markUnavailable(err);
        return [];
    }
}

// Import grezzo (es. snapshot lagerliste) su RawImports.
async function saveRawImport(sourceName, sourceType, payload) {
    if (!isEnabled()) return false;
    const seed = await ensureSeed();
    if (!seed) return false;
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('companyId', sql.Int, seed.companyId)
            .input('plantId', sql.Int, seed.plantId)
            .input('sourceName', sql.NVarChar(200), String(sourceName))
            .input('sourceType', sql.NVarChar(100), String(sourceType))
            .input('payload', sql.NVarChar(sql.MAX), JSON.stringify(payload))
            .query(`INSERT INTO dbo.RawImports (CompanyId, PlantId, SourceName, SourceType, RawPayload)
                    VALUES (@companyId, @plantId, @sourceName, @sourceType, @payload)`);
        logEvent('GOH-DATA RAWIMPORT: ' + sourceType + '/' + sourceName);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

// Documenti di stato condivisi (users, note ordini, soglie, ...) su dbo.AppState.
async function getAppState(key) {
    if (!isEnabled()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request()
            .input('key', sql.NVarChar(100), String(key))
            .query('SELECT Payload, UpdatedAt FROM dbo.AppState WHERE StateKey = @key');
        const row = result.recordset && result.recordset[0];
        if (!row) return null;
        return { payload: JSON.parse(row.Payload), updatedAt: row.UpdatedAt };
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function getAppStatesByPrefix(prefix) {
    if (!isEnabled()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const normalizedPrefix = String(prefix || '').slice(0, 100);
        const result = await pool.request()
            .input('prefix', sql.NVarChar(100), normalizedPrefix)
            .query(`SELECT StateKey, Payload, UpdatedAt
                    FROM dbo.AppState
                    WHERE LEFT(StateKey, LEN(@prefix)) = @prefix
                    ORDER BY StateKey`);
        const rows = [];
        for (const row of (result.recordset || [])) {
            try {
                rows.push({
                    key: String(row.StateKey || ''),
                    payload: JSON.parse(row.Payload),
                    updatedAt: row.UpdatedAt
                });
            } catch (parseError) {
                if (logEvent) logEvent('GOH APPSTATE JSON ERROR (' + String(row.StateKey || '?') + '): ' + parseError.message);
            }
        }
        return rows;
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function setAppState(key, payload) {
    if (!isEnabled()) return false;
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('key', sql.NVarChar(100), String(key))
            .input('payload', sql.NVarChar(sql.MAX), JSON.stringify(payload))
            .query(`MERGE dbo.AppState AS t
                USING (SELECT @key AS StateKey) AS s ON t.StateKey = s.StateKey
                WHEN MATCHED THEN UPDATE SET Payload = @payload, UpdatedAt = SYSUTCDATETIME()
                WHEN NOT MATCHED THEN INSERT (StateKey, Payload) VALUES (@key, @payload);`);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

module.exports = { configure, isEnabled, recordAftercalcSnapshot, getOrderTrend, saveRawImport, getAppState, getAppStatesByPrefix, setAppState, serverLabel: GOH_SERVER + '/' + GOH_DATABASE };
