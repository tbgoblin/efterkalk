const sql = require('mssql/msnodesqlv8');
const settingsService = require('./services/settingsService');

function buildConfig(profile) {
    return {
        server:            profile.server,
        database:          profile.database,
        driver:            'msnodesqlv8',
        connectionTimeout: 30000,
        requestTimeout:    90000,
        pool: { max: 20, min: 0, idleTimeoutMillis: 30000 },
        options: { trustedConnection: true, trustServerCertificate: true }
    };
}

let poolPromise   = null;
let activeServer  = null;
let activeDb      = null;

async function getConnection() {
    const profile = settingsService.getActiveProfile();
    // If profile changed, discard the old pool
    if (poolPromise && (activeServer !== profile.server || activeDb !== profile.database)) {
        try {
            const old = await poolPromise;
            old.close().catch(() => {});
        } catch { /* ignore */ }
        poolPromise = null;
    }
    if (!poolPromise) {
        activeServer = profile.server;
        activeDb     = profile.database;
        poolPromise  = new sql.ConnectionPool(buildConfig(profile))
            .connect()
            .catch(err => {
                poolPromise  = null;
                activeServer = null;
                activeDb     = null;
                console.error('Errore di connessione SQL:', err);
                throw err;
            });
    }
    return poolPromise;
}

/** Call this after switching profiles to force a fresh connection. */
function resetConnection() {
    if (poolPromise) {
        poolPromise.then(p => p.close().catch(() => {})).catch(() => {});
    }
    poolPromise  = null;
    activeServer = null;
    activeDb     = null;
}

module.exports = getConnection;
module.exports.resetConnection = resetConnection;
