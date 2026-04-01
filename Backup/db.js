const sql = require('mssql/msnodesqlv8');

// *** CONFIGURAZIONE ***
const SERVER = '10.2.0.3\\VISMA';     // IP + istanza nominata
const DATABASE = 'F0001';             // Nome del database

const config = {
    database: DATABASE,
    server: SERVER,
    driver: 'msnodesqlv8',
    connectionTimeout: 15000,
    requestTimeout: 45000,
    pool: {
        max: 10,
        min: 0,
        idleTimeoutMillis: 30000
    },
    options: {
        trustedConnection: true,
        trustServerCertificate: true,
    }
};

let poolPromise = null;

async function getConnection() {
    if (!poolPromise) {
        poolPromise = new sql.ConnectionPool(config)
            .connect()
            .catch(err => {
                poolPromise = null;
                console.error("Errore di connessione SQL:", err);
                throw err;
            });
    }

    return poolPromise;
}

module.exports = getConnection;