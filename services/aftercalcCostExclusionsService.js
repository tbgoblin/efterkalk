const fs = require('fs');
const path = require('path');
const gohData = require('./gohDataService');

const STATE_KEY_PREFIX = 'aftercalc_cost_exclusion:';

function resolveDefaultFile() {
    const explicitDir = String(process.env.GANTECH_DATA_DIR || process.env.GANTECH_NOTES_DIR || '').trim();
    if (explicitDir) return path.join(explicitDir, 'aftercalc_cost_exclusions.json');

    const localAppData = String(process.env.LOCALAPPDATA || '').trim();
    if (localAppData) return path.join(localAppData, 'Gantech Efterkalk', 'aftercalc_cost_exclusions.json');

    const portableDir = String(process.env.PORTABLE_EXECUTABLE_DIR || '').trim();
    if (portableDir) return path.join(portableDir, 'aftercalc_cost_exclusions.json');

    return path.join(__dirname, '..', 'aftercalc_cost_exclusions.json');
}

function createAftercalcCostExclusionsService(options = {}) {
    const fsImpl = options.fs || fs;
    const gohDataImpl = options.gohData || gohData;
    const stateFile = options.stateFile || resolveDefaultFile();
    let state = null;
    let hydrationPromise = null;
    let lastHydrationAttemptMs = 0;

    function normalizePositiveInteger(value, label) {
        const parsed = Number(value);
        if (!Number.isInteger(parsed) || parsed <= 0) {
            const error = new Error(label + ' ugyldig');
            error.statusCode = 400;
            throw error;
        }
        return parsed;
    }

    function loadLocal() {
        if (state !== null) return;
        try {
            const parsed = JSON.parse(fsImpl.readFileSync(stateFile, 'utf8'));
            state = parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
        } catch (_) {
            state = {};
        }
    }

    function saveLocal() {
        try {
            fsImpl.mkdirSync(path.dirname(stateFile), { recursive: true });
            fsImpl.writeFileSync(stateFile, JSON.stringify(state, null, 2) + '\n', 'utf8');
        } catch (error) {
            console.error('[aftercalcCostExclusions] local save error:', error.message);
        }
    }

    function stateKey(ordNo, lineNo) {
        return STATE_KEY_PREFIX + ordNo + ':' + lineNo;
    }

    function orderPrefix(ordNo) {
        return STATE_KEY_PREFIX + ordNo + ':';
    }

    function applyRows(rows, scopedOrdNo = null) {
        loadLocal();
        if (scopedOrdNo !== null) delete state[String(scopedOrdNo)];
        else state = {};

        for (const row of (Array.isArray(rows) ? rows : [])) {
            const match = String(row && row.key || '').match(/^aftercalc_cost_exclusion:(\d+):(\d+)$/);
            if (!match) continue;
            const ordNo = Number(match[1]);
            const lineNo = Number(match[2]);
            if (scopedOrdNo !== null && ordNo !== scopedOrdNo) continue;
            const payload = row && row.payload && typeof row.payload === 'object' ? row.payload : {};
            if (payload.excluded !== true) continue;
            if (!state[String(ordNo)]) state[String(ordNo)] = {};
            state[String(ordNo)][String(lineNo)] = {
                excluded: true,
                updatedAt: payload.updatedAt || row.updatedAt || null,
                updatedBy: String(payload.updatedBy || '')
            };
        }
        saveLocal();
    }

    function getLocalOrder(ordNo) {
        loadLocal();
        const entries = state[String(ordNo)] || {};
        return Object.keys(entries)
            .filter(lineNo => entries[lineNo] && entries[lineNo].excluded === true)
            .map(lineNo => ({
                lineNo: Number(lineNo),
                excluded: true,
                updatedAt: entries[lineNo].updatedAt || null,
                updatedBy: String(entries[lineNo].updatedBy || '')
            }))
            .filter(entry => Number.isInteger(entry.lineNo) && entry.lineNo > 0)
            .sort((a, b) => a.lineNo - b.lineNo);
    }

    async function hydrateFromDb() {
        if (hydrationPromise) return hydrationPromise;
        if (Date.now() - lastHydrationAttemptMs < 60 * 1000) return false;
        lastHydrationAttemptMs = Date.now();
        hydrationPromise = (async () => {
            const rows = await gohDataImpl.getAppStatesByPrefix(STATE_KEY_PREFIX);
            if (!Array.isArray(rows)) return false;
            applyRows(rows);
            return true;
        })();
        try {
            return await hydrationPromise;
        } finally {
            hydrationPromise = null;
        }
    }

    async function listForOrder(ordNoValue) {
        const ordNo = normalizePositiveInteger(ordNoValue, 'Ordrenummer');
        const rows = await gohDataImpl.getAppStatesByPrefix(orderPrefix(ordNo));
        if (Array.isArray(rows)) {
            applyRows(rows, ordNo);
            return { exclusions: getLocalOrder(ordNo), source: 'goh', shared: true };
        }
        return { exclusions: getLocalOrder(ordNo), source: 'local-fallback', shared: false };
    }

    async function setLine(ordNoValue, lineNoValue, excludedValue, updatedByValue) {
        const ordNo = normalizePositiveInteger(ordNoValue, 'Ordrenummer');
        const lineNo = normalizePositiveInteger(lineNoValue, 'Linjenummer');
        const excluded = excludedValue === true;
        const payload = {
            orderNo: ordNo,
            lineNo,
            excluded,
            updatedAt: new Date().toISOString(),
            updatedBy: String(updatedByValue || '').slice(0, 100)
        };

        const savedToGoh = await gohDataImpl.setAppState(stateKey(ordNo, lineNo), payload);
        if (!savedToGoh) {
            const error = new Error('GOH-databasen kunne ikke bekræfte ændringen');
            error.statusCode = 503;
            throw error;
        }

        loadLocal();
        if (excluded) {
            if (!state[String(ordNo)]) state[String(ordNo)] = {};
            state[String(ordNo)][String(lineNo)] = {
                excluded: true,
                updatedAt: payload.updatedAt,
                updatedBy: payload.updatedBy
            };
        } else if (state[String(ordNo)]) {
            delete state[String(ordNo)][String(lineNo)];
            if (Object.keys(state[String(ordNo)]).length === 0) delete state[String(ordNo)];
        }
        saveLocal();

        return { ok: true, shared: true, exclusion: payload };
    }

    function getExcludedLineKeysSync(ordNoValue) {
        const ordNo = Number(ordNoValue);
        if (!Number.isInteger(ordNo) || ordNo <= 0) return new Set();
        return new Set(getLocalOrder(ordNo).map(entry => 'line:' + entry.lineNo));
    }

    function getFingerprintSync(ordNoValue) {
        const ordNo = Number(ordNoValue);
        if (!Number.isInteger(ordNo) || ordNo <= 0) return '';
        return getLocalOrder(ordNo)
            .map(entry => String(entry.lineNo) + '@' + String(entry.updatedAt || ''))
            .join('|');
    }

    return {
        hydrateFromDb,
        ensureHydrated: hydrateFromDb,
        listForOrder,
        setLine,
        getExcludedLineKeysSync,
        getFingerprintSync,
        stateFile
    };
}

const singleton = createAftercalcCostExclusionsService();

module.exports = Object.assign(singleton, {
    createAftercalcCostExclusionsService,
    STATE_KEY_PREFIX
});
