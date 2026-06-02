/**
 * omsaetningThresholdsService.js
 * Persistent min/max threshold per customer for Omsaetning.
 * Stored in omsaetning_thresholds.json in a stable user-writable path
 * (GANTECH_NOTES_DIR or LOCALAPPDATA\Gantech Efterkalk) with legacy migration.
 *
 * Schema: {
 *   "75101026": { "warnThreshold": 3, "goodThreshold": 5, "updatedAt": "ISO8601" }
 * }
 */
const fs = require('fs');
const path = require('path');

const DEFAULT_WARN_THRESHOLD = 3;
const DEFAULT_GOOD_THRESHOLD = 5;

function resolveLegacyThresholdsFile() {
    return path.join(require('process').env.PORTABLE_EXECUTABLE_DIR || __dirname, '..', 'omsaetning_thresholds.json');
}

function resolveThresholdsBaseDir() {
    const explicitDir = String(process.env.GANTECH_NOTES_DIR || '').trim();
    if (explicitDir) return explicitDir;

    const localAppData = String(process.env.LOCALAPPDATA || '').trim();
    if (localAppData) return path.join(localAppData, 'Gantech Efterkalk');

    const portableDir = String(process.env.PORTABLE_EXECUTABLE_DIR || '').trim();
    if (portableDir) return portableDir;

    return path.join(__dirname, '..');
}

function resolveThresholdsFile() {
    return path.join(resolveThresholdsBaseDir(), 'omsaetning_thresholds.json');
}

function ensureThresholdsDir(thresholdsFile) {
    try {
        fs.mkdirSync(path.dirname(thresholdsFile), { recursive: true });
    } catch {
        // Ignore directory create errors; save/load will handle failures.
    }
}

function migrateLegacyThresholdsIfNeeded(thresholdsFile) {
    const legacyFile = resolveLegacyThresholdsFile();
    if (path.resolve(legacyFile) === path.resolve(thresholdsFile)) return;
    if (!fs.existsSync(legacyFile) || fs.existsSync(thresholdsFile)) return;
    try {
        ensureThresholdsDir(thresholdsFile);
        fs.copyFileSync(legacyFile, thresholdsFile);
    } catch {
        // Ignore migration failures and continue with normal load behavior.
    }
}

let _thresholds = null;

function normalizeThresholds(warnThreshold, goodThreshold) {
    const warn = Math.max(0, Number(warnThreshold));
    const baseWarn = Number.isFinite(warn) ? warn : DEFAULT_WARN_THRESHOLD;

    const good = Math.max(baseWarn, Number(goodThreshold));
    const baseGood = Number.isFinite(good) ? good : Math.max(baseWarn, DEFAULT_GOOD_THRESHOLD);

    return {
        warnThreshold: Number(baseWarn.toFixed(3)),
        goodThreshold: Number(baseGood.toFixed(3))
    };
}

function normalizeCustomerNo(custNo) {
    const key = String(custNo || '').trim();
    if (!key) return '';
    if (!/^\d{1,20}$/.test(key)) return '';
    return key;
}

function _load() {
    if (_thresholds !== null) return;
    const thresholdsFile = resolveThresholdsFile();
    migrateLegacyThresholdsIfNeeded(thresholdsFile);
    try {
        if (fs.existsSync(thresholdsFile)) {
            const raw = fs.readFileSync(thresholdsFile, 'utf8');
            _thresholds = JSON.parse(raw);
        } else {
            _thresholds = {};
        }
    } catch {
        _thresholds = {};
    }
}

function _save() {
    const thresholdsFile = resolveThresholdsFile();
    try {
        ensureThresholdsDir(thresholdsFile);
        fs.writeFileSync(thresholdsFile, JSON.stringify(_thresholds, null, 2), 'utf8');
    } catch (err) {
        console.error('[omsaetning-thresholds] save error:', err.message);
    }
}

function getThreshold(custNo) {
    _load();
    const key = normalizeCustomerNo(custNo);
    if (!key) return null;

    const existing = _thresholds[key];
    if (!existing) return null;

    const normalized = normalizeThresholds(existing.warnThreshold, existing.goodThreshold);
    return {
        warnThreshold: normalized.warnThreshold,
        goodThreshold: normalized.goodThreshold,
        updatedAt: existing.updatedAt || null
    };
}

function setThreshold(custNo, { warnThreshold, goodThreshold } = {}) {
    _load();
    const key = normalizeCustomerNo(custNo);
    if (!key) return null;

    const normalized = normalizeThresholds(warnThreshold, goodThreshold);
    _thresholds[key] = {
        warnThreshold: normalized.warnThreshold,
        goodThreshold: normalized.goodThreshold,
        updatedAt: new Date().toISOString()
    };

    _save();
    return {
        warnThreshold: normalized.warnThreshold,
        goodThreshold: normalized.goodThreshold,
        updatedAt: _thresholds[key].updatedAt
    };
}

function getStorageMeta() {
    return {
        filePath: resolveThresholdsFile(),
        defaultWarnThreshold: DEFAULT_WARN_THRESHOLD,
        defaultGoodThreshold: DEFAULT_GOOD_THRESHOLD
    };
}

module.exports = {
    DEFAULT_WARN_THRESHOLD,
    DEFAULT_GOOD_THRESHOLD,
    getThreshold,
    setThreshold,
    getStorageMeta
};
