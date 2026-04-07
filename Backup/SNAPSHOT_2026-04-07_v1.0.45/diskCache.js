/**
 * diskCache.js - File-based persistent cache for Gantech Efterkalkulation.
 *
 * Cache files are stored as JSON in the /cache/ folder.
 * Each file is human-readable and contains metadata + the cached data.
 *
 * Usage:
 *   const diskCache = require('./diskCache');
 *   diskCache.set('myKey', data, 30 * 60 * 1000);  // TTL in ms
 *   const data = diskCache.get('myKey');            // null if expired or missing
 *   diskCache.del('myKey');
 *   diskCache.list();                               // array of cache file summaries
 *   diskCache.clearAll();                           // delete all cache files
 */

const fs   = require('fs');
const os   = require('os');
const path = require('path');

function resolveWritableCacheDir() {
    const candidates = [
        process.env.GANTECH_CACHE_DIR,
        // Shared location on C:\ — works across all RDS users on the same machine
        'C:\\GantechCache',
        'C:\\cache\\Gantech',
        process.env.LOCALAPPDATA ? path.join(process.env.LOCALAPPDATA, 'Gantech Efterkalk', 'cache') : null,
        process.env.APPDATA ? path.join(process.env.APPDATA, 'Gantech Efterkalk', 'cache') : null,
        path.join(process.cwd(), 'cache'),
        path.join(os.tmpdir(), 'gantech-efterkalk-cache')
    ].filter(Boolean);

    for (const dir of candidates) {
        try {
            fs.mkdirSync(dir, { recursive: true });
            return dir;
        } catch (_) {
            // Try next candidate path.
        }
    }

    return path.join(os.tmpdir(), 'gantech-efterkalk-cache');
}

const CACHE_DIR = resolveWritableCacheDir();

function ensureCacheDir() {
    try {
        if (!fs.existsSync(CACHE_DIR)) {
            fs.mkdirSync(CACHE_DIR, { recursive: true });
        }
        return true;
    } catch {
        return false;
    }
}

function sanitizeKey(key) {
    return String(key).replace(/[^a-zA-Z0-9_\-]/g, '_').slice(0, 180);
}

function cacheFilePath(key) {
    return path.join(CACHE_DIR, sanitizeKey(key) + '.json');
}

/** Returns cached data if still fresh, otherwise null. */
function get(key) {
    try {
        const filePath = cacheFilePath(key);
        if (!fs.existsSync(filePath)) return null;
        const raw  = fs.readFileSync(filePath, 'utf8');
        const entry = JSON.parse(raw);
        if (!entry || !entry.cachedAtMs || !entry.ttlMs) return null;
        const age = Date.now() - entry.cachedAtMs;
        if (age > entry.ttlMs) return null; // expired
        return entry.data;
    } catch {
        return null;
    }
}

/** Returns cached data even if expired, null only if missing/invalid. */
function getStale(key) {
    try {
        const filePath = cacheFilePath(key);
        if (!fs.existsSync(filePath)) return null;
        const raw  = fs.readFileSync(filePath, 'utf8');
        const entry = JSON.parse(raw);
        if (!entry || entry.data === undefined) return null;
        return entry.data;
    } catch {
        return null;
    }
}

/** Writes data to disk cache with the given TTL (milliseconds). */
function set(key, data, ttlMs) {
    if (!ensureCacheDir()) return;
    const filePath = cacheFilePath(key);
    const entry = {
        key,
        cachedAt:  new Date().toISOString(),
        cachedAtMs: Date.now(),
        ttlMs,
        expiresAt: new Date(Date.now() + ttlMs).toISOString(),
        data
    };
    try {
        fs.writeFileSync(filePath, JSON.stringify(entry, null, 2), 'utf8');
    } catch (err) {
        console.error('[diskCache] Write error for key "' + key + '":', err.message);
    }
}

/** Deletes a single cache entry. */
function del(key) {
    try {
        const filePath = cacheFilePath(key);
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    } catch { /* ignore */ }
}

/** Returns a summary list of all cache files (for /cache-status). */
function list() {
    if (!ensureCacheDir()) return [];
    try {
        const files = fs.readdirSync(CACHE_DIR).filter(f => f.endsWith('.json'));
        return files
            .map(f => {
                try {
                    const raw   = fs.readFileSync(path.join(CACHE_DIR, f), 'utf8');
                    const entry = JSON.parse(raw);
                    const age   = Date.now() - entry.cachedAtMs;
                    const fresh = age <= entry.ttlMs;
                    return {
                        key:         entry.key,
                        cachedAt:    entry.cachedAt,
                        expiresAt:   entry.expiresAt,
                        fresh,
                        ageMinutes:  Math.floor(age / 60000),
                        ttlMinutes:  Math.floor(entry.ttlMs / 60000),
                        file:        f
                    };
                } catch {
                    return { file: f, error: 'parse error' };
                }
            })
            .sort((a, b) => (a.key || '').localeCompare(b.key || ''));
    } catch {
        return [];
    }
}

/** Deletes all cache files. Returns count of deleted files. */
function clearAll() {
    if (!ensureCacheDir()) return 0;
    try {
        const files = fs.readdirSync(CACHE_DIR).filter(f => f.endsWith('.json'));
        let deleted = 0;
        for (const f of files) {
            try { fs.unlinkSync(path.join(CACHE_DIR, f)); deleted++; } catch { /* ignore */ }
        }
        return deleted;
    } catch {
        return 0;
    }
}

module.exports = { get, getStale, set, del, list, clearAll };
