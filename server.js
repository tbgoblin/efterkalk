const express = require('express');
const sql = require('mssql/msnodesqlv8');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const getConnection = require('./db');
const diskCache = require('./diskCache');
const { createLogger } = require('./utils/logger');
const {
    isLaserLProduct,
    isGloballyExcludedProdNo,
    isExcludedOperationProdNo,
    isEstimatedOperationMinutesFallback,
    getEffectiveOperationMinutes,
    adjustOperationLinePricing
} = require('./utils/productRules');
const {
    isHttpUrl,
    isAbsoluteWindowsPath,
    normalizeWindowsPath,
    buildImageItems,
    isSupportedImagePath,
    getLatestDrawingByProdNo
} = require('./services/drawingService');
const { createAftercalcService } = require('./services/aftercalcService');
const { createApiRouter } = require('./routes/apiRoutes');

const CACHE_TTL_AFTERCALC_MS        = 8 * 60 * 60 * 1000;  // 8 hours - match background refresh cadence
const CACHE_TTL_PRODUCTION_SUMMARY_MS = 30 * 60 * 1000;  // 30 min
const CACHE_TTL_LASER_METRICS_MS    = 60 * 60 * 1000;  // 60 min
const CACHE_TTL_ORDER_MARGIN_MS     = 30 * 60 * 1000;  // 30 min
const AFTERCALC_CACHE_KEY_PREFIX = 'aftercalc_v21_';
const ORDER_MARGIN_CACHE_KEY_PREFIX = 'order_margin_v21_';
const LEGACY_AFTERCALC_CACHE_KEY_PREFIXES = ['aftercalc_v20_', 'aftercalc_v19_', 'aftercalc_v18_', 'aftercalc_v17_', 'aftercalc_'];

const app = express();
// Parser JSON globale 256kb, ma /bom/analyze-file ha il proprio parser 40mb a livello di route
const jsonBodyParser = express.json({ limit: '256kb' });
app.use((req, res, next) => {
    if (req.path === '/bom/analyze-file') return next();
    return jsonBodyParser(req, res, next);
});
app.use('/assets', express.static(path.join(__dirname, 'assets')));

// Read version from package.json
let pkgVersion = '1.0.0';
try {
    const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
    pkgVersion = pkg.version || '1.0.0';
} catch (e) {
    console.warn('Could not read package.json version');
}
const APP_VERSION = 'Gantech Operations Hub - v' + pkgVersion;

const { logEvent } = createLogger(APP_VERSION);
const ORDER_LIST_CACHE_TTL_MS = 8 * 60 * 60 * 1000;
const ORDER_LIST_MAX_ROWS = 150;
const ORDER_LIST_DAYS_BACK = 30;
const STARTUP_MARGIN_WARM_COUNT = ORDER_LIST_MAX_ROWS;
const BACKGROUND_WARM_INTERVAL_MS = 8 * 60 * 60 * 1000;
const BACKGROUND_AFTERCALC_WARM_COUNT = ORDER_LIST_MAX_ROWS;  // Warm full aftercalc for ALL orders in the list so clicking is instant
const BACKGROUND_WARM_DELAY_MS = 10;  // Small stagger between queue submissions to keep DB stable
const MAX_DB_CALC_CONCURRENCY = 2;  // Controlled parallel DB calculations for faster startup warmup
const AFTERCALC_QUERY_TIMEOUT_MS = 4 * 60 * 1000;

const orderListCache = {
    data: [],
    loadedAt: 0,
    loading: false,
    refreshPromise: null,
    lastError: null,
    lastModifiedTime: 0
};

const orderMarginCache = new Map();
const orderMarginInFlight = new Map();
const afterCalcInFlight = new Map();
const orderRefreshInFlight = new Map();
const orderRefreshStatus = new Map();
const dbCalcQueue = [];
let activeDbCalcs = 0;
let backgroundAftercalcWarmRunning = false;

const warmupProgress = {
    running: false,
    total: 0,
    cached: 0,
    loaded: 0,
    failed: 0,
    current: null,
    startedAt: null,
    completedAt: null
};

const {
    getAfterCalc,
    fetchOrderListBase,
    getProductionSummary
} = createAftercalcService({
    getConnection,
    sql,
    diskCache,
    logEvent,
    getLatestDrawingByProdNo,
    isGloballyExcludedProdNo,
    isExcludedOperationProdNo,
    isEstimatedOperationMinutesFallback,
    getEffectiveOperationMinutes,
    adjustOperationLinePricing,
    isLaserLProduct,
    orderListMaxRows: ORDER_LIST_MAX_ROWS,
    orderListDaysBack: ORDER_LIST_DAYS_BACK,
    cacheTtlProductionSummaryMs: CACHE_TTL_PRODUCTION_SUMMARY_MS
});

// Evita cache lato browser durante lo sviluppo: forza sempre il fetch dell'ultima UI/API.
app.use((req, res, next) => {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
    next();
});


function isOrderListCacheFresh() {
    return orderListCache.loadedAt > 0 && (Date.now() - orderListCache.loadedAt) < ORDER_LIST_CACHE_TTL_MS;
}

function runWithDbCalcLimit(task, priority = 'normal') {
    return new Promise((resolve, reject) => {
        const job = { task, resolve, reject };
        if (priority === 'high') {
            dbCalcQueue.unshift(job);
        } else {
            dbCalcQueue.push(job);
        }
        pumpDbCalcQueue();
    });
}

function pumpDbCalcQueue() {
    while (activeDbCalcs < MAX_DB_CALC_CONCURRENCY && dbCalcQueue.length > 0) {
        const job = dbCalcQueue.shift();
        activeDbCalcs += 1;
        Promise.resolve()
            .then(job.task)
            .then(job.resolve)
            .catch(job.reject)
            .finally(() => {
                activeDbCalcs -= 1;
                pumpDbCalcQueue();
            });
    }
}

function withTimeout(promise, timeoutMs, timeoutMessage) {
    let timeoutId;
    const timeoutPromise = new Promise((_, reject) => {
        timeoutId = setTimeout(() => reject(new Error(timeoutMessage)), timeoutMs);
    });

    return Promise.race([promise, timeoutPromise]).finally(() => {
        clearTimeout(timeoutId);
    });
}

function getAftercalcCacheWithFallback(ordNo, promoteToCurrent = false) {
    const numericOrdNo = Number(ordNo);
    if (!Number.isFinite(numericOrdNo)) return null;

    const currentKey = AFTERCALC_CACHE_KEY_PREFIX + numericOrdNo;
    const currentCached = diskCache.get(currentKey);
    if (currentCached) return currentCached;

    for (const prefix of LEGACY_AFTERCALC_CACHE_KEY_PREFIXES) {
        const legacyKey = prefix + numericOrdNo;
        const legacyCached = diskCache.get(legacyKey);
        if (legacyCached) {
            if (promoteToCurrent) {
                diskCache.set(currentKey, legacyCached, CACHE_TTL_AFTERCALC_MS);
            }
            return legacyCached;
        }
    }

    return null;
}

async function getOrComputeAftercalc(ordNo, options = {}) {
    const priority = options.priority || 'normal';
    const key = Number(ordNo);
    if (!Number.isFinite(key)) {
        throw new Error('Ordrenummer ugyldigt');
    }

    const cacheKey = AFTERCALC_CACHE_KEY_PREFIX + key;
    const cached = getAftercalcCacheWithFallback(key, true);
    if (cached) {
        logEvent('AFTERCALC CACHE HIT: ordNo=' + key);
        return cached;
    }

    let computePromise = afterCalcInFlight.get(key);
    if (computePromise) {
        logEvent('AFTERCALC IN-FLIGHT REUSE: ordNo=' + key);
    }
    if (!computePromise) {
        logEvent('AFTERCALC FRESH COMPUTE: ordNo=' + key + ', priority=' + priority);
        computePromise = runWithDbCalcLimit(async () => {
            const data = await withTimeout(
                getAfterCalc(key),
                AFTERCALC_QUERY_TIMEOUT_MS,
                'Aftercalc timeout for ordNo=' + key
            );
            if (!data.error) {
                diskCache.set(cacheKey, data, CACHE_TTL_AFTERCALC_MS);
            }
            return data;
        }, priority).finally(() => {
            afterCalcInFlight.delete(key);
        });
        afterCalcInFlight.set(key, computePromise);
    }

    return computePromise;
}


async function getOrComputeOrderMargin(ordNo, options = {}) {
    const forceRefresh = Boolean(options.forceRefresh);
    const key = Number(ordNo);
    if (!Number.isFinite(key)) {
        throw new Error('Ordrenummer ugyldigt');
    }

    if (!forceRefresh && orderMarginCache.has(key)) {
        return orderMarginCache.get(key);
    }

    if (!forceRefresh) {
        const diskMargin = diskCache.get(ORDER_MARGIN_CACHE_KEY_PREFIX + key)
            || diskCache.getStale(ORDER_MARGIN_CACHE_KEY_PREFIX + key)
            || diskCache.getStale('order_margin_v6_' + key);
        if (diskMargin && diskMargin.totalCost !== null && diskMargin.totalCost !== undefined) {
            const marginInfo = {
                ordNo: key,
                totalRevenue: Number(diskMargin.totalRevenue || 0),
                totalCost: Number(diskMargin.totalCost || 0),
                hasInvoiceWarning: diskMargin.hasInvoiceWarning === true,
                computedAt: Date.now()
            };
            orderMarginCache.set(key, marginInfo);
            return marginInfo;
        }
    }

    if (!forceRefresh && orderMarginInFlight.has(key)) {
        return orderMarginInFlight.get(key);
    }

    const computePromise = runWithDbCalcLimit(async () => {
        const data = await withTimeout(
            getAfterCalc(key),
            AFTERCALC_QUERY_TIMEOUT_MS,
            'Aftercalc timeout for ordNo=' + key
        );
        if (data.error) {
            throw new Error(data.error);
        }

        const marginInfo = {
            ordNo: key,
            totalRevenue: Number(data.summary.totalRevenue || 0),
            totalCost: Number(data.summary.totalCost || 0),
            hasInvoiceWarning: Boolean(data.summary.hasInvoiceWarning),
            computedAt: Date.now()
        };

        orderMarginCache.set(key, marginInfo);
        // Also save to persistent disk cache (24 hours) for faster startup next time
        diskCache.set(ORDER_MARGIN_CACHE_KEY_PREFIX + key, marginInfo, 24 * 60 * 60 * 1000);
        return marginInfo;
    }).finally(() => {
        orderMarginInFlight.delete(key);
    });

    orderMarginInFlight.set(key, computePromise);
    return computePromise;
}

function warmMarginsInBackground(ordNos) {
    if (!Array.isArray(ordNos) || ordNos.length === 0) return;
    logEvent('WARM-MARGIN: queueing ' + ordNos.length + ' orders');
    for (const ordNo of ordNos) {
        const numericOrdNo = Number(ordNo);
        if (!Number.isFinite(numericOrdNo)) continue;
        getOrComputeOrderMargin(numericOrdNo).catch(() => {});
    }
}

async function warmAftercalcInBackground(ordNos, sourceLabel, maxDelayMs = BACKGROUND_WARM_DELAY_MS) {
    if (!Array.isArray(ordNos) || ordNos.length === 0) {
        logEvent('WARM-AFTERCALC (' + sourceLabel + '): skipped (empty array)');
        return;
    }
    if (backgroundAftercalcWarmRunning) {
        logEvent('WARM-AFTERCALC (' + sourceLabel + '): skipped (previous run is still active)');
        return;
    }

    backgroundAftercalcWarmRunning = true;
    const startMs = Date.now();
    let total = 0;
    let alreadyCached = 0;
    let warmed = 0;
    let failed = 0;

    // Reset global progress tracker
    warmupProgress.running = true;
    warmupProgress.total = ordNos.filter(o => Number.isFinite(Number(o))).length;
    warmupProgress.cached = 0;
    warmupProgress.loaded = 0;
    warmupProgress.failed = 0;
    warmupProgress.current = null;
    warmupProgress.startedAt = Date.now();
    warmupProgress.completedAt = null;

    const queuedComputations = [];

    try {
        for (const ordNo of ordNos) {
            const numericOrdNo = Number(ordNo);
            if (!Number.isFinite(numericOrdNo)) continue;
            total += 1;

            if (getAftercalcCacheWithFallback(numericOrdNo, false)) {
                alreadyCached += 1;
                warmupProgress.cached += 1;
                warmupProgress.current = numericOrdNo;
                continue;
            }

            warmupProgress.current = numericOrdNo;
            const computePromise = getOrComputeAftercalc(numericOrdNo, { priority: 'normal' })
                .then(() => {
                    warmed += 1;
                    warmupProgress.loaded += 1;
                })
                .catch((err) => {
                    failed += 1;
                    warmupProgress.failed += 1;
                    logEvent('WARM-AFTERCALC ERROR ordNo=' + numericOrdNo + ': ' + err.message);
                });
            queuedComputations.push(computePromise);

            if (maxDelayMs > 0) {
                await new Promise(resolve => setTimeout(resolve, maxDelayMs));
            }
        }

        if (queuedComputations.length > 0) {
            await Promise.allSettled(queuedComputations);
        }
    } finally {
        backgroundAftercalcWarmRunning = false;
        warmupProgress.running = false;
        warmupProgress.current = null;
        warmupProgress.completedAt = Date.now();
        const sec = ((Date.now() - startMs) / 1000).toFixed(1);
        logEvent('WARM-AFTERCALC (' + sourceLabel + '): total=' + total + ', cached=' + alreadyCached + ', warmed=' + warmed + ', failed=' + failed + ', time=' + sec + 's');
    }
}

function tryLoadOrderListFromCache() {
    try {
        const cached = diskCache.get('order_list');
        if (cached && Array.isArray(cached) && cached.length > 0) {
            logEvent('ORDER-LIST: loaded ' + cached.length + ' rows from diskCache');
            return cached;
        }
    } catch (e) {
        logEvent('ORDER-LIST-CACHE READ ERROR: ' + e.message);
    }
    return null;
}

function preloadMarginsAndDetailsFromCache(ordNos) {
    let marginsLoaded = 0;
    let detailsLoaded = 0;
    
    for (const ordNo of ordNos) {
        try {
            const key = Number(ordNo);
            if (!Number.isFinite(key)) continue;
            
            // Preload margins
            const marginCached = diskCache.get(ORDER_MARGIN_CACHE_KEY_PREFIX + key);
            if (marginCached) {
                orderMarginCache.set(key, marginCached);
                marginsLoaded += 1;
            }
            
            // Preload aftercalc details
            const detailsCached = getAftercalcCacheWithFallback(key, false);
            if (detailsCached) {
                detailsLoaded += 1;
            }
        } catch (e) {
            // Silently skip errors during preload
        }
    }
    
    if (marginsLoaded > 0 || detailsLoaded > 0) {
        logEvent('PRELOAD: loaded ' + marginsLoaded + ' margins + ' + detailsLoaded + ' details from diskCache');
    }
}

async function refreshOrderListCache(force = false) {
    if (!force && isOrderListCacheFresh()) {
        return;
    }

    if (orderListCache.loading) {
        if (orderListCache.refreshPromise) {
            await orderListCache.refreshPromise;
        }
        return;
    }

    orderListCache.loading = true;
    logEvent('ORDER-LIST-REFRESH: start force=' + (force ? '1' : '0'));
    orderListCache.refreshPromise = (async () => {
        try {
            const rows = await fetchOrderListBase();
            logEvent('ORDER-LIST-REFRESH: fetched ' + rows.length + ' rows from DB');
            orderListCache.data = rows;
            orderListCache.loadedAt = Date.now();
            orderListCache.lastError = null;

            // Save to persistent disk cache (TTL: 24 hours) for startup speedup
            diskCache.set('order_list', rows, 24 * 60 * 60 * 1000);
            logEvent('ORDER-LIST-REFRESH: saved ' + rows.length + ' rows to diskCache');

            const warmOrdNos = rows.slice(0, STARTUP_MARGIN_WARM_COUNT).map(r => r.OrdNo);
            warmMarginsInBackground(warmOrdNos);
            // Trigger warmup if not already running (avoid double-start during initial DB refresh)
            if (!backgroundAftercalcWarmRunning) {
                const warmAftercalcOrdNos = rows.slice(0, BACKGROUND_AFTERCALC_WARM_COUNT).map(r => r.OrdNo);
                warmAftercalcInBackground(warmAftercalcOrdNos, force ? 'refresh-force' : 'refresh-auto', BACKGROUND_WARM_DELAY_MS);
            }
        } catch (err) {
            orderListCache.lastError = err.message;
            logEvent('ORDER-LIST-REFRESH ERROR: ' + err.message);
            throw err;
        } finally {
            orderListCache.loading = false;
            orderListCache.refreshPromise = null;
            logEvent('ORDER-LIST-REFRESH: done force=' + (force ? '1' : '0'));
        }
    })();

    await orderListCache.refreshPromise;
}

app.use(createApiRouter({
    getConnection,
    sql,
    fs,
    spawn,
    diskCache,
    logEvent,
    getOrComputeAftercalc,
    getOrComputeOrderMargin,
    getProductionSummary,
    AFTERCALC_CACHE_KEY_PREFIX,
    ORDER_MARGIN_CACHE_KEY_PREFIX,
    CACHE_TTL_ORDER_MARGIN_MS,
    CACHE_TTL_LASER_METRICS_MS,
    isHttpUrl,
    normalizeWindowsPath,
    isAbsoluteWindowsPath,
    isSupportedImagePath,
    buildImageItems,
    orderListCache,
    orderMarginCache,
    orderRefreshInFlight,
    orderRefreshStatus,
    orderMarginInFlight,
    afterCalcInFlight,
    warmupProgress,
    refreshOrderListCache,
    isOrderListCacheFresh,
    ORDER_LIST_DAYS_BACK,
    pkgVersion
}));

// Endpoint per HTML
app.get('/', (req, res) => {
    logEvent('HTTP GET /');
    res.send(`
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Gantech Operations Hub</title>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&display=swap');
            :root {
                --ink-900: #0f3560;
                --ink-800: #123f6f;
                --ink-700: #1f4f7f;
                --sky-050: #f7fbff;
                --sky-100: #eef5ff;
                --sky-200: #d8e7fb;
                --line-soft: #dce8f7;
                --text-900: #13253e;
                --text-700: #456383;
                --radius-m: 12px;
                --radius-l: 16px;
                --elev-soft: 0 10px 24px rgba(15,53,96,0.10);
            }
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { font-family: 'Sora', 'Segoe UI Variable', 'Segoe UI', sans-serif; background: radial-gradient(circle at 8% 0%, #f4f8ff 0%, #edf3fb 32%, #e8eef8 100%); padding: 20px; }
            .container { max-width: 1320px; margin: 0 auto; }
            .header-banner-wrapper { background: linear-gradient(135deg, #0f3560 0%, #14577e 62%, #123f6f 100%); color: #fff; font-weight: 800; font-size: 25px; padding: 10px 12px; border-radius: 10px; margin-bottom: 16px; letter-spacing: 0.2px; width: 100%; position: sticky; top: 0; z-index: 1200; display: flex; align-items: center; justify-content: space-between; gap: 12px; box-shadow: 0 12px 28px rgba(15,53,96,0.24); }
            .header-left-controls { display:flex; align-items:center; gap:8px; flex-shrink:0; }
            .header-nav-btn { background:rgba(255,255,255,0.18); border:none; border-radius:5px; color:#fff; font-size:20px; width:38px; height:38px; cursor:pointer; display:flex; align-items:center; justify-content:center; }
            .header-nav-btn:hover { background:rgba(255,255,255,0.28); }
            .side-menu-overlay { position:fixed; inset:0; background:rgba(10,20,35,0.42); backdrop-filter:blur(2px); z-index:15140; display:none; }
            .side-menu-overlay.open { display:block; }
            .side-menu-drawer { position:fixed; top:0; left:0; height:100vh; width:min(360px, 90vw); background:linear-gradient(180deg,#fdfefe 0%,#eef5ff 100%); border-right:1px solid #d5e6fb; box-shadow:14px 0 34px rgba(10,20,35,0.24); transform:translateX(-104%); transition:transform .24s ease; z-index:15150; display:flex; flex-direction:column; }
            .side-menu-overlay.open .side-menu-drawer { transform:translateX(0); }
            .side-menu-header { padding:14px 14px 10px 14px; border-bottom:1px solid #ddeafc; display:flex; align-items:center; justify-content:space-between; gap:10px; }
            .side-menu-title { color:#0f3560; font-size:15px; font-weight:800; }
            .side-menu-close { border:none; border-radius:999px; background:#dce9fb; color:#0f3560; width:30px; height:30px; cursor:pointer; font-size:18px; line-height:1; }
            .side-menu-content { padding:12px 14px 16px 14px; overflow:auto; display:flex; flex-direction:column; gap:12px; }
            .side-menu-section { background:#fff; border:1px solid #dbe8fa; border-radius:12px; padding:10px; }
            .side-menu-section h4 { margin:0 0 8px 0; font-size:13px; color:#0f3560; }
            .side-menu-section p { margin:0; font-size:12px; color:#4d6680; line-height:1.35; }
            .side-menu-login-row { display:flex; gap:8px; margin-top:8px; }
            .side-menu-login-row input { flex:1; border:1px solid #c8d9ef; border-radius:8px; padding:8px 10px; font-size:13px; }
            .side-menu-login-row button { border:none; border-radius:8px; background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); color:#fff; font-weight:700; padding:8px 11px; cursor:pointer; }
            .side-menu-auth-status { margin-top:7px; font-size:12px; color:#355675; }
            .side-menu-auth-status.ok { color:#1b5e20; }
            .side-menu-module-list { display:flex; flex-direction:column; gap:7px; }
            .side-menu-module-list button { text-align:left; border:1px solid #cfe0f6; border-radius:10px; background:#fff; color:#0f3560; font-weight:700; padding:9px 10px; cursor:pointer; }
            .side-menu-module-list button:hover { background:#f2f7ff; }
            .side-menu-module-list button[disabled] { background:#eef3f9; color:#7a8ca0; border-color:#d7e1ec; cursor:not-allowed; }
            .side-menu-module-list button[disabled]:hover { background:#eef3f9; }
            /* Personalehåndbog modal */
            #personalehåndbogsModal { position:fixed; inset:0; z-index:15200; display:none; align-items:stretch; justify-content:center; background:rgba(7,18,35,0.72); backdrop-filter:blur(6px); padding:14px; }
            #personalehåndbogsModal.open { display:flex; }
            .personalehåndbog-shell { width:min(1400px,100%); height:100%; background:#fff; border-radius:18px; box-shadow:0 24px 72px rgba(0,0,0,0.42); display:flex; flex-direction:column; overflow:hidden; }
            .personalehåndbog-header { display:flex; align-items:center; gap:10px; padding:12px 16px; background:linear-gradient(135deg,#0f3560 0%,#123f6f 56%,#0e2f4c 100%); color:#fff; flex-shrink:0; }
            .personalehåndbog-title { font-size:17px; font-weight:800; flex:1; }
            .personalehåndbog-search-wrap { display:flex; gap:6px; align-items:center; flex:0 0 auto; }
            .personalehåndbog-search-wrap input { border:none; border-radius:8px; padding:7px 12px; font-size:13px; width:240px; outline:none; }
            .personalehåndbog-search-wrap button { border:none; border-radius:8px; padding:7px 14px; background:#5ca646; color:#fff; font-weight:700; cursor:pointer; font-size:13px; white-space:nowrap; }
            .personalehåndbog-search-wrap button:hover { filter:brightness(1.08); }
            .personalehåndbog-close { border:none; border-radius:999px; background:rgba(255,255,255,0.18); color:#fff; width:32px; height:32px; cursor:pointer; font-size:18px; line-height:1; flex-shrink:0; }
            .personalehåndbog-close:hover { background:rgba(255,255,255,0.30); }
            .personalehåndbog-body { display:flex; flex:1; overflow:hidden; }
            .personalehåndbog-results { width:360px; flex-shrink:0; display:flex; flex-direction:column; background:#f5f7fb; border-right:1.5px solid #dde3ef; }
            .ph-results-header { padding:9px 12px 8px; font-size:11px; font-weight:800; color:#5a6785; letter-spacing:.06em; text-transform:uppercase; border-bottom:1px solid #dde3ef; display:flex; align-items:center; justify-content:space-between; gap:8px; flex-shrink:0; }
            .ph-reindex-btn { border:none; border-radius:6px; background:#e2e8f5; color:#3a5080; font-size:11px; font-weight:700; padding:4px 8px; cursor:pointer; }
            .ph-reindex-btn:hover { background:#cdd7ee; }
            .ph-results-list { flex:1; overflow-y:auto; padding:8px; display:flex; flex-direction:column; gap:5px; }
            .ph-status-msg { padding:24px 14px; font-size:13px; color:#8899bb; text-align:center; line-height:1.6; }
            .ph-result-item { background:#fff; border-radius:9px; border:1.5px solid #e0e7f3; padding:10px 12px; cursor:pointer; transition:border-color .13s,box-shadow .13s; }
            .ph-result-item:hover { border-color:#2563eb; box-shadow:0 0 0 2px rgba(37,99,235,.10); }
            .ph-result-item.ph-active { border-color:#2563eb; background:#eef3ff; }
            .ph-result-title { font-size:12px; font-weight:800; color:#1a2d55; margin-bottom:3px; }
            .ph-result-url { font-size:10px; color:#7a9abf; margin-bottom:4px; word-break:break-all; }
            .ph-result-snippet { font-size:11px; color:#4a5a7a; line-height:1.55; }
            .ph-result-snippet mark { background:#fff176; color:inherit; border-radius:2px; padding:0 1px; }
            .personalehåndbog-iframe { flex:1; border:none; width:100%; background:#fff; min-width:0; }
            /* Kvalitetsledelsessystem modal */
            #qmsModal { position:fixed; inset:0; z-index:15220; display:none; align-items:stretch; justify-content:center; background:rgba(6,16,29,0.74); backdrop-filter:blur(6px); padding:14px; }
            #qmsModal.open { display:flex; }
            .qms-shell { width:min(1420px,100%); height:100%; background:#fff; border-radius:18px; box-shadow:0 24px 72px rgba(0,0,0,0.44); display:flex; flex-direction:column; overflow:hidden; }
            .qms-header { display:flex; align-items:center; gap:10px; padding:12px 16px; background:linear-gradient(135deg,#0f3560 0%,#145083 56%,#0e2f4c 100%); color:#fff; flex-shrink:0; }
            .qms-title { font-size:17px; font-weight:800; flex:1; }
            .qms-search-wrap { display:flex; gap:6px; align-items:center; }
            .qms-search-wrap input { border:none; border-radius:8px; padding:7px 12px; font-size:13px; width:290px; outline:none; }
            .qms-search-wrap button { border:none; border-radius:8px; padding:7px 14px; background:#5ca646; color:#fff; font-weight:700; cursor:pointer; font-size:13px; white-space:nowrap; }
            .qms-close { border:none; border-radius:999px; background:rgba(255,255,255,0.18); color:#fff; width:32px; height:32px; cursor:pointer; font-size:18px; line-height:1; }
            .qms-close:hover { background:rgba(255,255,255,0.30); }
            .qms-body { display:flex; flex:1; overflow:hidden; }
            .qms-nav { width:390px; flex-shrink:0; display:flex; flex-direction:column; background:#f6f8fc; border-right:1.5px solid #dde3ef; }
            .qms-nav-header { padding:10px 12px; font-size:11px; font-weight:800; color:#5a6785; letter-spacing:.06em; text-transform:uppercase; border-bottom:1px solid #dde3ef; display:flex; align-items:center; justify-content:space-between; gap:8px; }
            .qms-nav-actions { display:flex; gap:6px; }
            .qms-nav-actions button { border:none; border-radius:6px; background:#e2e8f5; color:#304a77; font-size:11px; font-weight:700; padding:4px 8px; cursor:pointer; }
            .qms-nav-actions button:hover { background:#ced9ee; }
            .qms-list { flex:1; overflow-y:auto; padding:8px; display:flex; flex-direction:column; gap:5px; }
            .qms-item { background:#fff; border-radius:9px; border:1.5px solid #e0e7f3; padding:10px 12px; cursor:pointer; transition:border-color .13s,box-shadow .13s; }
            .qms-item:hover { border-color:#2563eb; box-shadow:0 0 0 2px rgba(37,99,235,.10); }
            .qms-item.active { border-color:#2563eb; background:#eef3ff; }
            .qms-item-title { font-size:12px; font-weight:800; color:#1a2d55; margin-bottom:3px; }
            .qms-item-meta { font-size:11px; color:#6b7ea2; }
            .qms-empty { padding:22px 14px; font-size:13px; color:#8899bb; text-align:center; line-height:1.6; }
            .qms-view { flex:1; min-width:0; overflow-y:auto; background:#fff; padding:18px 20px; }
            .qms-view h3 { margin:0 0 8px 0; border:none; color:#123f6f; font-size:22px; }
            .qms-view .qms-view-meta { color:#647a9a; font-size:12px; margin-bottom:12px; }
            .qms-view .qms-view-content { white-space:pre-wrap; color:#1f334f; font-size:14px; line-height:1.6; }
            .qms-view .qms-view-link { margin-top:14px; }
            .qms-view .qms-view-link a { color:#0f5ab7; font-weight:700; text-decoration:none; }
            .qms-view .qms-view-link a:hover { text-decoration:underline; }
            .qms-editor { display:flex; flex-direction:column; gap:10px; margin-top:4px; }
            .qms-editor label { font-size:12px; font-weight:700; color:#3a4f70; }
            .qms-editor input, .qms-editor textarea { width:100%; border:1px solid #cbd8ef; border-radius:8px; padding:8px 10px; font-size:13px; color:#1f334f; }
            .qms-editor textarea { min-height:180px; resize:vertical; }
            .qms-editor-actions { display:flex; gap:8px; }
            .qms-editor-actions button { border:none; border-radius:8px; padding:8px 12px; font-weight:700; cursor:pointer; }
            .qms-editor-actions .save { background:#2e7d32; color:#fff; }
            .qms-editor-actions .delete { background:#c62828; color:#fff; }
            .qms-editor-actions .cancel { background:#e2e8f5; color:#304a77; }
            .belastning-bars { display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:12px; align-items:start; grid-auto-rows:max-content; }
            .belastning-toolbar-note { font-size:11px; color:#4d6d94; padding:2px 2px 0; }
            .belastning-order-filter { border:1px solid #c7d9f3 !important; background:linear-gradient(180deg,#ffffff 0%,#f4f9ff 100%); }
            .belastning-bar-row { border:1px solid #dce5f5; border-radius:8px; padding:8px 10px; background:#fff; cursor:pointer; }
            .belastning-bar-row:hover { border-color:#8eb2e6; background:#f7fbff; }
            .belastning-bar-top { display:flex; justify-content:space-between; gap:8px; margin-bottom:6px; font-size:12px; font-weight:700; color:#1e4768; }
            .belastning-bar-track { height:10px; border-radius:999px; background:#e8f0fb; overflow:hidden; border:1px solid #d3e0f5; }
            .belastning-bar-fill { height:100%; border-radius:999px; background:linear-gradient(90deg,#1565c0 0%,#0f3560 100%); }
            .belastning-bar-meta { margin-top:5px; font-size:11px; color:#516d8d; display:flex; gap:10px; flex-wrap:wrap; }
            .belastning-chart-card { position:relative; overflow:hidden; border:1px solid #d7e5fa; border-radius:18px; background:linear-gradient(180deg,#ffffff 0%,#f5f9ff 100%); box-shadow:0 18px 40px rgba(15,53,96,0.10); }
            .belastning-chart-card::before { content:''; position:absolute; inset:0; background:radial-gradient(circle at top right, rgba(21,101,192,0.10), transparent 38%), radial-gradient(circle at bottom left, rgba(123,224,126,0.10), transparent 34%); pointer-events:none; }
            .belastning-chart-card .omsaetning-chart-head { position:relative; padding:14px 14px 10px; border-bottom:1px solid #dbe8f7; }
            .belastning-chart-card .omsaetning-chart-title { font-size:14px; letter-spacing:0.03em; }
            .belastning-chart-card .omsaetning-chart-sub { color:#5d7897; }
            .belastning-chart-grid { display:grid; grid-template-columns:1fr; gap:12px; }
            .belastning-svg-wrap { position:relative; margin-bottom:10px; border:1px solid #d6e3f7; border-radius:14px; background:linear-gradient(180deg,#fbfdff 0%,#f2f7ff 100%); padding:8px; overflow-x:auto; overflow-y:hidden; box-shadow:inset 0 1px 0 rgba(255,255,255,0.9); }
            .belastning-svg-wrap::after { content:''; position:absolute; inset:auto 12px 12px auto; width:80px; height:80px; border-radius:50%; background:radial-gradient(circle, rgba(21,101,192,0.14), transparent 70%); pointer-events:none; }
            .belastning-svg { width:100%; min-width:0; height:clamp(230px, 28vw, 320px); display:block; }
            .belastning-svg .axis { stroke:#aec2de; stroke-width:1.1; }
            .belastning-svg .grid { stroke:#e4edf9; stroke-width:1; stroke-dasharray:3 4; }
            .belastning-svg .label { fill:#3e5d84; font-size:10px; font-weight:700; }
            .belastning-day-band { fill:transparent; transition:fill 120ms ease; }
            .belastning-day-band:hover { fill:rgba(21,101,192,0.12); }
            .belastning-day-band.active { fill:rgba(21,101,192,0.18); }
            .belastning-series-kap { fill:#6d96dc; }
            .belastning-series-resv { fill:#ff6b57; }
            .belastning-series-aften { fill:#f4b8c8; }
            .belastning-resource-chart { min-width:0; max-width:100%; overflow:hidden; border:1px solid #d7e5fa; border-radius:12px; background:linear-gradient(180deg,#ffffff 0%,#f8fbff 100%); padding:10px 10px 9px; cursor:pointer; box-shadow:0 8px 18px rgba(15,53,96,0.05); transition:transform 120ms ease, box-shadow 120ms ease, border-color 120ms ease; }
            .belastning-resource-chart:hover { border-color:#9cb9e4; box-shadow:0 12px 24px rgba(15,53,96,0.10); transform:translateY(-1px); }
            .belastning-resource-chart.is-active { border-color:#3f82d5; box-shadow:0 14px 28px rgba(21,101,192,0.20); }
            .belastning-resource-chart[draggable="true"] { cursor:grab; }
            .belastning-resource-chart.is-dragging { opacity:0.55; transform:scale(0.985); box-shadow:0 18px 36px rgba(15,53,96,0.16); }
            .belastning-resource-chart.drag-target-before { border-top:3px solid #1565c0; }
            .belastning-resource-chart.drag-target-after { border-bottom:3px solid #1565c0; }
            .belastning-resource-title { font-size:12px; font-weight:800; color:#1f3f69; margin-bottom:6px; }
            .belastning-card-top { display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:6px; }
            .belastning-resource-chart h5 { margin:0 0 6px 0; color:#1e4768; font-size:13px; font-weight:800; letter-spacing:0.01em; }
            .belastning-resource-chart .belastning-card-top h5 { margin:0; flex:1; }
            .belastning-drag-chip { display:inline-flex; align-items:center; gap:4px; border:1px solid #c7d9f3; border-radius:999px; background:linear-gradient(180deg,#ffffff 0%,#edf5ff 100%); color:#476685; font-size:10px; font-weight:800; letter-spacing:0.04em; text-transform:uppercase; padding:4px 7px; white-space:nowrap; }
            .belastning-mini-meta { margin-top:4px; font-size:11px; color:#567395; display:flex; gap:8px; flex-wrap:wrap; }
            .belastning-section-title { margin:10px 0 6px 0; font-size:12px; font-weight:800; color:#1e4768; letter-spacing:0.02em; text-transform:uppercase; }
            .belastning-subrows { margin-top:4px; font-size:11px; line-height:1.4; color:#4b6584; background:#f5f9ff; border:1px dashed #c9d9f1; border-radius:6px; padding:6px 8px; }
            .belastning-order-shell { border:1px solid #cfe0f5; border-radius:14px; overflow:auto; background:linear-gradient(180deg,#fbfdff 0%,#eef5ff 100%); box-shadow:0 14px 30px rgba(15,53,96,0.08); }
            .belastning-order-table { width:100%; border-collapse:separate; border-spacing:0; min-width:920px; }
            .belastning-order-table thead th { position:sticky; top:0; z-index:3; background:linear-gradient(180deg,#12426f 0%,#0f3560 100%); color:#fff; font-size:9px; letter-spacing:0.05em; text-transform:uppercase; padding:6px 5px; border-bottom:1px solid rgba(255,255,255,0.16); }
            .belastning-order-table thead th:first-child { border-top-left-radius:14px; }
            .belastning-order-table thead th:last-child { border-top-right-radius:14px; }
            .belastning-order-table tbody td { padding:4px 5px; border-bottom:1px solid #dbe8f7; font-size:10px; color:#173452; vertical-align:top; line-height:1.2; }
            .belastning-order-table tbody tr:not(.belastning-date-row):not(.belastning-order-detail-row):hover td { background:#f2f8ff; }
            .belastning-date-row td { background:linear-gradient(90deg,#7be07e 0%,#8eed90 100%) !important; color:#103f16 !important; border-bottom:1px solid #67c86b; padding:6px 9px !important; font-weight:800; }
            .belastning-date-row.belastning-before-day td { background:linear-gradient(90deg,#b68557 0%,#8e5e34 100%) !important; color:#fff8ef !important; border-bottom:1px solid #7b4b22; }
            .belastning-date-row.belastning-day-focus td { box-shadow:inset 0 0 0 2px rgba(18,66,111,0.28); }
            .belastning-date-header { display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; }
            .belastning-date-meta { font-size:10px; opacity:0.92; font-weight:700; display:flex; align-items:center; gap:0; flex-wrap:wrap; }
            .belastning-date-meta-item { display:inline-flex; align-items:center; gap:3px; margin-right:14px; white-space:nowrap; }
            .belastning-date-meta-lbl { opacity:0.8; font-weight:600; }
            .belastning-date-meta-val { min-width:48px; text-align:right; font-weight:900; font-variant-numeric:tabular-nums; }
            .belastning-date-badge { display:inline-flex; align-items:center; gap:8px; font-weight:900; letter-spacing:0.02em; font-size:11px; }
            .belastning-order-row td { background:#fff8df; }
            .belastning-order-row:nth-of-type(odd) td { background:#fff1cc; }
            .belastning-order-row.belastning-before-day td { background:#f2dfce; }
            .belastning-order-row.belastning-before-day:nth-of-type(odd) td { background:#eacfb8; }
            .belastning-order-detail-row td { background:#dff3ff; }
            .belastning-order-detail-row.belastning-before-day td { background:#eddac8; }
            .belastning-order-detail-row .belastning-order-detail-cell { padding:10px 10px 12px !important; border-top:1px dashed #b9d6f0; }
            .belastning-order-toggle,
            .belastning-date-toggle { width:18px; height:18px; border-radius:5px; border:1px solid rgba(255,255,255,0.52); background:rgba(255,255,255,0.22); color:#fff; font-weight:800; cursor:pointer; font-size:10px; line-height:1; }
            .belastning-order-toggle { border-color:#8eb6df; background:linear-gradient(180deg,#eff6ff 0%,#dcecff 100%); color:#11406e; box-shadow:inset 0 1px 0 rgba(255,255,255,0.85); }
            .belastning-order-toggle:hover { background:linear-gradient(180deg,#e1efff 0%,#cde2fb 100%); }
            .belastning-order-id { font-weight:900; color:#0f3560; font-size:10px; }
            .belastning-order-sub { color:#58769a; font-size:10px; }
            .belastning-order-chip { display:inline-flex; align-items:center; border:1px solid #bcd3f2; border-radius:999px; padding:2px 8px; background:#f7fbff; font-size:11px; color:#1b4979; box-shadow:0 1px 0 rgba(255,255,255,0.8); }
            .belastning-order-shell td, .belastning-order-shell th { backdrop-filter:saturate(1.04); }
            .belastning-order-lines { display:flex; flex-direction:column; gap:6px; }
            .belastning-order-lines > div { background:rgba(255,255,255,0.72); border:1px solid #bfd9f0; border-radius:10px; padding:8px 10px; }
            .belastning-order-lines strong { color:#0f3560; }
            #belastningGrafiskWrap.omsaetning-charts { grid-template-columns:1fr; }
            .belastning-kunde-wrap { position:relative; }
            .belastning-kunde-pill { display:none; align-items:center; justify-content:space-between; gap:6px; padding:7px 8px 7px 10px; border:1px solid #c7d9f3; border-radius:8px; background:linear-gradient(180deg,#ffffff 0%,#f4f9ff 100%); font-size:13px; color:#1a3a5c; font-weight:600; min-height:36px; box-sizing:border-box; }
            .belastning-kunde-pill.active { display:flex; }
            .belastning-kunde-pill .pill-name { white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
            .belastning-kunde-clear { flex-shrink:0; width:22px; height:22px; border:none; border-radius:50%; background:#c7d9f3; color:#0f3560; font-size:14px; line-height:22px; text-align:center; cursor:pointer; padding:0; font-weight:bold; }
            .belastning-kunde-clear:hover { background:#9bbce8; color:#fff; }
            .belastning-kunde-dropdown { position:absolute; z-index:220; left:0; right:0; top:calc(100% + 2px); border:1px solid #c7d9f3; border-radius:8px; background:#fff; box-shadow:0 4px 16px rgba(15,53,96,0.13); max-height:200px; overflow-y:auto; }
            .belastning-kunde-option { padding:7px 10px; cursor:pointer; font-size:13px; color:#1a3a5c; border-bottom:1px solid #e8f0fb; }
            .belastning-kunde-option:last-child { border-bottom:none; }
            .belastning-kunde-option:hover { background:#ecf4ff; }
            .belastning-kunde-option .bko-sub { font-size:11px; color:#6b7f95; margin-left:6px; }
            .side-menu-actions { display:flex; flex-direction:column; gap:8px; }
            .side-menu-actions button { border:none; border-radius:10px; padding:9px 10px; font-weight:700; cursor:pointer; }
            .side-menu-actions .logout { background:linear-gradient(180deg,#b71c1c 0%,#8f1717 100%); color:#fff; }
            .side-menu-actions .logout[disabled] { opacity:0.5; cursor:not-allowed; }
            .header-brand { display:flex; align-items:center; gap:10px; min-width:0; flex:1; }
            .header-brand-logo { width:38px; height:38px; border-radius:8px; background:transparent; padding:5px; object-fit:contain; border:1px solid rgba(255,255,255,0.22); box-shadow:0 6px 16px rgba(4,16,30,0.25); filter:brightness(0) invert(1) contrast(1.08); flex-shrink:0; }
            .header-brand-text { min-width:0; font-size:20px; font-weight:800; letter-spacing:0.01em; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
            .header-user-greeting { display:inline-block; font-size:13px; font-weight:700; color:#0f3560; background:linear-gradient(180deg,#e9f3ff 0%,#dbeaff 100%); border:1px solid #c9defa; border-radius:999px; padding:7px 12px; white-space:nowrap; }
            #warmupBarWrap { display:none !important; align-items:center; gap:8px; background:rgba(0,0,0,0.15); border-radius:8px; padding:4px 10px; font-size:12px; color:#fff; white-space:nowrap; }
            #warmupBarWrap.active { display:flex; }
            #warmupBarBg { background:rgba(255,255,255,0.25); border-radius:999px; height:6px; width:110px; overflow:hidden; flex-shrink:0; }
            #warmupBarFill { background:#fff; height:100%; border-radius:999px; width:0%; transition:width 0.35s ease; }
            .search-box { background: linear-gradient(180deg, #ffffff 0%, #f4f9ff 100%); padding: 14px; margin-bottom: 18px; border-radius: 12px; box-shadow: 0 10px 24px rgba(15,53,96,0.10); border: 1px solid #d9e8f9; position: sticky; top: var(--search-sticky-top, 58px); z-index: 1100; display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
            .search-box.collapsed { padding: 8px 12px; height: 36px; }
            .search-box.collapsed > * { display: none; }
            .search-box.collapsed > #collapseToggleBtn { display: inline-block; }
            #collapseToggleBtn { background: linear-gradient(180deg, #5ca646 0%, #3f8d31 100%); color: #fff; border: 1px solid rgba(255,255,255,0.28); padding: 8px 12px; border-radius: 999px; cursor: pointer; font-weight: 700; font-size: 12px; box-shadow: 0 6px 14px rgba(63,141,49,0.28); }
            .build-badge { display: inline-block; font-size: 12px; color: #444; background: #f1f1f1; border: 1px solid #ddd; border-radius: 4px; padding: 4px 8px; }
            .build-banner { display: none; }
            .search-box input { padding: 8px 12px; font-size: 14px; width: 200px; border: 1px solid #c7d7ea; border-radius: 8px; }
            .search-box button { padding: 8px 14px; background: linear-gradient(180deg, #1565c0 0%, #0f3560 100%); color: white; border: 1px solid rgba(255,255,255,0.10); border-radius: 999px; cursor: pointer; margin-left: 0; font-weight: 700; letter-spacing: 0.01em; }
            .mode-btn { background: linear-gradient(180deg, #0f3560 0%, #0d2f53 100%) !important; color:#fff !important; border:1px solid rgba(255,255,255,0.08); }
            .list-toggle-btn { background: linear-gradient(180deg, #123f6f 0%, #0f3560 100%) !important; color: #fff !important; border:1px solid rgba(255,255,255,0.08); }
            .search-box button:hover { filter: brightness(1.04); }
            .search-box button:disabled { opacity: 0.55; cursor: not-allowed; }
            .report-open-btn { position: relative; padding-left: 36px !important; }
            .report-open-btn::before { content: '↗'; position: absolute; left: 10px; top: 50%; transform: translateY(-50%); width: 18px; height: 18px; border-radius: 999px; display: inline-flex; align-items: center; justify-content: center; background: #5ca646; color: #fff; font-size: 11px; box-shadow: 0 4px 10px rgba(92,166,70,0.40); }
            .report-open-btn:disabled::before { background: #9ca3af; box-shadow: none; }
            .filter-input { width: 260px !important; margin-left: 10px; }
            .order-value-filter-toggle { display:inline-flex; align-items:center; gap:6px; padding:6px 10px; border:1px solid #c9dbf2; border-radius:999px; background:#f6faff; color:#1d446b; font-size:12px; font-weight:700; }
            .order-value-filter-toggle input { width:auto !important; margin:0; }
            .order-value-filter-input { width:130px !important; }
            .order-value-filter-input:disabled { opacity:0.65; background:#f0f4f9; }
            .filter-select { width: 180px; padding: 8px 10px; border: 1px solid #ddd; border-radius: 3px; background: #fff; }
            .section { background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%); margin-bottom: 16px; border-radius: 16px; box-shadow: 0 12px 28px rgba(15,53,96,0.08); padding: 18px; border: 1px solid #d8e8fb; }
            .order-header { background: radial-gradient(680px 180px at 92% -20%, rgba(126,177,230,0.36) 0%, rgba(126,177,230,0.08) 45%, rgba(126,177,230,0) 70%), linear-gradient(135deg, #0f3560 0%, #12426f 56%, #1565c0 100%); color: white; padding: 20px; border-radius: 16px; margin-bottom: 18px; box-shadow: 0 16px 32px rgba(15,53,96,0.24); border: 1px solid rgba(255,255,255,0.15); }
            .order-header h2 { margin: 0 0 16px 0; font-size: clamp(22px, 2vw, 30px); font-weight: 800; letter-spacing: 0.01em; }
            .order-header-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; }
            .order-header-item { display: flex; flex-direction: column; background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.14); border-radius: 12px; padding: 10px 12px; backdrop-filter: blur(2px); }
            .order-header-label { font-size: 11px; font-weight: 700; opacity: 0.9; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 4px; }
            .order-header-value { font-size: 20px; font-weight: 800; color: #fff; }
            .invoice-status-badge { display: inline-block; font-size: 14px; font-weight: 700; padding: 4px 10px; border-radius: 6px; white-space: nowrap; }
            .status-in-production { background: rgba(255,255,255,0.15); color: #fff; border: 2px solid rgba(255,255,255,0.4); }
            .status-partial-invoiced { background: #e65100; color: #fff; }
            .status-fully-invoiced { background: #2e7d32; color: #fff; }
            h3 { color: var(--ink-900); margin-bottom: 14px; border-bottom: 2px solid #7eb1e6; padding-bottom: 10px; font-size: clamp(15px, 1.5vw, 19px); letter-spacing: 0.01em; }
            table { width: 100%; border-collapse: separate; border-spacing: 0; margin-top: 10px; background: #fff; border: 1px solid var(--line-soft); border-radius: var(--radius-m); overflow: hidden; box-shadow: 0 6px 16px rgba(15,53,96,0.05); }
            th, td { padding: 10px 11px; text-align: left; border-bottom: 1px solid #e8eff8; font-size: 13px; color: var(--text-900); }
            th { position: static; top: auto; z-index: 1; background: linear-gradient(180deg, #f4f9ff 0%, #eaf2ff 100%); font-weight: 700; color: var(--ink-900); text-transform: uppercase; font-size: 11px; letter-spacing: 0.04em; }
            .order-list-table th { top: auto; }
            tr:last-child td { border-bottom: none; }
            tr:nth-child(even) td { background: #fcfdff; }
            tr:hover td { background: #f3f8ff; }
            .micro-grid-table th,
            .micro-grid-table td { padding: 7px 10px; }
            .micro-grid-table td { line-height: 1.26; }
            .summary-row { font-weight: bold; background: #f2f7ff; }
            .summary-box { background: linear-gradient(160deg, #f9fbff 0%, #eef4ff 54%, #e8f0ff 100%); border: 1px solid #d4e3fb; box-shadow: 0 10px 24px rgba(15,53,96,0.10); padding: 16px 18px; border-radius: 14px; margin-top: 15px; }
            .summary-box div { margin: 8px 0; font-size: 14px; color: #0e2f4c; }
            .summary-box .total { font-size: 18px; color: #0f3560; font-weight: 800; }
            .order-list-summary-actions { display:flex; gap:8px; align-items:center; justify-content:flex-end; margin-top:10px; flex-wrap:wrap; }
            .order-detail-report { margin-top:16px; background:#fff; border:1px solid #d6e9ff; border-radius:8px; padding:14px 14px 16px 14px; }
            .order-report-toolbar { display:flex; justify-content:space-between; gap:10px; align-items:center; flex-wrap:wrap; margin-bottom:12px; }
            .order-report-toolbar .order-report-meta { font-size:13px; color:#4b5563; }
            .order-report-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin-top:12px; }
            .order-report-card { background:#f7fbff; border:1px solid #d6e9ff; border-radius:8px; padding:10px 12px; }
            .order-report-card .label { font-size:12px; color:#57718f; margin-bottom:4px; }
            .order-report-card .value { font-size:18px; font-weight:700; color:#0f3560; }
            .order-report-table { width:100%; border-collapse:collapse; margin-top:12px; font-size:13px; }
            .order-report-table th, .order-report-table td { border-bottom:1px solid #e5e7eb; padding:8px 10px; text-align:left; }
            .order-report-table th { background:#f3f8ff; color:#0f3560; }
            .order-report-table tr:last-child td { border-bottom:none; }
            .order-report-table .summary-row td { background:#f7fbff; font-weight:700; }
            .order-report-table tr { page-break-inside: avoid; }
            .print-preview-overlay { position:fixed; inset:0; background:rgba(0,0,0,0.55); z-index:15000; display:none; align-items:center; justify-content:center; padding:18px; }
            .print-preview-dialog { width:min(1320px, 96vw); max-height:92vh; background:#fff; border-radius:10px; box-shadow:0 18px 46px rgba(0,0,0,0.34); display:flex; flex-direction:column; overflow:hidden; outline:none; }
            .print-preview-header { display:flex; justify-content:space-between; align-items:center; gap:12px; padding:14px 16px; border-bottom:1px solid #e5e7eb; background:#f8fbff; }
            .print-preview-title { font-size:16px; font-weight:800; color:#0f3560; }
            .print-preview-actions { display:flex; gap:8px; flex-wrap:wrap; }
            .print-preview-body { padding:16px; overflow:auto; background:#fff; max-height:calc(92vh - 62px); }
            .print-preview-body .order-list-summary-actions,
            .print-preview-body .order-report-actions { display:none !important; }
            .print-preview-body .order-detail-report { display:block !important; margin-top:0; border:none; box-shadow:none; padding:0; }
            .print-preview-body .order-list-section { margin-bottom:0; box-shadow:none; }
            .print-preview-body .order-report-grid { grid-template-columns:repeat(4,minmax(0,1fr)); }
            body.print-preview-lock { overflow:hidden; }
            .order-detail-modal-overlay { position:fixed; inset:0; z-index:15100; display:none; align-items:stretch; justify-content:center; background:rgba(7,18,35,0.68); backdrop-filter:blur(6px); padding:16px; }
            .order-detail-modal-shell { width:min(1540px, 100%); height:100%; background:linear-gradient(180deg, #0f3560 0%, #123f6f 56%, #0e2f4c 100%); border-radius:18px; box-shadow:0 24px 72px rgba(0,0,0,0.42); display:flex; flex-direction:column; overflow:hidden; }
            .order-detail-modal-header { display:flex; align-items:center; justify-content:space-between; gap:12px; padding:16px 18px; border-bottom:1px solid rgba(255,255,255,0.10); color:#fff; }
            .order-detail-modal-title { display:flex; flex-direction:column; gap:4px; }
            .order-detail-modal-title strong { font-size:clamp(16px, 1.45vw, 21px); letter-spacing:0.2px; }
            .order-detail-modal-title span { font-size:12px; color:rgba(255,255,255,0.76); }
            .order-detail-modal-actions { display:flex; gap:10px; flex-wrap:wrap; }
            .order-detail-modal-actions .list-toggle-btn { background:linear-gradient(180deg, #5ca646 0%, #4e9440 100%) !important; color:#fff !important; border:1px solid rgba(255,255,255,0.16) !important; box-shadow:0 10px 20px rgba(92,166,70,0.18); }
            .order-detail-modal-actions .list-toggle-btn:hover { filter:brightness(1.04); }
            .order-detail-modal-body { flex:1; overflow:auto; background:linear-gradient(180deg, #f5f8fd 0%, #edf2f8 100%); padding:18px; }
            .order-detail-modal-body .order-detail-report { display:block !important; margin:0 auto; max-width:1480px; }
            .order-detail-modal-body .order-report-actions { display:flex !important; justify-content:flex-end; margin:8px 0 14px 0; }
            .order-detail-modal-body .order-report-toolbar { position:sticky; top:0; z-index:2; backdrop-filter:blur(10px); background:linear-gradient(135deg, rgba(7,18,35,0.92), rgba(19,45,79,0.92)); border:1px solid rgba(255,255,255,0.08); border-radius:18px; padding:18px 20px; margin-bottom:16px; box-shadow:0 18px 40px rgba(15,53,96,0.16); }
            .order-detail-modal-body .order-report-meta { display:flex; flex-direction:column; gap:8px; color:#fff; }
            .order-detail-modal-body .order-report-meta strong { font-size:24px; line-height:1.15; letter-spacing:0.2px; }
            .order-detail-modal-body .order-report-meta .report-subline { font-size:13px; color:rgba(255,255,255,0.78); }
            .order-detail-modal-body .report-badges { display:flex; gap:8px; flex-wrap:wrap; margin-top:4px; }
            .order-detail-modal-body .report-badge { display:inline-flex; align-items:center; gap:6px; padding:6px 10px; border-radius:999px; font-size:12px; font-weight:700; border:1px solid rgba(255,255,255,0.12); background:rgba(255,255,255,0.08); color:#fff; }
            .order-detail-modal-body .report-badge strong { font-weight:800; }
            .order-detail-modal-body .report-hero { background:linear-gradient(135deg, #0f3560 0%, #123f6f 56%, #0e2f4c 100%); border:1px solid rgba(255,255,255,0.10); border-radius:22px; padding:18px 20px; box-shadow:0 16px 36px rgba(15,53,96,0.18); margin-bottom:16px; position:relative; overflow:hidden; }
            .order-detail-modal-body .report-hero::after { content:''; position:absolute; inset:auto -70px -70px auto; width:220px; height:220px; border-radius:50%; background:radial-gradient(circle, rgba(92,166,70,0.28) 0%, rgba(92,166,70,0.06) 62%, rgba(92,166,70,0) 72%); pointer-events:none; }
            .order-detail-modal-body .report-hero-top { display:flex; justify-content:space-between; align-items:flex-start; gap:16px; flex-wrap:wrap; }
            .order-detail-modal-body .report-hero-title { display:flex; flex-direction:column; gap:6px; }
            .order-detail-modal-body .report-hero-title .eyebrow { font-size:11px; text-transform:uppercase; letter-spacing:0.12em; color:#bfe2b5; font-weight:800; }
            .order-detail-modal-body .report-hero-title h1 { margin:0; font-size:clamp(22px, 2.35vw, 32px); line-height:1.12; color:#ffffff; }
            .order-detail-modal-body .report-hero-title .context { font-size:13px; color:rgba(255,255,255,0.82); max-width:960px; }
            .order-detail-modal-body .report-hero-meta { display:flex; flex-direction:column; align-items:flex-end; gap:8px; min-width:220px; }
            .order-detail-modal-body .report-hero-meta .stamp { font-size:12px; color:#ffffff; text-align:right; background:rgba(255,255,255,0.10); border:1px solid rgba(255,255,255,0.16); border-radius:999px; padding:6px 10px; }
            .order-detail-modal-body .report-pill-row { display:flex; gap:8px; flex-wrap:wrap; margin-top:12px; }
            .order-detail-modal-body .report-pill { display:inline-flex; align-items:center; gap:8px; padding:7px 12px; border-radius:999px; background:rgba(255,255,255,0.10); color:#ffffff; border:1px solid rgba(255,255,255,0.16); font-size:12px; font-weight:700; }
            .order-detail-modal-body .report-pill.warn { background:rgba(92,166,70,0.16); border-color:rgba(92,166,70,0.32); color:#edf9e8; }
            .order-detail-modal-body .report-pill.ok { background:rgba(255,255,255,0.14); border-color:rgba(255,255,255,0.20); color:#ffffff; }
            .order-detail-modal-body .report-pill strong { font-size:13px; color:#ffffff; }
            .order-detail-modal-body .report-arrow { margin-top:6px; font-size:12px; font-weight:700; color:#dff4d8; display:inline-flex; align-items:center; gap:6px; }
            .order-detail-modal-body .report-arrow::before { content:'↗'; width:18px; height:18px; border-radius:999px; display:inline-flex; align-items:center; justify-content:center; background:#5ca646; color:#fff; font-size:11px; }
            .order-detail-modal-body .section { background:#fff; border:1px solid #dde9f6; border-radius:18px; box-shadow:0 12px 30px rgba(15,53,96,0.08); padding:16px 18px; margin-bottom:16px; }
            .order-detail-modal-body .section h3 { margin-top:0; color:#0f3560; font-size:16px; letter-spacing:0.1px; }
            .order-detail-modal-body .summary-box { background:linear-gradient(180deg, #f7fbff 0%, #eff6ff 100%); border:1px solid #d8e8fb; border-radius:16px; padding:14px 16px; }
            .order-detail-modal-body .order-report-grid { grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; }
            .order-detail-modal-body .order-report-card { background:linear-gradient(180deg, #ffffff 0%, #f7fbff 100%); border:1px solid #d7e6f8; border-radius:16px; box-shadow:0 10px 24px rgba(15,53,96,0.06); position:relative; overflow:hidden; }
            .order-detail-modal-body .order-report-card::before { content:''; position:absolute; inset:0 auto auto 0; width:4px; height:100%; background:linear-gradient(180deg, #1565c0, #4f8bdc); }
            .order-detail-modal-body .order-report-card .label { font-size:12px; color:#5c7590; margin-bottom:5px; text-transform:uppercase; letter-spacing:0.04em; }
            .order-detail-modal-body .order-report-card .value { font-size:21px; font-weight:800; color:#0f3560; }
            .order-detail-modal-body .order-report-table { margin-top:14px; border-radius:14px; overflow:hidden; border:1px solid #dfeaf7; }
            .order-detail-modal-body .order-report-table th { background:linear-gradient(180deg, #eef5ff 0%, #e2edff 100%); color:#0f3560; font-size:12px; text-transform:uppercase; letter-spacing:0.03em; }
            .order-detail-modal-body .order-report-table td { background:#fff; }
            .order-detail-modal-body .order-report-table tr:nth-child(even) td { background:#fafcff; }
            .order-detail-modal-body .order-report-table .summary-row td { background:#eff6ff; }
            .oversigt-launcher-section { background:linear-gradient(180deg, #ffffff 0%, #f7fbff 100%); border:1px solid #d8e8fb; }
            .oversigt-launcher-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }
            .oversigt-launcher-card { background:#ffffff; border:1px solid #d8e8fb; border-radius:16px; padding:14px; box-shadow:0 10px 24px rgba(15,53,96,0.06); display:flex; flex-direction:column; gap:10px; min-height:160px; }
            .oversigt-launcher-card h4 { margin:0; color:#0f3560; font-size:15px; }
            .oversigt-launcher-card .desc { color:#5c7590; font-size:12px; }
            .oversigt-launcher-kpi { background:linear-gradient(180deg, #f7fbff 0%, #eef5ff 100%); border:1px solid #d9e8fa; border-radius:12px; padding:10px 12px; font-size:13px; color:#0f3560; min-height:52px; display:flex; align-items:center; }
            .oversigt-launcher-card .list-toggle-btn { align-self:flex-start; margin-top:auto; }
            .oversigt-modal-overlay { position:fixed; inset:0; z-index:15160; display:none; align-items:center; justify-content:center; background:rgba(7,18,35,0.70); backdrop-filter:blur(4px); padding:14px; }
            .oversigt-modal-shell { width:min(1520px, 99vw); max-height:94vh; background:linear-gradient(180deg, #f5f9ff 0%, #edf3fb 100%); border-radius:18px; box-shadow:0 24px 72px rgba(0,0,0,0.42); overflow:hidden; display:flex; flex-direction:column; border:1px solid rgba(255,255,255,0.24); }
            .oversigt-modal-header { display:flex; justify-content:space-between; align-items:center; gap:10px; padding:14px 16px; background:linear-gradient(135deg, #0f3560 0%, #123f6f 56%, #0e2f4c 100%); color:#fff; }
            .oversigt-modal-title-wrap { display:flex; flex-direction:column; gap:4px; }
            .oversigt-modal-title-wrap strong { font-size:17px; letter-spacing:0.2px; }
            .oversigt-modal-title-wrap span { font-size:12px; color:rgba(255,255,255,0.82); }
            .oversigt-modal-actions { display:flex; gap:8px; flex-wrap:wrap; }
            .oversigt-modal-body { overflow:auto; padding:clamp(10px, 1.6vw, 18px); }
            .oversigt-modal-layout { display:grid; grid-template-columns:minmax(210px, 280px) minmax(0, 1fr); gap:10px; align-items:start; }
            .oversigt-panel { background:#fff; border:1px solid #dce9f7; border-radius:14px; box-shadow:0 8px 20px rgba(15,53,96,0.07); padding:12px; }
            .oversigt-panel h5 { margin:0 0 8px 0; font-size:13px; color:#0f3560; text-transform:uppercase; letter-spacing:0.05em; }
            .oversigt-panel.oversigt-kpi { padding:10px; }
            .oversigt-panel.oversigt-kpi .summary-box { padding:10px 12px; border-radius:10px; }
            .oversigt-panel.oversigt-kpi .summary-box div { margin:3px 0; font-size:12px; line-height:1.25; }
            .oversigt-panel .summary-box { margin:0; }
            .oversigt-panel table { margin-top:0; }
            .oversigt-details { overflow-x:hidden; overflow-y:auto; min-width:0; }
            .oversigt-details table { min-width:0; width:100%; table-layout:fixed; }
            .oversigt-details th,
            .oversigt-details td { padding:6px 8px; font-size:12px; white-space:normal; word-break:break-word; }
            .oversigt-details td:nth-child(3) { font-weight:700; color:#153b63; }
            .oversigt-details td:nth-child(n+5):nth-child(-n+11) { font-variant-numeric: tabular-nums; }
            .oversigt-table-laser th:nth-child(1) { width:10%; }
            .oversigt-table-laser th:nth-child(2) { width:10%; }
            .oversigt-table-laser th:nth-child(3) { width:16%; }
            .oversigt-table-laser th:nth-child(4) { width:6%; }
            .oversigt-table-laser th:nth-child(5) { width:7%; }
            .oversigt-table-laser th:nth-child(6),
            .oversigt-table-laser th:nth-child(7),
            .oversigt-table-laser th:nth-child(8),
            .oversigt-table-laser th:nth-child(9),
            .oversigt-table-laser th:nth-child(10),
            .oversigt-table-laser th:nth-child(11) { width:7%; }
            .oversigt-table-laser th:nth-child(12) { width:5%; }
            .oversigt-table-operation th:nth-child(1) { width:15%; }
            .oversigt-table-operation th:nth-child(2) { width:27%; }
            .oversigt-table-operation th:nth-child(3),
            .oversigt-table-operation th:nth-child(4),
            .oversigt-table-operation th:nth-child(5),
            .oversigt-table-operation th:nth-child(6),
            .oversigt-table-operation th:nth-child(7) { width:11%; }
            .order-margin-wrap { display:inline-flex; align-items:center; gap:6px; }
            .order-kpi-tone { width:18px; height:18px; border-radius:999px; display:inline-flex; align-items:center; justify-content:center; color:#fff; font-size:11px; font-weight:800; box-shadow:0 4px 10px rgba(15,53,96,0.20); }
            .order-kpi-tone.ok { background:#5ca646; }
            .order-kpi-tone.warn { background:#f59e0b; }
            .order-kpi-tone.bad { background:#d32f2f; }
            .order-kpi-tone.na { background:#607d8b; }
            body.report-modal-open { overflow:hidden; }
            @media print {
                @page { size: A4 portrait; margin: 12mm; }
                body.print-report-mode .header-banner-wrapper,
                body.print-report-mode .search-box,
                body.print-report-mode #orderList,
                body.print-report-mode #result > :not(#orderDetailReport) { display:none !important; }
                body.print-report-mode #result { display:block !important; }
                body.print-report-mode #orderDetailReport { display:block !important; margin:0 !important; border:none !important; box-shadow:none !important; }
                body.print-report-mode .order-report-toolbar button,
                body.print-report-mode .order-report-actions button { display:none !important; }
                body.print-report-mode .order-report-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
                body.print-report-mode .order-report-card,
                body.print-report-mode .order-report-table { break-inside: avoid; page-break-inside: avoid; }

                body.print-list-mode .header-banner-wrapper,
                body.print-list-mode .search-box,
                body.print-list-mode #result { display:none !important; }
                body.print-list-mode #orderList { display:block !important; }
                body.print-list-mode .order-list-summary-actions button { display:none !important; }

                body.print-preview-mode .header-banner-wrapper,
                body.print-preview-mode .search-box,
                body.print-preview-mode #orderList,
                body.print-preview-mode #result { display:none !important; }
                body.print-preview-mode .print-preview-overlay { display:flex !important; }
                body.print-preview-mode .print-preview-dialog { box-shadow:none; width:100%; max-height:none; }
                body.print-preview-mode .print-preview-header { display:none !important; }
                body.print-preview-mode .print-preview-body { padding:0; overflow:visible; }
                body.print-preview-mode .print-preview-body .order-list-summary-actions,
                body.print-preview-mode .print-preview-body .order-report-actions { display:none !important; }
            }
            .margin-positive { color: green; }
            .margin-negative { color: red; }
            .error { color: #9f1239; padding: 14px 16px; background: linear-gradient(180deg, #fff1f2 0%, #ffe4e6 100%); border: 1px solid #fecdd3; border-radius: 12px; font-weight: 600; }
            .loading { color: #0f3560; padding: 14px 16px; border-radius: 12px; border: 1px solid #d7e6fb; background: linear-gradient(110deg, #f8fbff 8%, #edf4ff 34%, #f8fbff 60%); background-size: 220% 100%; animation: softLoadingWave 1.25s ease-in-out infinite; font-weight: 600; }
            .prod-link { color: #1976D2; text-decoration: underline; cursor: pointer; }
            .prod-link:hover { color: #0D47A1; }
            .po-highlight { box-shadow: 0 0 0 3px #90CAF9; }
            .prodtp4-group { border: 1px solid #d8e7fb; border-radius: 12px; margin-bottom: 10px; overflow: hidden; box-shadow: 0 6px 16px rgba(15,53,96,0.05); background: #fff; }
            .prodtp4-header { background: linear-gradient(180deg, #f4f9ff 0%, #eaf3ff 100%); padding: 11px 12px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; font-weight: 800; }
            .prodtp4-header:hover { background: linear-gradient(180deg, #eef6ff 0%, #e2eeff 100%); }
            .prodtp4-label { color: #173b62; }
            .prodtp4-subtotal { color: #1565c0; font-weight: 800; }
            .prodtp4-body { padding: 10px 12px 12px; }
            .po-total-row { margin-top: 10px; padding: 10px 12px; border-top: 1px solid #d8e5f7; font-weight: 800; text-align: right; background: linear-gradient(180deg, #f8fbff 0%, #edf4ff 100%); color: var(--ink-900); border-radius: 8px; }
            .prodtp4-hint { color: #555; margin: 6px 0 10px; font-size: 13px; }
            .main-product-box { background: linear-gradient(180deg, #f8fbff 0%, #edf4ff 100%); border: 1px solid #bdd8f6; border-radius: 10px; padding: 10px 12px; margin: 8px 0 12px; box-shadow: 0 8px 20px rgba(15,53,96,0.08); }
            .main-product-box .value { font-size: 20px; font-weight: 800; color: #0d47a1; margin-top: 3px; }
            .inline-link { color: #1565c0; text-decoration: underline; cursor: pointer; }
            .inline-link:hover { color: #0d47a1; }
            .prod-no-link { color: #1565c0; text-decoration: underline; cursor: pointer; }
            .prod-no-link:hover { color: #0d47a1; }
            .modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.35); display: none; align-items: center; justify-content: center; z-index: 15220; }
            .modal-box { width: min(1400px, 97vw); max-height: 90vh; overflow: auto; background: linear-gradient(180deg, #ffffff 0%, #f6f9ff 100%); border: 1px solid #d9e8fb; border-radius: 14px; box-shadow: 0 18px 42px rgba(6,26,48,0.24); padding: 16px; }
            .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
            .modal-header-left { display: flex; align-items: center; gap: 8px; }
            .modal-content-wrap { display: grid; grid-template-columns: minmax(0, 1fr) minmax(340px, 32vw); gap: 16px; align-items: start; }
            #summaryModalBody { flex: 1; min-width: 0; }
            .modal-back { border: 1px solid #c9dcf9; background: linear-gradient(180deg, #ffffff 0%, #edf4ff 100%); border-radius: 999px; padding: 6px 12px; cursor: pointer; font-weight: 700; color: #0f3560; letter-spacing: 0.01em; }
            .modal-back.hidden { display: none; }
            .modal-close { border: 1px solid #c9dcf9; background: linear-gradient(180deg, #ffffff 0%, #edf4ff 100%); border-radius: 999px; padding: 6px 12px; cursor: pointer; color: #0f3560; letter-spacing: 0.01em; }
            .modal-loading { color: #0f3560; padding: 10px 12px; border-radius: 10px; border: 1px solid #d8e7fb; background: linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%); font-weight: 600; }
            .summary-image-panel { width: 100%; min-width: 0; max-height: calc(90vh - 140px); overflow: auto; border-left: 1px solid #e0e0e0; padding-left: 16px; position: sticky; top: 0; background: #fff; }
            .summary-image-panel.hidden { display: none; }
            .modal-content-wrap.image-focus { grid-template-columns: 1fr; }
            .modal-content-wrap.image-focus #summaryModalBody { display: none; }
            .modal-content-wrap.image-focus .summary-image-panel { border-left: none; padding-left: 0; max-height: calc(90vh - 160px); position: relative; top: auto; }
            .laser-summary-layout { display: flex; gap: 12px; align-items: flex-start; }
            .laser-image-panel { width: min(360px, 32vw); min-width: 240px; max-height: 70vh; overflow: auto; border: 1px solid #e0e0e0; border-radius: 8px; padding: 10px; position: sticky; top: 12px; background: #fff; }
            .laser-image-panel.hidden { display: none; }
            .summary-image-panel-header { display: flex; justify-content: space-between; align-items: center; gap: 10px; margin-bottom: 12px; }
            .summary-image-panel-title { font-size: 16px; font-weight: 700; color: #1f2937; }
            .summary-image-close { border: none; background: #efefef; border-radius: 4px; padding: 6px 10px; cursor: pointer; }
            .image-preview-btn { padding: 6px 11px; border: 1px solid rgba(255,255,255,0.15); border-radius: 999px; background: linear-gradient(180deg, #1565c0 0%, #0f3560 100%); color: #fff; cursor: pointer; font-size: 12px; font-weight: 700; }
            .image-preview-btn:hover { filter: brightness(1.05); }
            .image-preview-gallery { display: grid; gap: 12px; }
            .image-preview-card { border: 1px solid #dae8fb; border-radius: 10px; padding: 10px; background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%); }
            .image-preview-card img { display: block; width: 100%; max-height: 240px; object-fit: contain; background: #fff; border-radius: 4px; border: 1px solid #e5e7eb; cursor: zoom-in; }
            .image-preview-label { font-size: 12px; font-weight: 700; color: #374151; margin-bottom: 8px; }
            .image-preview-path { font-size: 11px; color: #6b7280; word-break: break-all; margin-top: 8px; }
            .image-preview-empty { font-size: 13px; color: #6b7280; padding: 8px 0; }
            .image-lightbox { position: fixed; inset: 0; background: rgba(17, 24, 39, 0.88); display: flex; align-items: center; justify-content: center; padding: 24px; z-index: 15300; }
            .image-lightbox.hidden { display: none; }
            .image-lightbox-dialog { width: min(1200px, 96vw); max-height: 92vh; background: #111827; color: #f9fafb; border-radius: 10px; box-shadow: 0 18px 40px rgba(0,0,0,0.35); padding: 16px; display: flex; flex-direction: column; gap: 12px; }
            .image-lightbox-header { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
            .image-lightbox-title { font-size: 15px; font-weight: 700; color: #f9fafb; }
            .image-lightbox-close { border: none; background: rgba(255,255,255,0.12); color: #fff; border-radius: 4px; padding: 6px 10px; cursor: pointer; }
            .image-lightbox-close:hover { background: rgba(255,255,255,0.2); }
            .image-lightbox-body { display: flex; align-items: center; justify-content: center; min-height: 0; overflow: auto; }
            .image-lightbox-body img { display: block; max-width: 100%; max-height: calc(92vh - 110px); object-fit: contain; border-radius: 6px; background: #fff; }
            .image-lightbox-path { font-size: 12px; color: #d1d5db; word-break: break-all; }
            .compact-image-modal { position: fixed; inset: 0; background: rgba(7,18,35,0.72); display: none; align-items: center; justify-content: center; padding: 18px; z-index: 15320; }
            .compact-image-modal.show { display: flex; }
            .compact-image-dialog { width: min(1080px, 95vw); max-height: 92vh; background: linear-gradient(180deg, #ffffff 0%, #f5f9ff 100%); border: 1px solid #d6e5fa; border-radius: 14px; box-shadow: 0 24px 52px rgba(5,23,42,0.36); display: flex; flex-direction: column; overflow: hidden; }
            .compact-image-header { display:flex; justify-content:space-between; align-items:center; gap:10px; padding:12px 14px; border-bottom:1px solid #d9e6f8; background:linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%); }
            .compact-image-title { font-size:17px; font-weight:800; color:#0f3560; }
            .compact-image-subtitle { font-size:12px; color:#4f6b86; margin-top:2px; }
            .compact-image-close { border:1px solid #c9dcf9; background:linear-gradient(180deg,#fff 0%,#edf4ff 100%); color:#0f3560; border-radius:999px; padding:6px 12px; cursor:pointer; font-weight:700; }
            .compact-image-body { padding:12px; overflow:auto; }
            .compact-image-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:10px; }
            .compact-image-card { border:1px solid #dce8f8; border-radius:10px; padding:10px; background:#fff; }
            .compact-image-label { font-size:12px; font-weight:700; color:#2f4f70; margin-bottom:6px; text-transform:uppercase; letter-spacing:0.02em; }
            .compact-image-card img { width:100%; max-height:360px; object-fit:contain; background:#fff; border:1px solid #e5ecf7; border-radius:8px; display:block; cursor:zoom-in; }
            .compact-image-path { margin-top:6px; font-size:11px; color:#5f738a; word-break:break-all; }
            @keyframes softLoadingWave {
                0% { background-position: 100% 0; }
                100% { background-position: -100% 0; }
            }
            .search-box button,
            #collapseToggleBtn,
            .list-toggle-btn,
            .mode-btn,
            .modal-close,
            .modal-back,
            .summary-image-close,
            .image-preview-btn,
            .order-detail-modal-actions .list-toggle-btn {
                transition: transform 0.18s ease, box-shadow 0.2s ease, filter 0.18s ease;
            }
            .search-box button:hover,
            #collapseToggleBtn:hover,
            .list-toggle-btn:hover,
            .mode-btn:hover,
            .modal-close:hover,
            .modal-back:hover,
            .summary-image-close:hover,
            .image-preview-btn:hover,
            .order-detail-modal-actions .list-toggle-btn:hover {
                transform: translateY(-1px);
                box-shadow: 0 8px 18px rgba(15,53,96,0.18);
            }
            .search-box button:active,
            #collapseToggleBtn:active,
            .list-toggle-btn:active,
            .mode-btn:active,
            .modal-close:active,
            .modal-back:active,
            .summary-image-close:active,
            .image-preview-btn:active,
            .order-detail-modal-actions .list-toggle-btn:active {
                transform: translateY(0);
            }
            @media (max-width: 1360px) {
                .modal-content-wrap { grid-template-columns: minmax(0, 1fr); }
                .summary-image-panel { border-left: none; border-top: 1px solid #e0e0e0; padding-left: 0; padding-top: 12px; position: relative; top: auto; max-height: 58vh; }
                .oversigt-modal-layout { grid-template-columns:1fr; }
                .oversigt-details table { min-width:0; }
            }
            @media (max-width: 1180px) {
                .belastning-bars { grid-template-columns:1fr; }
            }
            @media (max-width: 900px) {
                .modal-box { width: 99vw; max-height: 93vh; padding: 12px; }
                .modal-box th, .modal-box td { padding: 8px 6px; font-size: 13px; }
                .dashboard-category-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
                .dashboard-warmup-notice { flex-direction:column; align-items:flex-start; }
                .dashboard-warmup-progress { width:100%; justify-content:space-between; }
                .dashboard-warmup-track { flex:1; }
                .omsaetning-filters { grid-template-columns:1fr 1fr; }
                .omsaetning-field.omsaetning-accounts-field,
                .omsaetning-field.omsaetning-customer-field,
                .omsaetning-field.omsaetning-threshold-field,
                .omsaetning-actions { grid-column:span 2; }
                .omsaetning-kpis { grid-template-columns:1fr; }
                .omsaetning-charts { grid-template-columns:1fr; }
                #belastningGrafiskWrap.omsaetning-charts { grid-template-columns:1fr; }
                .belastning-bars { grid-template-columns:1fr; }
                .modal-content-wrap { grid-template-columns: 1fr; }
                .summary-image-panel { width: 100%; min-width: 0; max-height: 52vh; border-left: none; border-top: 1px solid #e0e0e0; padding-left: 0; padding-top: 12px; }
                .laser-summary-layout { flex-direction: column; }
                .laser-image-panel { width: 100%; min-width: 0; max-height: none; position: static; }
                .image-lightbox { padding: 12px; }
                .image-lightbox-dialog { width: 100%; max-height: 96vh; padding: 12px; }
                .image-lightbox-body img { max-height: calc(96vh - 110px); }
                .compact-image-grid { grid-template-columns:1fr; }
                .oversigt-launcher-grid { grid-template-columns:1fr; }
                .oversigt-modal-layout { grid-template-columns:1fr; }
                .oversigt-modal-header { flex-direction:column; align-items:flex-start; }
                .oversigt-modal-actions { width:100%; justify-content:flex-end; }
                .oversigt-details table { min-width:0; }
            }
            @media (max-width: 640px) {
                .dashboard-category-grid { grid-template-columns:1fr; }
                .order-detail-modal-overlay { padding: 6px; }
                .order-detail-modal-shell { border-radius: 10px; }
                .order-detail-modal-header { padding: 10px 12px; }
                .order-detail-modal-actions { gap: 6px; }
                .order-detail-modal-actions .list-toggle-btn { padding: 6px 10px !important; font-size: 12px; }
                .order-detail-modal-body { padding: 10px; }
                .order-detail-modal-body .order-report-meta strong { font-size: 20px; }
                .order-detail-modal-body .report-hero-title h1 { font-size: 24px; }
                .oversigt-modal-overlay { padding: 6px; }
                .oversigt-modal-shell { border-radius: 10px; max-height: 96vh; }
                .oversigt-modal-header { padding: 10px 12px; }
                .oversigt-modal-body { padding: 8px; }
                .oversigt-details table { min-width:0; }
                .order-list-section { padding: 12px; overflow-x: auto; }
                .order-list-table { min-width: 820px; }
                .belastning-order-table { min-width: 700px; }
            }
            .order-list-section { background: linear-gradient(180deg, #ffffff 0%, #f7fbff 100%); padding: 14px 16px; margin-bottom: 16px; border-radius: 14px; box-shadow: 0 10px 22px rgba(15,53,96,0.07); border: 1px solid #d9e8fb; }
            .order-list-section h3 { color: #0f3560; margin-bottom: 10px; border-bottom: 1px solid #cfe1f8; padding-bottom: 8px; }
            .order-list-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13px; border: 1px solid #dce8f8; border-radius: 12px; overflow: hidden; }
            .order-list-table th { background: linear-gradient(180deg, #eef5ff 0%, #e4efff 100%); color: #0f3560; padding: 8px 10px; text-align: left; border-bottom: 1px solid #d8e6f8; }
            .order-list-table td { padding: 8px 10px; border-bottom: 1px solid #e7eff9; cursor: pointer; }
            .order-list-table tr:hover td { background: #edf5ff; }
            .note-badge { display:inline-flex; align-items:center; gap:3px; font-size:11px; font-weight:700; padding:2px 7px; border-radius:10px; border:1px solid; cursor:pointer; white-space:nowrap; }
            .note-badge.ok  { background:#e8f5e9; color:#1b5e20; border-color:#a5d6a7; }
            .note-badge.error { background:#ffebee; color:#b71c1c; border-color:#ef9a9a; }
            .note-badge.check { background:#fff8e1; color:#f57f17; border-color:#ffe082; }
            .note-badge.text { background:#f3e5f5; color:#4a148c; border-color:#ce93d8; }
            .note-badge.credit { background:#e3f2fd; color:#0d47a1; border-color:#90caf9; }
            .note-popup-overlay { position:fixed; inset:0; background:rgba(0,0,0,0.45); z-index:15450; display:flex; align-items:center; justify-content:center; }
            .note-popup { background:#fff; border-radius:10px; padding:22px; width:min(480px,92vw); box-shadow:0 18px 42px rgba(0,0,0,0.28); }
            .note-popup h3 { margin:0 0 14px 0; font-size:16px; color:#1f2937; }
            .note-popup label { font-size:12px; font-weight:600; color:#555; display:block; margin-bottom:4px; }
            .note-popup select, .note-popup textarea { width:100%; padding:8px 10px; border:1px solid #d1d5db; border-radius:6px; font-size:14px; font-family:inherit; }
            .note-popup textarea { resize:vertical; min-height:80px; margin-bottom:12px; }
            .note-popup select { margin-bottom:12px; }
            .note-popup-actions { display:flex; gap:8px; justify-content:flex-end; margin-top:6px; }
            .note-popup-actions button { border:none; border-radius:6px; padding:8px 16px; font-weight:700; font-size:13px; cursor:pointer; }
            .btn-note-save { background:#1565c0; color:#fff; }
            .btn-note-delete { background:#b71c1c; color:#fff; }
            .btn-note-cancel { background:#e0e0e0; color:#333; }
            .order-note-banner { margin:0 0 14px 0; padding:10px 14px; border-radius:8px; font-size:13px; display:flex; align-items:flex-start; gap:10px; cursor:pointer; }
            .order-note-banner.ok    { background:#e8f5e9; border:1px solid #a5d6a7; color:#1b5e20; }
            .order-note-banner.error { background:#ffebee; border:1px solid #ef9a9a; color:#b71c1c; }
            .order-note-banner.check { background:#fff8e1; border:1px solid #ffe082; color:#e65100; }
            .order-note-banner.text  { background:#f3e5f5; border:1px solid #ce93d8; color:#4a148c; }
            .order-note-banner.credit { background:#e3f2fd; border:1px solid #90caf9; color:#0d47a1; }
            .order-note-banner .note-icon { font-size:18px; flex-shrink:0; }
            .order-note-banner .note-body { flex:1; }
            .order-list-summary { margin-top:12px; margin-bottom:14px; background:#f7fbff; border:1px solid #d6e9ff; border-radius:8px; padding:10px 12px; font-size:13px; color:#0f3560; }
            .order-list-summary strong { color:#0d47a1; }
            .access-gate-overlay { position: fixed; inset: 0; background: radial-gradient(920px 320px at 12% -8%, rgba(46,125,50,0.22) 0%, rgba(46,125,50,0.02) 46%, transparent 72%), linear-gradient(135deg, rgba(8,26,48,0.78) 0%, rgba(18,56,95,0.84) 55%, rgba(14,41,71,0.9) 100%); backdrop-filter: blur(2px); display: none; align-items: center; justify-content: center; z-index: 12000; }
            .access-gate-box { width: min(500px, 92vw); background: linear-gradient(180deg,#ffffff 0%,#f8fbff 100%); border: 1px solid #d8e8fb; border-radius: 16px; padding: 24px; box-shadow: 0 26px 52px rgba(3,14,28,0.34); }
            .access-gate-brand { display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:8px; }
            .access-gate-brand h3 { margin:0; border:none; padding:0; color:#0f3560; font-size:24px; letter-spacing:0.01em; }
            .access-gate-badge { font-size:11px; font-weight:800; color:#1f5e24; background:#e6f5e8; border:1px solid #b8dfbd; border-radius:999px; padding:4px 10px; }
            .access-gate-box p { margin: 0 0 16px 0; color: #4b6380; font-size:13px; }
            .access-gate-fields { display:flex; flex-direction:column; gap:10px; }
            .access-gate-field label { display:block; margin:0 0 4px 1px; font-size:12px; color:#355675; font-weight:700; }
            .access-gate-field input { width:100%; padding:10px 11px; border:1px solid #c9dbf2; border-radius:8px; font-size:15px; background:#fff; color:#173452; }
            .access-gate-field input:focus { outline:none; border-color:#1565c0; box-shadow:0 0 0 3px rgba(21,101,192,0.14); }
            .access-gate-row { display: flex; gap: 8px; margin-top:14px; }
            .access-gate-row button { border: none; border-radius: 999px; background: linear-gradient(180deg,#1565c0 0%,#0f3560 100%); color: #fff; font-weight: 800; padding: 10px 18px; cursor: pointer; min-width:130px; margin-left:auto; }
            .access-gate-row button:disabled { opacity:0.7; cursor:not-allowed; }
            .access-gate-error { margin-top: 10px; min-height: 18px; color: #b71c1c; font-weight: 600; font-size: 13px; }
            .main-dashboard { display:none; margin-bottom:16px; }
            .dashboard-shell { position:relative; overflow:hidden; background:radial-gradient(1100px 360px at 8% -12%, rgba(22,101,192,0.19) 0%, rgba(22,101,192,0.03) 40%, transparent 70%), linear-gradient(160deg, #ffffff 0%, #f3f8ff 62%, #edf4ff 100%); border:1px solid #d7e6fb; border-radius:16px; box-shadow:0 14px 30px rgba(15,53,96,0.10); padding:16px; }
            .dashboard-shell::before { content:''; position:absolute; width:240px; height:240px; right:-95px; top:-105px; border-radius:50%; background:radial-gradient(circle at 30% 30%, rgba(86,164,255,0.24), rgba(86,164,255,0)); pointer-events:none; }
            .dashboard-shell::after { content:''; position:absolute; width:360px; height:110px; left:-90px; bottom:-60px; transform:rotate(-8deg); background:linear-gradient(90deg, rgba(15,53,96,0.00), rgba(15,53,96,0.08), rgba(15,53,96,0.00)); pointer-events:none; }
            .dashboard-head h2 { margin:0; color:#0f3560; font-size:26px; letter-spacing:0.01em; }
            .dashboard-head p { margin:6px 0 0 0; color:#4d6680; font-size:13px; }
            .dashboard-update-notice { margin-top:12px; border:1px solid #cfe3fb; background:linear-gradient(180deg,#f8fbff 0%,#eef5ff 100%); border-radius:12px; padding:10px 12px; display:none; align-items:center; justify-content:space-between; gap:12px; }
            .dashboard-update-notice.active { display:flex; }
            .dashboard-update-copy { font-size:12px; color:#355675; }
            .dashboard-update-copy strong { color:#0f3560; }
            .dashboard-update-actions { display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
            .dashboard-update-actions button { border:none; border-radius:999px; padding:7px 12px; font-size:12px; font-weight:700; cursor:pointer; color:#fff; background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); }
            .dashboard-update-actions button.install { background:linear-gradient(180deg,#2e7d32 0%,#1f5e24 100%); }
            .dashboard-warmup-notice { margin-top:10px; border:1px solid #d4e7fb; background:linear-gradient(180deg,#fbfdff 0%,#f1f7ff 100%); border-radius:12px; padding:10px 12px; display:flex; align-items:center; justify-content:space-between; gap:12px; }
            .dashboard-warmup-notice.hidden { display:none; }
            .dashboard-warmup-copy { font-size:12px; color:#355675; }
            .dashboard-warmup-copy strong { color:#0f3560; }
            .dashboard-warmup-meta { opacity:0.92; }
            .dashboard-warmup-progress { display:flex; align-items:center; gap:8px; min-width:180px; justify-content:flex-end; }
            .dashboard-warmup-track { width:140px; height:8px; border-radius:999px; background:#d9e9fb; overflow:hidden; }
            .dashboard-warmup-track > div { height:100%; width:0%; border-radius:999px; background:linear-gradient(90deg,#1565c0 0%,#2e7d32 100%); transition:width .3s ease; }
            .dashboard-warmup-pct { min-width:40px; text-align:right; font-size:12px; font-weight:700; color:#0f3560; }
            .dashboard-grid { margin-top:14px; display:grid; grid-template-columns:1fr; gap:12px; }
            .dashboard-category { background:rgba(255,255,255,0.65); border:1px solid #d8e8fb; border-radius:14px; padding:10px; box-shadow:inset 0 1px 0 rgba(255,255,255,0.9); }
            .dashboard-category-head { display:flex; align-items:center; justify-content:space-between; gap:8px; margin:0 0 8px 0; }
            .dashboard-category-head h3 { margin:0; border:none; padding:0; color:#0f3560; font-size:15px; font-weight:800; }
            .dashboard-category-head span { font-size:11px; color:#5b7897; font-weight:700; }
            .dashboard-category-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; }
            .dash-card { position:relative; isolation:isolate; border:1px solid #d8e6fa; border-radius:14px; background:linear-gradient(180deg,#ffffff 0%,#f8fbff 100%); box-shadow:0 8px 18px rgba(15,53,96,0.08), inset 0 1px 0 rgba(255,255,255,0.90); padding:12px; display:flex; flex-direction:column; gap:8px; min-height:150px; transform:translateZ(0); transition:transform .2s ease, box-shadow .2s ease, border-color .2s ease; }
            .dash-card::before { content:''; position:absolute; inset:0; border-radius:inherit; background:linear-gradient(135deg, rgba(255,255,255,0.70), rgba(255,255,255,0.08)); z-index:-1; pointer-events:none; }
            .dash-card:hover { transform:translateY(-2px) scale(1.01); border-color:#c5dbf8; box-shadow:0 16px 26px rgba(15,53,96,0.16), inset 0 1px 0 rgba(255,255,255,0.95); }
            .dash-card h4 { margin:0; color:#0f3560; font-size:15px; }
            .dash-card p { margin:0; color:#5d7590; font-size:12px; line-height:1.35; }
            .dash-card .dash-chip { align-self:flex-start; font-size:11px; font-weight:700; color:#0f3560; background:#eaf3ff; border:1px solid #d0e2f9; border-radius:999px; padding:3px 8px; }
            .dash-card button { margin-top:auto; border:none; border-radius:999px; padding:8px 12px; font-weight:700; cursor:pointer; background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); color:#fff; }
            .dash-card button[disabled] { opacity:0.6; cursor:not-allowed; background:linear-gradient(180deg,#8aa6c5 0%,#5f7f9e 100%); }
            .main-omsaetning { display:none; margin-bottom:16px; }
            .omsaetning-shell { background:#fff; border:1px solid #d7e6fb; border-radius:14px; box-shadow:0 8px 20px rgba(15,53,96,0.10); padding:14px; }
            .omsaetning-head { display:flex; justify-content:space-between; align-items:flex-start; gap:10px; margin-bottom:12px; }
            .omsaetning-head h3 { margin:0; color:#0f3560; font-size:22px; }
            .omsaetning-head p { margin:4px 0 0 0; color:#4d6680; font-size:12px; }
            .omsaetning-head-actions { display:flex; align-items:center; gap:8px; flex-wrap:wrap; justify-content:flex-end; }
            .omsaetning-head-actions button { border:none; border-radius:999px; padding:8px 14px; font-weight:700; cursor:pointer; color:#fff; background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); }
            .omsaetning-head-actions select { border:1px solid #c9ddf8; border-radius:999px; padding:7px 10px; background:#fff; color:#0f3560; font-size:12px; font-weight:700; }
            .omsaetning-years { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:10px; }
            .omsaetning-year-note { margin:-2px 0 10px 0; color:#4f6d8c; font-size:12px; }
            .omsaetning-year-btn { border:1px solid #c9ddf8; background:linear-gradient(180deg,#fff 0%,#edf4ff 100%); color:#0f3560; border-radius:999px; padding:6px 10px; font-size:12px; font-weight:700; cursor:pointer; }
            .omsaetning-year-btn.active { background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); color:#fff; border-color:#0f3560; }
            .omsaetning-filters { display:grid; grid-template-columns:160px 160px minmax(260px,1fr) minmax(260px,1fr) 220px; gap:10px; align-items:end; }
            .omsaetning-field label { display:block; font-size:12px; color:#3f5875; font-weight:700; margin-bottom:4px; }
            .omsaetning-field input, .omsaetning-field select { width:100%; border:1px solid #cfe0f7; border-radius:8px; padding:8px 10px; font-size:13px; }
            .omsaetning-threshold-grid { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
            .omsaetning-accounts-head { display:flex; align-items:center; justify-content:space-between; gap:8px; margin-bottom:6px; }
            .omsaetning-accounts-toggle { border:1px solid #c6dcf8; border-radius:999px; background:#fff; color:#0f3560; padding:5px 10px; cursor:pointer; font-size:11px; font-weight:700; }
            .omsaetning-accounts-summary { font-size:11px; color:#4f6d8c; font-weight:700; }
            .omsaetning-accounts-active { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:6px; }
            .omsaetning-accounts-active .chip { display:inline-flex; align-items:center; border:1px solid #c9ddf8; border-radius:999px; background:#eef5ff; color:#0f3560; padding:3px 8px; font-size:10px; font-weight:700; }
            .omsaetning-accounts-active .chip.more { background:#f3f8ff; color:#496989; border-color:#d8e7f8; }
            .omsaetning-accounts-panel { border:1px solid #d3e4f8; border-radius:10px; background:#f9fcff; padding:8px; max-height:220px; overflow:auto; }
            .omsaetning-accounts-toolbar { display:flex; gap:6px; margin-bottom:8px; }
            .omsaetning-accounts-toolbar button { border:1px solid #c6dcf8; border-radius:999px; background:#fff; color:#0f3560; padding:4px 9px; cursor:pointer; font-size:11px; font-weight:700; }
            .omsaetning-accounts-list { display:flex; flex-direction:column; gap:4px; }
            .omsaetning-account-item { display:flex; align-items:center; gap:6px; padding:4px 6px; border-radius:6px; }
            .omsaetning-account-item:hover { background:#ecf4ff; }
            .omsaetning-account-item input { width:15px; height:15px; }
            .omsaetning-account-item span { font-size:12px; color:#244766; }
            .omsaetning-customer-results { border:1px solid #d3e4f8; border-radius:10px; background:#f9fcff; max-height:176px; overflow:auto; margin-top:6px; }
            .omsaetning-customer-empty { padding:7px 9px; font-size:12px; color:#6b7f95; }
            .omsaetning-customer-item { display:flex; align-items:center; justify-content:space-between; gap:8px; padding:7px 9px; border-bottom:1px solid #e3eefb; cursor:pointer; }
            .omsaetning-customer-item:last-child { border-bottom:none; }
            .omsaetning-customer-item:hover { background:#ecf4ff; }
            .omsaetning-customer-item .meta { display:flex; flex-direction:column; gap:2px; min-width:0; }
            .omsaetning-customer-item .meta strong { font-size:12px; color:#214867; }
            .omsaetning-customer-item .meta span { font-size:11px; color:#5f7892; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:320px; }
            .omsaetning-customer-item .pick { border:none; border-radius:999px; padding:4px 9px; font-size:11px; font-weight:700; cursor:pointer; color:#fff; background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); }
            .omsaetning-customer-item .pick.remove { background:linear-gradient(180deg,#8aa6c5 0%,#5f7f9e 100%); }
            .omsaetning-selected-customers { display:flex; flex-wrap:wrap; gap:6px; margin-top:8px; }
            .omsaetning-selected-chip { display:inline-flex; align-items:center; gap:6px; border:1px solid #c9ddf8; border-radius:999px; background:#eef5ff; color:#0f3560; padding:4px 10px; font-size:11px; font-weight:700; }
            .omsaetning-selected-chip button { border:none; background:transparent; color:#0f3560; font-size:12px; cursor:pointer; font-weight:800; }
            .omsaetning-customer-mode { margin-top:6px; font-size:11px; color:#4f6d8c; }
            .omsaetning-customer-thresholds { margin-top:6px; font-size:11px; color:#335170; display:flex; flex-direction:column; gap:3px; }
            .omsaetning-customer-threshold-row { display:flex; gap:6px; align-items:center; }
            .omsaetning-customer-threshold-row .cust { font-weight:700; color:#214867; }
            .omsaetning-customer-threshold-row .thr { color:#496989; }
            .omsaetning-persist-overlay { position:fixed; inset:0; z-index:2200; background:rgba(10,22,40,0.45); display:flex; align-items:center; justify-content:center; padding:16px; }
            .omsaetning-persist-dialog { width:min(560px, 100%); border-radius:12px; border:1px solid #c7daef; background:#ffffff; box-shadow:0 16px 42px rgba(10,35,65,0.28); overflow:hidden; }
            .omsaetning-persist-head { padding:12px 14px; background:#f1f7ff; border-bottom:1px solid #dbe8f9; }
            .omsaetning-persist-head h4 { margin:0; font-size:14px; color:#1e4768; }
            .omsaetning-persist-body { padding:12px 14px; font-size:12px; color:#355675; display:flex; flex-direction:column; gap:10px; }
            .omsaetning-persist-customer-list { max-height:130px; overflow:auto; border:1px solid #dbe8f9; border-radius:8px; background:#f8fbff; padding:8px 10px; font-size:11px; line-height:1.5; }
            .omsaetning-persist-customer-option { width:100%; border:1px solid #d5e6fa; border-radius:8px; background:#ffffff; color:#214867; padding:7px 9px; margin:0 0 6px 0; text-align:left; font-size:12px; font-weight:600; cursor:pointer; }
            .omsaetning-persist-customer-option:last-child { margin-bottom:0; }
            .omsaetning-persist-customer-option:hover { background:#eef6ff; border-color:#b9d3f1; }
            .omsaetning-persist-customer-option.active { background:#dfeeff; border-color:#1565c0; color:#0f3560; box-shadow:inset 0 0 0 1px rgba(21,101,192,0.20); }
            .omsaetning-persist-thr { font-weight:700; color:#214867; }
            .omsaetning-persist-pick { font-size:12px; color:#355675; }
            .omsaetning-persist-picked { font-weight:800; color:#0f3560; }
            .omsaetning-persist-actions { display:flex; gap:8px; flex-wrap:wrap; }
            .omsaetning-persist-actions button { border:none; border-radius:999px; padding:8px 12px; font-weight:700; cursor:pointer; }
            .omsaetning-persist-actions .primary { color:#fff; background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); }
            .omsaetning-persist-actions .ghost { color:#244a6d; background:#eaf3ff; border:1px solid #c7daef; }
            .omsaetning-persist-actions .danger { color:#7a1a1a; background:#ffecec; border:1px solid #f2bbbb; }
            .omsaetning-actions button { border:none; border-radius:999px; padding:8px 14px; font-weight:700; cursor:pointer; color:#fff; background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); }
            .omsaetning-kpis { margin-top:12px; display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:10px; }
            .omsaetning-kpi { border:1px solid #d6e7fb; border-radius:10px; background:#f8fbff; padding:10px; }
            .omsaetning-kpi .lbl { font-size:11px; font-weight:700; color:#4f6d8c; text-transform:uppercase; letter-spacing:0.03em; }
            .omsaetning-kpi .val { margin-top:4px; font-size:20px; font-weight:800; color:#0f3560; }
            .omsaetning-charts { margin-top:12px; display:grid; grid-template-columns:1fr; gap:10px; }
            .omsaetning-chart-card { border:1px solid #dbe8f9; border-radius:10px; background:linear-gradient(180deg,#ffffff 0%,#f6faff 100%); overflow:hidden; }
            .omsaetning-chart-head { display:flex; justify-content:space-between; align-items:center; gap:8px; padding:8px 10px; border-bottom:1px solid #dbe8f9; background:#f4f9ff; }
            .omsaetning-chart-title { font-size:12px; font-weight:800; color:#2f5475; }
            .omsaetning-chart-sub { font-size:11px; color:#5f7892; }
            .omsaetning-chart-body { padding:8px; overflow:auto; }
            .omsaetning-chart-svg { width:100%; min-width:680px; height:260px; display:block; }
            .ordreindgang-chart-svg { height:430px; min-height:400px; }
            .ordreindgang-legend-row { margin:0 0 8px 0; }
            .ordreindgang-legend-row .omsaetning-legend-item { font-size:12px; font-weight:700; }
            .omsaetning-legend { display:flex; flex-wrap:wrap; gap:6px 10px; margin-top:8px; }
            .omsaetning-legend-item { display:inline-flex; align-items:center; gap:6px; font-size:11px; color:#355675; }
            .omsaetning-legend-swatch { width:10px; height:10px; border-radius:3px; border:1px solid rgba(0,0,0,0.16); }
            .omsaetning-table-card { margin-top:12px; border:1px solid #dbe8f9; border-radius:10px; background:#fff; overflow:hidden; }
            .omsaetning-table-title { padding:8px 10px; font-size:12px; font-weight:700; color:#2f5475; background:#f4f9ff; border-bottom:1px solid #dbe8f9; }
            .omsaetning-table-title-row { display:flex; align-items:center; justify-content:space-between; gap:8px; }
            .omsaetning-collapse-btn { border:1px solid #bcd2eb; border-radius:999px; background:#ffffff; color:#1e4768; font-size:11px; font-weight:700; padding:5px 10px; cursor:pointer; }
            .omsaetning-collapse-btn:hover { background:#ebf4ff; }
            .omsaetning-table-wrap { margin-top:12px; overflow:auto; border:1px solid #dbe8f9; border-radius:10px; }
            .omsaetning-table { width:100%; border-collapse:collapse; min-width:760px; font-size:12px; }
            .omsaetning-table th { background:#1565c0; color:#fff; text-align:left; padding:8px 10px; position:static; top:auto; z-index:1; }
            .omsaetning-table td { padding:7px 10px; border-bottom:1px solid #e6eef9; }
            .ordreindgang-weekly-table { min-width:980px; table-layout:fixed; }
            .ordreindgang-weekly-table th:nth-child(n+2), .ordreindgang-weekly-table td:nth-child(n+2) { text-align:right; font-variant-numeric:tabular-nums; }
            .ordreindgang-weekly-table tbody tr:nth-child(even) { background:#f8fbff; }
            .ordreindgang-toggle { display:flex; align-items:center; gap:8px; min-height:34px; padding:6px 0; color:#244a6d; font-weight:600; }
            .ordreindgang-toggle input[type="checkbox"] { width:16px; height:16px; accent-color:#1565c0; }
            .omsaetning-cell-right { text-align:right; }
            .omsaetning-status { display:inline-flex; align-items:center; border-radius:999px; padding:2px 8px; font-size:11px; font-weight:700; }
            .omsaetning-status.good { color:#1b5e20; background:#e8f5e9; border:1px solid #a5d6a7; }
            .omsaetning-status.mid { color:#8d6e00; background:#fff8e1; border:1px solid #ffe082; }
            .omsaetning-status.low { color:#b71c1c; background:#ffebee; border:1px solid #ef9a9a; }
            .omsaetning-gauge-wrap { min-width:280px; }
            .omsaetning-gauge-meta { display:flex; justify-content:space-between; gap:8px; font-size:11px; color:#4f6d8c; margin-bottom:4px; }
            .omsaetning-gauge-track { position:relative; height:10px; border-radius:999px; background:linear-gradient(90deg,#ffe5e5 0%, #fff2cc 33%, #e8f5e9 66%, #d9eefb 100%); border:1px solid #d6e7fb; overflow:hidden; }
            .omsaetning-gauge-fill { position:absolute; top:1px; height:6px; border-radius:999px; }
            .omsaetning-gauge-fill.pos { background:#2e7d32; }
            .omsaetning-gauge-fill.neg { background:#c62828; }
            .omsaetning-gauge-marker { position:absolute; top:-2px; width:2px; height:14px; background:#355675; opacity:0.65; }
            .omsaetning-gauge-point { position:absolute; top:-3px; width:8px; height:8px; margin-left:-4px; border-radius:50%; background:#0f3560; box-shadow:0 0 0 2px rgba(255,255,255,0.95); }
            .omsaetning-gauge-legend { display:flex; justify-content:space-between; font-size:10px; color:#6a829b; margin-top:3px; }
            .omsaetning-gauge-delta { margin-top:4px; font-size:11px; color:#355675; }
            .omsaetning-gauge-delta strong { color:#0f3560; }
            .omsaetning-empty { margin-top:10px; padding:10px; border:1px dashed #c7daef; border-radius:8px; color:#4f6d8c; background:#f8fbff; }
            #mainWorkspace { display:none; }
            .warning-flag { display:inline-flex; align-items:center; justify-content:center; margin-left:6px; font-size:14px; line-height:1; cursor:help; vertical-align:middle; }
            .allocation-flag { display:inline-flex; align-items:center; justify-content:center; margin-left:4px; color:#b26a00; font-size:16px; font-weight:700; line-height:1; cursor:help; vertical-align:middle; }
            .invoice-status-banner { margin: 0 0 10px 0; padding: 8px 10px; border-radius: 6px; font-size: 13px; font-weight: 600; }
            .invoice-status-banner.ok { background: #e8f5e9; color: #1b5e20; border: 1px solid #c8e6c9; }
            .invoice-status-banner.warn { background: #fff8e1; color: #8d6e00; border: 1px solid #ffe082; }
            .manual-modal-body { display:grid; gap:12px; }
            .manual-card { border:1px solid #d8e6fa; border-radius:10px; background:#f8fbff; padding:10px 12px; }
            .manual-card h4 { margin:0 0 6px 0; color:#0f3560; }
            .manual-card p { margin:0 0 6px 0; color:#355675; font-size:13px; }
            .manual-card ul { margin:0; padding-left:18px; color:#23384f; font-size:13px; }
            .manual-meta { font-size:12px; color:#4f6d8c; }
        </style>
    </head>
    <body>
        <div id="accessGateOverlay" class="access-gate-overlay" style="display:flex;">
            <div class="access-gate-box">
                <div class="access-gate-brand">
                    <h3>Login Dashboard</h3>
                </div>
                <p>Indtast brugernavn og kode for at åbne dashboard og moduler.</p>
                <div class="access-gate-fields">
                    <div class="access-gate-field">
                        <label for="accessGateUserInput">Brugernavn</label>
                        <input id="accessGateUserInput" type="text" placeholder="fx Marco" autocomplete="off" />
                    </div>
                    <div class="access-gate-field">
                        <label for="accessGateInput">Kode</label>
                        <input id="accessGateInput" type="password" placeholder="Indtast kode" autocomplete="off" />
                    </div>
                </div>
                <div class="access-gate-row">
                    <button id="accessGateBtn" type="button" onclick="submitAccessCode()">Åbn dashboard</button>
                </div>
                <div id="accessGateError" class="access-gate-error"></div>
            </div>
        </div>
        <div id="sideMenuOverlay" class="side-menu-overlay" onclick="closeSideMenu(event)">
            <aside class="side-menu-drawer" onclick="event.stopPropagation()">
                <div class="side-menu-header">
                    <span class="side-menu-title">Navigation</span>
                    <button class="side-menu-close" type="button" onclick="closeSideMenu()">×</button>
                </div>
                <div class="side-menu-content">
                    <section class="side-menu-section">
                        <h4>Login</h4>
                        <p>Samme adgangskode som startskærm.</p>
                        <div class="side-menu-login-row">
                            <input id="sideMenuUserInput" type="text" placeholder="Navn (fx Marco)" autocomplete="off" />
                        </div>
                        <div class="side-menu-login-row">
                            <input id="sideMenuLoginInput" type="password" placeholder="Kode" autocomplete="off" />
                            <button id="sideMenuLoginBtn" type="button" onclick="submitAccessCodeFromSideMenu()">Åbn</button>
                        </div>
                        <div id="sideMenuAuthStatus" class="side-menu-auth-status">Ikke logget ind.</div>
                    </section>

                    <section class="side-menu-section">
                        <h4>Moduler</h4>
                        <div class="side-menu-module-list">
                            <button type="button" onclick="navigateFromSideMenu('dashboard')">🏠 Dashboard</button>
                            <button type="button" onclick="navigateFromSideMenu('efterkalk')">Efterkalkulation</button>
                            <button type="button" onclick="navigateFromSideMenu('omsaetning')">Omsætning</button>
                            <button type="button" onclick="navigateFromSideMenu('ordreindgang')">Ordreindgang</button>
                            <button type="button" disabled>Faktura - Kommer snart</button>
                            <button type="button" disabled>Ordreoversigt - Kommer snart</button>
                            <button type="button" onclick="window.location.href='/assets/bom-workspace-v2.html'">📊 BOMe+ Beregner</button>
                            <button type="button" disabled>APV - Kommer snart</button>
                            <button type="button" onclick="openModule('belastning')">Belastning</button>
                            <button type="button" onclick="navigateFromSideMenu('personalehåndbog')">Personalehåndbog</button>
                            <button type="button" onclick="navigateFromSideMenu('brugermanual')">Brugermanual</button>
                        </div>
                    </section>

                    <section class="side-menu-section">
                        <h4>Session</h4>
                        <div class="side-menu-actions">
                            <button id="sideMenuLogoutBtn" class="logout" type="button" onclick="logoutFromSideMenu()" disabled>Log ud</button>
                        </div>
                    </section>
                </div>
            </aside>
        </div>
        <div class="header-banner-wrapper">
            <div class="header-left-controls">
                <button id="menuBtn" class="header-nav-btn" onclick="toggleSideMenu()" title="Åbn menu">☰</button>
                <button id="homeBtn" class="header-nav-btn" onclick="goToDashboard()" title="Tilbage til dashboard">🏠</button>
            </div>
            <div class="header-brand">
                <img class="header-brand-logo" src="/assets/brand/logo-gantech.png" alt="Gantech logo" />
                <span class="header-brand-text">${APP_VERSION}</span>
            </div>
            <div id="warmupBarWrap" title="Forberegner ordredata i baggrunden">
                <div id="warmupBarBg"><div id="warmupBarFill"></div></div>
                <span id="warmupBarText">Forberegner...</span>
            </div>
            <span class="header-user-greeting" id="headerUserGreeting">Hej, Bruger</span>
        </div>
        <div class="container main-dashboard" id="mainDashboard">
            <section class="dashboard-shell">
                <div class="dashboard-head">
                    <h2>Gantech Operations Hub</h2>
                    <p>Vælg makrokategori og modul. Salg, Produktion og HR er klar til at blive udbygget.</p>
                    <div id="dashboardUpdateNotice" class="dashboard-update-notice" aria-live="polite">
                        <div class="dashboard-update-copy">
                            <strong id="dashboardUpdateTitle">Opdateringsstatus</strong>
                            <div id="dashboardUpdateText">Tjekker efter opdatering...</div>
                        </div>
                        <div class="dashboard-update-actions">
                            <button id="dashboardUpdateCheckBtn" type="button" onclick="checkDesktopUpdateNow()">Tjek nu</button>
                            <button id="dashboardUpdateInstallBtn" type="button" class="install" onclick="installDesktopUpdateNow()" style="display:none;">Installer nu</button>
                            <button id="dashboardClearCacheBtn" type="button" onclick="clearAppCache()">Ryd Efterkalk cache</button>
                        </div>
                    </div>
                    <div id="dashboardWarmupNotice" class="dashboard-warmup-notice hidden" aria-live="polite">
                        <div class="dashboard-warmup-copy">
                            <strong>Efterkalk warmup</strong>
                            <div id="dashboardWarmupText">Forbereder ordre-cache i baggrunden...</div>
                            <div id="dashboardWarmupMeta" class="dashboard-warmup-meta">Du kan bruge andre moduler imens.</div>
                        </div>
                        <div class="dashboard-warmup-progress">
                            <div class="dashboard-warmup-track"><div id="dashboardWarmupFill"></div></div>
                            <span id="dashboardWarmupPct" class="dashboard-warmup-pct">0%</span>
                        </div>
                    </div>
                </div>
                <div class="dashboard-grid">
                    <section class="dashboard-category">
                        <div class="dashboard-category-head">
                            <h3>Salg</h3>
                            <span>Ordre, kunde og faktura</span>
                        </div>
                        <div class="dashboard-category-grid">
                            <article class="dash-card">
                                <span class="dash-chip">Aktiv</span>
                                <h4>Efterkalkulation</h4>
                                <p>Ordreliste, kost, margin, produktion og rapportvisning.</p>
                                <button onclick="openModule('efterkalk')">Åbn Efterkalk</button>
                            </article>
                            <article class="dash-card">
                                <span class="dash-chip">Aktiv</span>
                                <h4>Omsætning</h4>
                                <p>Total omsætning, KPI-overblik og udvikling pr. periode/kunde.</p>
                                <button onclick="openModule('omsaetning')">Åbn Omsætning</button>
                            </article>
                            <article class="dash-card">
                                <span class="dash-chip">Aktiv</span>
                                <h4>Ordreindgang</h4>
                                <p>Ordreindgang fra SSRS: budget, ordre, tilbud og udvikling pr. uge/periode.</p>
                                <button onclick="openModule('ordreindgang')">Åbn Ordreindgang</button>
                            </article>
                            <article class="dash-card">
                                <span class="dash-chip">Planlagt</span>
                                <h4>Faktura</h4>
                                <p>Fakturastatus, kreditnota og opfølgning på åbne poster.</p>
                                <button type="button" disabled>Kommer snart</button>
                            </article>
                        </div>
                    </section>

                    <section class="dashboard-category">
                        <div class="dashboard-category-head">
                            <h3>Produktion</h3>
                            <span>Planlægning, BOM og ordreflow</span>
                        </div>
                        <div class="dashboard-category-grid">
                            <article class="dash-card">
                                <span class="dash-chip">Planlagt</span>
                                <h4>Ordreoversigt</h4>
                                <p>Samlet status for produktionsordrer, levering og kapacitet.</p>
                                <button type="button" disabled>Kommer snart</button>
                            </article>
                            <article class="dash-card">
                                <span class="dash-chip">Planlagt</span>
                                <h4>Bom</h4>
                                <p>Styklister, komponenter og versionering med sporbarhed.</p>
                                <button type="button" disabled>Kommer snart</button>
                            </article>
                            <article class="dash-card">
                                <span class="dash-chip">Aktiv</span>
                                <h4>Belastning</h4>
                                <p>Kapacitetsbelastning, ressourcer, ordreflyt og planlægningsudsving.</p>
                                <button onclick="openModule('belastning')">Åbn Belastning</button>
                            </article>
                        </div>
                    </section>

                    <section class="dashboard-category">
                        <div class="dashboard-category-head">
                            <h3>HR</h3>
                            <span>Arbejdsmiljø og medarbejderdata</span>
                        </div>
                        <div class="dashboard-category-grid">
                            <article class="dash-card">
                                <span class="dash-chip">Planlagt</span>
                                <h4>APV</h4>
                                <p>Arbejdsmiljøvurdering, opgaver, frister og opfølgning.</p>
                                <button type="button" disabled>Kommer snart</button>
                            </article>
                            <article class="dash-card">
                                <span class="dash-chip">Aktiv</span>
                                <h4>Personalehåndbog</h4>
                                <p>Intranet personalehåndbog med søgning.</p>
                                <button onclick="openPersonalehåndbog()">Åbn Personalehåndbog</button>
                            </article>
                            <article class="dash-card">
                                <span class="dash-chip">Planlagt</span>
                                <h4>Kvalitetsledelsessystem</h4>
                                <p>SharePoint startside for QMS procedurer og dokumenter.</p>
                                <button type="button" disabled>Kommer snart</button>
                            </article>
                        </div>
                    </section>
                </div>
            </section>
        </div>

        <div class="container main-omsaetning" id="mainOmsaetning">
            <section class="omsaetning-shell">
                <div class="omsaetning-head">
                    <div>
                        <h3>Omsætning</h3>
                        <p>SSRS-baseret oversigt fra AcTr, AcPr og Ac (kontogruppe 10_Omsætning).</p>
                    </div>
                    <div class="omsaetning-head-actions">
                        <button id="omsaetningLoadBtn" onclick="loadOmsaetningSummary({ persistThresholdsOnUpdate: true, forceRefresh: true })">Opdater</button>
                        <select id="omsaetningPrintOrientation" title="Print-layout">
                            <option value="auto" selected>Layout: Auto</option>
                            <option value="portrait">Layout: Stående</option>
                            <option value="landscape">Layout: Liggende</option>
                        </select>
                        <button id="omsaetningPrintBtn" onclick="printOmsaetningReport()">Print rapport</button>
                    </div>
                </div>
                <div id="omsaetningYears" class="omsaetning-years"></div>
                <p id="omsaetningYearNote" class="omsaetning-year-note">Regnskabsår (jul-jun). Vælg flere år for samlet periode.</p>
                <div class="omsaetning-filters">
                    <div class="omsaetning-field">
                        <label for="omsaetningFraMonth">Fra måned</label>
                        <input id="omsaetningFraMonth" type="month" onchange="scheduleOmsaetningAutoReload()" />
                    </div>
                    <div class="omsaetning-field">
                        <label for="omsaetningTilMonth">Til måned</label>
                        <input id="omsaetningTilMonth" type="month" onchange="scheduleOmsaetningAutoReload()" />
                    </div>
                    <div class="omsaetning-field omsaetning-accounts-field">
                        <label for="omsaetningAccountSearch">Kontoer (multi)</label>
                        <div class="omsaetning-accounts-head">
                            <button id="omsaetningAccountsToggleBtn" class="omsaetning-accounts-toggle" type="button" onclick="toggleOmsaetningAccountsPanel()">Vis konti</button>
                            <span id="omsaetningAccountsSummary" class="omsaetning-accounts-summary">0/0 valgt</span>
                        </div>
                        <div id="omsaetningAccountsActive" class="omsaetning-accounts-active"></div>
                        <input id="omsaetningAccountSearch" type="text" placeholder="Søg konto/navn..." oninput="filterOmsaetningAccounts()" style="display:none;" />
                        <div id="omsaetningAccountsPanel" class="omsaetning-accounts-panel" style="display:none;">
                            <div class="omsaetning-accounts-toolbar">
                                <button type="button" onclick="setAllOmsaetningAccounts(true)">Alle</button>
                                <button type="button" onclick="setAllOmsaetningAccounts(false)">Ingen</button>
                            </div>
                            <div id="omsaetningAccountsList" class="omsaetning-accounts-list"></div>
                        </div>
                    </div>
                    <div class="omsaetning-field omsaetning-customer-field">
                        <label for="omsaetningCustomerSearch">Kunde (prefix)</label>
                        <input id="omsaetningCustomerSearch" type="text" placeholder="Søg kunde fx logitr..." oninput="scheduleOmsaetningCustomerSearch()" />
                        <div id="omsaetningCustomerResults" class="omsaetning-customer-results">
                            <div class="omsaetning-customer-empty">Ingen kunde valgt: viser normal visning for valgte år og konti.</div>
                        </div>
                        <div id="omsaetningSelectedCustomers" class="omsaetning-selected-customers"></div>
                        <div id="omsaetningCustomerMode" class="omsaetning-customer-mode">Ingen kunde valgt: viser normal visning for valgte år og konti.</div>
                        <div id="omsaetningCustomerThresholds" class="omsaetning-customer-thresholds"></div>
                    </div>
                    <div class="omsaetning-field omsaetning-threshold-field">
                        <label>Tærskler (Mio)</label>
                        <div class="omsaetning-threshold-grid">
                            <input id="omsaetningWarnThreshold" type="number" step="0.1" min="0" value="3.0" title="Under denne værdi markeres lav" />
                            <input id="omsaetningGoodThreshold" type="number" step="0.1" min="0" value="5.0" title="Over denne værdi markeres god" />
                        </div>
                    </div>
                </div>
                <div class="omsaetning-kpis">
                    <div class="omsaetning-kpi"><div class="lbl">Omsætning (Mio)</div><div class="val" id="omsaetningTotalMio">-</div></div>
                    <div class="omsaetning-kpi"><div class="lbl">Rækker</div><div class="val" id="omsaetningRowsCount">-</div></div>
                    <div class="omsaetning-kpi"><div class="lbl">Perioder</div><div class="val" id="omsaetningPeriodsCount">-</div></div>
                </div>
                <div id="omsaetningChartsWrap" class="omsaetning-charts" style="display:none;">
                    <section class="omsaetning-chart-card">
                        <header class="omsaetning-chart-head">
                            <span id="omsaetningStackedTitle" class="omsaetning-chart-title">Omsætning pr. måned (stacked pr. konto)</span>
                            <span class="omsaetning-chart-sub">Mio DKK</span>
                        </header>
                        <div class="omsaetning-chart-body">
                            <svg id="omsaetningStackedChart" class="omsaetning-chart-svg" aria-label="Omsætning stacked chart"></svg>
                            <div id="omsaetningLegend" class="omsaetning-legend"></div>
                        </div>
                    </section>
                    <section class="omsaetning-chart-card">
                        <header class="omsaetning-chart-head">
                            <span class="omsaetning-chart-title">Trend total omsætning</span>
                            <span class="omsaetning-chart-sub">Månedlig udvikling</span>
                        </header>
                        <div class="omsaetning-chart-body">
                            <svg id="omsaetningTrendChart" class="omsaetning-chart-svg" aria-label="Omsætning trend chart"></svg>
                        </div>
                    </section>
                </div>
                <div id="omsaetningThresholdWrap" class="omsaetning-table-card" style="display:none;">
                    <div class="omsaetning-table-title">Månedstabel med tærskler</div>
                    <div id="omsaetningThresholdTable" class="omsaetning-table-wrap" style="margin-top:0;border:none;border-radius:0;"></div>
                </div>
                <div id="omsaetningDetailsWrap" class="omsaetning-table-card" style="display:none;">
                    <div class="omsaetning-table-title omsaetning-table-title-row">
                        <span>Måned/Kunde detaljer (tekst)</span>
                        <button id="omsaetningDetailsToggleBtn" class="omsaetning-collapse-btn" type="button" onclick="toggleOmsaetningDetails()">Vis detaljer</button>
                    </div>
                    <div id="omsaetningTableWrap" class="omsaetning-table-wrap" style="margin-top:0;border:none;border-radius:0;display:none;"></div>
                </div>
                <div id="omsaetningEmpty" class="omsaetning-empty">Vælg perioder og konti, og tryk Opdater.</div>
            </section>
        </div>

        <div class="container main-omsaetning" id="mainOrdreindgang">
            <section class="omsaetning-shell">
                <div class="omsaetning-head">
                    <div>
                        <h3>Ordreindgang</h3>
                        <p>SSRS-baseret oversigt over ordre og tilbud pr. uge.</p>
                    </div>
                    <div class="omsaetning-head-actions">
                        <button id="ordreindgangLoadBtn" onclick="loadOrdreindgangSummary({ forceRefresh: true })">Opdater</button>
                        <select id="ordreindgangPrintOrientation" title="Print-layout">
                            <option value="auto" selected>Layout: Auto</option>
                            <option value="portrait">Layout: Stående</option>
                            <option value="landscape">Layout: Liggende</option>
                        </select>
                        <button id="ordreindgangPrintBtn" onclick="printOrdreindgangReport()">Print rapport</button>
                    </div>
                </div>
                <div class="omsaetning-filters" style="grid-template-columns:180px 180px 140px minmax(180px,1fr);">
                    <div class="omsaetning-field">
                        <label for="ordreindgangFraWeek">Fra uge (YYYYWW)</label>
                        <input id="ordreindgangFraWeek" type="text" maxlength="6" placeholder="202601" onchange="scheduleOrdreindgangAutoReload()" />
                    </div>
                    <div class="omsaetning-field">
                        <label for="ordreindgangTilWeek">Til uge (YYYYWW)</label>
                        <input id="ordreindgangTilWeek" type="text" maxlength="6" placeholder="202612" onchange="scheduleOrdreindgangAutoReload()" />
                    </div>
                    <div class="omsaetning-field">
                        <label for="ordreindgangShowTilbud">Tilbud</label>
                        <label class="ordreindgang-toggle" for="ordreindgangShowTilbud">
                            <input id="ordreindgangShowTilbud" type="checkbox" onchange="renderOrdreindgangFromLastPayload()" />
                            <span>Vis tilbud i graf og tabel</span>
                        </label>
                    </div>
                    <div class="omsaetning-field">
                        <label>Status</label>
                        <div id="ordreindgangStatus" class="omsaetning-customer-mode">Vælg ugeperiode og tryk Opdater.</div>
                    </div>
                </div>
                <div class="omsaetning-kpis">
                    <div class="omsaetning-kpi"><div class="lbl">Total Ordre</div><div class="val" id="ordreindgangTotalOrd">-</div></div>
                    <div class="omsaetning-kpi"><div class="lbl">Total Tilbud</div><div class="val" id="ordreindgangTotalTilbud">-</div></div>
                    <div class="omsaetning-kpi"><div class="lbl">Gns. Ordre</div><div class="val" id="ordreindgangAvgOrd">-</div></div>
                    <div class="omsaetning-kpi"><div class="lbl">Tilbud → Ordre</div><div class="val" id="ordreindgangConv">-</div></div>
                </div>
                <div id="ordreindgangChartsWrap" class="omsaetning-charts" style="display:none;">
                    <section class="omsaetning-chart-card" style="grid-column:1 / -1;">
                        <header class="omsaetning-chart-head">
                            <span class="omsaetning-chart-title">Ugeudvikling (Ordre / Tilbud / Budget)</span>
                            <span class="omsaetning-chart-sub">DKK</span>
                        </header>
                        <div class="omsaetning-chart-body">
                            <div id="ordreindgangLegend" class="omsaetning-legend ordreindgang-legend-row"></div>
                            <svg id="ordreindgangTrendChart" class="omsaetning-chart-svg ordreindgang-chart-svg" aria-label="Ordreindgang trend chart"></svg>
                        </div>
                    </section>
                </div>
                <div id="ordreindgangWeeklyWrap" class="omsaetning-table-card" style="display:none;">
                    <div class="omsaetning-table-title omsaetning-table-title-row">
                        <span>Ugetabel</span>
                        <button id="ordreindgangWeeklyToggleBtn" class="omsaetning-collapse-btn" type="button" onclick="toggleOrdreindgangWeeklyTable()">Vis tabel</button>
                    </div>
                    <div id="ordreindgangWeeklyTable" class="omsaetning-table-wrap" style="margin-top:0;border:none;border-radius:0;display:none;"></div>
                </div>
                <div id="ordreindgangCustomersWrap" class="omsaetning-table-card" style="display:none;">
                    <div class="omsaetning-table-title omsaetning-table-title-row">
                        <span>Top kunder</span>
                        <button id="ordreindgangCustomersToggleBtn" class="omsaetning-collapse-btn" type="button" onclick="toggleOrdreindgangCustomersTable()">Vis tabel</button>
                    </div>
                    <div id="ordreindgangCustomersTable" class="omsaetning-table-wrap" style="margin-top:0;border:none;border-radius:0;display:none;"></div>
                </div>
                <div id="ordreindgangEmpty" class="omsaetning-empty">Vælg ugeperiode og tryk Opdater.</div>
            </section>
        </div>

        <div class="container main-omsaetning" id="mainBelastning">
            <section class="omsaetning-shell">
                <div class="omsaetning-head">
                    <div>
                        <h3>Belastning</h3>
                        <p>Grafisk kapacitetsbelastning med drillthrough til detaljer.</p>
                    </div>
                    <div class="omsaetning-head-actions">
                        <button id="belastningLoadBtn" onclick="loadBelastningGrafisk({ forceRefresh: true })">Opdater</button>
                    </div>
                </div>
                <div class="omsaetning-filters" style="grid-template-columns:150px 84px minmax(190px,1fr) minmax(150px,0.8fr) minmax(190px,1fr) minmax(180px,1fr);">
                    <div class="omsaetning-field">
                        <label for="belastningToDay">Start dato</label>
                        <input id="belastningToDay" type="date" onchange="scheduleBelastningAutoReload()" />
                    </div>
                    <div class="omsaetning-field">
                        <label for="belastningDage">Dage</label>
                        <input id="belastningDage" type="number" min="1" max="180" value="30" onchange="scheduleBelastningAutoReload()" />
                    </div>
                    <div class="omsaetning-field">
                        <label for="belastningResGr">Ressourcegrupper (kommasepareret)</label>
                        <input id="belastningResGr" type="text" placeholder="fx 11,21" onchange="scheduleBelastningAutoReload()" />
                    </div>
                    <div class="omsaetning-field">
                        <label for="belastningOrdre">Ordre (S/P)</label>
                        <input id="belastningOrdre" class="belastning-order-filter" type="text" inputmode="numeric" placeholder="fx 405374" oninput="scheduleBelastningAutoReload()" onchange="scheduleBelastningAutoReload()" />
                    </div>
                    <div class="omsaetning-field">
                        <label for="belastningKunde">Kunde</label>
                        <div class="belastning-kunde-wrap">
                            <input id="belastningKunde" class="belastning-order-filter" type="text" placeholder="fx Echolo"
                                autocomplete="off" spellcheck="false" autocorrect="off" autocapitalize="off"
                                oninput="scheduleBelastningAutoReload(); scheduleBelastningKundeSuggest()"
                                onchange="scheduleBelastningAutoReload()"
                                onfocus="scheduleBelastningKundeSuggest()"
                                onblur="hideBelastningKundeSuggestions()" />
                            <div id="belastningKundeDropdown" class="belastning-kunde-dropdown" style="display:none;"></div>
                        </div>
                    </div>
                    <div class="omsaetning-field">
                        <label>Status</label>
                        <div id="belastningStatus" class="omsaetning-customer-mode">Vælg filtre og tryk Opdater.</div>
                    </div>
                </div>
                <div class="belastning-toolbar-note">Klik en søjle i et kort for at åbne ordrer på den dag. Ordre- og kundefilter opdaterer både kort og detaljer.</div>

                <div id="belastningGrafiskWrap" class="omsaetning-charts" style="display:none;">
                    <section class="omsaetning-chart-card belastning-chart-card">
                        <header class="omsaetning-chart-head">
                            <span class="omsaetning-chart-title">Grafisk belastning</span>
                            <span class="omsaetning-chart-sub">Samlet visning. Klik på et kort for detaljer.</span>
                        </header>
                        <div class="omsaetning-chart-body">
                            <div id="belastningSvgCombined" class="belastning-svg-wrap"></div>
                            <div id="belastningBarsCombined" class="belastning-bars"></div>
                        </div>
                    </section>
                </div>

                <div id="belastningDetailWrap" class="omsaetning-table-card" style="display:none;">
                    <div class="omsaetning-table-title" id="belastningDetailTitle">Belastning detaljer</div>
                    <div id="belastningDetailSvg" class="belastning-svg-wrap" style="margin:0 10px 8px 10px;"></div>
                    <div id="belastningDetailTable" class="omsaetning-table-wrap" style="margin-top:0;border:none;border-radius:0;"></div>
                </div>

                <div id="belastningEmpty" class="omsaetning-empty">Tryk Opdater for at hente grafisk belastning.</div>
            </section>
        </div>

        <div class="container" id="mainWorkspace">
            <div class="search-box" id="searchBox">
                <button id="collapseToggleBtn" onclick="toggleSearchBox()" style="display:none;" title="Åbn søgefelt og filtre">▼ Søg</button>
                <input type="number" id="orderInput" placeholder="Indtast ordrenummer..." style="display:none;" />
                <button onclick="searchOrder()" title="Aabn detaljer for ordrenummeret" style="display:none;">Søg</button>
                <button id="refreshListBtn" class="list-toggle-btn" onclick="refreshOrderList()" title="Hent seneste ordreliste">Opdater liste</button>
                <button class="mode-btn" onclick="toggleMarginMode()" title="Skift hvordan margin beregnes i visningen">Skift marginberegning</button>
                <button id="listToggleBtn" class="list-toggle-btn" onclick="toggleOrderList()" title="Vis eller skjul kundelisten">Skjul kundeliste</button>
                <button id="clearCacheBtn" class="list-toggle-btn" onclick="clearAppCache()" style="background:#b71c1c !important;" title="DET TAGER LANG TID!!! Slet disk-cache og genindlaes data">Ryd cache</button>
                <select id="brugerFilterSelect" class="filter-select" onchange="setBrugerFilter()">
                    <option value="">Alle brugere</option>
                </select>
                <input type="text" id="customerFilterInput" class="filter-input" placeholder="Søg kunde i listen..." oninput="setOrderListFilter()" />
                <label class="order-value-filter-toggle" for="orderMinDkkEnabled">
                    <input type="checkbox" id="orderMinDkkEnabled" onchange="setOrderValueFilter()" />
                    <span>Skjul ordre &lt;</span>
                </label>
                <input type="number" id="orderMinDkkInput" class="order-value-filter-input" min="0" step="100" value="0" placeholder="DKK" oninput="setOrderValueFilter()" disabled />
                <button id="collapseExpandBtn" class="list-toggle-btn" onclick="toggleSearchBox()" style="margin-left:auto;" title="Skjul sogefelt og filtre">▲ Luk</button>
            </div>
            <div id="orderList"></div>
            <div id="result"></div>
        </div>

        <div id="summaryModal" class="modal-overlay" onclick="closeSummaryModal(event)">
            <div class="modal-box" onclick="event.stopPropagation()">
                <div class="modal-header">
                    <div class="modal-header-left">
                        <button id="summaryModalBackBtn" class="modal-back hidden" onclick="goSummaryModalBack()">←</button>
                        <h3 id="summaryModalTitle">Produktoversigt</h3>
                    </div>
                    <button class="modal-close" onclick="closeSummaryModal()">Luk</button>
                </div>
                <div class="modal-content-wrap">
                    <div id="summaryModalBody"></div>
                    <aside id="summaryImagePanel" class="summary-image-panel hidden"></aside>
                </div>
            </div>
        </div>

        <div id="imageLightbox" class="image-lightbox hidden" onclick="closeImageLightbox(event)">
            <div class="image-lightbox-dialog" onclick="event.stopPropagation()">
                <div class="image-lightbox-header">
                    <div id="imageLightboxTitle" class="image-lightbox-title">Billede</div>
                    <button class="image-lightbox-close" onclick="closeImageLightbox()">Luk</button>
                </div>
                <div class="image-lightbox-body">
                    <img id="imageLightboxImg" src="" alt="" />
                </div>
                <div id="imageLightboxPath" class="image-lightbox-path"></div>
            </div>
        </div>

        <div id="compactImageModal" class="compact-image-modal" onclick="closeCompactImageModal(event)">
            <div class="compact-image-dialog" onclick="event.stopPropagation()">
                <div class="compact-image-header">
                    <div>
                        <div id="compactImageTitle" class="compact-image-title">Billeder</div>
                        <div id="compactImageSubtitle" class="compact-image-subtitle"></div>
                    </div>
                    <button class="compact-image-close" onclick="closeCompactImageModal()">Luk</button>
                </div>
                <div id="compactImageBody" class="compact-image-body"></div>
            </div>
        </div>

        <div id="printPreviewOverlay" class="print-preview-overlay" role="dialog" aria-modal="true" aria-labelledby="printPreviewTitle" onclick="closePrintPreview(event)">
            <div class="print-preview-dialog" tabindex="-1" onclick="event.stopPropagation()">
                <div class="print-preview-header">
                    <div id="printPreviewTitle" class="print-preview-title">Forhåndsvisning</div>
                    <div class="print-preview-actions">
                        <button class="list-toggle-btn" onclick="closePrintPreview()">Luk</button>
                        <button class="list-toggle-btn" onclick="confirmPrintFromPreview()">Udskriv / PDF</button>
                    </div>
                </div>
                <div id="printPreviewBody" class="print-preview-body"></div>
            </div>
        </div>

        <div id="orderDetailModal" class="order-detail-modal-overlay" role="dialog" aria-modal="true" aria-labelledby="orderDetailModalTitle" onclick="closeOrderDetailModal(event)">
            <div class="order-detail-modal-shell" onclick="event.stopPropagation()">
                <div class="order-detail-modal-header">
                    <div class="order-detail-modal-title">
                        <strong id="orderDetailModalTitle">Ordre-rapport</strong>
                        <span id="orderDetailModalSubtitle">Manager-oversigt med produktion, cost og sporbarhed</span>
                    </div>
                    <div class="order-detail-modal-actions">
                            <button id="orderDetailModalBackBtn" class="list-toggle-btn" onclick="goBackFromReportToOrder()" style="display:none;">Tilbage til ordre</button>
                        <button class="list-toggle-btn" onclick="printOrderDetailReport()">Udskriv / PDF</button>
                        <button class="list-toggle-btn" onclick="closeOrderDetailModal()">Luk</button>
                    </div>
                </div>
                <div id="orderDetailModalBody" class="order-detail-modal-body"></div>
            </div>
        </div>

        <div id="personalehåndbogsModal" role="dialog" aria-modal="true" aria-label="Personalehåndbog">
            <div class="personalehåndbog-shell">
                <div class="personalehåndbog-header">
                    <span class="personalehåndbog-title">📖 Personalehåndbog</span>
                    <div class="personalehåndbog-search-wrap">
                        <input id="personalehåndbogsSearchInput" type="search" placeholder="Søg i personalehåndbog..." onkeydown="if(event.key==='Enter')searchPersonalehåndbog()" />
                        <button onclick="searchPersonalehåndbog()">Søg</button>
                    </div>
                    <button class="personalehåndbog-close" onclick="closePersonalehåndbog()" title="Luk">✕</button>
                </div>
                <div class="personalehåndbog-body">
                    <div class="personalehåndbog-results">
                        <div class="ph-results-header">
                            <span id="phResultsLabel">Resultater</span>
                            <button class="ph-reindex-btn" onclick="phReindex()" title="Genindekser sitet">↺ Genindekser</button>
                        </div>
                        <div id="phResultsList" class="ph-results-list">
                            <div class="ph-status-msg" id="phStatusMsg">Skriv en søgning og tryk Søg.</div>
                        </div>
                    </div>
                    <iframe id="personalehåndbogsIframe" class="personalehåndbog-iframe" src="" title="Personalehåndbog" sandbox="allow-same-origin allow-scripts allow-forms allow-popups"></iframe>
                </div>
            </div>
        </div>

        <div id="qmsModal" role="dialog" aria-modal="true" aria-label="Kvalitetsledelsessystem">
            <div class="qms-shell">
                <div class="qms-header">
                    <span class="qms-title">Kvalitetsledelsessystem</span>
                    <div class="qms-search-wrap">
                        <input id="qmsSearchInput" type="search" placeholder="Søg i QFP-sider..." onkeydown="if(event.key==='Enter')searchQmsPages()" />
                        <button onclick="searchQmsPages()">Søg</button>
                    </div>
                    <button class="qms-close" onclick="closeQmsModal()" title="Luk">✕</button>
                </div>
                <div class="qms-body">
                    <div class="qms-nav">
                        <div class="qms-nav-header">
                            <span id="qmsListLabel">QFP sider</span>
                            <div class="qms-nav-actions">
                                <button type="button" onclick="qmsCreateFolder()" title="Ny mappe">+ Mappe</button>
                                <button type="button" onclick="qmsCreateDocument()" title="Nyt dokument">+ Dokument</button>
                                <button type="button" onclick="toggleQmsEditMode()" id="qmsEditToggleBtn" title="Rediger valgt dokument">Rediger</button>
                            </div>
                        </div>
                        <div id="qmsList" class="qms-list"></div>
                    </div>
                    <div id="qmsView" class="qms-view">
                        <h3>Kvalitetsledelsessystem</h3>
                        <div class="qms-view-meta">Vælg en side i venstre menu for at se indhold.</div>
                        <div class="qms-view-content">Lokalt uddrag baseret på SharePoint-data.</div>
                    </div>
                </div>
            </div>
        </div>

        <div id="brugermanualModal" class="modal-overlay" onclick="closeBrugermanual(event)">
            <div class="modal-box" onclick="event.stopPropagation()">
                <div class="modal-header">
                    <div class="modal-header-left">
                        <h3>Brugermanual (kort)</h3>
                    </div>
                    <button class="modal-close" onclick="closeBrugermanual()">Luk</button>
                </div>
                <div id="brugermanualBody" class="manual-modal-body"></div>
            </div>
        </div>

        <div id="oversigtModal" class="oversigt-modal-overlay" role="dialog" aria-modal="true" aria-labelledby="oversigtModalTitle" onclick="closeOversigtModal(event)">
            <div class="oversigt-modal-shell" onclick="event.stopPropagation()">
                <div class="oversigt-modal-header">
                    <div class="oversigt-modal-title-wrap">
                        <strong id="oversigtModalTitle">Oversigt</strong>
                        <span id="oversigtModalSubtitle">Detaljeret produktionsanalyse</span>
                    </div>
                    <div class="oversigt-modal-actions">
                        <button class="list-toggle-btn" onclick="refreshActiveOversigtModal()">Opdater</button>
                        <button class="list-toggle-btn" onclick="closeOversigtModal()">Luk</button>
                    </div>
                </div>
                <div id="oversigtModalBody" class="oversigt-modal-body"></div>
            </div>
        </div>
        
        <script>
            function formatNumber(num) {
                const fixed = parseFloat(num).toFixed(2);
                const parts = fixed.split('.');
                const integerPart = parts[0];
                const decimalPart = parts[1];
                
                // Aggiungi punto come separatore migliaia da destra a sinistra
                let formatted = '';
                for (let i = integerPart.length - 1, count = 0; i >= 0; i--, count++) {
                    if (count > 0 && count % 3 === 0) {
                        formatted = '.' + formatted;
                    }
                    formatted = integerPart[i] + formatted;
                }
                
                return formatted + ',' + decimalPart;
            }

            function isLaserLProdNo(prodNo) {
                return String(prodNo || '').trim().toUpperCase().endsWith('L');
            }

            function isInvoiceTrackedProdNo(prodNo) {
                return String(prodNo || '').trim().toUpperCase().startsWith('U');
            }

            function shouldFilterChildSummary(prodTp4, prodNo, purcNo) {
                const displayKey = getDisplayProdTp4Key(prodTp4, prodNo, purcNo);
                return displayKey === '6' || displayKey === '9' || isInvoiceTrackedProdNo(prodNo);
            }

            function isProductionSummaryExcludedLine(line) {
                if (!line) return false;
                const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                return Number(line.LnNo || 0) === 1 || key === '0' || key === '3' || key === '5';
            }

            function getDisplayProdTp4Key(prodTp4, prodNo, purcNo) {
                const rawKey = (prodTp4 === null || prodTp4 === undefined) ? 'NA' : String(prodTp4);
                if (rawKey === '3') return '1';
                if (rawKey === '2' && !isLaserLProdNo(prodNo)) {
                    return Number(purcNo || 0) > 0 ? '9' : '4';
                }
                return rawKey;
            }

            function isExcludedOperationProdNo(prodNo) {
                const normalized = String(prodNo || '').trim().toUpperCase();
                return normalized === 'R1090' || normalized === 'R8200';
            }

            function escapeHtml(value) {
                return String(value || '')
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
            }

            function formatCount(value) {
                const num = Math.max(0, Math.trunc(Number(value) || 0));
                return new Intl.NumberFormat('da-DK', { maximumFractionDigits: 0 }).format(num);
            }

            function getResourceDisplayLabel(prodNo, descr) {
                const code = String(prodNo || '').trim();
                const name = String(descr || '').trim();
                if (!code) return name || '-';
                if (!name) return code;
                if (code.toUpperCase().startsWith('R')) {
                    return code + ' - ' + name;
                }
                return code;
            }

            function collectWarningMessages(item, fallbackText) {
                const unique = [];
                const pushValue = (value) => {
                    const chunks = String(value || '').split('|');
                    for (const chunk of chunks) {
                        const text = String(chunk || '').trim();
                        if (text && !unique.includes(text)) unique.push(text);
                    }
                };

                if (Array.isArray(item)) {
                    for (const entry of item) {
                        if (!entry) continue;
                        if (entry.WarningText) pushValue(entry.WarningText);
                        if (entry.warningText) pushValue(entry.warningText);
                    }
                } else if (item) {
                    if (item.WarningText) pushValue(item.WarningText);
                    if (item.warningText) pushValue(item.warningText);
                }

                if (unique.length === 0 && fallbackText) pushValue(fallbackText);
                return unique;
            }

            function getWarningIconMeta(message) {
                const text = String(message || '').trim().toLowerCase();
                if (text.includes('faktura') || text.includes('noinvo')) {
                    return { key: 'invoice', icon: '🧾' };
                }
                if (text.includes('tilknyttet produktionsordre') || text.includes('underliggende produktionsordre')) {
                    return { key: 'linked-order', icon: '🏭' };
                }
                if (text.includes('inkonsekvens') || text.includes('afvig')) {
                    return { key: 'consistency', icon: '⚠️' };
                }
                return { key: 'general', icon: '⚠️' };
            }

            function getWarningFlagHtml(item, fallbackText) {
                const hasWarning = Array.isArray(item)
                    ? item.some(entry => entry && (entry.HasWarning || entry.hasWarnings || entry.WarningText || entry.warningText))
                    : Boolean(item && (item.HasWarning || item.hasWarnings || item.WarningText || item.warningText));
                if (!hasWarning) return '';

                const messages = collectWarningMessages(item, fallbackText);
                if (messages.length === 0) return '';

                const grouped = new Map();
                for (const message of messages) {
                    const meta = getWarningIconMeta(message);
                    if (!grouped.has(meta.key)) {
                        grouped.set(meta.key, { icon: meta.icon, messages: [] });
                    }
                    grouped.get(meta.key).messages.push(message);
                }

                return Array.from(grouped.values()).map(group => {
                    const title = escapeHtml(group.messages.join(' | '));
                    return ' <span class="warning-flag" title="' + title + '">' + group.icon + '</span>';
                }).join('');
            }

            function getTimeAdjustmentFlagHtml(item, fallbackText) {
                if (!item || (!item.UsesEstimatedOperationTime && !item.hasEstimatedOperationTime)) return '';
                const title = escapeHtml(item.EstimatedTimeText || item.estimatedTimeText || fallbackText || 'Færdigmeldt minutter var 0 og er beregnet ud fra Stykliste Minutter.');
                return ' <span class="warning-flag" title="' + title + '">🕒</span>';
            }

            function getInvoiceStatusFlagHtml(item, forceShow = false) {
                if (!item) return '';
                const isTracked = Boolean(forceShow || item.IsInvoiceTracked || item.isInvoiceTracked || isInvoiceTrackedProdNo(item.ProdNo));
                if (!isTracked) return '';
                const noInvoValue = Number(item.NoInvo || 0);
                const noFinValue = Number(item.NoFin || 0);
                const hasMissing = Boolean(item.UsesMissingInvoiceFallback || item.usesMissingInvoiceFallback || String(item.MissingInvoiceText || item.missingInvoiceText || '').trim() || (noInvoValue === 0 && noFinValue > 0));
                const hasInvoice = item.HasInvoice === true || item.hasInvoice === true || noInvoValue > 0;
                const warningText = String(item.WarningText || item.warningText || '').toLowerCase();
                if (hasMissing && (warningText.includes('faktura') || warningText.includes('noinvo'))) {
                    return '';
                }
                const title = escapeHtml(item.InvoiceStatusText || item.invoiceStatusText || item.MissingInvoiceText || item.missingInvoiceText || (hasMissing
                    ? 'Mangler faktura; NoInvo er 0 og NoFin bruges til kostberegning.'
                    : (hasInvoice ? ('Faktura registreret: NoInvo = ' + noInvoValue + '.') : 'Ingen fakturainfo fundet.')));
                const icon = hasMissing ? '🧾' : (hasInvoice ? '📄' : '❔');
                return ' <span class="warning-flag" title="' + title + '">' + icon + '</span>';
            }

            function getInvoiceStatusSummaryHtml(lines, forceShow = false) {
                if (!Array.isArray(lines) || lines.length === 0) return '';
                const trackedLines = lines.filter(line => line && (forceShow || line.IsInvoiceTracked || line.isInvoiceTracked || isInvoiceTrackedProdNo(line.ProdNo)));
                if (trackedLines.length === 0) return '';
                const missingLine = trackedLines.find(line => {
                    if (!line) return false;
                    if (line.UsesMissingInvoiceFallback || line.usesMissingInvoiceFallback) return true;
                    const noInvoValue = Number(line.NoInvo || 0);
                    const noFinValue = Number(line.NoFin || 0);
                    return noInvoValue === 0 && noFinValue > 0;
                });
                const referenceLine = missingLine || trackedLines[0];
                const noInvoValue = Number((referenceLine && referenceLine.NoInvo) || 0);
                const hasInvoice = Boolean((referenceLine && (referenceLine.HasInvoice === true || referenceLine.hasInvoice === true)) || noInvoValue > 0);
                const cssClass = missingLine ? 'warn' : 'ok';
                const icon = missingLine ? '🧾' : (hasInvoice ? '📄' : '❔');
                const text = escapeHtml((referenceLine && (referenceLine.InvoiceStatusText || referenceLine.invoiceStatusText || referenceLine.MissingInvoiceText || referenceLine.missingInvoiceText)) || (missingLine
                    ? 'Mangler faktura; NoInvo er 0 og NoFin bruges til kostberegning.'
                    : (hasInvoice ? ('Faktura registreret: NoInvo = ' + noInvoValue + '.') : 'Ingen fakturainfo fundet.')));
                return '<div class="invoice-status-banner ' + cssClass + '">' + icon + ' ' + text + '</div>';
            }

            function getLaserAllocationFlagHtml(item, fallbackText) {
                if (!item || (!item.UsesLaserAllocationSpread && !item.usesLaserAllocationSpread)) return '';
                const title = escapeHtml(item.LaserAllocationText || item.laserAllocationText || fallbackText || 'Laserkosten er fordelt på et andet antal stk end denne ordrelinje, så pris pr. stk kan afvige.');
                return ' <span class="allocation-flag" title="' + title + '">*</span>';
            }

            const laserNestCostHints = new Map();

            function setLaserNestCostHint(ordNo, prodNo, nestingCost) {
                const numericOrdNo = Number(ordNo || 0);
                const normalizedProdNo = String(prodNo || '').trim().toUpperCase();
                const numericCost = Number(nestingCost || 0);
                if (!numericOrdNo || !normalizedProdNo || !(numericCost > 0)) return;
                laserNestCostHints.set(numericOrdNo + '|' + normalizedProdNo, numericCost);
            }

            function getLaserNestCostHint(ordNo, prodNo) {
                const numericOrdNo = Number(ordNo || 0);
                const normalizedProdNo = String(prodNo || '').trim().toUpperCase();
                if (!numericOrdNo || !normalizedProdNo) return null;
                const value = laserNestCostHints.get(numericOrdNo + '|' + normalizedProdNo);
                return Number.isFinite(Number(value)) && Number(value) > 0 ? Number(value) : null;
            }

            function toDrawingUrl(rawPath) {
                const value = String(rawPath || '').trim();
                if (!value) return '';
                const lower = value.toLowerCase();
                if (lower.startsWith('http://') || lower.startsWith('https://') || lower.startsWith('file://')) return value;

                const bs = String.fromCharCode(92);
                if (value.startsWith(bs + bs)) {
                    const uncPath = value.slice(2).split(bs).join('/');
                    return 'file://' + encodeURI(uncPath);
                }

                const normalized = value.split(bs).join('/');
                const hasDrivePrefix = normalized.length >= 3
                    && ((normalized[0] >= 'A' && normalized[0] <= 'Z') || (normalized[0] >= 'a' && normalized[0] <= 'z'))
                    && normalized[1] === ':'
                    && normalized[2] === '/';
                if (hasDrivePrefix) {
                    return 'file:///' + encodeURI(normalized);
                }

                return encodeURI(normalized);
            }

            function openDrawingPdf(pathOrMeta) {
                const meta = (pathOrMeta && typeof pathOrMeta === 'object') ? pathOrMeta : { path: pathOrMeta };
                const value = String(meta.path || '').trim();
                const prodNo = String(meta.prodNo || '').trim();
                const ordNo = String(meta.ordNo || '').trim();
                if (!value && !prodNo) return;
                fetch('/open-drawing', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ path: value, prodNo, ordNo })
                })
                .then(async (r) => {
                    if (r.ok) return;
                    let msg = 'Kunne ikke aabne tegning.';
                    try {
                        const d = await r.json();
                        if (d && d.message) msg = d.message;
                    } catch (_) {}
                    throw new Error(msg);
                })
                .catch((err) => {
                    const url = toDrawingUrl(value);
                    if (url) {
                        window.open(url, '_blank');
                    } else {
                        alert('Fejl ved åbning af tegning: ' + err.message);
                    }
                });
            }

            function toggleLaserOrderSummary() {
                const panel = document.getElementById('laserOrderSummaryPanel');
                const btn = document.getElementById('laserOrderSummaryToggleBtn');
                if (!panel || !btn) return;
                const isClosed = panel.style.display === 'none';
                panel.style.display = isClosed ? '' : 'none';
                if (!isClosed) {
                    const laserImagePanel = document.getElementById('laserImagePanel');
                    if (laserImagePanel) {
                        laserImagePanel.innerHTML = '';
                        laserImagePanel.classList.add('hidden');
                    }
                }
                btn.textContent = isClosed ? 'Skjul laseroversigt' : 'Vis laseroversigt';
            }

            function toggleOperationOrderSummary() {
                const panel = document.getElementById('operationOrderSummaryPanel');
                const btn = document.getElementById('operationOrderSummaryToggleBtn');
                if (!panel || !btn) return;
                const isClosed = panel.style.display === 'none';
                panel.style.display = isClosed ? '' : 'none';
                btn.textContent = isClosed ? 'Skjul operationer' : 'Vis operationer';
            }

            function buildOversigtModalView(type) {
                const isLaser = type === 'laser';
                const totalsId = isLaser ? 'laserOrderSummaryTotals' : 'operationOrderSummaryTotals';
                const bodyId = isLaser ? 'laserOrderSummaryBody' : 'operationOrderSummaryBody';
                const titleEl = document.getElementById('oversigtModalTitle');
                const subtitleEl = document.getElementById('oversigtModalSubtitle');
                const modalBody = document.getElementById('oversigtModalBody');
                const totals = document.getElementById(totalsId);
                const body = document.getElementById(bodyId);
                if (!titleEl || !subtitleEl || !modalBody || !totals || !body) return;

                titleEl.textContent = isLaser ? 'Laseroversigt (L-linjer)' : 'Operation Oversigt';
                subtitleEl.textContent = isLaser
                    ? 'Nesting, vægt og kost i samlet driftsvisning'
                    : 'Operationstid, kapacitet og kost i samlet driftsvisning';

                modalBody.innerHTML = ''
                    + '<div class="oversigt-modal-layout">'
                    +   '<section class="oversigt-panel oversigt-kpi"><h5>Samlede KPI</h5>' + totals.innerHTML + '</section>'
                    +   '<section class="oversigt-panel oversigt-details"><h5>Detaljer</h5>' + body.innerHTML + '</section>'
                    + '</div>';
                applyMicroTablePolish(modalBody);
            }

            function openOversigtModal(type) {
                currentOversigtModalType = (type === 'operation') ? 'operation' : 'laser';
                const modal = document.getElementById('oversigtModal');
                if (!modal) return;
                buildOversigtModalView(currentOversigtModalType);
                modal.style.display = 'flex';
            }

            function closeOversigtModal(event) {
                if (event && event.target && event.target.id !== 'oversigtModal') return;
                const modal = document.getElementById('oversigtModal');
                const body = document.getElementById('oversigtModalBody');
                if (modal) modal.style.display = 'none';
                if (body) body.innerHTML = '';
                currentOversigtModalType = null;
            }

            function refreshActiveOversigtModal() {
                if (!currentSearchOrderData) return;
                if (currentOversigtModalType === 'laser') {
                    loadSalesOrderLaserSummary(currentSearchOrderData);
                } else if (currentOversigtModalType === 'operation') {
                    loadSalesOrderOperationSummary(currentSearchOrderData);
                }
            }

            let currentMarginMode = 'classic';
            let orderListData = [];
            let orderListVisible = true;
            const ORDER_LIST_DAYS_BACK_CLIENT = 30;
            const AFTERCALC_CLIENT_CACHE_TTL_MS = 2 * 60 * 1000;
            let activeSearchRequestId = 0;
            let prefetchOrderDebounceTimer = null;
            let currentSearchOrderData = null;
            let lastOrderReportHtml = '';
            let lastOrderReportTitle = 'Rapport';
            let reportOriginState = null;
            let orderListFilter = '';
            let orderListBrugerFilter = '';
            let orderListMinDkkEnabled = false;
            let orderListMinDkkValue = 0;
            let marginStateByOrdNo = {};
            let marginJobQueue = [];
            let marginWorkerActiveCount = 0;
            let orderListRerenderTimer = null;
            let orderListLoading = false;
            let orderListAutoRefreshTimer = null;
            let orderListSortField = 'date';
            let orderListSortDir = 'desc';
            let marginSortRefreshTimer = null;
            let currentOversigtModalType = null;
            const aftercalcClientCache = new Map();
            const routeMetricsClientCache = new Map();
            const ROUTE_METRICS_CLIENT_CACHE_TTL_MS = 2 * 60 * 1000;

            function normalizeOrdNoValue(ordNo) {
                return String(ordNo || '').trim();
            }

            function pruneAftercalcClientCache() {
                if (aftercalcClientCache.size <= 80) return;
                const keys = Array.from(aftercalcClientCache.keys());
                for (let i = 0; i < keys.length - 80; i++) {
                    aftercalcClientCache.delete(keys[i]);
                }
            }

            function pruneRouteMetricsClientCache() {
                if (routeMetricsClientCache.size <= 100) return;
                const keys = Array.from(routeMetricsClientCache.keys());
                for (let i = 0; i < keys.length - 100; i++) {
                    routeMetricsClientCache.delete(keys[i]);
                }
            }

            async function requestRouteMetricsData(endpoint, options = {}) {
                const cacheKey = String(endpoint || '').trim();
                if (!cacheKey) throw new Error('Route metrics endpoint mangler');

                const forceReload = Boolean(options.forceReload);
                const now = Date.now();
                const existing = routeMetricsClientCache.get(cacheKey);

                if (!forceReload && existing) {
                    if (existing.data && (now - Number(existing.ts || 0)) < ROUTE_METRICS_CLIENT_CACHE_TTL_MS) {
                        return existing.data;
                    }
                    if (existing.promise) {
                        return existing.promise;
                    }
                }

                const fetchPromise = (async () => {
                    const response = await fetch(cacheKey);
                    const data = await response.json();
                    if (!response.ok || (data && data.error)) {
                        throw new Error((data && data.error) ? data.error : ('HTTP ' + response.status));
                    }
                    routeMetricsClientCache.set(cacheKey, { data, ts: Date.now(), promise: null });
                    pruneRouteMetricsClientCache();
                    return data;
                })();

                routeMetricsClientCache.set(cacheKey, { data: null, ts: now, promise: fetchPromise });
                try {
                    return await fetchPromise;
                } catch (err) {
                    routeMetricsClientCache.delete(cacheKey);
                    throw err;
                }
            }

            function buildLaserRouteMetricsEndpoint(ordine, route, prodNo, showAllRoutes) {
                return '/laser-route-metrics?ordine=' + encodeURIComponent(String(ordine || '').trim())
                    + (showAllRoutes ? '' : ('&route=' + encodeURIComponent(String(route || '').trim())))
                    + '&prodNo=' + encodeURIComponent(String(prodNo || '').trim())
                    + '&showAllRoutes=' + (showAllRoutes ? '1' : '0')
                    + (currentSalesOrderGr4 === 3 ? '&gr4=3' : '');
            }

            async function prefetchRouteMetricsForProduct(prodNo, ordNo, trInf2, trInf4, showAllRoutes) {
                if (!prodNo) return;
                const effectiveOrdine = String(ordNo || trInf2 || '').trim();
                if (!effectiveOrdine) return;

                let effectiveRoute = String(trInf4 || '').trim();
                if (!showAllRoutes && !effectiveRoute) {
                    try {
                        const fallbackResponse = await fetch('/nesting-detail/' + encodeURIComponent(effectiveOrdine) + '/' + encodeURIComponent(prodNo));
                        const fallbackRows = await fallbackResponse.json();
                        if (fallbackResponse.ok && Array.isArray(fallbackRows) && fallbackRows.length > 0) {
                            effectiveRoute = String(fallbackRows[0].TrInf4 || '').trim();
                        }
                    } catch (_) {}
                }

                if (!showAllRoutes && !effectiveRoute) return;
                const endpoint = buildLaserRouteMetricsEndpoint(effectiveOrdine, effectiveRoute, prodNo, showAllRoutes);
                requestRouteMetricsData(endpoint).catch(() => {});
            }

            async function requestAftercalcData(ordNo, options = {}) {
                const normalizedOrdNo = normalizeOrdNoValue(ordNo);
                if (!normalizedOrdNo) throw new Error('Ordrenummer mangler');

                const forceReload = Boolean(options.forceReload);
                const now = Date.now();
                const cacheKey = normalizedOrdNo;
                const existing = aftercalcClientCache.get(cacheKey);

                if (!forceReload && existing) {
                    if (existing.data && (now - Number(existing.ts || 0)) < AFTERCALC_CLIENT_CACHE_TTL_MS) {
                        return existing.data;
                    }
                    if (existing.promise) {
                        return existing.promise;
                    }
                }

                const fetchPromise = (async () => {
                    const response = await fetch('/aftercalc/' + encodeURIComponent(normalizedOrdNo));
                    const data = await response.json();
                    if (!response.ok) {
                        throw new Error((data && data.error) ? data.error : ('HTTP ' + response.status));
                    }
                    aftercalcClientCache.set(cacheKey, { data, ts: Date.now(), promise: null });
                    pruneAftercalcClientCache();
                    return data;
                })();

                aftercalcClientCache.set(cacheKey, { data: null, ts: now, promise: fetchPromise });

                try {
                    return await fetchPromise;
                } catch (err) {
                    aftercalcClientCache.delete(cacheKey);
                    throw err;
                }
            }

            function prefetchAftercalcData(ordNo) {
                const normalizedOrdNo = normalizeOrdNoValue(ordNo);
                if (!normalizedOrdNo) return;
                requestAftercalcData(normalizedOrdNo).catch(() => {});
            }

            function applyMicroTablePolish(rootEl) {
                const root = rootEl || document;
                const tables = Array.from(root.querySelectorAll('table'));
                if (!tables.length) return;

                const rightPattern = /færdigmeldt|minutter|min\.|kost|pris|margin|afvigelse|kg|%|beløb|antal|samlet|dkk|forbrugt|stykliste|icon vægt|nestkost|nestmulti/i;
                const centerPattern = /linje|rute|prodtp4|prod\.ordre|prodordre|nestingordre|ordre$/i;
                const leftPattern = /produkt|beskrivelse|kunde|type|linjer\\/ref|hvem|status|message|beskrivelse/i;

                for (const table of tables) {
                    table.classList.add('micro-grid-table');
                    const headerCells = Array.from(table.querySelectorAll('tr:first-child th'));
                    if (!headerCells.length) continue;

                    const alignByIndex = [];
                    for (let i = 0; i < headerCells.length; i++) {
                        const text = String(headerCells[i].textContent || '').trim().toLowerCase();
                        let align = 'left';
                        if (rightPattern.test(text)) {
                            align = 'right';
                        } else if (centerPattern.test(text)) {
                            align = 'center';
                        } else if (leftPattern.test(text)) {
                            align = 'left';
                        }
                        alignByIndex[i] = align;
                    }

                    const rows = Array.from(table.querySelectorAll('tr'));
                    for (const row of rows) {
                        const cells = Array.from(row.children);
                        for (let i = 0; i < cells.length; i++) {
                            const align = alignByIndex[i] || 'left';
                            cells[i].style.textAlign = align;
                            if (align === 'right') {
                                cells[i].style.fontVariantNumeric = 'tabular-nums';
                            }
                        }
                    }
                }
            }

            // ── ORDER NOTES ────────────────────────────────────────────────
            let orderNotesCache = {};  // ordNo(string) -> { status, text, updatedAt }

            async function loadAllNotes() {
                try {
                    const r = await fetch('/order-notes-all');
                    if (r.ok) orderNotesCache = await r.json();
                } catch {}
            }

            async function loadOrderNote(ordNo) {
                const numericOrdNo = Number(ordNo || 0);
                if (!numericOrdNo) return null;
                try {
                    const r = await fetch('/order-note/' + numericOrdNo);
                    if (!r.ok) return null;
                    const note = await r.json();
                    orderNotesCache[String(numericOrdNo)] = note || { status: '', text: '', updatedAt: null };
                    renderOrderNoteBanner(numericOrdNo);
                    updateOrderNoteCell(numericOrdNo);
                    return note;
                } catch {
                    return null;
                }
            }

            function getOrderNoteHtml(ordNo) {
                const note = orderNotesCache[String(ordNo)];
                if (!note || (!note.status && !note.text && !note.isCreditNote)) return '<span style="color:#bbb;font-size:12px;">-</span>';
                const icons = { ok: '✅', error: '❌', check: '⚠️', credit: '🧾' };
                const icon = note.isCreditNote ? icons.credit : (icons[note.status] || '📝');
                const cls = note.isCreditNote ? 'credit' : (note.status || 'text');
                const preview = note.text ? escapeHtmlFE(note.text.slice(0, 40)) + (note.text.length > 40 ? '…' : '') : '';
                return '<span class="note-badge ' + cls + '" onclick="event.stopPropagation();openNotePopup(' + Number(ordNo) + ')">'
                    + icon + (note.isCreditNote ? ' Kreditnota' : '') + (preview ? ' ' + preview : '') + '</span>';
            }

            function isOrderMarkedCreditNote(ordNo) {
                const note = orderNotesCache[String(ordNo)];
                return Boolean(note && note.isCreditNote === true);
            }

            function escapeHtmlFE(s) {
                return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
            }

            function openNotePopup(ordNo, fromOrderDetail = false) {
                const note = orderNotesCache[String(ordNo)] || { status: '', text: '' };
                const existing = document.getElementById('notePopupOverlay');
                if (existing) existing.remove();

                const overlay = document.createElement('div');
                overlay.id = 'notePopupOverlay';
                overlay.className = 'note-popup-overlay';
                overlay.innerHTML =
                    '<div class="note-popup">' +
                    '<h3>📝 Note for ordre <strong>' + ordNo + '</strong></h3>' +
                    '<label>Status</label>' +
                    '<select id="noteStatusSel">' +
                    '<option value="">— ingen status —</option>' +
                    '<option value="ok">✅ OK</option>' +
                    '<option value="error">❌ Fejl</option>' +
                    '<option value="check">⚠️ Tjek</option>' +
                    '</select>' +
                    '<label style="display:flex;align-items:center;gap:8px;margin:-2px 0 10px 0;font-weight:600;">' +
                    '<input id="noteCreditChk" type="checkbox" ' + (note.isCreditNote ? 'checked' : '') + ' style="width:16px;height:16px;" />' +
                    'Kreditnota (udeluk fra samlet resoconto)' +
                    '</label>' +
                    '<label>Note</label>' +
                    '<textarea id="noteTextArea" placeholder="Skriv en note til denne ordre...">' + escapeHtmlFE(note.text || '') + '</textarea>' +
                    (note.updatedAt ? '<div style="font-size:11px;color:#888;margin-bottom:10px;">Sidst opdateret: ' + note.updatedAt.slice(0,16).replace('T',' ') + '</div>' : '') +
                    '<div class="note-popup-actions">' +
                    '<button class="btn-note-delete" onclick="deleteOrderNote(' + ordNo + ',' + fromOrderDetail + ')">Slet</button>' +
                    '<button class="btn-note-cancel" onclick="document.getElementById(\\'notePopupOverlay\\').remove()">Annuller</button>' +
                    '<button class="btn-note-save" onclick="saveOrderNote(' + ordNo + ',' + fromOrderDetail + ')">Gem</button>' +
                    '</div></div>';

                document.body.appendChild(overlay);
                overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
                document.getElementById('noteStatusSel').value = note.status || '';
            }

            async function saveOrderNote(ordNo, fromOrderDetail) {
                const status = document.getElementById('noteStatusSel').value;
                const text = document.getElementById('noteTextArea').value.trim();
                const isCreditNote = Boolean(document.getElementById('noteCreditChk') && document.getElementById('noteCreditChk').checked);
                try {
                    const r = await fetch('/order-note/' + ordNo, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ status, text, isCreditNote })
                    });
                    if (r.ok) {
                        const note = await r.json();
                        orderNotesCache[String(ordNo)] = note;
                    }
                } catch {}
                document.getElementById('notePopupOverlay').remove();
                updateOrderNoteCell(ordNo);
                if (fromOrderDetail) renderOrderNoteBanner(ordNo);
                if (orderListVisible) renderOrderList();
            }

            async function deleteOrderNote(ordNo, fromOrderDetail) {
                try {
                    await fetch('/order-note/' + ordNo, { method: 'DELETE' });
                    delete orderNotesCache[String(ordNo)];
                } catch {}
                document.getElementById('notePopupOverlay').remove();
                updateOrderNoteCell(ordNo);
                if (fromOrderDetail) renderOrderNoteBanner(ordNo);
                if (orderListVisible) renderOrderList();
            }

            function updateOrderNoteCell(ordNo) {
                const listEl = document.getElementById('orderList');
                if (!listEl) return;
                const cells = listEl.querySelectorAll('.order-note-cell[data-ordno="' + ordNo + '"]');
                const html = getOrderNoteHtml(ordNo);
                for (const cell of cells) { cell.innerHTML = html; }
                updateOrderListSummaryPanel();
            }

            function renderOrderNoteBanner(ordNo) {
                const el = document.getElementById('order-note-banner-' + ordNo);
                if (!el) return;
                const note = orderNotesCache[String(ordNo)];
                if (!note || (!note.status && !note.text && !note.isCreditNote)) { el.style.display = 'none'; return; }
                const icons = { ok: '✅', error: '❌', check: '⚠️', credit: '🧾' };
                const icon = note.isCreditNote ? icons.credit : (icons[note.status] || '📝');
                const cls = note.isCreditNote ? 'credit' : (note.status || 'text');
                el.className = 'order-note-banner ' + cls;
                el.style.display = 'flex';
                const label = note.isCreditNote
                    ? 'Kreditnota'
                    : (note.status === 'ok' ? 'OK' : note.status === 'error' ? 'Fejl' : note.status === 'check' ? 'Tjek' : 'Note');
                el.innerHTML = '<span class="note-icon">' + icon + '</span><div class="note-body"><strong>' +
                    label +
                    '</strong>' + (note.text ? ': ' + escapeHtmlFE(note.text) : '') + '</div>' +
                    '<span style="font-size:11px;opacity:0.7;margin-left:auto;cursor:pointer;" onclick="openNotePopup(' + ordNo + ',true)">✏️ Rediger</span>';
            }
            let summaryModalHistory = [];
            let summaryImageRegistry = {};
            let summaryImageRegistryCounter = 0;
            const ACCESS_CODE = '12345';
            let accessGranted = false;
            let loggedUserDisplayName = 'Bruger';
            let sideMenuOpen = false;
            let dashboardUpdatePollTimer = null;
            const MARGIN_MAX_CONCURRENT = 2;
            const MARGIN_QUEUE_DELAY_MS = 120;
            const MARGIN_FETCH_TIMEOUT_MS = 20000;
            const MARGIN_PREFETCH_ROWS = 150;
            const ORDER_LIST_AUTO_REFRESH_MS = 2 * 60 * 1000;
            let lastOrderListCheckTime = 0;
            let lastOrderListRemoteTime = 0;
            let omsaetningInitialized = false;
            let omsaetningAccounts = [];
            let omsaetningSelectedAccounts = new Set();
            let omsaetningCustomerResults = [];
            let omsaetningSelectedCustomers = new Map();
            let omsaetningCustomerSearchToken = 0;
            let omsaetningCustomerSearchTimer = null;
            let omsaetningSelectedFiscalYears = new Set();
            let omsaetningAutoReloadTimer = null;
            let omsaetningThresholdsByCustomer = new Map();
            let omsaetningDetailsCollapsed = true;
            let omsaetningAccountsPanelOpen = false;
            const OMSAETNING_SSRS_DEFAULT_ACCOUNTS = new Set(['11012', '11015', '11040']);
            const OMSAETNING_AUTO_RELOAD_DELAY_MS = 280;
            const OMSAETNING_SUMMARY_CACHE_TTL_MS = 15 * 60 * 1000;
            const OMSAETNING_CUSTOMER_SEARCH_CACHE_TTL_MS = 120000;
            const OMSAETNING_CACHE_MAX_ITEMS = 30;
            const OMSAETNING_SHOW_THRESHOLD_SECTION = false;
            let omsaetningThresholdLoadToken = 0;
            const OMSAETNING_DEFAULT_WARN_THRESHOLD = 3;
            const OMSAETNING_DEFAULT_GOOD_THRESHOLD = 5;
            let omsaetningSummaryCache = new Map();
            let omsaetningSummaryInFlight = new Map();
            let omsaetningCustomerSearchCache = new Map();
            let omsaetningCustomerSearchInFlight = new Map();
            let ordreindgangInitialized = false;
            let ordreindgangAutoReloadTimer = null;
            let ordreindgangSummaryCache = new Map();
            let ordreindgangSummaryInFlight = new Map();
            let ordreindgangLastPayload = null;
            let ordreindgangWeeklyCollapsed = true;
            let ordreindgangCustomersCollapsed = true;
            let ordreindgangResizeTimer = null;
            const ORDREINDGANG_AUTO_RELOAD_DELAY_MS = 280;
            const ORDREINDGANG_SUMMARY_CACHE_TTL_MS = 15 * 60 * 1000;
            let belastningInitialized = false;
            let belastningAutoReloadTimer = null;
            let belastningPeriodicTimer = null;
            let belastningLastPayload = null;
            let belastningSelectedDayKey = '';
            let belastningDetailContext = { resGr: '', parity: 1 };
            let belastningDraggedCardKey = '';
            const BELASTNING_FILTER_DEBOUNCE_MS = 280;
            let _belastningKundeSuggestTimer = null;
            let _belastningKundeResults = [];

            function scheduleBelastningKundeSuggest() {
                if (_belastningKundeSuggestTimer) clearTimeout(_belastningKundeSuggestTimer);
                _belastningKundeSuggestTimer = setTimeout(doBelastningKundeSuggest, 250);
            }

            function hideBelastningKundeSuggestions() {
                setTimeout(function() {
                    var d = document.getElementById('belastningKundeDropdown');
                    if (d) d.style.display = 'none';
                }, 200);
            }

            async function doBelastningKundeSuggest() {
                var inp = document.getElementById('belastningKunde');
                var q = inp ? inp.value.trim() : '';
                var d = document.getElementById('belastningKundeDropdown');
                if (!d) return;
                if (q.length < 2) { d.style.display = 'none'; return; }
                try {
                    var resp = await fetch('/omsaetning/customers?q=' + encodeURIComponent(q) + '&limit=15');
                    if (!resp.ok) return;
                    var data = await resp.json();
                    var results = Array.isArray(data.customers) ? data.customers : [];
                    if (!results.length) { d.style.display = 'none'; return; }
                    _belastningKundeResults = results;
                    d.innerHTML = results.map(function(r, i) {
                        var nm = escapeHtmlFE(String(r.name || ''));
                        var no = escapeHtmlFE(String(r.custNo || ''));
                        return '<div class="belastning-kunde-option" onmousedown="selectBelastningKundeOption(' + i + ',event)">'
                            + nm + '<span class="bko-sub">' + no + '</span></div>';
                    }).join('');
                    d.style.display = 'block';
                } catch(e) { /* silent */ }
            }

            function selectBelastningKundeOption(idx, e) {
                if (e) e.preventDefault();
                var r = _belastningKundeResults[idx];
                var name = r ? String(r.name || '') : '';
                var d = document.getElementById('belastningKundeDropdown');
                if (d) d.style.display = 'none';
                var inp = document.getElementById('belastningKunde');
                if (inp) {
                    inp.value = name;
                    inp.blur();
                }
                scheduleBelastningAutoReload();
            }
            const BELASTNING_PERIODIC_REFRESH_MS = 15 * 60 * 1000;

            function sanitizeDisplayName(name) {
                const safe = String(name || '').trim();
                return safe ? safe.slice(0, 32) : 'Bruger';
            }

            function updateHeaderGreeting() {
                const greeting = document.getElementById('headerUserGreeting');
                if (!greeting) return;
                greeting.textContent = 'Hej, ' + sanitizeDisplayName(loggedUserDisplayName);
            }

            function setLoggedUserDisplayName(name, persist = true) {
                loggedUserDisplayName = sanitizeDisplayName(name);
                if (persist) {
                    try {
                        localStorage.setItem('afterkalk_logged_user_name', loggedUserDisplayName);
                    } catch {}
                }
                updateHeaderGreeting();
            }

            function toggleSideMenu() {
                if (sideMenuOpen) {
                    closeSideMenu();
                } else {
                    openSideMenu();
                }
            }

            function openSideMenu() {
                const overlay = document.getElementById('sideMenuOverlay');
                if (!overlay) return;
                overlay.classList.add('open');
                sideMenuOpen = true;
                refreshSideMenuAuthState();
                const input = document.getElementById('sideMenuLoginInput');
                if (!accessGranted && input) {
                    setTimeout(() => input.focus(), 30);
                }
            }

            function closeSideMenu(event) {
                if (event && event.target && event.target.id !== 'sideMenuOverlay') return;
                const overlay = document.getElementById('sideMenuOverlay');
                if (!overlay) return;
                overlay.classList.remove('open');
                sideMenuOpen = false;
            }

            function refreshSideMenuAuthState() {
                const userInput = document.getElementById('sideMenuUserInput');
                const input = document.getElementById('sideMenuLoginInput');
                const loginBtn = document.getElementById('sideMenuLoginBtn');
                const status = document.getElementById('sideMenuAuthStatus');
                const logoutBtn = document.getElementById('sideMenuLogoutBtn');
                if (!status) return;

                if (accessGranted) {
                    status.textContent = 'Logget ind som ' + sanitizeDisplayName(loggedUserDisplayName) + '.';
                    status.classList.add('ok');
                    if (userInput) userInput.disabled = true;
                    if (input) {
                        input.value = '';
                        input.disabled = true;
                    }
                    if (loginBtn) loginBtn.disabled = true;
                    if (logoutBtn) logoutBtn.disabled = false;
                } else {
                    status.textContent = 'Ikke logget ind.';
                    status.classList.remove('ok');
                    if (userInput) userInput.disabled = false;
                    if (input) input.disabled = false;
                    if (loginBtn) loginBtn.disabled = false;
                    if (logoutBtn) logoutBtn.disabled = true;
                }
            }

            function submitAccessCodeFromSideMenu() {
                const sideUserInput = document.getElementById('sideMenuUserInput');
                const sideInput = document.getElementById('sideMenuLoginInput');
                const gateInput = document.getElementById('accessGateInput');
                if (sideUserInput) {
                    const desiredName = sanitizeDisplayName(sideUserInput.value);
                    setLoggedUserDisplayName(desiredName);
                }
                if (sideInput && gateInput) {
                    gateInput.value = sideInput.value || '';
                }
                submitAccessCode();
            }

            function navigateFromSideMenu(target) {
                if (target === 'dashboard') {
                    goToDashboard();
                    closeSideMenu();
                    return;
                }
                if (target === 'brugermanual') {
                    openBrugermanual();
                    closeSideMenu();
                    return;
                }
                if (target === 'personalehåndbog') {
                    openPersonalehåndbog();
                    closeSideMenu();
                    return;
                }
                openModule(target);
                closeSideMenu();
            }

            function logoutFromSideMenu() {
                accessGranted = false;
                setLoggedUserDisplayName('Bruger');
                closeSideMenu();
                goToDashboard();
                showAccessGate();
                const gateInput = document.getElementById('accessGateInput');
                if (gateInput) gateInput.value = '';
                refreshSideMenuAuthState();
            }

            function openBrugermanual() {
                const modal = document.getElementById('brugermanualModal');
                const body = document.getElementById('brugermanualBody');
                if (!modal || !body) return;
                body.innerHTML = ''
                    + '<section class="manual-card">'
                    + '<h4>1. Dashboard</h4>'
                    + '<p>Overblik over makrokategorier og hurtig adgang til moduler.</p>'
                    + '<ul><li>Brug kortene til at åbne modul.</li><li>Brug "Ryd Efterkalk cache" kun ved dataproblemer.</li><li>Warmup-status viser baggrundsindlæsning.</li></ul>'
                    + '</section>'
                    + '<section class="manual-card">'
                    + '<h4>2. Efterkalkulation</h4>'
                    + '<p>Ordreliste, margin, produktion og rapportdetaljer.</p>'
                    + '<ul><li>Klik på ordrelinje for fuld rapport.</li><li>"Opdater" på en ordre rydder cache for netop den ordre og henter frisk beregning.</li><li>Hvis tal ikke ændrer sig, er kilde-data sandsynligvis uændret.</li></ul>'
                    + '</section>'
                    + '<section class="manual-card">'
                    + '<h4>3. Omsætning</h4>'
                    + '<p>Periode-, konto- og kundebaseret omsætningsanalyse.</p>'
                    + '<ul><li>Vælg periode og konti.</li><li>Tryk "Opdater" for nye tal.</li><li>Print fra modulet efter opdatering.</li></ul>'
                    + '</section>'
                    + '<section class="manual-card">'
                    + '<h4>4. Ordreindgang</h4>'
                    + '<p>Ugevis ordre- og tilbudsoverblik.</p>'
                    + '<ul><li>Vælg ugeinterval (YYYYWW).</li><li>Tryk "Opdater".</li><li>Brug tabeller/grafer til opfølgning.</li></ul>'
                    + '</section>'
                    + '<section class="manual-card">'
                    + '<h4>5. Datadifferencer (NestKost)</h4>'
                    + '<p>NestKost pr. stk kan afvige, hvis færdigmeldt antal på ordrelinjen ikke matcher nesting-fordeling.</p>'
                    + '<ul><li>Pris pr. stk beregnes fra samme kilde som linjens totale kost.</li><li>Routedetaljer kan vise et andet antal pga. fordeling/split på ruter.</li></ul>'
                    + '</section>'
                    + '<div class="manual-meta">Tip: Brug side-menuen (☰) til hurtig navigation mellem moduler og manual.</div>';
                modal.style.display = 'flex';
            }

            function closeBrugermanual(event) {
                if (event && event.target && event.target.id !== 'brugermanualModal') return;
                const modal = document.getElementById('brugermanualModal');
                if (modal) modal.style.display = 'none';
            }

            const PH_BASE_URL = 'http://apv/GHB/';

            function openPersonalehåndbog() {
                const modal = document.getElementById('personalehåndbogsModal');
                const iframe = document.getElementById('personalehåndbogsIframe');
                const input = document.getElementById('personalehåndbogsSearchInput');
                if (!modal || !iframe) return;
                if (!iframe.src || iframe.src === 'about:blank' || iframe.src === window.location.href) {
                    iframe.src = PH_BASE_URL;
                }
                modal.classList.add('open');
                document.body.style.overflow = 'hidden';
                phCheckStatus();
                if (input) setTimeout(() => input.focus(), 150);
            }

            function closePersonalehåndbog() {
                const modal = document.getElementById('personalehåndbogsModal');
                const iframe = document.getElementById('personalehåndbogsIframe');
                if (modal) modal.classList.remove('open');
                if (iframe) iframe.src = '';
                document.body.style.overflow = '';
            }

            let qmsDataset = null;
            let qmsFlatDocs = [];
            let qmsSelectedDocId = null;
            let qmsEditMode = false;

            function escapeHtml(str) {
                return String(str || '')
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
            }

            function makeQmsId(prefix) {
                return String(prefix || 'id') + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
            }

            async function loadQmsDataset(force = false) {
                if (qmsDataset && !force) return qmsDataset;
                const r = await fetch('/qms/dataset');
                const data = await r.json();
                if (!data.ok || !data.dataset) throw new Error(data.error || 'QMS dataset fejl');
                qmsDataset = data.dataset;
                return qmsDataset;
            }

            async function saveQmsDataset() {
                const r = await fetch('/qms/dataset', {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ dataset: qmsDataset })
                });
                const data = await r.json();
                if (!data.ok) throw new Error(data.error || 'Kunne ikke gemme dataset');
                qmsDataset = data.dataset;
                return qmsDataset;
            }

            function flattenQmsDataset() {
                const out = [];
                if (!qmsDataset || !Array.isArray(qmsDataset.folders)) return out;
                for (const folder of qmsDataset.folders) {
                    const docs = Array.isArray(folder.documents) ? folder.documents : [];
                    for (const doc of docs) {
                        out.push({
                            folderId: folder.id,
                            folderName: folder.name,
                            folderDescription: folder.description || '',
                            id: doc.id,
                            title: doc.title,
                            url: doc.url || '',
                            content: doc.content || '',
                            tags: Array.isArray(doc.tags) ? doc.tags : []
                        });
                    }
                }
                qmsFlatDocs = out;
                return out;
            }

            function getSelectedQmsDoc() {
                return qmsFlatDocs.find(d => d.id === qmsSelectedDocId) || null;
            }

            function renderQmsView(doc) {
                const view = document.getElementById('qmsView');
                if (!view) return;
                if (!doc) {
                    view.innerHTML = '<h3>Kvalitetsledelsessystem</h3><div class="qms-view-meta">Vælg et dokument i venstre side.</div>';
                    return;
                }
                if (qmsEditMode) {
                    view.innerHTML = ''
                        + '<h3>Rediger dokument</h3>'
                        + '<div class="qms-view-meta">' + escapeHtml(doc.folderName) + '</div>'
                        + '<div class="qms-editor">'
                        + '<label>Titel</label><input id="qmsEditTitle" value="' + escapeHtml(doc.title) + '" />'
                        + '<label>URL (valgfri)</label><input id="qmsEditUrl" value="' + escapeHtml(doc.url) + '" />'
                        + '<label>Indhold</label><textarea id="qmsEditContent">' + escapeHtml(doc.content) + '</textarea>'
                        + '<div class="qms-editor-actions">'
                        + '<button class="save" onclick="qmsSaveCurrentDoc()">Gem dokument</button>'
                        + '<button class="delete" onclick="qmsDeleteCurrentDoc()">Slet dokument</button>'
                        + '<button class="cancel" onclick="toggleQmsEditMode(false)">Afslut redigering</button>'
                        + '</div>'
                        + '</div>';
                    return;
                }
                view.innerHTML = ''
                    + '<h3>' + escapeHtml(doc.title) + '</h3>'
                    + '<div class="qms-view-meta">' + escapeHtml(doc.folderName) + '</div>'
                    + '<div class="qms-view-content">' + escapeHtml(doc.content) + '</div>'
                    + (doc.url ? '<div class="qms-view-link"><a href="' + escapeHtml(doc.url) + '" target="_blank" rel="noopener noreferrer">Åbn original reference</a></div>' : '');
            }

            function renderQmsList(query = '') {
                const list = document.getElementById('qmsList');
                const label = document.getElementById('qmsListLabel');
                if (!list || !label) return;
                const q = String(query || '').trim().toLowerCase();
                const docs = flattenQmsDataset().filter(doc => {
                    if (!q) return true;
                    return (doc.title + ' ' + doc.folderName + ' ' + doc.content + ' ' + doc.tags.join(' ')).toLowerCase().includes(q);
                });
                label.textContent = docs.length + ' dokumenter';
                if (docs.length === 0) {
                    list.innerHTML = '<div class="qms-empty">Ingen dokumenter matcher din søgning.</div>';
                    renderQmsView(null);
                    return;
                }
                list.innerHTML = docs.map(doc => (
                    '<div class="qms-item" data-doc-id="' + escapeHtml(doc.id) + '" onclick="openQmsPage(this)">' +
                    '<div class="qms-item-title">' + escapeHtml(doc.title) + '</div>' +
                    '<div class="qms-item-meta">' + escapeHtml(doc.folderName) + '</div>' +
                    '</div>'
                )).join('');
                if (!qmsSelectedDocId || !docs.some(d => d.id === qmsSelectedDocId)) {
                    qmsSelectedDocId = docs[0].id;
                }
                const active = list.querySelector('.qms-item[data-doc-id="' + CSS.escape(qmsSelectedDocId) + '"]') || list.querySelector('.qms-item');
                if (active) openQmsPage(active);
            }

            async function openKvalitetsledelsessystem() {
                const modal = document.getElementById('qmsModal');
                const input = document.getElementById('qmsSearchInput');
                if (!modal) return;
                modal.classList.add('open');
                document.body.style.overflow = 'hidden';
                try {
                    await loadQmsDataset(false);
                    renderQmsList('');
                } catch (err) {
                    const list = document.getElementById('qmsList');
                    if (list) list.innerHTML = '<div class="qms-empty">Kunne ikke læse QMS dataset: ' + escapeHtml(err.message || '') + '</div>';
                    renderQmsView(null);
                }
                if (input) {
                    input.value = '';
                    setTimeout(() => input.focus(), 120);
                }
            }

            function closeQmsModal() {
                const modal = document.getElementById('qmsModal');
                if (modal) modal.classList.remove('open');
                document.body.style.overflow = '';
            }

            function searchQmsPages() {
                const input = document.getElementById('qmsSearchInput');
                renderQmsList(input ? input.value : '');
            }

            function openQmsPage(el) {
                const docId = el && el.getAttribute ? el.getAttribute('data-doc-id') : '';
                if (!docId) return;
                qmsSelectedDocId = docId;
                document.querySelectorAll('#qmsList .qms-item').forEach(x => x.classList.remove('active'));
                el.classList.add('active');
                renderQmsView(getSelectedQmsDoc());
            }

            function toggleQmsEditMode(force) {
                if (typeof force === 'boolean') {
                    qmsEditMode = force;
                } else {
                    qmsEditMode = !qmsEditMode;
                }
                const btn = document.getElementById('qmsEditToggleBtn');
                if (btn) btn.textContent = qmsEditMode ? 'Visning' : 'Rediger';
                renderQmsView(getSelectedQmsDoc());
            }

            async function qmsSaveCurrentDoc() {
                const doc = getSelectedQmsDoc();
                if (!doc) return;
                const title = document.getElementById('qmsEditTitle');
                const url = document.getElementById('qmsEditUrl');
                const content = document.getElementById('qmsEditContent');
                const folder = (qmsDataset.folders || []).find(f => f.id === doc.folderId);
                if (!folder) return;
                const target = (folder.documents || []).find(d => d.id === doc.id);
                if (!target) return;
                target.title = String(title && title.value || '').trim() || target.title;
                target.url = String(url && url.value || '').trim();
                target.content = String(content && content.value || '').trim();
                try {
                    await saveQmsDataset();
                    renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
                } catch (err) {
                    alert('Kunne ikke gemme: ' + (err.message || err));
                }
            }

            async function qmsDeleteCurrentDoc() {
                const doc = getSelectedQmsDoc();
                if (!doc) return;
                if (!confirm('Slet dokumentet "' + doc.title + '"?')) return;
                const folder = (qmsDataset.folders || []).find(f => f.id === doc.folderId);
                if (!folder) return;
                folder.documents = (folder.documents || []).filter(d => d.id !== doc.id);
                qmsSelectedDocId = null;
                try {
                    await saveQmsDataset();
                    renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
                } catch (err) {
                    alert('Kunne ikke slette: ' + (err.message || err));
                }
            }

            async function qmsCreateFolder() {
                try {
                    await loadQmsDataset(false);
                    const name = prompt('Navn på ny mappe:');
                    if (!name || !name.trim()) return;
                    qmsDataset.folders.push({
                        id: makeQmsId('folder'),
                        name: name.trim(),
                        description: '',
                        documents: []
                    });
                    await saveQmsDataset();
                    renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
                } catch (err) {
                    alert('Kunne ikke oprette mappe: ' + (err.message || err));
                }
            }

            async function qmsCreateDocument() {
                try {
                    await loadQmsDataset(false);
                    if (!Array.isArray(qmsDataset.folders) || qmsDataset.folders.length === 0) {
                        alert('Opret først en mappe.');
                        return;
                    }
                    const title = prompt('Titel på nyt dokument:');
                    if (!title || !title.trim()) return;
                    let folder = (qmsDataset.folders || []).find(f => f.id === qmsSelectedDocId) || null;
                    const selected = getSelectedQmsDoc();
                    if (selected) {
                        folder = (qmsDataset.folders || []).find(f => f.id === selected.folderId) || null;
                    }
                    if (!folder) folder = qmsDataset.folders[0];
                    folder.documents = Array.isArray(folder.documents) ? folder.documents : [];
                    const docId = makeQmsId('doc');
                    folder.documents.push({
                        id: docId,
                        title: title.trim(),
                        url: '',
                        content: '',
                        tags: []
                    });
                    qmsSelectedDocId = docId;
                    await saveQmsDataset();
                    renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
                    qmsEditMode = true;
                    toggleQmsEditMode(true);
                } catch (err) {
                    alert('Kunne ikke oprette dokument: ' + (err.message || err));
                }
            }

            function phSetStatus(msg) {
                const lbl = document.getElementById('phResultsLabel');
                const msgEl = document.getElementById('phStatusMsg');
                const list = document.getElementById('phResultsList');
                if (list) list.innerHTML = '<div class="ph-status-msg" id="phStatusMsg">' + msg + '</div>';
                if (lbl) lbl.textContent = 'Resultater';
            }

            async function phCheckStatus() {
                try {
                    const r = await fetch('/ph/status');
                    const d = await r.json();
                    if (d.status === 'indexing') {
                        phSetStatus('⏳ Indekserer sitet, vent venligst…');
                        setTimeout(phCheckStatus, 2000);
                    } else if (d.status === 'ready') {
                        phSetStatus('Skriv en søgning og tryk Søg.<br><small style="color:#9aabcc">' + d.count + ' sider indekseret</small>');
                    } else if (d.status === 'idle') {
                        phSetStatus('Indeks ikke klar. Tryk ↺ Genindekser.');
                    } else if (d.status === 'error') {
                        phSetStatus('⚠️ Fejl ved indeksering: ' + (d.error || ''));
                    }
                } catch { phSetStatus('Kunne ikke kontakte serveren.'); }
            }

            async function phReindex() {
                phSetStatus('⏳ Indekserer sitet, vent venligst…');
                try {
                    await fetch('/ph/reindex', { method: 'POST' });
                    setTimeout(phCheckStatus, 1500);
                } catch { phSetStatus('⚠️ Fejl ved genindeksering.'); }
            }

            function phEscapeHtml(s) {
                return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
            }

            function phHighlight(text, terms) {
                let out = phEscapeHtml(text);
                for (const t of terms) {
                    if (!t) continue;
                    const lower = out.toLowerCase();
                    const tl = t.toLowerCase();
                    let i = 0, result = '', pos;
                    while ((pos = lower.indexOf(tl, i)) !== -1) {
                        result += out.slice(i, pos) + '<mark>' + out.slice(pos, pos + t.length) + '</mark>';
                        i = pos + t.length;
                    }
                    out = result + out.slice(i);
                }
                return out;
            }

            async function searchPersonalehåndbog() {
                const input = document.getElementById('personalehåndbogsSearchInput');
                const list  = document.getElementById('phResultsList');
                const lbl   = document.getElementById('phResultsLabel');
                if (!input || !list) return;
                const q = input.value.trim();
                if (!q) { phCheckStatus(); return; }
                phSetStatus('🔍 Søger…');
                try {
                    const r = await fetch('/ph/search?q=' + encodeURIComponent(q));
                    const d = await r.json();
                    if (d.status === 'indexing') { phSetStatus('⏳ Indekserer endnu, prøv igen om lidt…'); return; }
                    if (!d.results || d.results.length === 0) {
                        phSetStatus('Ingen resultater for <strong>' + phEscapeHtml(q) + '</strong>.');
                        if (lbl) lbl.textContent = '0 resultater';
                        return;
                    }
                    if (lbl) lbl.textContent = d.results.length + ' resultat' + (d.results.length !== 1 ? 'er' : '');
                    const terms = q.toLowerCase().split(/\\s+/).filter(Boolean);
                    list.innerHTML = d.results.map((res, i) => {
                        const title = phHighlight(res.title || res.url, terms);
                        const snip  = phHighlight(res.snippet, terms);
                        const safeUrl = phEscapeHtml(res.url);
                        return '<div class="ph-result-item" data-url="' + safeUrl + '" onclick="phOpenResult(this)" title="' + safeUrl + '">'
                            + '<div class="ph-result-title">' + title + '</div>'
                            + '<div class="ph-result-url">' + safeUrl + '</div>'
                            + '<div class="ph-result-snippet">' + snip + '</div>'
                            + '</div>';
                    }).join('');
                    // Auto-load first result
                    const first = list.querySelector('.ph-result-item');
                    if (first) phOpenResult(first);
                } catch { phSetStatus('⚠️ Søgefejl. Prøv igen.'); }
            }

            function phOpenResult(el) {
                const url = el.getAttribute('data-url');
                if (!url) return;
                const iframe = document.getElementById('personalehåndbogsIframe');
                if (iframe) iframe.src = url;
                document.querySelectorAll('.ph-result-item').forEach(e => e.classList.remove('ph-active'));
                el.classList.add('ph-active');
            }

            function setOmsaetningCacheEntry(cacheMap, key, value) {
                cacheMap.set(key, {
                    ts: Date.now(),
                    value
                });
                if (cacheMap.size > OMSAETNING_CACHE_MAX_ITEMS) {
                    const oldestKey = cacheMap.keys().next().value;
                    if (oldestKey !== undefined) cacheMap.delete(oldestKey);
                }
            }

            function getOmsaetningCacheEntry(cacheMap, key, ttlMs) {
                const hit = cacheMap.get(key);
                if (!hit) return null;
                if ((Date.now() - Number(hit.ts || 0)) > ttlMs) {
                    cacheMap.delete(key);
                    return null;
                }
                return hit.value;
            }

            function buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, selectedCustomers) {
                const accounts = Array.from(new Set((Array.isArray(selectedAccounts) ? selectedAccounts : [])
                    .map(v => String(v || '').trim())
                    .filter(Boolean))).sort();
                const customers = Array.from(new Set((Array.isArray(selectedCustomers) ? selectedCustomers : [])
                    .map(v => String(v || '').trim())
                    .filter(Boolean))).sort();
                return JSON.stringify({ fra, til, accounts, customers });
            }

            async function fetchOmsaetningSummaryCached(fra, til, selectedAccounts, selectedCustomers, options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const forceRefresh = safeOptions.forceRefresh === true;
                const cacheKey = buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, selectedCustomers);
                if (!forceRefresh) {
                    const cached = getOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, OMSAETNING_SUMMARY_CACHE_TTL_MS);
                    if (cached) {
                        return cached;
                    }
                }

                const customerFilters = Array.isArray(selectedCustomers)
                    ? Array.from(new Set(selectedCustomers.map(v => String(v || '').trim()).filter(Boolean))).sort()
                    : [];
                if (!forceRefresh && customerFilters.length > 0) {
                    const allCustomersKey = buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, []);
                    const allCustomersCached = getOmsaetningCacheEntry(omsaetningSummaryCache, allCustomersKey, OMSAETNING_SUMMARY_CACHE_TTL_MS);
                    if (allCustomersCached && Array.isArray(allCustomersCached.rows)) {
                        const selectedSet = new Set(customerFilters);
                        const filteredRows = allCustomersCached.rows.filter(row => {
                            const custNo = row && row.custNo !== null && row.custNo !== undefined
                                ? String(row.custNo).trim()
                                : '';
                            return custNo && selectedSet.has(custNo);
                        });
                        const derivedPayload = {
                            ok: true,
                            filters: {
                                fra,
                                til,
                                accounts: Array.isArray(allCustomersCached.filters && allCustomersCached.filters.accounts)
                                    ? allCustomersCached.filters.accounts
                                    : [],
                                customers: customerFilters
                            },
                            totalRevenueMio: filteredRows.reduce((sum, row) => sum + Number((row && row.revenueMio) || 0), 0),
                            rows: filteredRows
                        };
                        setOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, derivedPayload);
                        return derivedPayload;
                    }
                }

                if (!forceRefresh) {
                    const inFlight = omsaetningSummaryInFlight.get(cacheKey);
                    if (inFlight) {
                        return await inFlight;
                    }
                }

                const query = new URLSearchParams({
                    fra,
                    til,
                    accounts: selectedAccounts.join(','),
                    customers: selectedCustomers.join(',')
                });

                const reqPromise = (async () => {
                    const response = await fetch('/omsaetning/summary?' + query.toString());
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    setOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, payload);
                    return payload;
                })();

                omsaetningSummaryInFlight.set(cacheKey, reqPromise);
                try {
                    return await reqPromise;
                } finally {
                    omsaetningSummaryInFlight.delete(cacheKey);
                }
            }

            async function searchOmsaetningCustomersCached(q) {
                const key = String(q || '').trim().toLowerCase();
                if (!key) return { customers: [] };

                const cached = getOmsaetningCacheEntry(omsaetningCustomerSearchCache, key, OMSAETNING_CUSTOMER_SEARCH_CACHE_TTL_MS);
                if (cached) {
                    return cached;
                }

                const inFlight = omsaetningCustomerSearchInFlight.get(key);
                if (inFlight) {
                    return await inFlight;
                }

                const reqPromise = (async () => {
                    const response = await fetch('/omsaetning/customers?q=' + encodeURIComponent(key) + '&limit=25');
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    setOmsaetningCacheEntry(omsaetningCustomerSearchCache, key, payload);
                    return payload;
                })();

                omsaetningCustomerSearchInFlight.set(key, reqPromise);
                try {
                    return await reqPromise;
                } finally {
                    omsaetningCustomerSearchInFlight.delete(key);
                }
            }

            function applyOmsaetningDetailsCollapsedState() {
                const tableWrap = document.getElementById('omsaetningTableWrap');
                const toggleBtn = document.getElementById('omsaetningDetailsToggleBtn');
                if (!tableWrap || !toggleBtn) return;
                tableWrap.style.display = omsaetningDetailsCollapsed ? 'none' : 'block';
                toggleBtn.textContent = omsaetningDetailsCollapsed ? 'Vis detaljer' : 'Skjul detaljer';
            }

            function renderOmsaetningAccountsSummary() {
                const summaryEl = document.getElementById('omsaetningAccountsSummary');
                const activeEl = document.getElementById('omsaetningAccountsActive');
                if (!summaryEl) return;
                const selectedCount = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean).length;
                const totalCount = Array.isArray(omsaetningAccounts) ? omsaetningAccounts.length : 0;
                summaryEl.textContent = String(selectedCount) + '/' + String(totalCount) + ' valgt';

                if (activeEl) {
                    const selectedValues = Array.from(omsaetningSelectedAccounts.values())
                        .map(v => String(v || '').trim())
                        .filter(Boolean)
                        .sort((a, b) => a.localeCompare(b));

                    if (selectedValues.length === 0) {
                        activeEl.innerHTML = '<span class="chip more">Ingen konti valgt</span>';
                    } else {
                        const visible = selectedValues.slice(0, 5).map(acNo => {
                            const account = Array.isArray(omsaetningAccounts)
                                ? omsaetningAccounts.find(a => String(a && a.acNo || '').trim() === acNo)
                                : null;
                            const label = account ? (acNo + ' ' + String(account.name || '').trim()) : acNo;
                            return '<span class="chip" title="' + escapeHtmlFE(label) + '">' + escapeHtmlFE(acNo) + '</span>';
                        });
                        if (selectedValues.length > 5) {
                            visible.push('<span class="chip more">+' + escapeHtmlFE(String(selectedValues.length - 5)) + '</span>');
                        }
                        activeEl.innerHTML = visible.join('');
                    }
                }
            }

            function formatSigned(value, digits) {
                const n = Number(value || 0);
                const fixed = n.toFixed(Number.isFinite(digits) ? digits : 1);
                return (n > 0 ? '+' : '') + fixed;
            }

            function buildOmsaetningGaugeData(amountMio, warnThreshold, goodThreshold) {
                const amount = Number(amountMio || 0);
                const warn = Number(warnThreshold || 0);
                const good = Math.max(warn + 0.0001, Number(goodThreshold || 0));
                const span = Math.max(0.0001, good - warn);

                const marginPct = ((amount - warn) / span) * 30;
                const scaleMin = -30;
                const scaleMax = 60;
                const toLeft = value => ((value - scaleMin) / (scaleMax - scaleMin)) * 100;

                const zeroLeft = toLeft(0);
                const targetLeft = toLeft(30);
                const clamped = Math.max(scaleMin, Math.min(scaleMax, marginPct));
                const pointLeft = toLeft(clamped);

                return {
                    marginPct,
                    pointLeft,
                    zeroLeft,
                    targetLeft,
                    fillLeft: Math.min(pointLeft, zeroLeft),
                    fillWidth: Math.abs(pointLeft - zeroLeft),
                    fillClass: pointLeft >= zeroLeft ? 'pos' : 'neg',
                    deltaWarn: amount - warn,
                    deltaGood: amount - good
                };
            }

            function applyOmsaetningAccountsPanelState() {
                const panel = document.getElementById('omsaetningAccountsPanel');
                const search = document.getElementById('omsaetningAccountSearch');
                const btn = document.getElementById('omsaetningAccountsToggleBtn');
                if (panel) panel.style.display = omsaetningAccountsPanelOpen ? 'block' : 'none';
                if (search) search.style.display = omsaetningAccountsPanelOpen ? 'block' : 'none';
                if (btn) btn.textContent = omsaetningAccountsPanelOpen ? 'Skjul konti' : 'Vis konti';
            }

            function toggleOmsaetningAccountsPanel() {
                omsaetningAccountsPanelOpen = !omsaetningAccountsPanelOpen;
                applyOmsaetningAccountsPanelState();
            }

            function toggleOmsaetningDetails() {
                omsaetningDetailsCollapsed = !omsaetningDetailsCollapsed;
                applyOmsaetningDetailsCollapsedState();
            }

            function formatMio(value) {
                const numeric = Number(value || 0);
                return numeric.toLocaleString('da-DK', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
            }

            function formatDkkFromMio(valueMio) {
                const numeric = Number(valueMio || 0) * 1000000;
                return numeric.toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            }

            function formatMonthDa(dateValue) {
                if (!dateValue) return '-';
                const dt = new Date(dateValue);
                if (Number.isNaN(dt.getTime())) return String(dateValue);
                return dt.toLocaleDateString('da-DK', { month: 'short', year: 'numeric' });
            }

            function normalizeOmsaetningMonthKey(dateValue) {
                const dt = new Date(dateValue);
                if (Number.isNaN(dt.getTime())) return String(dateValue || '').trim();
                const year = dt.getFullYear();
                const month = String(dt.getMonth() + 1).padStart(2, '0');
                return String(year) + '-' + month + '-01';
            }

            function parseMonthInputToPeriod(monthValue) {
                const raw = String(monthValue || '').trim();
                const match = raw.match(/^(\\d{4})-(\\d{2})$/);
                if (!match) return null;
                const year = Number(match[1]);
                const month = Number(match[2]);
                if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return null;
                return {
                    year,
                    month,
                    period: year * 100 + month
                };
            }

            function calendarMonthToFiscalYrPr(year, month) {
                const y = Number(year);
                const m = Number(month);
                if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) return null;
                if (m >= 7) {
                    return (y * 100) + (m - 6);
                }
                return ((y - 1) * 100) + (m + 6);
            }

            function getCurrentFiscalYearStart(referenceDate) {
                const now = referenceDate instanceof Date ? referenceDate : new Date();
                const year = now.getFullYear();
                const month = now.getMonth() + 1;
                return month >= 7 ? year : (year - 1);
            }

            function getFiscalYearRange(yearValue) {
                const startYear = Number(yearValue);
                if (!Number.isFinite(startYear)) return null;
                return {
                    fromMonth: String(startYear) + '-07',
                    toMonth: String(startYear + 1) + '-06',
                    fra: String(startYear * 100 + 7),
                    til: String((startYear + 1) * 100 + 7)
                };
            }

            function applySelectedFiscalYearsToInputs() {
                const fromEl = document.getElementById('omsaetningFraMonth');
                const toEl = document.getElementById('omsaetningTilMonth');
                const selectedYears = Array.from(omsaetningSelectedFiscalYears.values())
                    .map(y => Number(y))
                    .filter(y => Number.isFinite(y))
                    .sort((a, b) => a - b);
                if (selectedYears.length === 0) return;

                const firstRange = getFiscalYearRange(selectedYears[0]);
                const lastRange = getFiscalYearRange(selectedYears[selectedYears.length - 1]);
                if (!firstRange || !lastRange) return;

                if (fromEl) fromEl.value = firstRange.fromMonth;
                if (toEl) toEl.value = lastRange.toMonth;
            }

            function buildOmsaetningPeriodRange() {
                const fromEl = document.getElementById('omsaetningFraMonth');
                const toEl = document.getElementById('omsaetningTilMonth');
                const fromMeta = parseMonthInputToPeriod(fromEl ? fromEl.value : '');
                const toMeta = parseMonthInputToPeriod(toEl ? toEl.value : '');
                if (!fromMeta || !toMeta) return null;

                const fromDate = new Date(fromMeta.year, fromMeta.month - 1, 1);
                const toDate = new Date(toMeta.year, toMeta.month - 1, 1);
                if (fromDate.getTime() > toDate.getTime()) return null;

                const exclusiveToDate = new Date(toMeta.year, toMeta.month, 1);
                const fraFiscal = calendarMonthToFiscalYrPr(fromMeta.year, fromMeta.month);
                const tilFiscal = calendarMonthToFiscalYrPr(exclusiveToDate.getFullYear(), exclusiveToDate.getMonth() + 1);
                if (!Number.isFinite(fraFiscal) || !Number.isFinite(tilFiscal)) return null;
                return {
                    fra: String(fraFiscal),
                    til: String(tilFiscal)
                };
            }

            function buildOmsaetningMonthKeys(periodRange) {
                if (!periodRange) return [];
                const fraNum = Number(periodRange.fra);
                const tilNum = Number(periodRange.til);
                if (!Number.isFinite(fraNum) || !Number.isFinite(tilNum) || fraNum >= tilNum) return [];

                let year = Math.floor(fraNum / 100);
                let month = fraNum % 100;
                const keys = [];

                while ((year * 100 + month) < tilNum) {
                    const mm = String(month).padStart(2, '0');
                    keys.push(String(year) + '-' + mm + '-01');
                    month += 1;
                    if (month > 12) {
                        month = 1;
                        year += 1;
                    }
                }
                return keys;
            }

            function renderOmsaetningYearChips(centerYear) {
                const wrap = document.getElementById('omsaetningYears');
                if (!wrap) return;

                const currentFiscalYear = Number(centerYear || getCurrentFiscalYearStart());
                const years = [currentFiscalYear - 3, currentFiscalYear - 2, currentFiscalYear - 1, currentFiscalYear, currentFiscalYear + 1];
                wrap.innerHTML = years.map(fiscalYearStart => {
                    const activeCls = omsaetningSelectedFiscalYears.has(fiscalYearStart) ? ' active' : '';
                    const label = String(fiscalYearStart) + '/' + String(fiscalYearStart + 1).slice(-2);
                    return '<button type="button" class="omsaetning-year-btn' + activeCls + '" onclick="toggleOmsaetningFiscalYear(' + fiscalYearStart + ')">' +
                        (fiscalYearStart === currentFiscalYear ? 'Nu ' : '') + escapeHtmlFE(label) +
                        '</button>';
                }).join('');
            }

            function toggleOmsaetningFiscalYear(year) {
                const yearValue = Number(year);
                if (!Number.isFinite(yearValue)) return;
                if (omsaetningSelectedFiscalYears.has(yearValue) && omsaetningSelectedFiscalYears.size > 1) {
                    omsaetningSelectedFiscalYears.delete(yearValue);
                } else {
                    omsaetningSelectedFiscalYears.add(yearValue);
                }
                applySelectedFiscalYearsToInputs();
                renderOmsaetningYearChips(getCurrentFiscalYearStart());
                loadOmsaetningSummary();
            }

            function filterOmsaetningAccounts() {
                const qEl = document.getElementById('omsaetningAccountSearch');
                const q = String((qEl && qEl.value) || '').trim().toLowerCase();
                const rows = document.querySelectorAll('#omsaetningAccountsList .omsaetning-account-item');
                for (const row of rows) {
                    const text = String(row.getAttribute('data-search') || '').toLowerCase();
                    row.style.display = (!q || text.includes(q)) ? '' : 'none';
                }
            }

            function setAllOmsaetningAccounts(checked) {
                const list = document.getElementById('omsaetningAccountsList');
                if (!list) return;
                const boxes = list.querySelectorAll('input[type="checkbox"][data-accno]');
                for (const box of boxes) {
                    box.checked = !!checked;
                    const value = String(box.getAttribute('data-accno') || '').trim();
                    if (!value) continue;
                    if (checked) omsaetningSelectedAccounts.add(value);
                    else omsaetningSelectedAccounts.delete(value);
                }
                renderOmsaetningAccountsSummary();
                scheduleOmsaetningAutoReload();
            }

            function renderOmsaetningAccountsList() {
                const list = document.getElementById('omsaetningAccountsList');
                if (!list) return;
                if (!Array.isArray(omsaetningAccounts) || omsaetningAccounts.length === 0) {
                    list.innerHTML = '<div class="omsaetning-account-item"><span>Ingen konti fundet</span></div>';
                    return;
                }

                list.innerHTML = omsaetningAccounts.map(acc => {
                    const value = String(acc.acNo || '').trim();
                    const checked = omsaetningSelectedAccounts.has(value) ? ' checked' : '';
                    const search = (value + ' ' + String(acc.name || '')).replace(/"/g, '&quot;');
                    return '<label class="omsaetning-account-item" data-search="' + search + '">' +
                        '<input type="checkbox" data-accno="' + escapeHtmlFE(value) + '"' + checked + ' onchange="toggleOmsaetningAccount(this)" />' +
                        '<span>' + escapeHtmlFE(value + ' - ' + String(acc.name || '')) + '</span>' +
                        '</label>';
                }).join('');
                renderOmsaetningAccountsSummary();
                applyOmsaetningAccountsPanelState();
            }

            function toggleOmsaetningAccount(inputEl) {
                if (!inputEl) return;
                const value = String(inputEl.getAttribute('data-accno') || '').trim();
                if (!value) return;
                if (inputEl.checked) omsaetningSelectedAccounts.add(value);
                else omsaetningSelectedAccounts.delete(value);
                renderOmsaetningAccountsSummary();
                scheduleOmsaetningAutoReload();
            }

            function renderOmsaetningSelectedCustomers() {
                const wrap = document.getElementById('omsaetningSelectedCustomers');
                if (!wrap) return;

                const entries = Array.from(omsaetningSelectedCustomers.entries());
                if (entries.length === 0) {
                    wrap.innerHTML = '';
                    return;
                }

                wrap.innerHTML = entries.map(([custNo, name]) => (
                    '<span class="omsaetning-selected-chip">' +
                    escapeHtmlFE(String(custNo) + ' - ' + String(name || '')) +
                    '<button type="button" title="Fjern kunde" onclick="removeOmsaetningCustomer(' + Number(custNo) + ')">×</button>' +
                    '</span>'
                )).join('');
            }

            function removeOmsaetningCustomer(custNo) {
                const key = String(custNo || '').trim();
                if (!key) return;
                omsaetningSelectedCustomers.delete(key);
                renderOmsaetningSelectedCustomers();
                onOmsaetningSelectedCustomersChanged();
                renderOmsaetningCustomerResults();
            }

            function clearOmsaetningCustomerSearchUi() {
                const qEl = document.getElementById('omsaetningCustomerSearch');
                if (qEl) {
                    qEl.value = '';
                    qEl.blur();
                }
                omsaetningCustomerResults = [];
                renderOmsaetningCustomerResults();
            }

            function renderOmsaetningCustomerMode() {
                const modeEl = document.getElementById('omsaetningCustomerMode');
                if (!modeEl) return;

                const selectedCount = omsaetningSelectedCustomers.size;
                if (selectedCount === 0) {
                    modeEl.textContent = 'Ingen kunde valgt: viser normal visning for valgte år og konti.';
                    return;
                }

                if (selectedCount === 1) {
                    modeEl.textContent = '1 kunde valgt: rapporten filtreres på kunden.';
                    return;
                }

                modeEl.textContent = String(selectedCount) + ' kunder valgt: søjlediagram sammenligner kunder måned for måned.';
            }

            function toggleOmsaetningCustomer(custNo, name) {
                const key = String(custNo || '').trim();
                if (!key) return;
                const wasSelected = omsaetningSelectedCustomers.has(key);
                if (omsaetningSelectedCustomers.has(key)) {
                    omsaetningSelectedCustomers.delete(key);
                } else {
                    omsaetningSelectedCustomers.set(key, String(name || '').trim());
                }
                renderOmsaetningSelectedCustomers();
                onOmsaetningSelectedCustomersChanged();
                if (!wasSelected) clearOmsaetningCustomerSearchUi();
                renderOmsaetningCustomerResults();
            }

            function toggleOmsaetningCustomerByButton(buttonEl) {
                if (!buttonEl) return;
                const custNo = String(buttonEl.getAttribute('data-custno') || '').trim();
                const custName = String(buttonEl.getAttribute('data-custname') || '').trim();
                toggleOmsaetningCustomer(custNo, custName);
            }

            function renderOmsaetningCustomerResults() {
                const wrap = document.getElementById('omsaetningCustomerResults');
                if (!wrap) return;
                const qEl = document.getElementById('omsaetningCustomerSearch');
                const q = String((qEl && qEl.value) || '').trim();
                const selectedCount = omsaetningSelectedCustomers.size;

                if (q.length < 2) {
                    if (selectedCount === 0) {
                        wrap.innerHTML = '<div class="omsaetning-customer-empty">Ingen kunde valgt: viser normal visning for valgte år og konti.</div>';
                    } else {
                        wrap.innerHTML = '<div class="omsaetning-customer-empty">Skriv mindst 2 tegn for at tilføje flere kunder.</div>';
                    }
                    return;
                }

                if (!Array.isArray(omsaetningCustomerResults) || omsaetningCustomerResults.length === 0) {
                    wrap.innerHTML = '<div class="omsaetning-customer-empty">Ingen kunder matcher søgningen.</div>';
                    return;
                }

                wrap.innerHTML = omsaetningCustomerResults.map(row => {
                    const custNo = String(row.custNo || '').trim();
                    const custName = String(row.name || '').trim();
                    const selected = omsaetningSelectedCustomers.has(custNo);
                    return '<div class="omsaetning-customer-item">' +
                        '<div class="meta"><strong>' + escapeHtmlFE(custName || '(uden navn)') + '</strong><span>' + escapeHtmlFE(custNo) + '</span></div>' +
                        '<button type="button" class="pick' + (selected ? ' remove' : '') + '" data-custno="' + escapeHtmlFE(custNo) + '" data-custname="' + escapeHtmlFE(custName) + '" onclick="toggleOmsaetningCustomerByButton(this)">' + (selected ? 'Valgt' : 'Vælg') + '</button>' +
                        '</div>';
                }).join('');
            }

            function scheduleOmsaetningCustomerSearch() {
                if (omsaetningCustomerSearchTimer) {
                    clearTimeout(omsaetningCustomerSearchTimer);
                }
                omsaetningCustomerSearchTimer = setTimeout(() => {
                    searchOmsaetningCustomers();
                }, 220);
            }

            async function searchOmsaetningCustomers() {
                const qEl = document.getElementById('omsaetningCustomerSearch');
                const q = String((qEl && qEl.value) || '').trim();
                const token = ++omsaetningCustomerSearchToken;

                if (q.length < 2) {
                    omsaetningCustomerResults = [];
                    renderOmsaetningCustomerResults();
                    return;
                }

                try {
                    const payload = await searchOmsaetningCustomersCached(q);
                    if (token !== omsaetningCustomerSearchToken) return;
                    const rows = Array.isArray(payload.customers) ? payload.customers : [];
                    omsaetningCustomerResults = rows;
                    renderOmsaetningCustomerResults();
                } catch (err) {
                    if (token !== omsaetningCustomerSearchToken) return;
                    omsaetningCustomerResults = [];
                    const wrap = document.getElementById('omsaetningCustomerResults');
                    if (wrap) {
                        wrap.innerHTML = '<div class="omsaetning-customer-empty">Fejl ved kundesøgning.</div>';
                    }
                }
            }

            function getOmsaetningStatusClass(valueMio, warnThreshold, goodThreshold) {
                const n = Number(valueMio || 0);
                if (n >= goodThreshold) return 'good';
                if (n >= warnThreshold) return 'mid';
                return 'low';
            }

            function getOmsaetningThresholdInputs() {
                const warnThresholdInput = document.getElementById('omsaetningWarnThreshold');
                const goodThresholdInput = document.getElementById('omsaetningGoodThreshold');

                const warnRaw = Number((warnThresholdInput && warnThresholdInput.value) || OMSAETNING_DEFAULT_WARN_THRESHOLD);
                const warnThreshold = Number.isFinite(warnRaw) ? Math.max(0, warnRaw) : OMSAETNING_DEFAULT_WARN_THRESHOLD;

                const goodRaw = Number((goodThresholdInput && goodThresholdInput.value) || OMSAETNING_DEFAULT_GOOD_THRESHOLD);
                const goodThreshold = Number.isFinite(goodRaw)
                    ? Math.max(warnThreshold, goodRaw)
                    : Math.max(warnThreshold, OMSAETNING_DEFAULT_GOOD_THRESHOLD);

                return { warnThreshold, goodThreshold };
            }

            function applyOmsaetningThresholdInputs(warnThreshold, goodThreshold) {
                const warnThresholdInput = document.getElementById('omsaetningWarnThreshold');
                const goodThresholdInput = document.getElementById('omsaetningGoodThreshold');
                if (!warnThresholdInput || !goodThresholdInput) return;

                const warnRaw = Number(warnThreshold);
                const normalizedWarn = Number.isFinite(warnRaw) ? Math.max(0, warnRaw) : OMSAETNING_DEFAULT_WARN_THRESHOLD;

                const goodRaw = Number(goodThreshold);
                const normalizedGood = Number.isFinite(goodRaw)
                    ? Math.max(normalizedWarn, goodRaw)
                    : Math.max(normalizedWarn, OMSAETNING_DEFAULT_GOOD_THRESHOLD);

                warnThresholdInput.value = normalizedWarn.toFixed(1);
                goodThresholdInput.value = normalizedGood.toFixed(1);
            }

            async function loadOmsaetningThresholdForCustomer(custNo) {
                const key = String(custNo || '').trim();
                if (!/^\\d{1,20}$/.test(key)) return;
                const token = ++omsaetningThresholdLoadToken;

                try {
                    const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(key));
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    if (token !== omsaetningThresholdLoadToken) return;
                    omsaetningThresholdsByCustomer.set(key, {
                        warnThreshold: Number(payload.warnThreshold || OMSAETNING_DEFAULT_WARN_THRESHOLD),
                        goodThreshold: Number(payload.goodThreshold || OMSAETNING_DEFAULT_GOOD_THRESHOLD)
                    });
                    applyOmsaetningThresholdInputs(payload.warnThreshold, payload.goodThreshold);
                    renderOmsaetningCustomerThresholds();
                } catch (err) {
                    if (token !== omsaetningThresholdLoadToken) return;
                    console.warn('loadOmsaetningThresholdForCustomer failed:', err && err.message ? err.message : err);
                }
            }

            function renderOmsaetningCustomerThresholds() {
                const wrap = document.getElementById('omsaetningCustomerThresholds');
                if (!wrap) return;

                const selectedEntries = Array.from(omsaetningSelectedCustomers.entries());
                if (selectedEntries.length === 0) {
                    wrap.innerHTML = '';
                    return;
                }

                const rows = selectedEntries.map(([custNo, custName]) => {
                    const key = String(custNo || '').trim();
                    const threshold = omsaetningThresholdsByCustomer.get(key);
                    if (!threshold) {
                        return '<div class="omsaetning-customer-threshold-row"><span class="cust">' +
                            escapeHtmlFE(String(custName || key)) + ' (' + escapeHtmlFE(key) + ')' +
                            '</span><span class="thr">tærskler: indlæser...</span></div>';
                    }
                    return '<div class="omsaetning-customer-threshold-row"><span class="cust">' +
                        escapeHtmlFE(String(custName || key)) + ' (' + escapeHtmlFE(key) + ')' +
                        '</span><span class="thr">tærskler: ' + escapeHtmlFE(Number(threshold.warnThreshold).toFixed(1)) +
                        ' / ' + escapeHtmlFE(Number(threshold.goodThreshold).toFixed(1)) + '</span></div>';
                });

                wrap.innerHTML = rows.join('');
            }

            async function refreshOmsaetningThresholdsForSelectedCustomers(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys())
                    .map(v => String(v || '').trim())
                    .filter(v => /^\\d{1,20}$/.test(v));

                if (selectedCustomers.length === 0) {
                    omsaetningThresholdLoadToken += 1;
                    omsaetningThresholdsByCustomer = new Map();
                    renderOmsaetningCustomerThresholds();
                    return;
                }

                const token = ++omsaetningThresholdLoadToken;
                renderOmsaetningCustomerThresholds();

                const results = await Promise.all(selectedCustomers.map(async custNo => {
                    try {
                        const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(custNo));
                        if (!response.ok) throw new Error('HTTP ' + response.status);
                        const payload = await response.json();
                        return {
                            custNo,
                            warnThreshold: Number(payload.warnThreshold || OMSAETNING_DEFAULT_WARN_THRESHOLD),
                            goodThreshold: Number(payload.goodThreshold || OMSAETNING_DEFAULT_GOOD_THRESHOLD)
                        };
                    } catch (err) {
                        console.warn('refresh threshold failed for', custNo, err && err.message ? err.message : err);
                        return {
                            custNo,
                            warnThreshold: OMSAETNING_DEFAULT_WARN_THRESHOLD,
                            goodThreshold: OMSAETNING_DEFAULT_GOOD_THRESHOLD
                        };
                    }
                }));

                if (token !== omsaetningThresholdLoadToken) return;

                omsaetningThresholdsByCustomer = new Map(results.map(item => [item.custNo, {
                    warnThreshold: item.warnThreshold,
                    goodThreshold: item.goodThreshold
                }]));

                if (safeOptions.applySingleSelectionToInputs === true && selectedCustomers.length === 1) {
                    const single = omsaetningThresholdsByCustomer.get(selectedCustomers[0]);
                    if (single) {
                        applyOmsaetningThresholdInputs(single.warnThreshold, single.goodThreshold);
                    }
                }

                renderOmsaetningCustomerThresholds();
            }

            async function persistOmsaetningThresholdsForCustomers(customerNos, warnThreshold, goodThreshold) {
                const customerKeys = Array.from(new Set((Array.isArray(customerNos) ? customerNos : [])
                    .map(v => String(v || '').trim())
                    .filter(v => /^\\d{1,20}$/.test(v))));
                if (customerKeys.length === 0) return;

                await Promise.all(customerKeys.map(async custNo => {
                    const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(custNo), {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ warnThreshold, goodThreshold })
                    });
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                }));
            }

            function showOmsaetningThresholdPersistDialog(customerKeys, warnThreshold, goodThreshold) {
                return new Promise(resolve => {
                    const listHtml = customerKeys.map(custNo => {
                        const name = String(omsaetningSelectedCustomers.get(custNo) || '').trim();
                        const label = custNo + (name ? (' - ' + name) : '');
                        return '<button type="button" class="omsaetning-persist-customer-option" data-cust="' + escapeHtmlFE(custNo) + '">' + escapeHtmlFE(label) + '</button>';
                    }).join('');

                    const defaultCustomer = String(customerKeys[0] || '');
                    let selectedCustomer = defaultCustomer;
                    const overlay = document.createElement('div');
                    overlay.className = 'omsaetning-persist-overlay';
                    overlay.innerHTML =
                        '<div class="omsaetning-persist-dialog" role="dialog" aria-modal="true" aria-label="Gem tærskler">' +
                            '<div class="omsaetning-persist-head"><h4>Gem tærskler for valgte kunder</h4></div>' +
                            '<div class="omsaetning-persist-body">' +
                                '<div>Flere kunder er valgt. Vælg hvor de nye tærskler skal gemmes.</div>' +
                                '<div class="omsaetning-persist-customer-list">' + listHtml + '</div>' +
                                '<div class="omsaetning-persist-thr">Nye tærskler: ' + escapeHtmlFE(Number(warnThreshold).toFixed(1)) + ' / ' + escapeHtmlFE(Number(goodThreshold).toFixed(1)) + '</div>' +
                                '<div class="omsaetning-persist-pick">Valgt kunde: <span id="omsaetningPersistPicked" class="omsaetning-persist-picked">' + escapeHtmlFE(defaultCustomer) + '</span></div>' +
                                '</div>' +
                                '<div class="omsaetning-persist-actions">' +
                                    '<button type="button" class="primary" data-action="all">Gem alle</button>' +
                                    '<button type="button" class="ghost" data-action="single">Gem valgt kunde</button>' +
                                    '<button type="button" class="danger" data-action="none">Gem ikke</button>' +
                                '</div>' +
                            '</div>' +
                        '</div>';

                    const closeWith = value => {
                        overlay.remove();
                        resolve(value);
                    };

                    overlay.addEventListener('click', ev => {
                        if (ev.target === overlay) closeWith('NONE');
                    });

                    overlay.querySelector('[data-action="all"]').addEventListener('click', () => closeWith('ALL'));
                    overlay.querySelector('[data-action="none"]').addEventListener('click', () => closeWith('NONE'));
                    overlay.querySelector('[data-action="single"]').addEventListener('click', () => {
                        closeWith(String(selectedCustomer || '').trim());
                    });

                    const pickedEl = overlay.querySelector('#omsaetningPersistPicked');
                    const customerButtons = Array.from(overlay.querySelectorAll('.omsaetning-persist-customer-option'));

                    const renderPicked = () => {
                        for (const btn of customerButtons) {
                            const btnCust = String(btn.getAttribute('data-cust') || '');
                            btn.classList.toggle('active', btnCust === selectedCustomer);
                        }
                        if (pickedEl) pickedEl.textContent = selectedCustomer || '-';
                    };

                    for (const btn of customerButtons) {
                        btn.addEventListener('click', () => {
                            selectedCustomer = String(btn.getAttribute('data-cust') || '').trim();
                            renderPicked();
                        });
                    }

                    overlay.addEventListener('keydown', ev => {
                        if (ev.key === 'Enter') {
                            ev.preventDefault();
                            closeWith(String(selectedCustomer || '').trim());
                        }
                        if (ev.key === 'Escape') {
                            ev.preventDefault();
                            closeWith('NONE');
                        }
                    });

                    document.body.appendChild(overlay);
                    renderPicked();
                    if (customerButtons.length > 0) {
                        customerButtons[0].focus();
                    }
                });
            }

            async function resolveOmsaetningThresholdPersistTargets(selectedCustomers, warnThreshold, goodThreshold, options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const customerKeys = Array.from(new Set((Array.isArray(selectedCustomers) ? selectedCustomers : [])
                    .map(v => String(v || '').trim())
                    .filter(v => /^\\d{1,20}$/.test(v))));

                if (customerKeys.length === 0) return [];
                if (safeOptions.silentValidation === true) return [];
                if (safeOptions.persistThresholdsOnUpdate !== true) return [];

                if (customerKeys.length === 1) {
                    return customerKeys;
                }

                const answerRaw = await showOmsaetningThresholdPersistDialog(customerKeys, warnThreshold, goodThreshold);
                const answer = String(answerRaw || '').trim().toUpperCase();

                if (!answer || answer === 'NONE' || answer === 'NO' || answer === 'N') {
                    return [];
                }
                if (answer === 'ALL' || answer === 'A') {
                    return customerKeys;
                }

                const exact = customerKeys.find(c => c.toUpperCase() === answer);
                if (exact) return [exact];

                alert('Ugyldigt valg for tærskel-gemning. Ingen tærskler blev gemt.');
                return [];
            }

            function onOmsaetningSelectedCustomersChanged() {
                const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys()).filter(Boolean);
                renderOmsaetningCustomerMode();
                if (selectedCustomers.length === 0) {
                    omsaetningThresholdLoadToken += 1;
                    omsaetningThresholdsByCustomer = new Map();
                    applyOmsaetningThresholdInputs(OMSAETNING_DEFAULT_WARN_THRESHOLD, OMSAETNING_DEFAULT_GOOD_THRESHOLD);
                    renderOmsaetningCustomerThresholds();
                    scheduleOmsaetningAutoReload();
                    return;
                }

                refreshOmsaetningThresholdsForSelectedCustomers({ applySingleSelectionToInputs: true });
                scheduleOmsaetningAutoReload();
            }

            function scheduleOmsaetningAutoReload() {
                if (omsaetningAutoReloadTimer) {
                    clearTimeout(omsaetningAutoReloadTimer);
                }
                omsaetningAutoReloadTimer = setTimeout(() => {
                    omsaetningAutoReloadTimer = null;
                    loadOmsaetningSummary({ silentValidation: true });
                }, OMSAETNING_AUTO_RELOAD_DELAY_MS);
            }

            function getOmsaetningStatusLabel(statusClass) {
                if (statusClass === 'good') return 'Over mål';
                if (statusClass === 'mid') return 'Nær mål';
                return 'Under mål';
            }

            function getOmsaetningColor(index) {
                const palette = ['#1565c0', '#00acc1', '#00897b', '#7b1fa2', '#ef6c00', '#5e35b1', '#43a047', '#c62828'];
                return palette[index % palette.length];
            }

            function renderOmsaetningCharts(rows, forcedMonthKeys) {
                const chartsWrap = document.getElementById('omsaetningChartsWrap');
                const stackedSvg = document.getElementById('omsaetningStackedChart');
                const trendSvg = document.getElementById('omsaetningTrendChart');
                const legend = document.getElementById('omsaetningLegend');
                const stackedTitle = document.getElementById('omsaetningStackedTitle');
                if (!chartsWrap || !stackedSvg || !trendSvg || !legend) return;

                const safeRows = Array.isArray(rows) ? rows : [];
                const safeForcedMonths = Array.isArray(forcedMonthKeys) ? forcedMonthKeys.map(v => String(v || '').trim()).filter(Boolean) : [];

                if (safeRows.length === 0 && safeForcedMonths.length === 0) {
                    chartsWrap.style.display = 'none';
                    stackedSvg.innerHTML = '';
                    trendSvg.innerHTML = '';
                    legend.innerHTML = '';
                    if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';
                    return;
                }

                const monthMap = new Map();
                const accountOrder = [];
                const seenAccounts = new Set();
                for (const row of safeRows) {
                    const monthKey = normalizeOmsaetningMonthKey(row.date);
                    if (!monthMap.has(monthKey)) monthMap.set(monthKey, new Map());
                    const accountKey = String(row.acNo || '');
                    if (!seenAccounts.has(accountKey)) {
                        seenAccounts.add(accountKey);
                        accountOrder.push({ acNo: accountKey, name: String(row.name || '') });
                    }
                    const monthAcc = monthMap.get(monthKey);
                    monthAcc.set(accountKey, (monthAcc.get(accountKey) || 0) + Number(row.revenueMio || 0));
                }

                const monthKeys = (safeForcedMonths.length > 0 ? safeForcedMonths : Array.from(monthMap.keys())).sort((a, b) => String(a).localeCompare(String(b)));
                if (monthKeys.length === 0) {
                    chartsWrap.style.display = 'none';
                    stackedSvg.innerHTML = '';
                    trendSvg.innerHTML = '';
                    legend.innerHTML = '';
                    if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';
                    return;
                }
                const monthlyTotals = monthKeys.map(key => {
                    const m = monthMap.get(key) || new Map();
                    let t = 0;
                    for (const value of m.values()) t += Number(value || 0);
                    return t;
                });

                function buildScale(values, fallbackAbs) {
                    const safeValues = Array.isArray(values) ? values : [];
                    let min = 0;
                    let max = 0;
                    for (const raw of safeValues) {
                        const value = Number(raw);
                        if (!Number.isFinite(value)) continue;
                        if (value < min) min = value;
                        if (value > max) max = value;
                    }
                    if (min === 0 && max === 0) {
                        max = Number.isFinite(fallbackAbs) && fallbackAbs > 0 ? fallbackAbs : 0.1;
                    }
                    if (Math.abs(max - min) < 0.000001) {
                        if (max > 0) min = 0;
                        else if (min < 0) max = 0;
                        else max = 0.1;
                    }
                    return { min, max, span: max - min };
                }

                const compareCustomers = Array.from(omsaetningSelectedCustomers.entries())
                    .map(([custNo, name]) => ({ custNo: String(custNo || '').trim(), name: String(name || '').trim() }))
                    .filter(c => c.custNo);
                const showCustomerComparison = compareCustomers.length > 1;

                const leftPad = 48;
                const topPad = 16;
                const bottomPad = 42;
                const chartHeight = 190;
                const innerHeight = chartHeight - topPad - bottomPad;
                const barWidth = 34;
                const barGap = 18;
                const innerWidth = Math.max(560, monthKeys.length * (barWidth + barGap));
                const viewWidth = leftPad + innerWidth + 20;
                const viewHeight = chartHeight;

                function toY(value, scale) {
                    const safeScale = scale || { min: 0, max: 1, span: 1 };
                    const safeValue = Number.isFinite(Number(value)) ? Number(value) : 0;
                    const ratio = (safeScale.max - safeValue) / (safeScale.span || 1);
                    return topPad + (ratio * innerHeight);
                }

                function appendGrid(svgHtml, width, scale) {
                    let out = svgHtml;
                    const ticks = 4;
                    for (let i = 0; i <= ticks; i++) {
                        const ratio = i / ticks;
                        const tickValue = scale.max - (ratio * scale.span);
                        const y = topPad + (innerHeight * ratio);
                        out += '<line x1="' + leftPad + '" y1="' + y + '" x2="' + (leftPad + width) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
                        out += '<text x="' + (leftPad - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="#5f7892">' + escapeHtmlFE(formatMio(tickValue)) + '</text>';
                    }
                    const zeroY = toY(0, scale);
                    out += '<line x1="' + leftPad + '" y1="' + zeroY + '" x2="' + (leftPad + width) + '" y2="' + zeroY + '" stroke="#8fa8c2" stroke-width="1.2" />';
                    return out;
                }

                let stackedSvgHtml = '<g>';
                if (showCustomerComparison) {
                    if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (kunde-sammenligning)';

                    const byMonthCustomer = new Map();
                    for (const row of safeRows) {
                        const monthKey = normalizeOmsaetningMonthKey(row.date);
                        const custNo = String(row.custNo || '').trim();
                        if (!monthKey || !custNo) continue;
                        if (!byMonthCustomer.has(monthKey)) byMonthCustomer.set(monthKey, new Map());
                        const map = byMonthCustomer.get(monthKey);
                        map.set(custNo, (map.get(custNo) || 0) + Number(row.revenueMio || 0));
                    }

                    const groupedValues = [];
                    for (const monthKey of monthKeys) {
                        const values = byMonthCustomer.get(monthKey) || new Map();
                        for (const customer of compareCustomers) {
                            const v = Number(values.get(customer.custNo) || 0);
                            groupedValues.push(v);
                        }
                    }
                    const groupedScale = buildScale(groupedValues, 0.1);
                    const groupedZeroY = toY(0, groupedScale);

                    const groupedBarWidth = 14;
                    const groupedBarGap = 5;
                    const monthGroupGap = 18;
                    const perMonthGroupWidth = (compareCustomers.length * groupedBarWidth) + ((compareCustomers.length - 1) * groupedBarGap);
                    const groupedInnerWidth = Math.max(560, monthKeys.length * (perMonthGroupWidth + monthGroupGap));

                    stackedSvgHtml = appendGrid(stackedSvgHtml, groupedInnerWidth, groupedScale);

                    monthKeys.forEach((monthKey, monthIndex) => {
                        const monthX = leftPad + monthIndex * (perMonthGroupWidth + monthGroupGap);
                        const values = byMonthCustomer.get(monthKey) || new Map();
                        compareCustomers.forEach((customer, customerIndex) => {
                            const value = Number(values.get(customer.custNo) || 0);
                            const yValue = toY(value, groupedScale);
                            const y = value >= 0 ? yValue : groupedZeroY;
                            const h = Math.max(1, Math.abs(groupedZeroY - yValue));
                            const x = monthX + customerIndex * (groupedBarWidth + groupedBarGap);
                            const titleText = formatMonthDa(monthKey) + ' - ' + String(customer.name || customer.custNo) + ' (' + customer.custNo + '): ' + formatMio(value) + ' Mio DKK (' + formatDkkFromMio(value) + ' DKK)';
                            stackedSvgHtml += '<rect x="' + x + '" y="' + y + '" width="' + groupedBarWidth + '" height="' + h + '" fill="' + getOmsaetningColor(customerIndex) + '" rx="2"><title>' + escapeHtmlFE(titleText) + '</title></rect>';
                        });

                        const labelX = monthX + (perMonthGroupWidth / 2);
                        stackedSvgHtml += '<text x="' + labelX + '" y="' + (topPad + innerHeight + 14) + '" text-anchor="middle" font-size="10" fill="#47617c">' + escapeHtmlFE(formatMonthDa(monthKey)) + '</text>';
                    });

                    stackedSvgHtml += '</g>';
                    stackedSvg.setAttribute('viewBox', '0 0 ' + (leftPad + groupedInnerWidth + 20) + ' ' + viewHeight);
                    stackedSvg.innerHTML = stackedSvgHtml;

                    legend.innerHTML = compareCustomers.map((customer, idx) =>
                        '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:' + getOmsaetningColor(idx) + ';"></span>' +
                        escapeHtmlFE(String(customer.name || customer.custNo)) + ' (' + escapeHtmlFE(customer.custNo) + ')</span>'
                    ).join('');
                } else {
                    if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';

                    const stackedMonthTotals = monthKeys.map(monthKey => {
                        const values = monthMap.get(monthKey) || new Map();
                        let pos = 0;
                        let neg = 0;
                        for (const value of values.values()) {
                            const n = Number(value || 0);
                            if (n >= 0) pos += n;
                            else neg += n;
                        }
                        return { pos, neg };
                    });
                    const stackedScale = buildScale(
                        stackedMonthTotals.flatMap(t => [t.pos, t.neg]),
                        0.1
                    );
                    const stackedZeroY = toY(0, stackedScale);

                    stackedSvgHtml = appendGrid(stackedSvgHtml, innerWidth, stackedScale);

                    monthKeys.forEach((monthKey, monthIndex) => {
                        const x = leftPad + monthIndex * (barWidth + barGap);
                        const values = monthMap.get(monthKey) || new Map();
                        let positiveStack = 0;
                        let negativeStack = 0;
                        accountOrder.forEach((acc, accIndex) => {
                            const value = Number(values.get(acc.acNo) || 0);
                            if (value === 0) return;

                            let startValue;
                            let endValue;
                            if (value > 0) {
                                startValue = positiveStack;
                                endValue = positiveStack + value;
                                positiveStack = endValue;
                            } else {
                                startValue = negativeStack;
                                endValue = negativeStack + value;
                                negativeStack = endValue;
                            }

                            const yStart = toY(startValue, stackedScale);
                            const yEnd = toY(endValue, stackedScale);
                            const y = Math.min(yStart, yEnd);
                            const h = Math.max(1, Math.abs(yEnd - yStart));
                            const titleText = formatMonthDa(monthKey) + ' - ' + String(acc.acNo) + ' ' + String(acc.name || '') + ': ' + formatMio(value) + ' Mio DKK (' + formatDkkFromMio(value) + ' DKK)';
                            stackedSvgHtml += '<rect x="' + x + '" y="' + y + '" width="' + barWidth + '" height="' + h + '" fill="' + getOmsaetningColor(accIndex) + '" rx="2"><title>' + escapeHtmlFE(titleText) + '</title></rect>';
                        });

                        if (Math.abs(positiveStack) < 0.000001 && Math.abs(negativeStack) < 0.000001) {
                            stackedSvgHtml += '<line x1="' + x + '" y1="' + stackedZeroY + '" x2="' + (x + barWidth) + '" y2="' + stackedZeroY + '" stroke="#cddced" stroke-width="1" />';
                        }
                        stackedSvgHtml += '<text x="' + (x + barWidth / 2) + '" y="' + (topPad + innerHeight + 14) + '" text-anchor="middle" font-size="10" fill="#47617c">' + escapeHtmlFE(formatMonthDa(monthKey)) + '</text>';
                    });

                    stackedSvgHtml += '</g>';
                    stackedSvg.setAttribute('viewBox', '0 0 ' + viewWidth + ' ' + viewHeight);
                    stackedSvg.innerHTML = stackedSvgHtml;

                    legend.innerHTML = accountOrder.map((acc, idx) =>
                        '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:' + getOmsaetningColor(idx) + ';"></span>' +
                        escapeHtmlFE(String(acc.acNo)) + ' ' + escapeHtmlFE(acc.name || '') + '</span>'
                    ).join('');
                }

                const trendLeftPad = 42;
                const trendTopPad = 16;
                const trendBottomPad = 28;
                const trendHeight = 190;
                const trendInnerHeight = trendHeight - trendTopPad - trendBottomPad;
                const trendInnerWidth = Math.max(560, monthKeys.length * 54);
                const trendViewWidth = trendLeftPad + trendInnerWidth + 16;
                const trendViewHeight = trendHeight;
                const trendScale = buildScale(monthlyTotals, 0.1);

                function toTrendY(value) {
                    const safeValue = Number.isFinite(Number(value)) ? Number(value) : 0;
                    const ratio = (trendScale.max - safeValue) / (trendScale.span || 1);
                    return trendTopPad + (ratio * trendInnerHeight);
                }

                let trendSvgHtml = '<g>';
                for (let i = 0; i <= 4; i++) {
                    const ratio = i / 4;
                    const y = trendTopPad + (trendInnerHeight * ratio);
                    const tickValue = trendScale.max - (ratio * trendScale.span);
                    trendSvgHtml += '<line x1="' + trendLeftPad + '" y1="' + y + '" x2="' + (trendLeftPad + trendInnerWidth) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
                    trendSvgHtml += '<text x="' + (trendLeftPad - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="#5f7892">' + escapeHtmlFE(formatMio(tickValue)) + '</text>';
                    }

                const trendZeroY = toTrendY(0);
                trendSvgHtml += '<line x1="' + trendLeftPad + '" y1="' + trendZeroY + '" x2="' + (trendLeftPad + trendInnerWidth) + '" y2="' + trendZeroY + '" stroke="#8fa8c2" stroke-width="1.2" />';

                const points = monthKeys.map((monthKey, idx) => {
                    const x = trendLeftPad + (trendInnerWidth * (monthKeys.length === 1 ? 0.5 : (idx / (monthKeys.length - 1))));
                    const y = toTrendY(monthlyTotals[idx]);
                    return { x, y, monthKey, total: monthlyTotals[idx] };
                });
                if (points.length === 0) {
                    chartsWrap.style.display = 'none';
                    trendSvg.innerHTML = '';
                    return;
                }
                const linePath = points.map((p, idx) => (idx === 0 ? 'M' : 'L') + p.x + ' ' + p.y).join(' ');
                const areaPath = linePath + ' L ' + points[points.length - 1].x + ' ' + trendZeroY + ' L ' + points[0].x + ' ' + trendZeroY + ' Z';
                trendSvgHtml += '<path d="' + areaPath + '" fill="rgba(21,101,192,0.12)" />';
                trendSvgHtml += '<path d="' + linePath + '" fill="none" stroke="#1565c0" stroke-width="3" />';
                points.forEach(p => {
                    const pointTitle = formatMonthDa(p.monthKey) + ': ' + formatMio(p.total) + ' Mio DKK (' + formatDkkFromMio(p.total) + ' DKK)';
                    trendSvgHtml += '<circle cx="' + p.x + '" cy="' + p.y + '" r="3.5" fill="#0f3560"><title>' + escapeHtmlFE(pointTitle) + '</title></circle>';
                });
                trendSvgHtml += '</g>';

                trendSvg.setAttribute('viewBox', '0 0 ' + trendViewWidth + ' ' + trendViewHeight);
                trendSvg.innerHTML = trendSvgHtml;
                chartsWrap.style.display = 'grid';
            }

            function normalizeWeekKeyInput(value) {
                return String(value || '').replace(/[^0-9]/g, '').slice(0, 6);
            }

            function parseWeekKeyMeta(value) {
                const raw = normalizeWeekKeyInput(value);
                if (!/^[0-9]{6}$/.test(raw)) return null;
                const year = Number(raw.slice(0, 4));
                const week = Number(raw.slice(4, 6));
                if (!Number.isFinite(year) || !Number.isFinite(week) || week < 1 || week > 53) return null;
                return { raw, year, week };
            }

            function getIsoWeekMeta(dateValue) {
                const d = new Date(Date.UTC(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()));
                d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
                const isoYear = d.getUTCFullYear();
                const yearStart = new Date(Date.UTC(isoYear, 0, 1));
                const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
                return {
                    isoYear,
                    week,
                    weekKey: String(isoYear) + String(week).padStart(2, '0')
                };
            }

            function getIsoWeekStartDate(isoYear, week) {
                const jan4 = new Date(Date.UTC(isoYear, 0, 4));
                const jan4Day = jan4.getUTCDay() || 7;
                const mondayWeek1 = new Date(jan4);
                mondayWeek1.setUTCDate(jan4.getUTCDate() - jan4Day + 1);
                const weekStart = new Date(mondayWeek1);
                weekStart.setUTCDate(mondayWeek1.getUTCDate() + ((week - 1) * 7));
                return weekStart;
            }

            function formatWeekLabel(weekKey) {
                const meta = parseWeekKeyMeta(weekKey);
                if (!meta) return String(weekKey || '-');
                return String(meta.year) + '-W' + String(meta.week).padStart(2, '0');
            }

            function formatDkkDa(value) {
                return Number(value || 0).toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            }

            function formatPctDa(value) {
                if (value === null || value === undefined || !Number.isFinite(Number(value))) return '-';
                return Number(value).toLocaleString('da-DK', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + '%';
            }

            function shouldShowOrdreindgangTilbudLine() {
                const el = document.getElementById('ordreindgangShowTilbud');
                return !!el && el.checked === true;
            }

            function setOrdreindgangStatus(message) {
                const el = document.getElementById('ordreindgangStatus');
                if (el) el.textContent = String(message || '');
            }

            function buildModulePrintStyles(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const orientation = safeOptions.orientation === 'landscape' ? 'landscape' : 'portrait';
                const reportMaxWidth = orientation === 'landscape' ? '277mm' : '190mm';
                return '<style>' +
                    '@page { size: A4 ' + orientation + '; margin: 12mm; }' +
                    'body { font-family: Segoe UI, Arial, sans-serif; margin:0; color:#172b3c; background:#fff; }' +
                    '.report { max-width: ' + reportMaxWidth + '; margin:0 auto; }' +
                    '.report-head { border-bottom:2px solid #d9e6f5; padding:0 0 8px 0; margin:0 0 10px 0; }' +
                    '.report-title { font-size:22px; font-weight:800; color:#0f3560; margin:0; }' +
                    '.report-sub { margin:3px 0 0 0; font-size:12px; color:#4b6783; }' +
                    '.report-meta { display:flex; gap:8px 16px; flex-wrap:wrap; margin-top:8px; font-size:12px; color:#355675; }' +
                    '.report-meta strong { color:#0f3560; }' +
                    '.kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:8px; margin:10px 0 12px 0; }' +
                    '.kpi { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#f8fbff; }' +
                    '.kpi .lbl { font-size:11px; color:#4f6d8c; text-transform:uppercase; font-weight:700; }' +
                    '.kpi .val { margin-top:3px; font-size:16px; font-weight:800; color:#0f3560; }' +
                    '.section { margin-top:10px; page-break-inside:avoid; }' +
                    '.section h3 { margin:0 0 6px 0; font-size:14px; color:#214867; border-bottom:1px solid #e2ebf7; padding-bottom:4px; }' +
                    '.chart-box { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#fff; }' +
                    '.chart-box svg { width:100%; height:auto; display:block; }' +
                    '.legend-line { margin:0 0 6px 0; font-size:12px; color:#355675; font-weight:700; }' +
                    '.omsaetning-legend-item,.ordreindgang-legend-item { display:inline-flex; align-items:center; gap:6px; margin-right:10px; font-size:12px; font-weight:700; color:#355675; }' +
                    '.omsaetning-legend-swatch,.ordreindgang-legend-swatch { width:12px; height:12px; border-radius:3px; display:inline-block; }' +
                    '.context-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; }' +
                    '.context-card { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#f8fbff; }' +
                    '.context-card h4 { margin:0 0 6px 0; font-size:12px; color:#214867; text-transform:uppercase; }' +
                    '.context-card p { margin:0; font-size:12px; color:#355675; }' +
                    '.context-line { margin-top:4px; font-size:12px; color:#355675; }' +
                    '.pill-list { display:flex; flex-wrap:wrap; gap:4px; margin-top:6px; }' +
                    '.pill { display:inline-block; padding:2px 7px; border-radius:999px; background:#e8f1fc; color:#1f476a; font-size:11px; font-weight:700; }' +
                    '.report-landscape .kpis { grid-template-columns:repeat(4,minmax(0,1fr)); }' +
                    '.report-landscape table { font-size:11px; }' +
                    '.table-wrap { border:1px solid #dbe8f9; border-radius:8px; overflow:hidden; }' +
                    'table { width:100%; border-collapse:collapse; font-size:12px; }' +
                    'th { background:#1565c0; color:#fff; text-align:left; padding:7px 8px; }' +
                    'td { border-bottom:1px solid #e7eef8; padding:6px 8px; }' +
                    'td[style*="text-align:right"], th[style*="text-align:right"] { text-align:right !important; }' +
                    '.muted { color:#6a829b; font-size:11px; }' +
                    '</style>';
            }

            function openModulePrintWindow(title, subtitle, metaHtml, kpiHtml, sectionsHtml, options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const orientation = safeOptions.orientation === 'landscape' ? 'landscape' : 'portrait';
                const html = '<!doctype html><html><head><meta charset="utf-8" />' +
                    '<title>' + escapeHtmlFE(title) + '</title>' +
                    buildModulePrintStyles({ orientation }) +
                    '</head><body>' +
                    '<div class="report report-' + orientation + '">' +
                    '<header class="report-head">' +
                    '<h1 class="report-title">' + escapeHtmlFE(title) + '</h1>' +
                    '<p class="report-sub">' + escapeHtmlFE(subtitle) + '</p>' +
                    '<div class="report-meta">' + metaHtml + '</div>' +
                    '</header>' +
                    '<section class="kpis">' + kpiHtml + '</section>' +
                    sectionsHtml +
                    '</div>' +
                    '</body></html>';

                const existing = document.getElementById('modulePrintFrame');
                if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

                const frame = document.createElement('iframe');
                frame.id = 'modulePrintFrame';
                frame.setAttribute('aria-hidden', 'true');
                frame.style.position = 'fixed';
                frame.style.right = '0';
                frame.style.bottom = '0';
                frame.style.width = '1px';
                frame.style.height = '1px';
                frame.style.border = '0';
                frame.style.opacity = '0';
                document.body.appendChild(frame);

                const cleanup = () => {
                    const current = document.getElementById('modulePrintFrame');
                    if (current && current.parentNode) current.parentNode.removeChild(current);
                };

                const runPrint = () => {
                    const w = frame.contentWindow;
                    if (!w) {
                        alert('Kunne ikke oprette print-visning. Prøv igen.');
                        cleanup();
                        return;
                    }

                    w.onafterprint = cleanup;
                    w.focus();
                    w.print();
                    setTimeout(cleanup, 2000);
                };

                frame.onload = () => {
                    setTimeout(runPrint, 120);
                };

                // srcdoc avoids popup blockers and is more reliable in Electron than window.open.
                frame.srcdoc = html;
            }

            function getOrientationLabelDa(orientation) {
                return orientation === 'landscape' ? 'Liggende' : 'Stående';
            }

            function getPrintOrientationPreference(moduleKey) {
                const id = moduleKey === 'ordreindgang' ? 'ordreindgangPrintOrientation' : 'omsaetningPrintOrientation';
                const raw = String((document.getElementById(id) || {}).value || 'auto').toLowerCase();
                if (raw === 'portrait' || raw === 'landscape') return raw;
                return 'auto';
            }

            function resolvePrintOrientation(preference, autoOrientation) {
                if (preference === 'portrait' || preference === 'landscape') return preference;
                return autoOrientation === 'landscape' ? 'landscape' : 'portrait';
            }

            function getOrientationSourceLabelDa(preference) {
                return preference === 'auto' ? 'Auto' : 'Manuel';
            }

            function chooseOmsaetningPrintOrientation(context) {
                const c = context && typeof context === 'object' ? context : {};
                if (c.detailsOpen || c.thresholdOpen) return 'landscape';
                if (Number(c.selectedCustomers || 0) > 1) return 'landscape';
                if (Number(c.selectedAccounts || 0) > 6) return 'landscape';
                if (Number(c.periods || 0) > 12) return 'landscape';
                return 'portrait';
            }

            function chooseOrdreindgangPrintOrientation(context) {
                const c = context && typeof context === 'object' ? context : {};
                if (c.weeklyOpen || c.customersOpen) return 'landscape';
                if (Number(c.weeks || 0) > 20) return 'landscape';
                return 'portrait';
            }

            function printOmsaetningReport() {
                const chartsWrap = document.getElementById('omsaetningChartsWrap');
                if (!chartsWrap || chartsWrap.style.display === 'none') {
                    alert('Ingen Omsætning-data at printe. Tryk Opdater først.');
                    return;
                }

                const detailsWrapEl = document.getElementById('omsaetningTableWrap');
                const thresholdWrapEl = document.getElementById('omsaetningThresholdWrap');
                const detailsOpen = !!detailsWrapEl && detailsWrapEl.style.display !== 'none';
                const thresholdOpen = !!thresholdWrapEl && thresholdWrapEl.style.display !== 'none';

                const fra = String((document.getElementById('omsaetningFraMonth') || {}).value || '-');
                const til = String((document.getElementById('omsaetningTilMonth') || {}).value || '-');
                const total = String((document.getElementById('omsaetningTotalMio') || {}).textContent || '-').trim();
                const rows = String((document.getElementById('omsaetningRowsCount') || {}).textContent || '-').trim();
                const periods = String((document.getElementById('omsaetningPeriodsCount') || {}).textContent || '-').trim();
                const stackedSvg = (document.getElementById('omsaetningStackedChart') || {}).outerHTML || '';
                const trendSvg = (document.getElementById('omsaetningTrendChart') || {}).outerHTML || '';
                const legend = (document.getElementById('omsaetningLegend') || {}).innerHTML || '';
                const thresholdTable = (document.getElementById('omsaetningThresholdTable') || {}).innerHTML || '';
                const detailsTable = (document.getElementById('omsaetningTableWrap') || {}).innerHTML || '<div class="muted">Ingen detaljetabel tilgængelig.</div>';
                const customerModeText = String((document.getElementById('omsaetningCustomerMode') || {}).textContent || '').trim();
                const customerThresholdsHtml = (document.getElementById('omsaetningCustomerThresholds') || {}).innerHTML || '';
                const selectedAccounts = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean).map(v => String(v));
                const selectedCustomerEntries = Array.from(omsaetningSelectedCustomers.entries());
                const numericPeriods = Number(periods.replace(/\./g, '').replace(',', '.')) || 0;
                const orientationPreference = getPrintOrientationPreference('omsaetning');
                const autoOrientation = chooseOmsaetningPrintOrientation({
                    detailsOpen,
                    thresholdOpen,
                    selectedCustomers: selectedCustomerEntries.length,
                    selectedAccounts: selectedAccounts.length,
                    periods: numericPeriods
                });
                const orientation = resolvePrintOrientation(orientationPreference, autoOrientation);

                const accountPills = selectedAccounts.slice(0, 18).map(v => '<span class="pill">' + escapeHtmlFE(v) + '</span>').join('');
                const accountRest = selectedAccounts.length > 18
                    ? '<span class="pill">+' + escapeHtmlFE(String(selectedAccounts.length - 18)) + '</span>'
                    : '';
                const customerPills = selectedCustomerEntries.slice(0, 12).map(([custNo, custName]) => {
                    const label = String(custName || '').trim() || String(custNo || '').trim();
                    return '<span class="pill">' + escapeHtmlFE(label + ' (' + String(custNo || '') + ')') + '</span>';
                }).join('');
                const customerRest = selectedCustomerEntries.length > 12
                    ? '<span class="pill">+' + escapeHtmlFE(String(selectedCustomerEntries.length - 12)) + '</span>'
                    : '';

                const metaHtml =
                    '<div><strong>Periode:</strong> ' + escapeHtmlFE(fra + ' → ' + til) + '</div>' +
                    '<div><strong>Layoutvalg:</strong> ' + escapeHtmlFE(getOrientationSourceLabelDa(orientationPreference)) + '</div>' +
                    '<div><strong>Layout:</strong> ' + escapeHtmlFE(getOrientationLabelDa(orientation)) + '</div>' +
                    '<div><strong>Udskrevet:</strong> ' + escapeHtmlFE(new Date().toLocaleString('da-DK')) + '</div>';

                const kpiHtml =
                    '<div class="kpi"><div class="lbl">Omsætning (Mio)</div><div class="val">' + escapeHtmlFE(total) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Rækker</div><div class="val">' + escapeHtmlFE(rows) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Perioder</div><div class="val">' + escapeHtmlFE(periods) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Modul</div><div class="val">Omsætning</div></div>';

                const contextSection =
                    '<section class="section"><h3>Aktive filtre og visning</h3>' +
                        '<div class="context-grid">' +
                            '<div class="context-card">' +
                                '<h4>Konti</h4>' +
                                '<p>' + escapeHtmlFE(String(selectedAccounts.length)) + ' aktiv</p>' +
                                '<div class="pill-list">' + accountPills + accountRest + '</div>' +
                            '</div>' +
                            '<div class="context-card">' +
                                '<h4>Kunder</h4>' +
                                '<p>' + escapeHtmlFE(String(selectedCustomerEntries.length)) + ' valgt</p>' +
                                '<div class="pill-list">' + (selectedCustomerEntries.length > 0 ? (customerPills + customerRest) : '<span class="pill">Ingen kunde</span>') + '</div>' +
                                '<div class="context-line"><strong>Visning:</strong> ' + escapeHtmlFE(customerModeText || 'Standardvisning') + '</div>' +
                            '</div>' +
                        '</div>' +
                    '</section>';

                const customerCompareSection = selectedCustomerEntries.length > 1
                    ? '<section class="section"><h3>Kunde-sammenligning (aktiv)</h3>' +
                        '<div class="context-card"><p>Flere kunder er valgt. Graf og tabeller er baseret på sammenligning pr. måned.</p></div>' +
                        '<div class="context-line"></div>' +
                        '<div class="table-wrap">' + (customerThresholdsHtml || '<div class="muted" style="padding:8px;">Ingen kundetærskler tilgængelig.</div>') + '</div>' +
                    '</section>'
                    : '';

                const thresholdSection = thresholdTable
                    ? '<section class="section"><h3>Tærskel-tabel</h3><div class="table-wrap">' + thresholdTable + '</div></section>'
                    : '';

                const sectionsHtml =
                    contextSection +
                    customerCompareSection +
                    '<section class="section"><h3>Stacked graf</h3><div class="chart-box"><div class="legend-line">' + legend + '</div>' + stackedSvg + '</div></section>' +
                    '<section class="section"><h3>Trend graf</h3><div class="chart-box">' + trendSvg + '</div></section>' +
                    (thresholdOpen ? thresholdSection : '') +
                    (detailsOpen ? '<section class="section"><h3>Detaljer</h3><div class="table-wrap">' + detailsTable + '</div></section>' : '');

                openModulePrintWindow(
                    'Gantech Operations Hub - Omsætning',
                    'Rapportudskrift (' + getOrientationLabelDa(orientation).toLowerCase() + ')',
                    metaHtml,
                    kpiHtml,
                    sectionsHtml,
                    { orientation }
                );
            }

            function printOrdreindgangReport() {
                const chartsWrap = document.getElementById('ordreindgangChartsWrap');
                if (!chartsWrap || chartsWrap.style.display === 'none') {
                    alert('Ingen Ordreindgang-data at printe. Tryk Opdater først.');
                    return;
                }

                const weeklyWrapEl = document.getElementById('ordreindgangWeeklyTable');
                const customersWrapEl = document.getElementById('ordreindgangCustomersTable');
                const weeklyOpen = !!weeklyWrapEl && weeklyWrapEl.style.display !== 'none';
                const customersOpen = !!customersWrapEl && customersWrapEl.style.display !== 'none';

                const fraWeek = String((document.getElementById('ordreindgangFraWeek') || {}).value || '-');
                const tilWeek = String((document.getElementById('ordreindgangTilWeek') || {}).value || '-');
                const totalOrd = String((document.getElementById('ordreindgangTotalOrd') || {}).textContent || '-').trim();
                const totalTilbud = String((document.getElementById('ordreindgangTotalTilbud') || {}).textContent || '-').trim();
                const avgOrd = String((document.getElementById('ordreindgangAvgOrd') || {}).textContent || '-').trim();
                const conv = String((document.getElementById('ordreindgangConv') || {}).textContent || '-').trim();
                const trendSvg = (document.getElementById('ordreindgangTrendChart') || {}).outerHTML || '';
                const legend = (document.getElementById('ordreindgangLegend') || {}).innerHTML || '';
                const weeklyTable = (document.getElementById('ordreindgangWeeklyTable') || {}).innerHTML || '<div class="muted">Ingen ugetabel tilgængelig.</div>';
                const customerTable = (document.getElementById('ordreindgangCustomersTable') || {}).innerHTML || '<div class="muted">Ingen kundetabel tilgængelig.</div>';
                const statusText = String((document.getElementById('ordreindgangStatus') || {}).textContent || '').trim();
                const tilbudEnabled = !!((document.getElementById('ordreindgangShowTilbud') || {}).checked);
                const weekCount = Array.isArray(ordreindgangLastPayload && ordreindgangLastPayload.weeklyRows)
                    ? ordreindgangLastPayload.weeklyRows.length
                    : 0;
                const orientationPreference = getPrintOrientationPreference('ordreindgang');
                const autoOrientation = chooseOrdreindgangPrintOrientation({
                    weeklyOpen,
                    customersOpen,
                    weeks: weekCount
                });
                const orientation = resolvePrintOrientation(orientationPreference, autoOrientation);

                const metaHtml =
                    '<div><strong>Periode:</strong> ' + escapeHtmlFE(fraWeek + ' → ' + tilWeek) + '</div>' +
                    '<div><strong>Tilbud-linje:</strong> ' + (tilbudEnabled ? 'Aktiv' : 'Skjult') + '</div>' +
                    '<div><strong>Layoutvalg:</strong> ' + escapeHtmlFE(getOrientationSourceLabelDa(orientationPreference)) + '</div>' +
                    '<div><strong>Layout:</strong> ' + escapeHtmlFE(getOrientationLabelDa(orientation)) + '</div>' +
                    '<div><strong>Udskrevet:</strong> ' + escapeHtmlFE(new Date().toLocaleString('da-DK')) + '</div>';

                const kpiHtml =
                    '<div class="kpi"><div class="lbl">Total Ordre</div><div class="val">' + escapeHtmlFE(totalOrd) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Total Tilbud</div><div class="val">' + escapeHtmlFE(totalTilbud) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Gns. Ordre</div><div class="val">' + escapeHtmlFE(avgOrd) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Tilbud → Ordre</div><div class="val">' + escapeHtmlFE(conv) + '</div></div>';

                const contextSection =
                    '<section class="section"><h3>Aktive filtre og visning</h3>' +
                        '<div class="context-grid">' +
                            '<div class="context-card">' +
                                '<h4>Periode</h4>' +
                                '<p>' + escapeHtmlFE(fraWeek + ' → ' + tilWeek) + '</p>' +
                            '</div>' +
                            '<div class="context-card">' +
                                '<h4>Tilbud-linje</h4>' +
                                '<p>' + (tilbudEnabled ? 'Aktiv (vises i graf)' : 'Skjult (kun ordre)') + '</p>' +
                                '<div class="context-line"><strong>Status:</strong> ' + escapeHtmlFE(statusText || 'Ingen status') + '</div>' +
                            '</div>' +
                        '</div>' +
                    '</section>';

                const sectionsHtml =
                    contextSection +
                    '<section class="section"><h3>Ugeudvikling</h3><div class="chart-box"><div class="legend-line">' + legend + '</div>' + trendSvg + '</div></section>' +
                    (weeklyOpen ? '<section class="section"><h3>Ugetabel</h3><div class="table-wrap">' + weeklyTable + '</div></section>' : '') +
                    (customersOpen ? '<section class="section"><h3>Topkunder</h3><div class="table-wrap">' + customerTable + '</div></section>' : '');

                openModulePrintWindow(
                    'Gantech Operations Hub - Ordreindgang',
                    'Rapportudskrift (' + getOrientationLabelDa(orientation).toLowerCase() + ')',
                    metaHtml,
                    kpiHtml,
                    sectionsHtml,
                    { orientation }
                );
            }

            function applyOrdreindgangDefaultWeeks() {
                const fraEl = document.getElementById('ordreindgangFraWeek');
                const tilEl = document.getElementById('ordreindgangTilWeek');
                if (!fraEl || !tilEl) return;

                const today = new Date();
                const toWeek = getIsoWeekMeta(today);
                const fromYear = toWeek.isoYear - 1;
                fraEl.value = String(fromYear) + '04';
                tilEl.value = toWeek.weekKey;
            }

            function buildOrdreindgangRange() {
                const fraEl = document.getElementById('ordreindgangFraWeek');
                const tilEl = document.getElementById('ordreindgangTilWeek');
                const fraMeta = parseWeekKeyMeta(fraEl ? fraEl.value : '');
                const tilMeta = parseWeekKeyMeta(tilEl ? tilEl.value : '');
                if (fraEl) fraEl.value = normalizeWeekKeyInput(fraEl.value);
                if (tilEl) tilEl.value = normalizeWeekKeyInput(tilEl.value);
                if (!fraMeta || !tilMeta) return null;

                const fromStart = getIsoWeekStartDate(fraMeta.year, fraMeta.week);
                const toStart = getIsoWeekStartDate(tilMeta.year, tilMeta.week);
                if (toStart.getTime() < fromStart.getTime()) return null;

                return {
                    fraWeek: fraMeta.raw,
                    tilWeek: tilMeta.raw
                };
            }

            function buildOrdreindgangSummaryCacheKey(range) {
                return JSON.stringify({
                    fraWeek: String(range.fraWeek || ''),
                    tilWeek: String(range.tilWeek || '')
                });
            }

            async function fetchOrdreindgangSummaryCached(range, options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const forceRefresh = safeOptions.forceRefresh === true;
                const cacheKey = buildOrdreindgangSummaryCacheKey(range);

                if (!forceRefresh) {
                    const cached = getOmsaetningCacheEntry(ordreindgangSummaryCache, cacheKey, ORDREINDGANG_SUMMARY_CACHE_TTL_MS);
                    if (cached) return cached;
                }

                if (!forceRefresh) {
                    const inFlight = ordreindgangSummaryInFlight.get(cacheKey);
                    if (inFlight) return await inFlight;
                }

                const reqPromise = (async () => {
                    const query = new URLSearchParams({
                        fraWeek: String(range.fraWeek),
                        tilWeek: String(range.tilWeek)
                    });
                    const response = await fetch('/ordreindgang/summary?' + query.toString());
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    setOmsaetningCacheEntry(ordreindgangSummaryCache, cacheKey, payload);
                    return payload;
                })();

                ordreindgangSummaryInFlight.set(cacheKey, reqPromise);
                try {
                    return await reqPromise;
                } finally {
                    ordreindgangSummaryInFlight.delete(cacheKey);
                }
            }

            function renderOrdreindgangTrendChart(rows) {
                const wrap = document.getElementById('ordreindgangChartsWrap');
                const svg = document.getElementById('ordreindgangTrendChart');
                const legendEl = document.getElementById('ordreindgangLegend');
                const safeRows = Array.isArray(rows) ? rows : [];
                const showTilbudLine = shouldShowOrdreindgangTilbudLine();
                if (!wrap || !svg) return;

                if (safeRows.length === 0) {
                    wrap.style.display = 'none';
                    svg.innerHTML = '';
                    if (legendEl) legendEl.innerHTML = '';
                    return;
                }

                const labels = safeRows.map(r => {
                    const key = String(r.weekKey || '');
                    if (/^\\d{6}$/.test(key)) return key.slice(0, 4) + '-' + key.slice(4, 6);
                    return formatWeekLabel(key);
                });
                const ordValues = safeRows.map(r => Number(r.totalOrd || 0));
                const tilbudValues = safeRows.map(r => Number(r.totalTilbud || 0));
                const avgValue = Number(safeRows[0] && safeRows[0].avgOrd || 0);
                const allValues = showTilbudLine
                    ? ordValues.concat(tilbudValues).concat([avgValue])
                    : ordValues.concat([avgValue]);
                const rawMax = Math.max(...allValues, 0);
                const chartMax = rawMax <= 0 ? 1000 : (Math.ceil(rawMax / 1000) * 1000);

                const leftPad = 72;
                const topPad = 36;
                const bottomPad = 108;
                const viewportWidth = Math.max(760, (wrap.clientWidth || 0) - 24);
                const visibleWeeksTarget = Math.max(12, Math.min(24, safeRows.length || 12));
                const slot = Math.max(18, Math.min(36, Math.floor(viewportWidth / visibleWeeksTarget)));
                const width = Math.max(viewportWidth, safeRows.length * slot);
                const height = Math.max(320, Math.min(440, Math.round(window.innerHeight * 0.46)));
                const innerHeight = height - topPad - bottomPad;
                const viewWidth = leftPad + width + 20;
                const toY = val => topPad + ((chartMax - Math.max(0, Number(val || 0))) / chartMax) * innerHeight;
                const toXCenter = idx => leftPad + (idx * slot) + (slot / 2);
                const yZero = toY(0);

                let html = '<g>';
                for (let i = 0; i <= 6; i++) {
                    const ratio = i / 6;
                    const y = topPad + (innerHeight * ratio);
                    const tickVal = chartMax * (1 - ratio);
                    html += '<line x1="' + leftPad + '" y1="' + y + '" x2="' + (leftPad + width) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
                    html += '<text x="' + (leftPad - 8) + '" y="' + (y + 5) + '" text-anchor="end" font-size="12" fill="#5f7892">' + escapeHtmlFE(Math.round(tickVal).toLocaleString('da-DK')) + '</text>';
                }

                const ordBarWidth = showTilbudLine ? 8 : 12;
                const tilbudBarWidth = 7;

                const labelStep = safeRows.length > 70 ? 4 : (safeRows.length > 52 ? 3 : (safeRows.length > 36 ? 2 : 1));
                safeRows.forEach((row, idx) => {
                    const centerX = toXCenter(idx);
                    const ordY = toY(ordValues[idx]);
                    const ordH = Math.max(1, yZero - ordY);

                    if (showTilbudLine) {
                        const tilbudY = toY(tilbudValues[idx]);
                        const tilbudH = Math.max(1, yZero - tilbudY);
                        const tilbudX = centerX - tilbudBarWidth - 1;
                        html += '<rect x="' + tilbudX + '" y="' + tilbudY + '" width="' + tilbudBarWidth + '" height="' + tilbudH + '" fill="#8ec3f7" rx="1"><title>' +
                            escapeHtmlFE(labels[idx] + ' Tilbud: ' + formatDkkDa(tilbudValues[idx])) + '</title></rect>';
                    }

                    const ordX = showTilbudLine ? (centerX + 1) : (centerX - (ordBarWidth / 2));
                    html += '<rect x="' + ordX + '" y="' + ordY + '" width="' + ordBarWidth + '" height="' + ordH + '" fill="#2f5ea5" rx="1"><title>' +
                        escapeHtmlFE(labels[idx] + ' Ordre: ' + formatDkkDa(ordValues[idx])) + '</title></rect>';

                    if ((idx % labelStep) === 0 || idx === safeRows.length - 1) {
                        html += '<text x="' + centerX + '" y="' + (topPad + innerHeight + 50) + '" text-anchor="middle" transform="rotate(-90 ' + centerX + ' ' + (topPad + innerHeight + 50) + ')" font-size="12" fill="#5f7892">' + escapeHtmlFE(labels[idx]) + '</text>';
                    }
                });

                const avgY = toY(avgValue);
                html += '<line x1="' + leftPad + '" y1="' + avgY + '" x2="' + (leftPad + width) + '" y2="' + avgY + '" stroke="#3d6eb5" stroke-width="3" />';
                html += '</g>';

                svg.setAttribute('viewBox', '0 0 ' + viewWidth + ' ' + height);
                svg.innerHTML = html;
                if (legendEl) {
                    let legendHtml = '';
                    legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#2f5ea5"></span>Ordre</span>';
                    if (showTilbudLine) {
                        legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#8ec3f7"></span>Tilbud</span>';
                    }
                    legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#3d6eb5"></span>Gennem.Ordre</span>';
                    legendEl.innerHTML = legendHtml;
                }
                wrap.style.display = 'grid';
            }

            function applyOrdreindgangWeeklyCollapsedState() {
                const tableWrap = document.getElementById('ordreindgangWeeklyTable');
                const toggleBtn = document.getElementById('ordreindgangWeeklyToggleBtn');
                if (!tableWrap || !toggleBtn) return;
                tableWrap.style.display = ordreindgangWeeklyCollapsed ? 'none' : 'block';
                toggleBtn.textContent = ordreindgangWeeklyCollapsed ? 'Vis tabel' : 'Skjul tabel';
            }

            function applyOrdreindgangCustomersCollapsedState() {
                const tableWrap = document.getElementById('ordreindgangCustomersTable');
                const toggleBtn = document.getElementById('ordreindgangCustomersToggleBtn');
                if (!tableWrap || !toggleBtn) return;
                tableWrap.style.display = ordreindgangCustomersCollapsed ? 'none' : 'block';
                toggleBtn.textContent = ordreindgangCustomersCollapsed ? 'Vis tabel' : 'Skjul tabel';
            }

            function toggleOrdreindgangWeeklyTable() {
                ordreindgangWeeklyCollapsed = !ordreindgangWeeklyCollapsed;
                applyOrdreindgangWeeklyCollapsedState();
            }

            function toggleOrdreindgangCustomersTable() {
                ordreindgangCustomersCollapsed = !ordreindgangCustomersCollapsed;
                applyOrdreindgangCustomersCollapsedState();
            }

            function renderOrdreindgangWeeklyTable(rows) {
                const wrapCard = document.getElementById('ordreindgangWeeklyWrap');
                const wrap = document.getElementById('ordreindgangWeeklyTable');
                const safeRows = Array.isArray(rows) ? rows : [];
                const showTilbudColumn = shouldShowOrdreindgangTilbudLine();
                if (!wrapCard || !wrap) return;
                if (safeRows.length === 0) {
                    wrapCard.style.display = 'none';
                    wrap.innerHTML = '';
                    return;
                }

                const body = safeRows.map(row => {
                    const cells = [
                        '<td>' + escapeHtmlFE(formatWeekLabel(row.weekKey)) + '</td>',
                        '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalOrd)) + '</td>'
                    ];
                    if (showTilbudColumn) {
                        cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalTilbud)) + '</td>');
                    }
                    cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalBudget)) + '</td>');
                    cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.avgOrd)) + '</td>');
                    return '<tr>' +
                        cells.join('') +
                        '</tr>';
                }).join('');

                const headers = [
                    '<th>Uge</th>',
                    '<th class="omsaetning-cell-right">Ordre</th>'
                ];
                if (showTilbudColumn) {
                    headers.push('<th class="omsaetning-cell-right">Tilbud</th>');
                }
                headers.push('<th class="omsaetning-cell-right">Budget</th>');
                headers.push('<th class="omsaetning-cell-right">Gns. ordre</th>');

                const colgroup = showTilbudColumn
                    ? '<colgroup><col style="width:16%;" /><col style="width:21%;" /><col style="width:23%;" /><col style="width:20%;" /><col style="width:20%;" /></colgroup>'
                    : '<colgroup><col style="width:18%;" /><col style="width:28%;" /><col style="width:27%;" /><col style="width:27%;" /></colgroup>';

                wrap.innerHTML = '<table class="omsaetning-table ordreindgang-weekly-table">' +
                    colgroup +
                    '<thead><tr>' + headers.join('') + '</tr></thead>' +
                    '<tbody>' + body + '</tbody></table>';
                wrapCard.style.display = 'block';
                applyOrdreindgangWeeklyCollapsedState();
            }

            function renderOrdreindgangCustomersTable(rows) {
                const wrapCard = document.getElementById('ordreindgangCustomersWrap');
                const wrap = document.getElementById('ordreindgangCustomersTable');
                const safeRows = Array.isArray(rows) ? rows : [];
                if (!wrapCard || !wrap) return;
                if (safeRows.length === 0) {
                    wrapCard.style.display = 'none';
                    wrap.innerHTML = '';
                    return;
                }

                const body = safeRows.map(row => {
                    const label = String(row.customerName || '').trim() || String(row.custNo || '-');
                    return '<tr>' +
                        '<td>' + escapeHtmlFE(label) + '</td>' +
                        '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.ordSum)) + '</td>' +
                        '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.tilbudSum)) + '</td>' +
                        '<td style="text-align:right;">' + escapeHtmlFE(formatPctDa(row.conversionPct)) + '</td>' +
                        '</tr>';
                }).join('');

                wrap.innerHTML = '<table class="omsaetning-table">' +
                    '<thead><tr><th>Kunde</th><th>Ordre</th><th>Tilbud</th><th>Tilbud → Ordre</th></tr></thead>' +
                    '<tbody>' + body + '</tbody></table>';
                wrapCard.style.display = 'block';
                applyOrdreindgangCustomersCollapsedState();
            }

            function scheduleOrdreindgangAutoReload() {
                if (ordreindgangAutoReloadTimer) {
                    clearTimeout(ordreindgangAutoReloadTimer);
                }
                ordreindgangAutoReloadTimer = setTimeout(() => {
                    ordreindgangAutoReloadTimer = null;
                    loadOrdreindgangSummary({ silentValidation: true });
                }, ORDREINDGANG_AUTO_RELOAD_DELAY_MS);
            }

            async function initializeOrdreindgangIfNeeded() {
                if (ordreindgangInitialized) return;
                ordreindgangInitialized = true;
                window.addEventListener('resize', () => {
                    if (ordreindgangResizeTimer) clearTimeout(ordreindgangResizeTimer);
                    ordreindgangResizeTimer = setTimeout(() => {
                        ordreindgangResizeTimer = null;
                        renderOrdreindgangFromLastPayload();
                    }, 120);
                });
                applyOrdreindgangDefaultWeeks();
                await loadOrdreindgangSummary({ forceRefresh: true });
            }

            function renderOrdreindgangFromLastPayload() {
                if (!ordreindgangLastPayload || !Array.isArray(ordreindgangLastPayload.weeklyRows)) return;
                renderOrdreindgangTrendChart(ordreindgangLastPayload.weeklyRows);
                renderOrdreindgangWeeklyTable(ordreindgangLastPayload.weeklyRows);
            }

            async function loadOrdreindgangSummary(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const emptyEl = document.getElementById('ordreindgangEmpty');
                const loadBtn = document.getElementById('ordreindgangLoadBtn');
                const range = buildOrdreindgangRange();

                if (!range) {
                    setOrdreindgangStatus('Ugyldig ugeperiode. Brug format YYYYWW.');
                    if (safeOptions.silentValidation !== true) {
                        alert('Ugyldig ugeperiode. Brug format YYYYWW.');
                    }
                    return;
                }

                if (loadBtn) loadBtn.disabled = true;
                if (emptyEl) {
                    emptyEl.style.display = 'block';
                    emptyEl.textContent = 'Henter ordreindgang...';
                }
                setOrdreindgangStatus('Henter data...');

                try {
                    const payload = await fetchOrdreindgangSummaryCached(range, { forceRefresh: safeOptions.forceRefresh === true });
                    ordreindgangLastPayload = payload;
                    const kpis = payload && payload.kpis ? payload.kpis : {};
                    const weeklyRows = Array.isArray(payload && payload.weeklyRows) ? payload.weeklyRows : [];
                    const customerRows = Array.isArray(payload && payload.customerRows) ? payload.customerRows : [];

                    const ordEl = document.getElementById('ordreindgangTotalOrd');
                    const tilbudEl = document.getElementById('ordreindgangTotalTilbud');
                    const avgEl = document.getElementById('ordreindgangAvgOrd');
                    const convEl = document.getElementById('ordreindgangConv');

                    if (ordEl) ordEl.textContent = formatDkkDa(kpis.totalOrdSum || 0);
                    if (tilbudEl) tilbudEl.textContent = formatDkkDa(kpis.totalTilbudSum || 0);
                    if (avgEl) avgEl.textContent = formatDkkDa(kpis.avgSumOrd || 0);
                    if (convEl) convEl.textContent = formatPctDa(kpis.conversionPct);

                    renderOrdreindgangTrendChart(weeklyRows);
                    renderOrdreindgangWeeklyTable(weeklyRows);
                    renderOrdreindgangCustomersTable(customerRows);

                    if (emptyEl) {
                        if (weeklyRows.length === 0) {
                            emptyEl.style.display = 'block';
                            emptyEl.textContent = 'Ingen data i valgt ugeperiode.';
                        } else {
                            emptyEl.style.display = 'none';
                        }
                    }

                    setOrdreindgangStatus('Periode: ' + formatWeekLabel(range.fraWeek) + ' til ' + formatWeekLabel(range.tilWeek) + ' · rækker: ' + String(weeklyRows.length));
                } catch (err) {
                    setOrdreindgangStatus('Fejl ved hentning af ordreindgang.');
                    if (emptyEl) {
                        emptyEl.style.display = 'block';
                        emptyEl.textContent = 'Fejl: ' + (err && err.message ? err.message : 'ukendt fejl');
                    }
                } finally {
                    if (loadBtn) loadBtn.disabled = false;
                }
            }

            function setBelastningStatus(message) {
                const el = document.getElementById('belastningStatus');
                if (el) el.textContent = String(message || '');
            }

            function scrollBelastningDetailIntoView() {
                const detailEl = document.getElementById('belastningDetailSvg') || document.getElementById('belastningDetailWrap');
                if (!detailEl) return;
                const rect = detailEl.getBoundingClientRect();
                const headerOffset = 86;
                const top = Math.max(0, (window.scrollY || 0) + rect.top - headerOffset);
                window.scrollTo({ top, behavior: 'smooth' });
            }

            function getBelastningFilters() {
                const todayInput = document.getElementById('belastningToDay');
                const daysInput = document.getElementById('belastningDage');
                const resGrInput = document.getElementById('belastningResGr');
                const orderInput = document.getElementById('belastningOrdre');
                const customerInput = document.getElementById('belastningKunde');
                const today = String(todayInput && todayInput.value || '').trim() || new Date().toISOString().slice(0, 10);
                const daysRaw = Number(daysInput && daysInput.value || 30);
                const dage = Number.isFinite(daysRaw) ? Math.max(1, Math.min(180, Math.round(daysRaw))) : 30;
                const resGr = String(resGrInput && resGrInput.value || '').trim();
                const ord = String(orderInput && orderInput.value || '').replace(/\\D+/g, '').slice(0, 12);
                const kunde = String(customerInput && customerInput.value || '').trim().replace(/\\s+/g, ' ').slice(0, 80);
                if (orderInput && orderInput.value !== ord) {
                    orderInput.value = ord;
                }
                if (customerInput && customerInput.value !== kunde) {
                    customerInput.value = kunde;
                }
                return { today, dage, resGr, ord, kunde };
            }

            function escapeJsSingle(value) {
                return String(value || '').split('\\\\').join('\\\\\\\\').split("'").join("\\\\'");
            }

            function formatBelastningMinutes(value) {
                const num = Number(value || 0);
                if (!Number.isFinite(num)) return '0';
                return new Intl.NumberFormat('da-DK', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(Math.round(num));
            }

            function getBelastningLayoutStorageKey() {
                return 'afterkalk_belastning_layout_v1_' + sanitizeDisplayName(loggedUserDisplayName).toLowerCase();
            }

            function getBelastningCardKey(item) {
                const rg = String(item && item.resGr || '').trim();
                const parity = item && Number(item.parity) === 0 ? 0 : 1;
                return rg ? (String(parity) + ':' + rg) : '';
            }

            function readBelastningLayoutOrder() {
                try {
                    const raw = localStorage.getItem(getBelastningLayoutStorageKey());
                    const parsed = JSON.parse(String(raw || '[]'));
                    return Array.isArray(parsed) ? parsed.map(v => String(v || '').trim()).filter(Boolean) : [];
                } catch {
                    return [];
                }
            }

            function writeBelastningLayoutOrder(keys) {
                try {
                    localStorage.setItem(getBelastningLayoutStorageKey(), JSON.stringify(Array.from(new Set((Array.isArray(keys) ? keys : []).map(v => String(v || '').trim()).filter(Boolean)))));
                } catch {}
            }

            function sortBelastningItemsBySavedLayout(items) {
                const safeItems = Array.isArray(items) ? items.slice() : [];
                const order = readBelastningLayoutOrder();
                const orderIndex = new Map(order.map((key, index) => [key, index]));
                return safeItems.sort((a, b) => {
                    const keyA = getBelastningCardKey(a);
                    const keyB = getBelastningCardKey(b);
                    const idxA = orderIndex.has(keyA) ? orderIndex.get(keyA) : Number.MAX_SAFE_INTEGER;
                    const idxB = orderIndex.has(keyB) ? orderIndex.get(keyB) : Number.MAX_SAFE_INTEGER;
                    if (idxA !== idxB) return idxA - idxB;
                    return String(a && a.resGr || '').localeCompare(String(b && b.resGr || ''), 'da');
                });
            }

            function clearBelastningDragMarkers(root) {
                const host = root || document;
                host.querySelectorAll('.belastning-resource-chart.drag-target-before, .belastning-resource-chart.drag-target-after').forEach(el => {
                    el.classList.remove('drag-target-before', 'drag-target-after');
                });
            }

            function persistBelastningLayoutFromDom(targetId) {
                const wrap = document.getElementById(targetId);
                if (!wrap) return;
                const visibleKeys = Array.from(wrap.querySelectorAll('.belastning-resource-chart[data-belastning-key]'))
                    .map(el => String(el.getAttribute('data-belastning-key') || '').trim())
                    .filter(Boolean);
                if (visibleKeys.length === 0) return;
                const previousKeys = readBelastningLayoutOrder().filter(key => !visibleKeys.includes(key));
                writeBelastningLayoutOrder([...visibleKeys, ...previousKeys]);
            }

            function attachBelastningDragAndDrop(targetId) {
                const wrap = document.getElementById(targetId);
                if (!wrap || wrap.dataset.dragReady === '1') return;
                wrap.dataset.dragReady = '1';

                wrap.addEventListener('dragstart', event => {
                    const card = event.target && event.target.closest ? event.target.closest('.belastning-resource-chart[data-belastning-key]') : null;
                    if (!card) return;
                    belastningDraggedCardKey = String(card.getAttribute('data-belastning-key') || '').trim();
                    if (!belastningDraggedCardKey) return;
                    card.classList.add('is-dragging');
                    if (event.dataTransfer) {
                        event.dataTransfer.effectAllowed = 'move';
                        event.dataTransfer.setData('text/plain', belastningDraggedCardKey);
                    }
                });

                wrap.addEventListener('dragover', event => {
                    if (!belastningDraggedCardKey) return;
                    event.preventDefault();
                    const target = event.target && event.target.closest ? event.target.closest('.belastning-resource-chart[data-belastning-key]') : null;
                    clearBelastningDragMarkers(wrap);
                    if (!target || String(target.getAttribute('data-belastning-key') || '').trim() === belastningDraggedCardKey) return;
                    const rect = target.getBoundingClientRect();
                    const insertBefore = event.clientY < rect.top + (rect.height / 2);
                    target.classList.add(insertBefore ? 'drag-target-before' : 'drag-target-after');
                });

                wrap.addEventListener('drop', event => {
                    if (!belastningDraggedCardKey) return;
                    event.preventDefault();
                    const dragged = wrap.querySelector('.belastning-resource-chart[data-belastning-key="' + cssEscape(belastningDraggedCardKey) + '"]');
                    const target = event.target && event.target.closest ? event.target.closest('.belastning-resource-chart[data-belastning-key]') : null;
                    clearBelastningDragMarkers(wrap);
                    if (!dragged || !target || dragged === target) return;
                    const rect = target.getBoundingClientRect();
                    const insertBefore = event.clientY < rect.top + (rect.height / 2);
                    if (insertBefore) {
                        wrap.insertBefore(dragged, target);
                    } else {
                        wrap.insertBefore(dragged, target.nextSibling);
                    }
                    persistBelastningLayoutFromDom(targetId);
                });

                wrap.addEventListener('dragend', () => {
                    clearBelastningDragMarkers(wrap);
                    wrap.querySelectorAll('.belastning-resource-chart.is-dragging').forEach(el => el.classList.remove('is-dragging'));
                    belastningDraggedCardKey = '';
                });
            }

            function cssEscape(value) {
                const bs = String.fromCharCode(92);
                return String(value || '')
                    .split(bs).join(bs + bs)
                    .split('"').join(bs + '"');
            }

            function normalizeBelastningDateKey(rawDate, dateLabel) {
                const label = String(dateLabel || '').trim();
                const fullDa = label.match(/^(\\d{1,2})\\.(\\d{1,2})\\.(\\d{4})$/);
                if (fullDa) {
                    const dd = fullDa[1].padStart(2, '0');
                    const mm = fullDa[2].padStart(2, '0');
                    const yy = fullDa[3];
                    return yy + '-' + mm + '-' + dd;
                }
                const parsed = rawDate ? new Date(rawDate) : null;
                if (parsed && !Number.isNaN(parsed.getTime())) {
                    const y = parsed.getFullYear();
                    const m = String(parsed.getMonth() + 1).padStart(2, '0');
                    const d = String(parsed.getDate()).padStart(2, '0');
                    return y + '-' + m + '-' + d;
                }
                return '';
            }

            function normalizeBelastningDisplayDate(rawDate, dateLabel) {
                const label = String(dateLabel || '').trim();
                if (/^\\d{1,2}\\.\\d{1,2}\\.\\d{4}$/.test(label)) return label;
                const parsed = rawDate ? new Date(rawDate) : null;
                if (parsed && !Number.isNaN(parsed.getTime())) {
                    return parsed.toLocaleDateString('da-DK');
                }
                return label || '-';
            }

            function getBelastningDateSortValue(dateKey) {
                if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(String(dateKey || ''))) return Number.MAX_SAFE_INTEGER;
                const ts = new Date(dateKey + 'T00:00:00').getTime();
                return Number.isNaN(ts) ? Number.MAX_SAFE_INTEGER : ts;
            }

            async function onBelastningDayColumnClick(resGr, parity, dayKey, event) {
                if (event && typeof event.stopPropagation === 'function') {
                    event.stopPropagation();
                }
                const safeDayKey = String(dayKey || '').trim();
                if (!safeDayKey) return;
                await loadBelastningDetail(resGr, parity, { focusDayKey: safeDayKey });
            }

            function scheduleBelastningAutoReload() {
                if (belastningAutoReloadTimer) {
                    clearTimeout(belastningAutoReloadTimer);
                }
                belastningAutoReloadTimer = setTimeout(() => {
                    belastningAutoReloadTimer = null;
                    loadBelastningGrafisk({ forceRefresh: true, silentValidation: true });
                }, BELASTNING_FILTER_DEBOUNCE_MS);
            }

            function startBelastningPeriodicRefresh() {
                if (belastningPeriodicTimer) return;
                belastningPeriodicTimer = setInterval(() => {
                    const root = document.getElementById('mainBelastning');
                    const isVisible = root && root.style.display !== 'none';
                    if (isVisible) {
                        loadBelastningGrafisk({ forceRefresh: true, silentValidation: true });
                    }
                }, BELASTNING_PERIODIC_REFRESH_MS);
            }

            function buildBelastningClusterSvg(dayRows, opts) {
                const options = opts && typeof opts === 'object' ? opts : {};
                const clickable = options.clickable === true;
                const chartResGr = String(options.resGr || '').trim();
                const chartParity = options.parity === 0 ? 0 : 1;
                const activeDayKey = String(options.activeDayKey || '').trim();
                const sourceRows = Array.isArray(dayRows) ? dayRows : [];
                const collapseBeforeToday = options.collapseBeforeToday !== false;
                const activeDateInput = document.getElementById('belastningToDay');
                const todayRaw = (activeDateInput && activeDateInput.value)
                    ? activeDateInput.value
                    : new Date().toISOString().slice(0, 10);
                const todayCut = (() => {
                    const t = new Date(todayRaw + 'T00:00:00').getTime();
                    return Number.isNaN(t) ? Date.now() : t;
                })();

                let rows = [];
                if (collapseBeforeToday) {
                    const groupedByDay = new Map();
                    sourceRows.forEach((row, index) => {
                        const dateKey = normalizeBelastningDateKey(row && row.Dato, row && row.DatoX);
                        const dateSort = getBelastningDateSortValue(dateKey);
                        const isNullDate = !row.Dato && !String(row.DatoX || '').trim();
                        const isBefore = isNullDate || dateSort < todayCut;
                        const bucketKey = isBefore ? '__before__' : (dateKey || ('__unknown__' + index));
                        if (!groupedByDay.has(bucketKey)) {
                            groupedByDay.set(bucketKey, {
                                Dato: isBefore ? null : (row && row.Dato),
                                DatoX: isBefore ? '-' : normalizeBelastningDisplayDate(row && row.Dato, row && row.DatoX),
                                Kap: 0,
                                Resv: 0,
                                Aften: 0,
                                __dayKey: isBefore ? 'before' : dateKey,
                                __dateLabel: isBefore ? '-' : normalizeBelastningDisplayDate(row && row.Dato, row && row.DatoX),
                                __sort: isBefore ? Number.MIN_SAFE_INTEGER : dateSort
                            });
                        }
                        const bucket = groupedByDay.get(bucketKey);
                        bucket.Kap += isBefore ? 0 : Number(row && row.Kap || 0);
                        bucket.Resv += Number(row && row.Resv || 0);
                        bucket.Aften += Number(row && row.Aften || 0);
                    });
                    rows = Array.from(groupedByDay.values()).sort((a, b) => Number(a.__sort || 0) - Number(b.__sort || 0));
                } else {
                    rows = sourceRows.slice().map((row, index) => {
                        const dateKey = normalizeBelastningDateKey(row && row.Dato, row && row.DatoX);
                        return {
                            ...row,
                            __dayKey: dateKey || ('__unknown__' + index),
                            __dateLabel: normalizeBelastningDisplayDate(row && row.Dato, row && row.DatoX),
                            __sort: getBelastningDateSortValue(dateKey)
                        };
                    }).sort((a, b) => Number(a.__sort || 0) - Number(b.__sort || 0));
                }

                if (rows.length === 0) return '';

                const leftPad = 40;
                const rightPad = 12;
                const topPad = 22;
                const bottomPad = 86;
                const parentWidth = Math.max(420, (window.innerWidth || 1200) * 0.42);
                const targetDaysOnScreen = Math.max(12, Math.min(30, rows.length));
                const groupW = Math.max(16, Math.min(30, Math.floor(parentWidth / Math.max(1, targetDaysOnScreen))));
                const innerW = Math.max(parentWidth, rows.length * groupW);
                const innerH = 180;
                const svgW = leftPad + innerW + rightPad;
                const svgH = topPad + innerH + bottomPad;
                const maxVal = rows.reduce((max, row) => {
                    return Math.max(max, Number(row.Kap || 0), Number(row.Resv || 0), Number(row.Aften || 0));
                }, 1);

                const yFor = (v) => topPad + innerH - (Math.max(0, Number(v || 0)) / maxVal) * innerH;
                const xFor = (i) => leftPad + i * groupW + (groupW / 2);
                const barW = Math.max(3, Math.min(7, groupW / 3.4));

                const grid = [];
                for (let i = 0; i <= 4; i++) {
                    const y = topPad + (innerH * i / 4);
                    const val = formatCount(Math.round(maxVal * (1 - i / 4)));
                    grid.push('<line class="grid" x1="' + leftPad + '" y1="' + y + '" x2="' + (svgW - rightPad) + '" y2="' + y + '"></line>');
                    grid.push('<text class="label" x="' + (leftPad - 4) + '" y="' + (y + 3) + '" text-anchor="end">' + val + '</text>');
                }

                const bars = rows.map((row, i) => {
                    const x = xFor(i);
                    const kap = Number(row.Kap || 0);
                    const resv = Number(row.Resv || 0);
                    const aften = Number(row.Aften || 0);
                    const dateLabel = String(row.__dateLabel || normalizeBelastningDisplayDate(row.Dato, row.DatoX));
                    const dayKey = String(row.__dayKey || normalizeBelastningDateKey(row.Dato, row.DatoX));
                    const ky = yFor(kap);
                    const ry = yFor(resv);
                    const ay = yFor(aften);
                    const kH = topPad + innerH - ky;
                    const rH = topPad + innerH - ry;
                    const aH = topPad + innerH - ay;
                    const hitX = x - (barW * 1.9);
                    const hitW = Math.max(barW * 3.8, 10);
                    const canClick = clickable && chartResGr && dayKey;
                    const clickAttr = canClick
                        ? (' onclick="onBelastningDayColumnClick(\\'' + escapeJsSingle(chartResGr) + '\\',' + chartParity + ',\\'' + escapeJsSingle(dayKey) + '\\', event)"')
                        : '';
                    const bandClass = 'belastning-day-band' + (activeDayKey && dayKey === activeDayKey ? ' active' : '');
                    return ''
                        + '<rect class="' + bandClass + '" x="' + hitX + '" y="' + topPad + '" width="' + hitW + '" height="' + innerH + '"' + clickAttr + '></rect>'
                        + '<rect class="belastning-series-kap" x="' + (x - barW * 1.5) + '" y="' + ky + '" width="' + barW + '" height="' + kH + '"><title>' + escapeHtmlFE(dateLabel + ' Kapacitet: ' + formatBelastningMinutes(kap)) + '</title></rect>'
                        + '<rect class="belastning-series-resv" x="' + (x - barW * 0.5) + '" y="' + ry + '" width="' + barW + '" height="' + rH + '"><title>' + escapeHtmlFE(dateLabel + ' Reservationer: ' + formatBelastningMinutes(resv)) + '</title></rect>'
                        + '<rect class="belastning-series-aften" x="' + (x + barW * 0.5) + '" y="' + ay + '" width="' + barW + '" height="' + aH + '"><title>' + escapeHtmlFE(dateLabel + ' Rest Aften: ' + formatBelastningMinutes(aften)) + '</title></rect>';
                }).join('');

                // Show every day label; dates are already rotated to reduce overlap.
                const labelStep = 1;
                const labels = rows.map((row, i) => {
                    const isToday = row.__dayKey === todayRaw;
                    if (i % labelStep !== 0 && i !== rows.length - 1 && !isToday) return '';
                    const txt = escapeHtmlFE(String(row.__dateLabel || normalizeBelastningDisplayDate(row.Dato, row.DatoX) || ''));
                    const labelY = topPad + innerH + 44;
                    const cx = xFor(i);
                    return '<text class="label" x="' + cx + '" y="' + labelY + '" text-anchor="middle" transform="rotate(-90 ' + cx + ' ' + labelY + ')">' + txt + '</text>';
                }).join('');

                const legendX = leftPad + 4;
                const legendY = 10;
                const legend = ''
                    + '<rect class="belastning-series-kap" x="' + legendX + '" y="' + legendY + '" width="10" height="10"></rect><text class="label" x="' + (legendX + 14) + '" y="' + (legendY + 9) + '">Kapacitet</text>'
                    + '<rect class="belastning-series-resv" x="' + (legendX + 88) + '" y="' + legendY + '" width="10" height="10"></rect><text class="label" x="' + (legendX + 102) + '" y="' + (legendY + 9) + '">Reservationer</text>'
                    + '<rect class="belastning-series-aften" x="' + (legendX + 196) + '" y="' + legendY + '" width="10" height="10"></rect><text class="label" x="' + (legendX + 210) + '" y="' + (legendY + 9) + '">Rest Aften</text>';

                return '<svg class="belastning-svg" style="min-width:' + svgW + 'px" viewBox="0 0 ' + svgW + ' ' + svgH + '" preserveAspectRatio="xMinYMin meet">'
                    + '<line class="axis" x1="' + leftPad + '" y1="' + topPad + '" x2="' + leftPad + '" y2="' + (topPad + innerH) + '"></line>'
                    + '<line class="axis" x1="' + leftPad + '" y1="' + (topPad + innerH) + '" x2="' + (svgW - rightPad) + '" y2="' + (topPad + innerH) + '"></line>'
                    + grid.join('')
                    + bars
                    + labels
                    + legend
                    + '</svg>';
            }

            function renderBelastningDetailSvg(rows, opts) {
                const wrap = document.getElementById('belastningDetailSvg');
                if (!wrap) return;
                const safeRows = Array.isArray(rows) ? rows : [];
                if (safeRows.length === 0) {
                    wrap.innerHTML = '';
                    return;
                }
                wrap.innerHTML = buildBelastningClusterSvg(safeRows, opts);
            }

            function renderBelastningBars(targetId, items, allRows) {
                const wrap = document.getElementById(targetId);
                if (!wrap) return;
                const rows = sortBelastningItemsBySavedLayout(items);
                const fullRows = Array.isArray(allRows) ? allRows : [];
                if (rows.length === 0) {
                    wrap.innerHTML = '<div class="qms-empty">Ingen data i valgt periode.</div>';
                    const svgTarget = targetId === 'belastningBarsCombined' ? 'belastningSvgCombined' : null;
                    const svgWrap = document.getElementById(svgTarget);
                    if (svgWrap) svgWrap.innerHTML = '';
                    return;
                }

                const grouped = new Map();
                for (const row of fullRows) {
                    const key = String(row && row.ResGr || '').trim();
                    if (!key) continue;
                    if (!grouped.has(key)) grouped.set(key, []);
                    grouped.get(key).push(row);
                }

                wrap.innerHTML = rows.map(item => {
                    const kap = Number(item.totalKap || 0);
                    const resv = Number(item.totalResv || 0);
                    const aften = Number(item.totalAften || 0);
                    const loadPct = kap > 0 ? Math.min(160, (resv / kap) * 100) : 0;
                    const rg = String(item.resGr || '').trim();
                    const chartRows = grouped.get(rg) || [];
                    const itemParity = item && item.parity === 0 ? 0 : 1;
                    const isActive = belastningDetailContext
                        && String(belastningDetailContext.resGr || '') === rg
                        && Number(belastningDetailContext.parity) === itemParity;
                    const cardKey = getBelastningCardKey(item);
                    return ''
                        + '<div class="belastning-resource-chart' + (isActive ? ' is-active' : '') + '" draggable="true" data-belastning-key="' + escapeHtmlFE(cardKey) + '" onclick="loadBelastningDetail(\\'' + escapeHtmlFE(rg) + '\\',' + itemParity + ')">'
                        + '<div class="belastning-card-top">'
                        + '<h5>Kapacitetsbelastning: ' + escapeHtmlFE(rg + ' ' + String(item.nm || '')) + '</h5>'
                        + '<span class="belastning-drag-chip" title="Træk kortet for at gemme din egen rækkefølge">Flyt</span>'
                        + '</div>'
                        + buildBelastningClusterSvg(chartRows, {
                            clickable: true,
                            resGr: rg,
                            parity: itemParity,
                            activeDayKey: isActive ? belastningSelectedDayKey : ''
                        })
                        + '<div class="belastning-mini-meta">'
                        + '<span>Belastning: ' + escapeHtmlFE(formatPctDa(loadPct)) + '</span> · '
                        + '<span>Resv: ' + escapeHtmlFE(formatBelastningMinutes(resv)) + '</span>'
                        + '<span>Kap: ' + escapeHtmlFE(formatBelastningMinutes(kap)) + '</span>'
                        + '<span>Aften: ' + escapeHtmlFE(formatBelastningMinutes(aften)) + '</span>'
                        + '</div></div>';
                }).join('');

                    attachBelastningDragAndDrop(targetId);

                const svgTarget = targetId === 'belastningBarsCombined' ? 'belastningSvgCombined' : null;
                const svgWrap = document.getElementById(svgTarget);
                if (svgWrap) svgWrap.innerHTML = '';
            }

            function renderBelastningDetailTable(detailData, resGr, parity) {
                const wrap = document.getElementById('belastningDetailTable');
                const title = document.getElementById('belastningDetailTitle');
                const card = document.getElementById('belastningDetailWrap');
                if (!wrap || !title || !card) return;
                const safeRows = Array.isArray(detailData)
                    ? detailData
                    : (detailData && Array.isArray(detailData.rows) ? detailData.rows : []);
                const orderRows = detailData && Array.isArray(detailData.orderRows) ? detailData.orderRows : [];
                const subOrderRows = detailData && Array.isArray(detailData.subOrderRows) ? detailData.subOrderRows : [];
                const orderLineRows = detailData && Array.isArray(detailData.orderLineRows) ? detailData.orderLineRows : [];
                const resourceLookup = belastningLastPayload
                    ? [
                        ...(belastningLastPayload.odd && Array.isArray(belastningLastPayload.odd.resources) ? belastningLastPayload.odd.resources : []),
                        ...(belastningLastPayload.even && Array.isArray(belastningLastPayload.even.resources) ? belastningLastPayload.even.resources : [])
                    ]
                    : [];
                const resourceName = (() => {
                    const match = resourceLookup.find(item => String(item && item.resGr || '').trim() === String(resGr || '').trim());
                    return match ? String(match.nm || '').trim() : '';
                })();

                if (safeRows.length === 0 && orderRows.length === 0) {
                    card.style.display = 'none';
                    wrap.innerHTML = '';
                    return;
                }
                const groupLabel = String(resGr || '').trim() === '51'
                    ? 'Robotsvejs'
                    : '';
                const titleParts = ['Ordreoverblik: ' + String(resGr || '') + (resourceName ? (' ' + resourceName) : '')];
                if (groupLabel) titleParts.push(groupLabel);
                if (detailData && detailData.ord) titleParts.push('Ordre ' + String(detailData.ord));
                if (detailData && detailData.kunde) titleParts.push('Kunde ' + String(detailData.kunde));
                title.textContent = titleParts.join(' · ');

                const subOrderMap = new Map();
                for (const subRow of subOrderRows) {
                    const key = String(subRow && subRow.SubOrdNo || '').trim();
                    if (!key) continue;
                    if (!subOrderMap.has(key)) subOrderMap.set(key, []);
                    subOrderMap.get(key).push(subRow);
                }

                const orderLineMap = new Map();
                for (const line of orderLineRows) {
                    const key = String(line && line.OrdNo || '').trim();
                    if (!key) continue;
                    if (!orderLineMap.has(key)) orderLineMap.set(key, []);
                    orderLineMap.get(key).push(line);
                }

                const groupedByDate = new Map();
                const activeDateInput = document.getElementById('belastningToDay');
                const todayCut = (() => {
                    const raw = activeDateInput && activeDateInput.value
                        ? activeDateInput.value
                        : new Date().toISOString().slice(0, 10);
                    const t = new Date(raw + 'T00:00:00').getTime();
                    return Number.isNaN(t) ? Date.now() : t;
                })();
                const readResv = row => Number((row && (row.Resv !== undefined ? row.Resv : row.ResvRaw)) || 0);
                const readRest = row => Number((row && (row.RestResv !== undefined ? row.RestResv : row.ResvNet)) || 0);
                const readAften = row => Number((row && (row.RestAften !== undefined ? row.RestAften : (row.Aften !== undefined ? row.Aften : row.AftenRaw))) || 0);

                for (const row of orderRows) {
                    const dateKey = normalizeBelastningDateKey(row && row.Dato, row && row.DatoX);
                    const dateSort = getBelastningDateSortValue(dateKey);
                    const dateTxt = normalizeBelastningDisplayDate(row && row.Dato, row && row.DatoX);
                    const mapKey = dateKey || dateTxt;
                    if (!groupedByDate.has(mapKey)) {
                        groupedByDate.set(mapKey, {
                            dateKey,
                            dateTxt,
                            dateSort,
                            rows: [],
                            totalResv: 0,
                            totalRest: 0,
                            totalAften: 0
                        });
                    }
                    const bucket = groupedByDate.get(mapKey);
                    bucket.rows.push(row);
                    bucket.totalResv += readResv(row);
                    bucket.totalRest += readRest(row);
                    bucket.totalAften += readAften(row);
                }

                const sortedDateGroups = Array.from(groupedByDate.values()).sort((a, b) => a.dateSort - b.dateSort);
                const fmtDate = value => {
                    const parsed = value ? new Date(value) : null;
                    return parsed && !Number.isNaN(parsed.getTime())
                        ? parsed.toLocaleDateString('da-DK')
                        : '-';
                };

                const tableRows = sortedDateGroups.map((dateGroup, groupIndex) => {
                    const dateKey = 'beldate_' + groupIndex;
                    const groupsBySOrdre = new Map();
                    for (const row of dateGroup.rows) {
                        const sOrdre = String((row && (row.SOrdre || row.OrdNo)) || '-').trim() || '-';
                        if (!groupsBySOrdre.has(sOrdre)) {
                            groupsBySOrdre.set(sOrdre, {
                                sOrdre,
                                rows: [],
                                totalResv: 0,
                                totalRest: 0,
                                totalAften: 0
                            });
                        }
                        const bucket = groupsBySOrdre.get(sOrdre);
                        bucket.rows.push(row);
                        bucket.totalResv += readResv(row);
                        bucket.totalRest += readRest(row);
                        bucket.totalAften += readAften(row);
                    }

                    const orderRowsHtml = Array.from(groupsBySOrdre.values()).map((orderGroup, orderIndex) => {
                        const firstRow = orderGroup.rows[0] || {};
                        const detailKey = dateKey + '_ord_' + orderIndex;
                        const beforeClass = dateGroup.dateSort < todayCut ? ' belastning-before-day' : '';

                        const groupsByPOrdre = new Map();
                        for (const row of orderGroup.rows) {
                            const pOrdre = String((row && (row.POrdre || row.PurcNo || row.OrdNo)) || '-').trim() || '-';
                            if (!groupsByPOrdre.has(pOrdre)) {
                                groupsByPOrdre.set(pOrdre, {
                                    pOrdre,
                                    rows: [],
                                    totalResv: 0,
                                    totalRest: 0,
                                    totalAften: 0
                                });
                            }
                            const pBucket = groupsByPOrdre.get(pOrdre);
                            pBucket.rows.push(row);
                            pBucket.totalResv += readResv(row);
                            pBucket.totalRest += readRest(row);
                            pBucket.totalAften += readAften(row);
                        }

                        const pOrderGroups = Array.from(groupsByPOrdre.values());
                        const hasChildren = pOrderGroups.length > 0;
                        const routeList = Array.from(new Set(orderGroup.rows.map(x => String(x && x.Opr || '').trim()).filter(Boolean))).join(' ');
                        const parentRow = '<tr class="belastning-order-row' + beforeClass + '" data-parent-date="' + dateKey + '" data-parent-order="' + detailKey + '">'
                            + '<td style="text-align:center;">'
                            + (hasChildren
                                ? ('<button type="button" class="belastning-order-toggle" data-order-key="' + detailKey + '" data-collapsed="1" onclick="toggleBelastningOrderNode(\\'' + detailKey + '\\', this)">+</button>')
                                : '<span class="belastning-order-sub">-</span>')
                            + '</td>'
                            + '<td>' + escapeHtmlFE(dateGroup.dateTxt) + '</td>'
                            + '<td style="text-align:right;"><span class="belastning-order-id">' + escapeHtmlFE(String(firstRow.SOrdre || orderGroup.sOrdre || '-')) + '</span></td>'
                            + '<td style="text-align:right;"><span class="belastning-order-sub">intern</span></td>'
                            + '<td>' + escapeHtmlFE(String(firstRow.Kunde || '-')) + '</td>'
                            + '<td>' + escapeHtmlFE(routeList || '-') + '</td>'
                            + '<td>' + escapeHtmlFE(String(firstRow.LevMode || '-')) + '</td>'
                            + '<td>' + escapeHtmlFE(fmtDate(firstRow.LevDato)) + '</td>'
                            + '<td>' + escapeHtmlFE(fmtDate(firstRow.ULDato)) + '</td>'
                            + '<td style="text-align:right;">' + escapeHtmlFE(formatBelastningMinutes(orderGroup.totalResv)) + '</td>'
                            + '<td style="text-align:right;">' + escapeHtmlFE(formatBelastningMinutes(orderGroup.totalRest)) + '</td>'
                            + '<td style="text-align:right;">' + escapeHtmlFE(formatBelastningMinutes(orderGroup.totalAften)) + '</td>'
                            + '</tr>';

                        const childRowsHtml = pOrderGroups.map((pOrderGroup, childIndex) => {
                            const childKey = detailKey + '_child_' + childIndex;
                            const childRow = pOrderGroup.rows[0] || {};
                            const routeText = String(childRow.Opr || '').trim() || '-';
                            return '<tr class="belastning-order-detail-row' + beforeClass + '" data-parent-date="' + dateKey + '" data-parent-order="' + detailKey + '" data-child-key="' + childKey + '" style="display:none;">'
                                + '<td></td>'
                                + '<td>' + escapeHtmlFE(dateGroup.dateTxt) + '</td>'
                                + '<td style="text-align:right;">' + escapeHtmlFE(String(childRow.SOrdre || orderGroup.sOrdre || '-')) + '</td>'
                                + '<td style="text-align:right;">' + escapeHtmlFE(String(pOrderGroup.pOrdre || '-')) + '</td>'
                                + '<td>' + escapeHtmlFE(String(childRow.Kunde || '-')) + '</td>'
                                + '<td>' + escapeHtmlFE(routeText) + '</td>'
                                + '<td>-</td>'
                                + '<td>-</td>'
                                + '<td>' + escapeHtmlFE(fmtDate(childRow.ULDato)) + '</td>'
                                + '<td style="text-align:right;">' + escapeHtmlFE(formatBelastningMinutes(pOrderGroup.totalResv)) + '</td>'
                                + '<td style="text-align:right;">-</td>'
                                + '<td style="text-align:right;">-</td>'
                                + '</tr>';
                        }).join('');

                        return parentRow + childRowsHtml;
                    }).join('');

                    const groupHeaderClass = dateGroup.dateSort < todayCut ? 'belastning-date-row belastning-before-day' : 'belastning-date-row';
                    const groupHeader = '<tr class="' + groupHeaderClass + '" data-day-key="' + escapeHtmlFE(String(dateGroup.dateKey || '')) + '">'
                        + '<td colspan="12">'
                        + '<div class="belastning-date-header">'
                        + '<button type="button" class="belastning-date-toggle" data-date-key="' + dateKey + '" data-collapsed="0" onclick="toggleBelastningDateGroup(\\'' + dateKey + '\\', this)">-</button>'
                        + '<span class="belastning-date-badge">Dato: ' + escapeHtmlFE(dateGroup.dateTxt) + '</span>'
                        + '<span class="belastning-date-meta">'
                        + '<span class="belastning-date-meta-item"><span class="belastning-date-meta-lbl">S-Ordre:</span><span class="belastning-date-meta-val">' + escapeHtmlFE(String(groupsBySOrdre.size)) + '</span></span>'
                        + '<span class="belastning-date-meta-item"><span class="belastning-date-meta-lbl">Minutter:</span><span class="belastning-date-meta-val">' + escapeHtmlFE(formatBelastningMinutes(dateGroup.totalResv)) + '</span></span>'
                        + '<span class="belastning-date-meta-item"><span class="belastning-date-meta-lbl">Rest:</span><span class="belastning-date-meta-val">' + escapeHtmlFE(formatBelastningMinutes(dateGroup.totalRest)) + '</span></span>'
                        + '<span class="belastning-date-meta-item"><span class="belastning-date-meta-lbl">Aften:</span><span class="belastning-date-meta-val">' + escapeHtmlFE(formatBelastningMinutes(dateGroup.totalAften)) + '</span></span>'
                        + '</span>'
                        + '</div></td></tr>';

                    return groupHeader + orderRowsHtml;
                }).join('');

                const orderTable = sortedDateGroups.length > 0
                    ? ('<div class="belastning-section-title">Ordreoverblik</div>'
                        + '<div class="belastning-order-shell">'
                        + '<table class="belastning-order-table">'
                        + '<thead><tr>'
                        + '<th></th><th>Dato</th><th>S-Ordre</th><th>P-Ordre</th><th>Kunde</th><th>Rute</th><th>Lev.måde</th><th>Lev.dato</th><th>U-dato</th><th>Minutter</th><th>Rest</th><th>Aften</th>'
                        + '</tr></thead>'
                        + '<tbody>' + tableRows + '</tbody></table>'
                        + '</div>')
                    : '<div class="qms-empty">Ingen ordrelinjer for valgt ressourcegruppe/periode.</div>';

                renderBelastningDetailSvg(safeRows, {
                    clickable: true,
                    resGr: String(resGr || '').trim(),
                    parity: parity === 0 ? 0 : 1,
                    activeDayKey: belastningSelectedDayKey
                });
                wrap.innerHTML = orderTable;
                card.style.display = 'block';
                focusBelastningDayInTable(belastningSelectedDayKey);
            }

            function focusBelastningDayInTable(dayKey) {
                const key = String(dayKey || '').trim();
                const isBeforeFocus = key === 'before';
                const headers = Array.from(document.querySelectorAll('tr.belastning-date-row'));
                headers.forEach(row => row.classList.remove('belastning-day-focus'));
                if (!key) return;

                if (isBeforeFocus) {
                    headers.filter(row => row.classList.contains('belastning-before-day')).forEach(row => row.classList.add('belastning-day-focus'));
                } else {
                    const targetHeader = document.querySelector('tr.belastning-date-row[data-day-key="' + key + '"]');
                    if (!targetHeader) return;
                    targetHeader.classList.add('belastning-day-focus');
                }

                const toggles = Array.from(document.querySelectorAll('.belastning-date-toggle'));
                toggles.forEach(btn => {
                    const internalKey = String(btn.getAttribute('data-date-key') || '');
                    const groupRows = Array.from(document.querySelectorAll('tr[data-parent-date="' + internalKey + '"]'));
                    const headerRow = btn.closest('tr.belastning-date-row');
                    const isTarget = headerRow && (isBeforeFocus
                        ? headerRow.classList.contains('belastning-before-day')
                        : headerRow.getAttribute('data-day-key') === key);
                    if (isTarget) {
                        btn.setAttribute('data-collapsed', '0');
                        btn.textContent = '-';
                        groupRows.forEach(row => {
                            if (!row.classList.contains('belastning-order-detail-row')) {
                                row.style.display = 'table-row';
                                return;
                            }
                            const orderKey = row.getAttribute('data-parent-order') || '';
                            const orderBtn = document.querySelector('.belastning-order-toggle[data-order-key="' + orderKey + '"]');
                            const orderCollapsed = !orderBtn || orderBtn.getAttribute('data-collapsed') === '1';
                            row.style.display = orderCollapsed ? 'none' : 'table-row';
                        });
                    } else {
                        btn.setAttribute('data-collapsed', '1');
                        btn.textContent = '+';
                        groupRows.forEach(row => {
                            row.style.display = 'none';
                        });
                    }
                });
            }

            function toggleBelastningDateGroup(dateKey, buttonEl) {
                const collapsed = buttonEl && buttonEl.getAttribute('data-collapsed') === '1';
                const nextCollapsed = !collapsed;
                if (buttonEl) {
                    buttonEl.setAttribute('data-collapsed', nextCollapsed ? '1' : '0');
                    buttonEl.textContent = nextCollapsed ? '+' : '-';
                }

                const rows = document.querySelectorAll('tr[data-parent-date="' + dateKey + '"]');
                rows.forEach(row => {
                    if (nextCollapsed) {
                        row.style.display = 'none';
                        return;
                    }
                    if (row.classList.contains('belastning-order-detail-row')) {
                        const orderKey = row.getAttribute('data-parent-order') || '';
                        const orderBtn = document.querySelector('.belastning-order-toggle[data-order-key="' + orderKey + '"]');
                        const orderCollapsed = !orderBtn || orderBtn.getAttribute('data-collapsed') === '1';
                        row.style.display = orderCollapsed ? 'none' : 'table-row';
                        return;
                    }
                    row.style.display = 'table-row';
                });
            }

            function toggleBelastningOrderNode(orderKey, buttonEl) {
                const detailRows = Array.from(document.querySelectorAll('tr.belastning-order-detail-row[data-parent-order="' + orderKey + '"]'));
                if (detailRows.length === 0) return;

                const collapsed = buttonEl && buttonEl.getAttribute('data-collapsed') === '1';
                const nextCollapsed = !collapsed;
                if (buttonEl) {
                    buttonEl.setAttribute('data-collapsed', nextCollapsed ? '1' : '0');
                    buttonEl.textContent = nextCollapsed ? '+' : '-';
                }

                const dateKey = detailRows[0].getAttribute('data-parent-date') || '';
                const dateBtn = document.querySelector('.belastning-date-toggle[data-date-key="' + dateKey + '"]');
                const dateCollapsed = dateBtn && dateBtn.getAttribute('data-collapsed') === '1';
                detailRows.forEach(row => {
                    row.style.display = (nextCollapsed || dateCollapsed) ? 'none' : 'table-row';
                });
            }

            async function loadBelastningDetail(resGr, parity, options) {
                try {
                    const safeOptions = options && typeof options === 'object' ? options : {};
                    const filters = getBelastningFilters();
                    const query = new URLSearchParams({
                        toDay: filters.today,
                        dage: String(filters.dage),
                        resGr: String(resGr || ''),
                        parity: String(parity === 0 ? 0 : 1),
                        ord: filters.ord,
                        kunde: filters.kunde
                    });
                    belastningDetailContext = { resGr: String(resGr || '').trim(), parity: parity === 0 ? 0 : 1 };
                    belastningSelectedDayKey = String(safeOptions.focusDayKey || '').trim();
                    const response = await fetch('/belastning/detail?' + query.toString());
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    if (!payload.ok) throw new Error(payload.error || 'Belastning detail fejl');
                    renderBelastningDetailTable(payload, resGr, parity === 0 ? 0 : 1);
                    if (belastningLastPayload) {
                        const oddItems = belastningLastPayload.odd && Array.isArray(belastningLastPayload.odd.resources)
                            ? belastningLastPayload.odd.resources.map(item => ({ ...item, parity: 1 }))
                            : [];
                        const evenItems = belastningLastPayload.even && Array.isArray(belastningLastPayload.even.resources)
                            ? belastningLastPayload.even.resources.map(item => ({ ...item, parity: 0 }))
                            : [];
                        const oddRows = belastningLastPayload.odd && Array.isArray(belastningLastPayload.odd.rows) ? belastningLastPayload.odd.rows : [];
                        const evenRows = belastningLastPayload.even && Array.isArray(belastningLastPayload.even.rows) ? belastningLastPayload.even.rows : [];
                        renderBelastningBars('belastningBarsCombined', [...oddItems, ...evenItems], [...oddRows, ...evenRows]);
                    }
                    scrollBelastningDetailIntoView();
                    setTimeout(scrollBelastningDetailIntoView, 160);
                } catch (err) {
                    setBelastningStatus('Fejl ved detailhentning: ' + (err && err.message ? err.message : 'ukendt'));
                }
            }

            async function initializeBelastningIfNeeded() {
                if (belastningInitialized) return;
                belastningInitialized = true;
                const todayInput = document.getElementById('belastningToDay');
                if (todayInput && !todayInput.value) {
                    todayInput.value = new Date().toISOString().slice(0, 10);
                }
                startBelastningPeriodicRefresh();
                setBelastningStatus('Henter data...');
                await loadBelastningGrafisk({ forceRefresh: true });
            }

            async function loadBelastningGrafisk(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const emptyEl = document.getElementById('belastningEmpty');
                const loadBtn = document.getElementById('belastningLoadBtn');
                const graphWrap = document.getElementById('belastningGrafiskWrap');
                const detailWrap = document.getElementById('belastningDetailWrap');
                if (loadBtn) loadBtn.disabled = true;
                if (detailWrap) detailWrap.style.display = 'none';
                if (emptyEl) {
                    emptyEl.style.display = 'block';
                    emptyEl.textContent = 'Henter belastning...';
                }

                try {
                    const filters = getBelastningFilters();
                    const query = new URLSearchParams({
                        toDay: filters.today,
                        dage: String(filters.dage),
                        resGr: filters.resGr,
                        ord: filters.ord,
                        kunde: filters.kunde
                    });
                    const response = await fetch('/belastning/grafisk?' + query.toString());
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    if (!payload.ok) throw new Error(payload.error || 'Belastning grafisk fejl');
                    belastningLastPayload = payload;
                    const oddItems = payload.odd && Array.isArray(payload.odd.resources) ? payload.odd.resources.map(item => ({ ...item, parity: 1 })) : [];
                    const evenItems = payload.even && Array.isArray(payload.even.resources) ? payload.even.resources.map(item => ({ ...item, parity: 0 })) : [];
                    const oddRows = payload.odd && Array.isArray(payload.odd.rows) ? payload.odd.rows : [];
                    const evenRows = payload.even && Array.isArray(payload.even.rows) ? payload.even.rows : [];
                    const combinedItems = [...oddItems, ...evenItems];
                    const combinedRows = [...oddRows, ...evenRows];
                    renderBelastningBars('belastningBarsCombined', combinedItems, combinedRows);
                    if (graphWrap) graphWrap.style.display = combinedItems.length ? 'grid' : 'none';
                    if (emptyEl) {
                        if (combinedItems.length === 0) {
                            emptyEl.style.display = 'block';
                            emptyEl.textContent = 'Ingen belastningsdata i valgt periode.';
                        } else {
                            emptyEl.style.display = 'none';
                        }
                    }
                    const orderStatus = filters.ord ? (' · Ordre-filter: ' + filters.ord) : '';
                    const customerStatus = filters.kunde ? (' · Kunde-filter: ' + filters.kunde) : '';
                    setBelastningStatus('Periode: ' + filters.today + ' + ' + filters.dage + ' dage · Ressourcer: ' + combinedItems.length + orderStatus + customerStatus + ' · Auto-opdater: 15 min');
                } catch (err) {
                    setBelastningStatus('Fejl ved hentning af belastning.');
                    if (emptyEl) {
                        emptyEl.style.display = 'block';
                        emptyEl.textContent = 'Fejl: ' + (err && err.message ? err.message : 'ukendt fejl');
                    }
                    if (graphWrap) graphWrap.style.display = 'none';
                } finally {
                    if (loadBtn) loadBtn.disabled = false;
                }
            }

            function showAccessGate() {
                const overlay = document.getElementById('accessGateOverlay');
                const userInput = document.getElementById('accessGateUserInput');
                const input = document.getElementById('accessGateInput');
                const err = document.getElementById('accessGateError');
                if (!overlay) return;
                if (err) err.textContent = '';
                if (userInput) {
                    const currentName = sanitizeDisplayName(loggedUserDisplayName);
                    if (!String(userInput.value || '').trim() && currentName && currentName !== 'Bruger') {
                        userInput.value = currentName;
                    }
                }
                overlay.style.display = 'flex';
                refreshSideMenuAuthState();
                setTimeout(() => {
                    if (userInput && !String(userInput.value || '').trim()) {
                        userInput.focus();
                        return;
                    }
                    if (input) input.focus();
                }, 30);
            }

            function hideAccessGate() {
                const overlay = document.getElementById('accessGateOverlay');
                if (!overlay) return;
                overlay.style.display = 'none';
                refreshSideMenuAuthState();
            }

            function submitAccessCode() {
                const userInput = document.getElementById('accessGateUserInput');
                const input = document.getElementById('accessGateInput');
                const err = document.getElementById('accessGateError');
                const btn = document.getElementById('accessGateBtn');
                const userName = sanitizeDisplayName(userInput ? userInput.value : '');
                const value = input ? String(input.value || '').trim() : '';
                if (value !== ACCESS_CODE) {
                    if (err) err.textContent = 'Forkert kode.';
                    if (input) {
                        input.select();
                        input.focus();
                    }
                    return;
                }

                if (err) err.textContent = 'Åbner...';
                if (btn) {
                    btn.disabled = true;
                    btn.textContent = 'Åbner...';
                }

                setTimeout(() => {
                    try {
                        if (userName && userName !== 'Bruger') {
                            setLoggedUserDisplayName(userName);
                        }
                        accessGranted = true;
                        hideAccessGate();
                        refreshSideMenuAuthState();
                        initializeAfterAccess();
                    } catch (e) {
                        accessGranted = false;
                        showAccessGate();
                        if (err) err.textContent = 'Fejl ved åbning: ' + (e && e.message ? e.message : 'ukendt fejl');
                    } finally {
                        if (btn) {
                            btn.disabled = false;
                            btn.textContent = 'Åbn';
                        }
                    }
                }, 0);
            }

            function initializeAfterAccess() {
                startWarmupPolling();
                loadOrderList(false);
                setTimeout(() => {
                    if (!orderListData || orderListData.length === 0) {
                        loadOrderList(true);
                    }
                }, 2500);
                startOrderListAutoRefresh();
                startDashboardUpdatePolling();

                const params = new URLSearchParams(window.location.search);
                if (params.has('ord')) {
                    document.getElementById('orderInput').value = params.get('ord');
                    openModule('efterkalk');
                    searchOrder();
                    return;
                }
                goToDashboard();
            }

            function openModule(moduleKey) {
                const dashboard = document.getElementById('mainDashboard');
                const workspace = document.getElementById('mainWorkspace');
                const omsaetning = document.getElementById('mainOmsaetning');
                const ordreindgang = document.getElementById('mainOrdreindgang');
                const belastning = document.getElementById('mainBelastning');

                if (moduleKey === 'efterkalk') {
                    if (!warmupCombinedReady) {
                        const msg = warmupCombinedTotal > 0
                            ? ('Efterkalk er ikke klar endnu (' + warmupCombinedDone + '/' + warmupCombinedTotal + ', ' + warmupCombinedPct + '%). Vent til warmup er færdig.')
                            : 'Efterkalk er ikke klar endnu. Vent et øjeblik til warmup/calculations er færdige.';
                        alert(msg);
                        return;
                    }
                    if (dashboard) dashboard.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'none';
                    if (ordreindgang) ordreindgang.style.display = 'none';
                    if (belastning) belastning.style.display = 'none';
                    if (workspace) workspace.style.display = 'block';
                    closeSideMenu();
                    goBackToList();
                    setTimeout(syncStickyOffsets, 0);
                    return;
                }

                if (moduleKey === 'omsaetning') {
                    if (dashboard) dashboard.style.display = 'none';
                    if (workspace) workspace.style.display = 'none';
                    if (ordreindgang) ordreindgang.style.display = 'none';
                    if (belastning) belastning.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'block';
                    closeSideMenu();
                    initializeOmsaetningIfNeeded();
                    setTimeout(syncStickyOffsets, 0);
                    return;
                }

                if (moduleKey === 'ordreindgang') {
                    if (dashboard) dashboard.style.display = 'none';
                    if (workspace) workspace.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'none';
                    if (belastning) belastning.style.display = 'none';
                    if (ordreindgang) ordreindgang.style.display = 'block';
                    closeSideMenu();
                    initializeOrdreindgangIfNeeded();
                    setTimeout(syncStickyOffsets, 0);
                    return;
                }

                if (moduleKey === 'belastning') {
                    if (dashboard) dashboard.style.display = 'none';
                    if (workspace) workspace.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'none';
                    if (ordreindgang) ordreindgang.style.display = 'none';
                    if (belastning) belastning.style.display = 'block';
                    closeSideMenu();
                    initializeBelastningIfNeeded();
                    setTimeout(syncStickyOffsets, 0);
                    return;
                }

                if (moduleKey !== 'efterkalk') {
                    alert('Dette modul er klar til næste fase. Når du sender logikken, bygger vi det visuelt og funktionelt.');
                    return;
                }
            }

            function goToDashboard() {
                const dashboard = document.getElementById('mainDashboard');
                const workspace = document.getElementById('mainWorkspace');
                const omsaetning = document.getElementById('mainOmsaetning');
                const ordreindgang = document.getElementById('mainOrdreindgang');
                const belastning = document.getElementById('mainBelastning');
                closeSideMenu();
                if (workspace) workspace.style.display = 'none';
                if (omsaetning) omsaetning.style.display = 'none';
                if (ordreindgang) ordreindgang.style.display = 'none';
                if (belastning) belastning.style.display = 'none';
                if (dashboard) dashboard.style.display = 'block';
                const detailModal = document.getElementById('orderDetailModal');
                const detailBody = document.getElementById('orderDetailModalBody');
                if (detailModal) detailModal.style.display = 'none';
                if (detailBody) detailBody.innerHTML = '';
                document.body.classList.remove('report-modal-open');
                setTimeout(syncStickyOffsets, 0);
            }

            async function initializeOmsaetningIfNeeded() {
                if (omsaetningInitialized) return;
                omsaetningInitialized = true;

                const currentFiscalYear = getCurrentFiscalYearStart();
                omsaetningSelectedFiscalYears = new Set([currentFiscalYear]);
                applySelectedFiscalYearsToInputs();
                renderOmsaetningYearChips(currentFiscalYear);
                applyOmsaetningThresholdInputs(OMSAETNING_DEFAULT_WARN_THRESHOLD, OMSAETNING_DEFAULT_GOOD_THRESHOLD);
                renderOmsaetningCustomerMode();
                renderOmsaetningCustomerResults();

                await loadOmsaetningAccounts();
                await loadOmsaetningSummary();
            }

            async function loadOmsaetningAccounts() {
                const list = document.getElementById('omsaetningAccountsList');
                if (!list) return;
                list.innerHTML = '<div class="omsaetning-account-item"><span>Indlæser konti...</span></div>';
                try {
                    const response = await fetch('/omsaetning/accounts');
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    const accounts = Array.isArray(payload.accounts) ? payload.accounts : [];
                    omsaetningAccounts = accounts;
                    const allAccounts = accounts.map(acc => String(acc.acNo || '').trim()).filter(Boolean);
                    const ssrsPreset = allAccounts.filter(acNo => OMSAETNING_SSRS_DEFAULT_ACCOUNTS.has(acNo));
                    if (ssrsPreset.length === 0) {
                        omsaetningSelectedAccounts = new Set(allAccounts);
                    } else {
                        omsaetningSelectedAccounts = new Set(ssrsPreset);
                    }
                    renderOmsaetningAccountsList();
                } catch (err) {
                    list.innerHTML = '<div class="omsaetning-account-item"><span>Fejl ved konti</span></div>';
                    console.error('loadOmsaetningAccounts failed:', err);
                }
            }

            async function loadOmsaetningSummary(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const silentValidation = safeOptions.silentValidation === true;
                const loadBtn = document.getElementById('omsaetningLoadBtn');
                const empty = document.getElementById('omsaetningEmpty');
                const tableWrap = document.getElementById('omsaetningTableWrap');
                const detailsWrap = document.getElementById('omsaetningDetailsWrap');
                const thresholdWrap = document.getElementById('omsaetningThresholdWrap');
                const thresholdTable = document.getElementById('omsaetningThresholdTable');
                const chartsWrap = document.getElementById('omsaetningChartsWrap');
                const totalEl = document.getElementById('omsaetningTotalMio');
                const rowsEl = document.getElementById('omsaetningRowsCount');
                const periodsEl = document.getElementById('omsaetningPeriodsCount');

                const periodRange = buildOmsaetningPeriodRange();
                if (!periodRange) {
                    if (!silentValidation) {
                        alert('Vælg gyldig periode (Fra måned skal være før eller lig Til måned).');
                    }
                    return;
                }

                const fra = periodRange.fra;
                const til = periodRange.til;
                const monthKeysForPeriod = [];

                const selected = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean);
                if (selected.length === 0) {
                    if (!silentValidation) {
                        alert('Vælg mindst én konto.');
                    }
                    return;
                }
                const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys()).filter(Boolean);

                const thresholdInputs = getOmsaetningThresholdInputs();
                const warnThreshold = thresholdInputs.warnThreshold;
                const goodThreshold = thresholdInputs.goodThreshold;
                applyOmsaetningThresholdInputs(warnThreshold, goodThreshold);

                if (loadBtn) {
                    loadBtn.disabled = true;
                    loadBtn.textContent = 'Indlæser...';
                }

                try {
                    const payload = await fetchOmsaetningSummaryCached(fra, til, selected, selectedCustomers, safeOptions);

                    const persistTargets = await resolveOmsaetningThresholdPersistTargets(
                        selectedCustomers,
                        warnThreshold,
                        goodThreshold,
                        safeOptions
                    );

                    if (persistTargets.length > 0) {
                        persistOmsaetningThresholdsForCustomers(persistTargets, warnThreshold, goodThreshold)
                            .catch(err => console.warn('persistOmsaetningThresholdsForCustomers failed:', err && err.message ? err.message : err));

                        for (const custNo of persistTargets) {
                            omsaetningThresholdsByCustomer.set(custNo, {
                                warnThreshold,
                                goodThreshold
                            });
                        }
                        renderOmsaetningCustomerThresholds();
                    }

                    const rows = Array.isArray(payload.rows) ? payload.rows : [];
                    const uniquePeriods = new Set(monthKeysForPeriod.length > 0 ? monthKeysForPeriod : rows.map(r => normalizeOmsaetningMonthKey(r.date)));

                    const monthTotals = new Map();
                    for (const monthKey of monthKeysForPeriod) {
                        monthTotals.set(String(monthKey), 0);
                    }
                    for (const row of rows) {
                        const monthKey = normalizeOmsaetningMonthKey(row.date);
                        const prev = monthTotals.get(monthKey) || 0;
                        monthTotals.set(monthKey, prev + Number(row.revenueMio || 0));
                    }

                    if (totalEl) totalEl.textContent = formatMio(payload.totalRevenueMio || 0);
                    if (rowsEl) rowsEl.textContent = formatCount(rows.length);
                    if (periodsEl) periodsEl.textContent = formatCount(uniquePeriods.size);

                    const sortedMonths = (monthKeysForPeriod.length > 0
                        ? monthKeysForPeriod.map(monthKey => [monthKey, Number(monthTotals.get(String(monthKey)) || 0)])
                        : Array.from(monthTotals.entries()).sort((a, b) => String(a[0]).localeCompare(String(b[0]))));
                    let thresholdHtml = '<table class="omsaetning-table"><thead><tr>' +
                        '<th>Måned</th><th class="omsaetning-cell-right">Omsætning (Mio)</th><th>Tærskel</th>' +
                        '</tr></thead><tbody>';
                    for (const [monthKey, amountMio] of sortedMonths) {
                        const statusClass = getOmsaetningStatusClass(amountMio, warnThreshold, goodThreshold);
                        const monthMioLabel = formatMio(amountMio);
                        const monthDkkLabel = formatDkkFromMio(amountMio);
                        const gauge = buildOmsaetningGaugeData(amountMio, warnThreshold, goodThreshold);
                        const marginPctLabel = formatSigned(gauge.marginPct, 1) + '%';
                        const deltaWarnLabel = formatSigned(gauge.deltaWarn, 3) + ' Mio';
                        const deltaGoodLabel = formatSigned(gauge.deltaGood, 3) + ' Mio';
                        thresholdHtml += '<tr>' +
                            '<td>' + escapeHtmlFE(formatMonthDa(monthKey)) + '</td>' +
                            '<td class="omsaetning-cell-right" title="' + escapeHtmlFE(monthDkkLabel + ' DKK') + '">' + escapeHtmlFE(monthMioLabel) + '</td>' +
                            '<td>' +
                                '<div class="omsaetning-gauge-wrap">' +
                                    '<div class="omsaetning-gauge-meta">' +
                                        '<span class="omsaetning-status ' + statusClass + '">' + escapeHtmlFE(getOmsaetningStatusLabel(statusClass)) + '</span>' +
                                        '<strong>' + escapeHtmlFE(marginPctLabel) + '</strong>' +
                                    '</div>' +
                                    '<div class="omsaetning-gauge-track">' +
                                        '<div class="omsaetning-gauge-fill ' + gauge.fillClass + '" style="left:' + gauge.fillLeft.toFixed(2) + '%;width:' + gauge.fillWidth.toFixed(2) + '%;"></div>' +
                                        '<span class="omsaetning-gauge-marker" style="left:' + gauge.zeroLeft.toFixed(2) + '%;"></span>' +
                                        '<span class="omsaetning-gauge-marker" style="left:' + gauge.targetLeft.toFixed(2) + '%;"></span>' +
                                        '<span class="omsaetning-gauge-point" style="left:' + gauge.pointLeft.toFixed(2) + '%;"></span>' +
                                    '</div>' +
                                    '<div class="omsaetning-gauge-legend"><span>-30%</span><span>0% (3)</span><span>30% (5)</span><span>60%</span></div>' +
                                    '<div class="omsaetning-gauge-delta">vs 3: <strong>' + escapeHtmlFE(deltaWarnLabel) + '</strong> · vs 5: <strong>' + escapeHtmlFE(deltaGoodLabel) + '</strong></div>' +
                                '</div>' +
                            '</td>' +
                            '</tr>';
                    }
                    thresholdHtml += '</tbody></table>';

                    let html = '<table class="omsaetning-table"><thead><tr>' +
                        '<th>Måned</th><th>Konto</th><th>Navn</th><th>Kunde</th><th>Kundenavn</th><th style="text-align:right;">Omsætning (Mio)</th>' +
                        '</tr></thead><tbody>';

                    if (rows.length === 0) {
                        html += '<tr><td colspan="6" style="color:#5f7892;">Ingen bevægelser i perioden (0 for alle måneder).</td></tr>';
                    } else {
                        for (const row of rows) {
                            const rowMioLabel = formatMio(row.revenueMio || 0);
                            const rowDkkLabel = formatDkkFromMio(row.revenueMio || 0);
                            html += '<tr>' +
                                '<td>' + escapeHtmlFE(formatMonthDa(row.date)) + '</td>' +
                                '<td>' + escapeHtmlFE(String(row.acNo || '')) + '</td>' +
                                '<td>' + escapeHtmlFE(String(row.name || '')) + '</td>' +
                                '<td>' + escapeHtmlFE(row.custNo === null || row.custNo === undefined ? '' : String(row.custNo)) + '</td>' +
                                '<td>' + escapeHtmlFE(String(row.customerName || '')) + '</td>' +
                                '<td style="text-align:right;" title="' + escapeHtmlFE(rowDkkLabel + ' DKK') + '">' + escapeHtmlFE(rowMioLabel) + '</td>' +
                                '</tr>';
                        }
                    }

                    html += '</tbody></table>';
                    if (OMSAETNING_SHOW_THRESHOLD_SECTION) {
                        if (thresholdTable) thresholdTable.innerHTML = thresholdHtml;
                        if (thresholdWrap) thresholdWrap.style.display = 'block';
                    } else {
                        if (thresholdTable) thresholdTable.innerHTML = '';
                        if (thresholdWrap) thresholdWrap.style.display = 'none';
                    }
                    if (tableWrap) {
                        tableWrap.innerHTML = html;
                    }
                    if (detailsWrap) {
                        detailsWrap.style.display = 'block';
                        applyOmsaetningDetailsCollapsedState();
                    }
                    renderOmsaetningCharts(rows, monthKeysForPeriod);
                    if (empty) empty.style.display = 'none';
                } catch (err) {
                    if (tableWrap) {
                        tableWrap.style.display = 'none';
                        tableWrap.innerHTML = '';
                    }
                    if (detailsWrap) detailsWrap.style.display = 'none';
                    if (thresholdWrap) thresholdWrap.style.display = 'none';
                    if (thresholdTable) thresholdTable.innerHTML = '';
                    if (chartsWrap) chartsWrap.style.display = 'none';
                    if (empty) {
                        empty.style.display = 'block';
                        empty.textContent = 'Fejl ved indlæsning: ' + (err && err.message ? err.message : 'ukendt fejl');
                    }
                    console.error('loadOmsaetningSummary failed:', err);
                } finally {
                    if (loadBtn) {
                        loadBtn.disabled = false;
                        loadBtn.textContent = 'Opdater';
                    }
                }
            }

            async function checkOrderListFreshness() {
                const now = Date.now();
                if (now - lastOrderListCheckTime < 30000) return;
                lastOrderListCheckTime = now;

                try {
                    const r = await fetch('/order-list-check-time');
                    if (!r.ok) return;
                    const d = await r.json();
                    const remoteMaxDate = Number(d.lastModifiedDate || 0);
                    
                    if (remoteMaxDate > 0 && remoteMaxDate !== lastOrderListRemoteTime) {
                        console.info('ORDER-LIST: Database has new/changed order (date=' + remoteMaxDate + ')');
                        lastOrderListRemoteTime = remoteMaxDate;
                        await loadOrderList(true);
                    }
                } catch (err) {
                    console.warn('checkOrderListFreshness failed:', err.message);
                }
            }

            function registerSummaryImageData(title, items) {
                if (!Array.isArray(items) || items.length === 0) return '';
                summaryImageRegistryCounter += 1;
                const key = 'img-' + summaryImageRegistryCounter;
                summaryImageRegistry[key] = {
                    title: title || 'Billeder',
                    items: items
                };
                return key;
            }

            function getSummaryImageSrc(item) {
                if (!item) return '';
                if (item.type === 'url') return item.value;
                return '/image-file?path=' + encodeURIComponent(item.value || '');
            }

            function selectCompactImageItems(items) {
                const source = Array.isArray(items) ? items : [];
                const clean = source.filter(item => item && String(item.value || '').trim());
                if (clean.length <= 2) return clean;

                const pickByLabel = (pattern) => clean.find(item => pattern.test(String(item.label || '')));
                const picked = [];
                const pushUnique = (item) => {
                    if (!item) return;
                    const value = String(item.value || '').trim().toLowerCase();
                    if (!value) return;
                    if (picked.some(x => String(x.value || '').trim().toLowerCase() === value)) return;
                    picked.push(item);
                };

                pushUnique(pickByLabel(/webpg|nesting/i));
                pushUnique(pickByLabel(/pictfnm|icon/i));
                for (const item of clean) {
                    pushUnique(item);
                    if (picked.length >= 2) break;
                }
                return picked.slice(0, 2);
            }

            function openCompactImageModal(imageKey) {
                const entry = summaryImageRegistry[imageKey];
                const modal = document.getElementById('compactImageModal');
                const titleEl = document.getElementById('compactImageTitle');
                const subtitleEl = document.getElementById('compactImageSubtitle');
                const bodyEl = document.getElementById('compactImageBody');
                if (!entry || !modal || !titleEl || !bodyEl) return;

                const items = selectCompactImageItems(entry.items);
                if (!items.length) return;

                titleEl.textContent = entry.title || 'Billeder';
                subtitleEl.textContent = 'Viser nesting + hovedbilleder';

                let html = '<div class="compact-image-grid">';
                for (const item of items) {
                    const src = getSummaryImageSrc(item);
                    html += '<div class="compact-image-card">';
                    html += '<div class="compact-image-label">' + escapeHtml(item.label || 'Billede') + '</div>';
                    html += '<img class="image-preview-zoomable" src="' + escapeHtml(src) + '" alt="' + escapeHtml(item.label || entry.title || 'Billede') + '" loading="lazy" data-fullsrc="' + escapeHtml(src) + '" data-title="' + escapeHtml(item.label || entry.title || 'Billede') + '" data-path="' + escapeHtml(item.value || '') + '" />';
                    html += '<div class="compact-image-path">' + escapeHtml(item.value || '') + '</div>';
                    html += '</div>';
                }
                html += '</div>';
                bodyEl.innerHTML = html;
                modal.classList.add('show');
            }

            function closeCompactImageModal(event) {
                if (event && event.target && event.target.id !== 'compactImageModal') return;
                const modal = document.getElementById('compactImageModal');
                const bodyEl = document.getElementById('compactImageBody');
                if (!modal) return;
                modal.classList.remove('show');
                if (bodyEl) bodyEl.innerHTML = '';
            }

            function closeSummaryImagePanel() {
                const panels = [
                    document.getElementById('summaryImagePanel'),
                    document.getElementById('laserImagePanel')
                ];
                for (const panel of panels) {
                    if (!panel) continue;
                    panel.innerHTML = '';
                    panel.classList.add('hidden');
                }
                updateSummaryImagePanelLayout();
            }

            function updateSummaryImagePanelLayout() {
                const wrap = document.querySelector('#summaryModal .modal-content-wrap');
                const summaryPanel = document.getElementById('summaryImagePanel');
                if (!wrap || !summaryPanel) return;
                const hasImages = !summaryPanel.classList.contains('hidden') && summaryPanel.innerHTML.trim() !== '';
                const shouldFocus = hasImages && window.matchMedia('(max-width: 1440px)').matches;
                wrap.classList.toggle('image-focus', shouldFocus);
            }

            function openSummaryImagePanel(imageKey, preferredPanelId) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const laserPanelWrap = document.getElementById('laserOrderSummaryPanel');
                const laserPanel = document.getElementById('laserImagePanel');
                const summaryPanel = document.getElementById('summaryImagePanel');
                const isLaserVisible = laserPanelWrap && laserPanelWrap.style.display !== 'none';
                const isVisible = (el) => {
                    if (!el) return false;
                    const s = getComputedStyle(el);
                    return s.display !== 'none' && s.visibility !== 'hidden' && el.getClientRects().length > 0;
                };
                let panel = null;
                if (preferredPanelId === 'laserImagePanel') {
                    panel = isVisible(laserPanel) ? laserPanel : summaryPanel;
                } else if (preferredPanelId === 'summaryImagePanel') {
                    panel = summaryPanel;
                } else {
                    panel = (isLaserVisible && laserPanel) ? laserPanel : summaryPanel;
                }
                const entry = summaryImageRegistry[imageKey];
                if (!panel || !entry || !Array.isArray(entry.items) || entry.items.length === 0) {
                    closeSummaryImagePanel();
                    return;
                }

                if (panel.id === 'summaryImagePanel' && title) {
                    title.textContent = entry.title || 'Billeder';
                }
                if (panel.id === 'summaryImagePanel' && modal && modal.style.display !== 'flex') {
                    modal.style.display = 'flex';
                }

                let html = '<div class="summary-image-panel-header">';
                html += '<div class="summary-image-panel-title">' + escapeHtml(entry.title) + '</div>';
                html += '<button class="summary-image-close" onclick="closeSummaryImagePanel()">Luk</button>';
                html += '</div>';
                html += '<div class="image-preview-gallery">';

                for (const item of entry.items) {
                    const src = getSummaryImageSrc(item);
                    html += '<div class="image-preview-card">';
                    html += '<div class="image-preview-label">' + escapeHtml(item.label || 'Billede') + '</div>';
                    html += '<img class="image-preview-zoomable" src="' + escapeHtml(src) + '" alt="' + escapeHtml(entry.title) + '" loading="lazy" data-fullsrc="' + escapeHtml(src) + '" data-title="' + escapeHtml(item.label || entry.title || 'Billede') + '" data-path="' + escapeHtml(item.value || '') + '" />';
                    html += '<div class="image-preview-path">' + escapeHtml(item.value || '') + '</div>';
                    html += '</div>';
                }

                html += '</div>';
                panel.innerHTML = html;
                panel.classList.remove('hidden');
                updateSummaryImagePanelLayout();
            }

            function openImageLightbox(src, title, pathText) {
                const lightbox = document.getElementById('imageLightbox');
                const img = document.getElementById('imageLightboxImg');
                const titleEl = document.getElementById('imageLightboxTitle');
                const pathEl = document.getElementById('imageLightboxPath');
                if (!lightbox || !img) return;

                img.src = src || '';
                img.alt = title || 'Billede';
                if (titleEl) titleEl.textContent = title || 'Billede';
                if (pathEl) pathEl.textContent = pathText || '';
                lightbox.classList.remove('hidden');
            }

            function closeImageLightbox(event) {
                if (event && event.target && event.target.id !== 'imageLightbox') return;
                const lightbox = document.getElementById('imageLightbox');
                const img = document.getElementById('imageLightboxImg');
                const pathEl = document.getElementById('imageLightboxPath');
                if (!lightbox || lightbox.classList.contains('hidden')) return;

                lightbox.classList.add('hidden');
                if (img) {
                    img.src = '';
                    img.alt = '';
                }
                if (pathEl) pathEl.textContent = '';
            }

            function updateSummaryModalBackBtn() {
                const backBtn = document.getElementById('summaryModalBackBtn');
                if (!backBtn) return;
                backBtn.classList.toggle('hidden', summaryModalHistory.length === 0);
            }

            function pushSummaryModalState() {
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');
                const imagePanel = document.getElementById('summaryImagePanel');
                if (!title || !body) return;
                summaryModalHistory.push({
                    title: title.textContent,
                    bodyHtml: body.innerHTML,
                    imageHtml: imagePanel ? imagePanel.innerHTML : '',
                    imageHidden: imagePanel ? imagePanel.classList.contains('hidden') : true
                });
                updateSummaryModalBackBtn();
            }

            function goSummaryModalBack() {
                if (summaryModalHistory.length === 0) return;
                const prev = summaryModalHistory.pop();
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');
                const imagePanel = document.getElementById('summaryImagePanel');
                if (title) title.textContent = prev.title;
                if (body) body.innerHTML = prev.bodyHtml;
                if (imagePanel) {
                    imagePanel.innerHTML = prev.imageHtml || '';
                    imagePanel.classList.toggle('hidden', prev.imageHidden !== false);
                }
                updateSummaryImagePanelLayout();
                updateSummaryModalBackBtn();
            }

            function setSystemStatus(text, bgColor, textColor) {
                // Header right area now shows user greeting instead of system status.
            }

            // Warmup progress bar polling
            let warmupPollTimer = null;
            let warmupTopBarHideScheduled = false;
            let warmupCombinedReady = false;
            let warmupCombinedPct = 0;
            let warmupCombinedDone = 0;
            let warmupCombinedTotal = 0;
            let showDashboardWarmupNotice = false;
            function startWarmupPolling() {
                if (warmupPollTimer) return;
                const wrap = document.getElementById('warmupBarWrap');
                const fill = document.getElementById('warmupBarFill');
                const txt  = document.getElementById('warmupBarText');
                const dashText = document.getElementById('dashboardWarmupText');
                const dashMeta = document.getElementById('dashboardWarmupMeta');
                const dashFill = document.getElementById('dashboardWarmupFill');
                const dashPct = document.getElementById('dashboardWarmupPct');
                const dashWrap = document.getElementById('dashboardWarmupNotice');
                if (!wrap && !dashText) return;

                warmupPollTimer = setInterval(async () => {
                    try {
                        const r = await fetch('/warmup-status');
                        if (!r.ok) return;
                        const d = await r.json();

                        const totalCombined = Number(d.combinedTotal || d.total || 0);
                        const doneCombined = Number(d.combinedDone || d.done || 0);
                        const pctCombined = Number(d.combinedPct || d.pct || 0);
                        const readyCombined = d.ready === true;

                        warmupCombinedPct = Math.max(0, Math.min(100, pctCombined));
                        warmupCombinedDone = Math.max(0, doneCombined);
                        warmupCombinedTotal = Math.max(0, totalCombined);
                        warmupCombinedReady = readyCombined || (!d.running && warmupCombinedTotal === 0);

                        if (dashFill) dashFill.style.width = String(Math.max(0, Math.min(100, pctCombined))) + '%';
                        if (dashPct) dashPct.textContent = String(Math.max(0, Math.min(100, pctCombined))) + '%';

                        // Keep warmup hidden on initial dashboard screen, unless user explicitly triggered cache reset.
                        if (dashWrap) {
                            const shouldShowDashWarmup = d.running || (!readyCombined && totalCombined > 0) || showDashboardWarmupNotice;
                            dashWrap.classList.toggle('hidden', !shouldShowDashWarmup);
                        }

                        if (dashText) {
                            if (d.running) {
                                dashText.textContent = 'Forbereder ' + doneCombined + '/' + totalCombined + ' ordredata...';
                                if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                            } else if (readyCombined && totalCombined > 0) {
                                dashText.textContent = 'Klar! Efterkalk-data er forberedt.';
                                if (dashMeta) dashMeta.textContent = 'Åbn Efterkalk når som helst.';
                                if (dashWrap) {
                                    setTimeout(() => {
                                        dashWrap.classList.add('hidden');
                                        showDashboardWarmupNotice = false;
                                    }, 1800);
                                }
                            } else if (totalCombined > 0) {
                                dashText.textContent = 'Afventer baggrundsjob...';
                                if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                            } else {
                                dashText.textContent = 'Venter på warmup-status...';
                                if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                            }
                        }

                        if (d.total === 0) {
                            if (wrap) wrap.classList.remove('active');
                            return;
                        }

                        if (wrap) wrap.classList.add('active');
                        if (fill) fill.style.width = d.pct + '%';

                        if (d.running) {
                            if (txt) txt.textContent = 'Forberegner ' + d.done + '/' + d.total + ' ordrer...';
                            warmupTopBarHideScheduled = false;
                        } else {
                            if (txt) txt.textContent = 'Klar! ' + d.loaded + ' nye + ' + d.cached + ' fra cache';
                            if (fill) fill.style.width = '100%';
                            if (!warmupTopBarHideScheduled) {
                                warmupTopBarHideScheduled = true;
                                setTimeout(() => {
                                    if (wrap) wrap.classList.remove('active');
                                    warmupTopBarHideScheduled = false;
                                }, 3000);
                            }
                        }
                    } catch(e) {
                        // ignore polling errors silently
                    }
                }, 800);
            }
            function updateSystemStatusFromOrders(orders) {
                if (!orders || orders.length === 0) {
                    setSystemStatus('System klar', '#e8f5e9', '#1b5e20');
                    return;
                }

                const visibleOrders = orders.slice(0, MARGIN_PREFETCH_ROWS);
                const total = visibleOrders.length;
                let completed = 0;

                for (const o of visibleOrders) {
                    const state = getMarginState(o.OrdNo);
                    if (state && (state.status === 'success' || state.status === 'error')) {
                        completed += 1;
                    }
                }

                if (completed >= total) {
                    setSystemStatus('System klar', '#e8f5e9', '#1b5e20');
                    return;
                }

                setSystemStatus('System indlæser... ' + completed + '/' + total, '#fff3cd', '#8a6d3b');
            }

            function getMarginModeLabel() {
                return currentMarginMode === 'new'
                    ? 'Ny (Salg/Kost x 100)'
                    : 'Klassisk ((Salg-Kost)/Salg x 100)';
            }

            function calculateOrderMarginPercent(revenue, cost) {
                if (currentMarginMode === 'new') {
                    return cost > 0 ? ((revenue / cost) * 100) : 0;
                }
                return revenue > 0 ? (((revenue - cost) / revenue) * 100) : 0;
            }

            function calculateLineMarginPercent(salesPrice, lineCost) {
                if (currentMarginMode === 'new') {
                    return lineCost > 0 ? ((salesPrice / lineCost) * 100) : 0;
                }
                return salesPrice > 0 ? (((salesPrice - lineCost) / salesPrice) * 100) : 0;
            }

            function toggleMarginMode() {
                currentMarginMode = currentMarginMode === 'new' ? 'classic' : 'new';
                renderOrderList();
            }

            function scheduleOrderListRerender() {
                if (orderListRerenderTimer) return;
                orderListRerenderTimer = setTimeout(() => {
                    orderListRerenderTimer = null;
                    renderOrderList();
                }, 120);
            }

            function getMarginState(ordNo) {
                return marginStateByOrdNo[String(ordNo)] || null;
            }

            function getOrderInvoiceStatusHtml(ordNo) {
                const marginState = getMarginState(ordNo);
                if (!marginState || marginState.status !== 'success') return '<span style="color:#999;">-</span>';
                if (marginState.hasInvoiceWarning) {
                    return '<span title="En eller flere linjer mangler faktura (NoInvo=0); kostberegning bruger NoFin som fallback." style="display:inline-flex;align-items:center;gap:4px;background:#fff3e0;color:#e65100;font-size:12px;font-weight:600;padding:2px 7px;border-radius:10px;border:1px solid #ffcc80;">🧾 Mangler</span>';
                }
                return '<span style="color:#388e3c;font-size:13px;" title="Alle fakturaer registreret.">✓</span>';
            }

            function updateOrderInvoiceCell(ordNo) {
                const listEl = document.getElementById('orderList');
                if (!listEl) return;
                const cells = listEl.querySelectorAll('.order-invoice-cell[data-ordno="' + ordNo + '"]');
                const html = getOrderInvoiceStatusHtml(ordNo);
                for (const cell of cells) { cell.innerHTML = html; }
            }

            function getOrderMarginHtml(ordNo) {
                const marginState = getMarginState(ordNo);
                let marginHtml = '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>';
                let toneClass = 'na';
                if (marginState && marginState.status === 'loading') {
                    marginHtml = '<span style="background:#546e7a; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">...</span>';
                    toneClass = 'na';
                } else if (marginState && marginState.status === 'success') {
                    const marginValue = calculateOrderMarginPercent(marginState.totalRevenue || 0, marginState.totalCost || 0);
                    const margin = marginValue.toFixed(2);
                    marginHtml = getMarginBadge(margin);
                    if (currentMarginMode === 'new') {
                        toneClass = marginValue >= 125 ? 'ok' : (marginValue >= 105 ? 'warn' : 'bad');
                    } else {
                        toneClass = marginValue > 20 ? 'ok' : (marginValue >= 5 ? 'warn' : 'bad');
                    }
                }
                return '<span class="order-margin-wrap"><span class="order-kpi-tone ' + toneClass + '" title="KPI tone fra Rapport 2.0">↗</span>' + marginHtml + '</span>';
            }

            function updateOrderMarginCell(ordNo) {
                const listEl = document.getElementById('orderList');
                if (!listEl) return;
                const cells = listEl.querySelectorAll('.order-margin-cell[data-ordno="' + ordNo + '"]');
                if (!cells || cells.length === 0) return;
                const marginHtml = getOrderMarginHtml(ordNo);
                for (const cell of cells) {
                    cell.innerHTML = marginHtml;
                }
                updateOrderListSummaryPanel();
            }

            function refreshOrderListStatus() {
                if (!orderListVisible) return;
                const visibleOrders = getFilteredOrders().slice(0, MARGIN_PREFETCH_ROWS);
                updateSystemStatusFromOrders(visibleOrders);
            }

            function scheduleMarginSortRefresh() {
                if (orderListSortField !== 'margin') return;
                if (marginSortRefreshTimer) return;
                marginSortRefreshTimer = setTimeout(() => {
                    marginSortRefreshTimer = null;
                    renderOrderList();
                }, 350);
            }

            function hydrateMarginStateFromOrderList(orders) {
                marginStateByOrdNo = {};
                for (const o of orders) {
                    const ordNo = Number(o.OrdNo);
                    if (!Number.isFinite(ordNo)) continue;

                    if (o.TotalCost !== null && o.TotalCost !== undefined) {
                        marginStateByOrdNo[String(ordNo)] = {
                            status: 'success',
                            totalRevenue: Number(o.InvoAm || 0),
                            totalCost: Number(o.TotalCost || 0)
                        };
                    }
                }
            }

            function queueMarginLoad(ordNos) {
                for (const ordNo of ordNos) {
                    const key = String(ordNo);
                    const existing = marginStateByOrdNo[key];
                    if (existing && (existing.status === 'success' || existing.status === 'loading')) {
                        continue;
                    }

                    marginStateByOrdNo[key] = { status: 'loading' };
                    marginJobQueue.push(Number(ordNo));
                    updateOrderMarginCell(ordNo);
                }
                pumpMarginQueue();
                refreshOrderListStatus();
            }

            async function loadSingleOrderMargin(ordNo) {
                const key = String(ordNo);
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), MARGIN_FETCH_TIMEOUT_MS);
                try {
                    const response = await fetch('/order-margin/' + ordNo, { signal: controller.signal });
                    let data = null;
                    try {
                        data = await response.json();
                    } catch {
                        data = { error: 'Invalid JSON response' };
                    }
                    if (!response.ok || data.error) {
                        marginStateByOrdNo[key] = { status: 'error' };
                        updateOrderMarginCell(ordNo);
                        refreshOrderListStatus();
                        scheduleMarginSortRefresh();
                        return;
                    }

                    marginStateByOrdNo[key] = {
                        status: 'success',
                        totalRevenue: Number(data.totalRevenue || 0),
                        totalCost: Number(data.totalCost || 0),
                        hasInvoiceWarning: Boolean(data.hasInvoiceWarning)
                    };
                    updateOrderMarginCell(ordNo);
                    updateOrderInvoiceCell(ordNo);
                    refreshOrderListStatus();
                    scheduleMarginSortRefresh();
                } catch (err) {
                    marginStateByOrdNo[key] = { status: 'error' };
                    updateOrderMarginCell(ordNo);
                    refreshOrderListStatus();
                    scheduleMarginSortRefresh();
                } finally {
                    clearTimeout(timeoutId);
                }
            }

            function pumpMarginQueue() {
                while (marginWorkerActiveCount < MARGIN_MAX_CONCURRENT && marginJobQueue.length > 0) {
                    const ordNo = marginJobQueue.shift();
                    marginWorkerActiveCount += 1;

                    loadSingleOrderMargin(ordNo)
                        .finally(() => {
                            marginWorkerActiveCount -= 1;
                            setTimeout(pumpMarginQueue, MARGIN_QUEUE_DELAY_MS);
                        });
                }
            }

            function toggleOrderList() {
                orderListVisible = !orderListVisible;
                renderOrderList();
            }

            function setOrderListFilter() {
                const input = document.getElementById('customerFilterInput');
                orderListFilter = (input && input.value ? input.value : '').trim().toLowerCase();
                if (!orderListVisible && orderListFilter) {
                    orderListVisible = true;
                }
                renderOrderList();
            }

            function setBrugerFilter() {
                const input = document.getElementById('brugerFilterSelect');
                orderListBrugerFilter = (input && input.value ? input.value : '').trim();
                if (!orderListVisible && orderListBrugerFilter) {
                    orderListVisible = true;
                }
                renderOrderList();
            }

            function setOrderValueFilter() {
                const enabledInput = document.getElementById('orderMinDkkEnabled');
                const thresholdInput = document.getElementById('orderMinDkkInput');
                orderListMinDkkEnabled = !!(enabledInput && enabledInput.checked);

                const raw = Number(thresholdInput && thresholdInput.value || 0);
                orderListMinDkkValue = Number.isFinite(raw) ? Math.max(0, Math.round(raw)) : 0;

                if (thresholdInput) {
                    thresholdInput.disabled = !orderListMinDkkEnabled;
                    if (!thresholdInput.disabled && String(thresholdInput.value || '') !== String(orderListMinDkkValue)) {
                        thresholdInput.value = String(orderListMinDkkValue);
                    }
                }

                if (!orderListVisible && orderListMinDkkEnabled && orderListMinDkkValue > 0) {
                    orderListVisible = true;
                }
                renderOrderList();
            }

            function populateBrugerFilterOptions() {
                const select = document.getElementById('brugerFilterSelect');
                if (!select) return;

                const selectedValue = orderListBrugerFilter;
                const users = Array.from(new Set(
                    orderListData
                        .map(o => String(o.SellerUsr || '').trim())
                        .filter(v => v)
                )).sort((a, b) => a.localeCompare(b));

                let html = '<option value="">Alle brugere</option>';
                for (const user of users) {
                    html += '<option value="' + user + '">' + user + '</option>';
                }
                select.innerHTML = html;
                select.value = selectedValue;
            }

            function setOrderListSort(field) {
                if (orderListSortField === field) {
                    orderListSortDir = orderListSortDir === 'asc' ? 'desc' : 'asc';
                } else {
                    orderListSortField = field;
                    orderListSortDir = field === 'date' || field === 'ordno' || field === 'belob' || field === 'margin' ? 'desc' : 'asc';
                }
                renderOrderList();
            }

            function getMarginValue(ordNo) {
                const state = getMarginState(ordNo);
                if (!state || state.status !== 'success') return null;
                return calculateOrderMarginPercent(state.totalRevenue || 0, state.totalCost || 0);
            }

            function getFilteredOrders() {
                const filtered = orderListData.filter(o => {
                    const bruger = String(o.SellerUsr || '').trim();
                    const customer = String(o.CustomerName || '').toLowerCase();
                    const ord = String(o.OrdNo || '');
                    const matchesText = !orderListFilter || customer.includes(orderListFilter) || ord.includes(orderListFilter);
                    const matchesBruger = !orderListBrugerFilter || bruger === orderListBrugerFilter;
                    const invoDkk = Number(o.InvoAm || 0);
                    const matchesMinDkk = !orderListMinDkkEnabled || invoDkk >= orderListMinDkkValue;
                    return matchesText && matchesBruger && matchesMinDkk;
                });

                const dir = orderListSortDir === 'asc' ? 1 : -1;
                filtered.sort((a, b) => {
                    switch (orderListSortField) {
                        case 'bruger': {
                            const cmp = String(a.SellerUsr || '').localeCompare(String(b.SellerUsr || ''));
                            return cmp * dir || Number(b.LstInvDt || 0) - Number(a.LstInvDt || 0);
                        }
                        case 'ordno':
                            return (Number(a.OrdNo || 0) - Number(b.OrdNo || 0)) * dir;
                        case 'kunde': {
                            const cmp = String(a.CustomerName || '').localeCompare(String(b.CustomerName || ''));
                            return cmp * dir || Number(b.OrdNo || 0) - Number(a.OrdNo || 0);
                        }
                        case 'date': {
                            const d = (Number(a.LstInvDt || 0) - Number(b.LstInvDt || 0)) * dir;
                            return d || (Number(b.OrdNo || 0) - Number(a.OrdNo || 0)) * dir;
                        }
                        case 'belob':
                            return (Number(a.InvoAm || 0) - Number(b.InvoAm || 0)) * dir;
                        case 'margin': {
                            const ma = getMarginValue(a.OrdNo);
                            const mb = getMarginValue(b.OrdNo);
                            if (ma === null && mb === null) return 0;
                            if (ma === null) return 1;
                            if (mb === null) return -1;
                            return (ma - mb) * dir;
                        }
                        default:
                            return Number(b.LstInvDt || 0) - Number(a.LstInvDt || 0);
                    }
                });
                return filtered;
            }

            function getMarginBadge(marginPercent) {
                const margin = parseFloat(marginPercent);
                if (currentMarginMode === 'new') {
                    if (margin >= 125) {
                        return '<span style="background:#2e7d32; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">✅ ' + marginPercent + '%</span>';
                    } else if (margin >= 105) {
                        return '<span style="background:#ff9800; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">⚠️ ' + marginPercent + '%</span>';
                    }
                    return '<span style="background:#d32f2f; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">❌ ' + marginPercent + '%</span>';
                }

                if (margin > 20) {
                    return '<span style="background:#2e7d32; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">✅ ' + marginPercent + '%</span>';
                } else if (margin >= 5) {
                    return '<span style="background:#ff9800; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">⚠️ ' + marginPercent + '%</span>';
                }
                return '<span style="background:#d32f2f; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">❌ ' + marginPercent + '%</span>';
            }

            function getFilteredOrderSummary(orders) {
                const safeOrders = Array.isArray(orders) ? orders : [];
                let considered = 0;
                let excludedCredit = 0;
                let pendingMargin = 0;
                let totalRevenue = 0;
                let totalCost = 0;

                for (const o of safeOrders) {
                    if (isOrderMarkedCreditNote(o.OrdNo)) {
                        excludedCredit += 1;
                        continue;
                    }
                    const state = getMarginState(o.OrdNo);
                    if (!state || state.status !== 'success') {
                        pendingMargin += 1;
                        continue;
                    }
                    considered += 1;
                    totalRevenue += Number(state.totalRevenue || 0);
                    totalCost += Number(state.totalCost || 0);
                }

                const marginAmount = totalRevenue - totalCost;
                const marginPct = totalCost > 0 ? calculateOrderMarginPercent(totalRevenue, totalCost).toFixed(2) : '0.00';
                return {
                    considered,
                    excludedCredit,
                    pendingMargin,
                    totalRevenue,
                    totalCost,
                    marginAmount,
                    marginPct
                };
            }

            function buildOrderListSummaryHtml(orders) {
                const listSummary = getFilteredOrderSummary(orders);
                const activeFilters = [];
                if (orderListFilter) activeFilters.push('kunde/søgning: "' + escapeHtml(orderListFilter) + '"');
                if (orderListBrugerFilter) activeFilters.push('bruger: "' + escapeHtml(orderListBrugerFilter) + '"');
                if (orderListMinDkkEnabled) activeFilters.push('minimum fakturabeløb: ' + formatNumber(orderListMinDkkValue) + ' DKK');
                const filterText = activeFilters.length > 0 ? activeFilters.join(', ') : 'ingen aktive filtre';
                let html = '<div><strong>Filtreret ordrelisteoversigt</strong> (vist: ' + orders.length + ', medtaget: ' + listSummary.considered + ', kreditnota udelukket: ' + listSummary.excludedCredit + ', mangler margin: ' + listSummary.pendingMargin + ')</div>';
                html += '<div style="margin-top:4px; font-size:12px; color:#57718f;">Genereret: ' + escapeHtml(new Date().toLocaleString('da-DK')) + ' • Filtre: ' + filterText + '</div>';
                html += '<div style="margin-top:6px;display:flex;gap:18px;flex-wrap:wrap;">';
                html += '<span>Samlet omsætning: <strong>' + formatNumber(listSummary.totalRevenue) + ' DKK</strong></span>';
                html += '<span>Samlet kost: <strong>' + formatNumber(listSummary.totalCost) + ' DKK</strong></span>';
                html += '<span>Margin: <strong>' + formatNumber(listSummary.marginAmount) + ' DKK (' + listSummary.marginPct + '%)</strong></span>';
                html += '</div>';
                html += '<div class="order-list-summary-actions">';
                html += '<button class="list-toggle-btn" onclick="openOrderListPrintPreview()" title="Vis forhåndsvisning af den filtrerede ordreliste">Forhåndsvisning / PDF</button>';
                html += '</div>';
                return html;
            }

            function buildOrderDetailReportHtml(orderData, orderMarginPercent, costToDateFromProduction) {
                const orderHeader = (orderData && orderData.orderHeader) || {};
                const productionOrders = Array.isArray(orderData && orderData.productionOrders) ? orderData.productionOrders : [];
                const salesOrderLines = Array.isArray(orderData && orderData.salesOrderLines) ? orderData.salesOrderLines : [];
                const salesLines = Array.isArray(orderData && orderData.salesLines) ? orderData.salesLines : [];

                let totalPlannedMinutes = 0;
                let totalUsedMinutes = 0;
                let totalOperationCost = 0;
                let totalLaserCost = 0;
                const rows = [];
                const exceptionMap = new Map();

                const pushException = (type, prodOrdNo, ref, message) => {
                    const text = String(message || '').trim();
                    if (!text) return;
                    const normalizedType = String(type || '-').trim() || '-';
                    const normalizedProdOrdNo = String(prodOrdNo || '-').trim() || '-';
                    const normalizedRef = String(ref || '-').trim() || '-';
                    const key = normalizedType + '|' + normalizedProdOrdNo + '|' + normalizedRef + '|' + text;
                    const existing = exceptionMap.get(key);
                    if (existing) {
                        existing.count += 1;
                        return;
                    }
                    exceptionMap.set(key, {
                        type: normalizedType,
                        prodOrdNo: normalizedProdOrdNo,
                        ref: normalizedRef,
                        message: text,
                        count: 1
                    });
                };

                for (const prodOrder of productionOrders) {
                    const lines = Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : [];
                    let plannedMinutes = 0;
                    let usedMinutes = 0;
                    let operationCost = 0;
                    let laserCost = 0;
                    let materialCost = 0;
                    let productLabel = '-';

                    for (const line of lines) {
                        const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                        const lnNo = Number((line && line.LnNo) || 0);
                        if (lnNo === 1 && line && line.ProdNo) {
                            productLabel = String(line.ProdNo || '-') + ' - ' + String(line.Descr || '');
                        }
                        if (lnNo === 1 || key === '0' || key === '3' || key === '5') continue;

                        const totalCost = Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                        if (key === '1') {
                            plannedMinutes += Number((line && line.NoOrg) || 0);
                            usedMinutes += Number((line && (line.EffectiveOperationMinutes ?? line.NoFin)) || 0);
                            operationCost += totalCost;
                        } else if (key === '2') {
                            laserCost += totalCost;
                            if (!isLaserLProdNo(line && line.ProdNo)) {
                                materialCost += totalCost;
                            }
                        }

                        if (line && line.HasWarning && line.WarningText) {
                            pushException('Advarsel', prodOrder.ordNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.WarningText);
                        }
                        if (line && line.IsInvoiceTracked && line.UsesMissingInvoiceFallback) {
                            pushException('Faktura', prodOrder.ordNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.MissingInvoiceText || 'Mangler faktura');
                        }
                        if (line && line.UsesEstimatedOperationTime) {
                            pushException('Tid', prodOrder.ordNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.EstimatedTimeText || 'Færdigmeldt minutter var 0 og blev estimeret');
                        }
                    }

                    totalPlannedMinutes += plannedMinutes;
                    totalUsedMinutes += usedMinutes;
                    totalOperationCost += operationCost;
                    totalLaserCost += laserCost;
                    rows.push({
                        ordNo: Number(prodOrder && prodOrder.ordNo) || 0,
                        productLabel,
                        plannedMinutes,
                        usedMinutes,
                        deltaMinutes: usedMinutes - plannedMinutes,
                        operationCost,
                        laserCost,
                        materialCost,
                        totalCost: Number(prodOrder && prodOrder.totalCost) || 0
                    });
                }

                for (const line of salesOrderLines) {
                    if (line && line.HasWarning && line.WarningText) {
                        pushException('Salgsordre', line.PurcNo || orderHeader.OrdNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.WarningText);
                    }
                    if (line && line.IsInvoiceTracked && line.UsesMissingInvoiceFallback) {
                        pushException('Faktura', line.PurcNo || orderHeader.OrdNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.MissingInvoiceText || 'Mangler faktura');
                    }
                    if (line && line.UsesEstimatedOperationTime) {
                        pushException('Tid', line.PurcNo || orderHeader.OrdNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.EstimatedTimeText || 'Færdigmeldt minutter var 0 og blev estimeret');
                    }
                }

                for (const line of salesLines) {
                    if (line && line.HasWarning && line.WarningText) {
                        pushException('Ekstra linje', line.PurcNo || orderHeader.OrdNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.WarningText);
                    }
                }

                const exceptionRows = Array.from(exceptionMap.values())
                    .sort((a, b) => {
                        const aProd = Number(a.prodOrdNo) || 0;
                        const bProd = Number(b.prodOrdNo) || 0;
                        if (aProd !== bProd) return aProd - bProd;
                        if (a.type !== b.type) return a.type.localeCompare(b.type, 'da');
                        if (a.ref !== b.ref) return a.ref.localeCompare(b.ref, 'da');
                        return a.message.localeCompare(b.message, 'da');
                    });

                const exceptionCompactMap = new Map();
                for (const ex of exceptionRows) {
                    const prodOrdNo = String(ex.prodOrdNo || '-');
                    const message = String(ex.message || '').trim() || '-';
                    const compactKey = prodOrdNo + '|' + message;
                    if (!exceptionCompactMap.has(compactKey)) {
                        exceptionCompactMap.set(compactKey, {
                            prodOrdNo,
                            message,
                            typeSet: new Set(),
                            refSet: new Set(),
                            count: 0
                        });
                    }
                    const row = exceptionCompactMap.get(compactKey);
                    row.typeSet.add(String(ex.type || '-'));
                    row.refSet.add(String(ex.ref || '-'));
                    row.count += Number(ex.count || 0);
                }

                const exceptionCompactRows = Array.from(exceptionCompactMap.values())
                    .map(row => ({
                        prodOrdNo: row.prodOrdNo,
                        types: Array.from(row.typeSet.values()).sort((a, b) => a.localeCompare(b, 'da')),
                        refs: Array.from(row.refSet.values()).sort((a, b) => a.localeCompare(b, 'da')),
                        message: row.message,
                        count: row.count
                    }))
                    .sort((a, b) => {
                        const aProd = Number(a.prodOrdNo) || 0;
                        const bProd = Number(b.prodOrdNo) || 0;
                        if (aProd !== bProd) return aProd - bProd;
                        if (b.count !== a.count) return b.count - a.count;
                        return a.message.localeCompare(b.message, 'da');
                    });

                const exceptionOccurrenceCount = exceptionRows.reduce((sum, item) => sum + Number(item.count || 0), 0);
                const exceptionGroupCount = exceptionCompactRows.length;

                const marginAmount = Number((orderData && orderData.summary && orderData.summary.margin) || 0);
                const marginPct = orderMarginPercent || '0.00';
                const revenue = Number((orderData && orderData.summary && orderData.summary.totalRevenue) || 0);
                const cost = Number((orderData && orderData.summary && orderData.summary.totalCost) || 0);
                const generatedAt = new Date().toLocaleString('da-DK');
                const orderTypeLabel = Number(orderHeader.Gr4 || 0) === 3 ? 'Multiordre' : 'Ordre';
                const statusLabel = Number(orderHeader.InvoAm || 0) === 0
                    ? 'I produktion'
                    : (Number(orderHeader.DInvoIF || 0) <= 0 ? 'Komplet faktureret' : 'Delvist faktureret');

                let html = '<div id="orderDetailReport" class="order-detail-report">';
                html += '<div class="order-report-toolbar">';
                html += '<div class="order-report-meta"><strong>' + escapeHtml(orderTypeLabel) + ' ' + escapeHtml(String(orderHeader.OrdNo || '-')) + '</strong><div class="report-subline">' + escapeHtml(String(orderHeader.CustomerName || '-')) + '</div><div class="report-badges"><span class="report-badge">Status: <strong>' + escapeHtml(statusLabel) + '</strong></span><span class="report-badge">Margin: <strong>' + escapeHtml(marginPct) + '%</strong></span></div></div>';
                html += '</div>';
                html += '<div class="report-hero">';
                html += '<div class="report-hero-top">';
                html += '<div class="report-hero-title">';
                html += '<div class="eyebrow">Ledelsesoverblik</div>';
                html += '<h1>Rapport for ordre ' + escapeHtml(String(orderHeader.OrdNo || '-')) + '</h1>';
                html += '<div class="context">' + escapeHtml(String(orderHeader.CustomerName || '-')) + ' • Genereret ' + escapeHtml(generatedAt) + '</div>';
                html += '<div class="report-arrow">Forbedringsretning i grøn KPI-tone</div>';
                html += '</div>';
                html += '<div class="report-hero-meta">';
                html += '<div class="stamp">' + escapeHtml(orderTypeLabel) + ' • ' + escapeHtml(statusLabel) + '</div>';
                html += '<div class="stamp">' + escapeHtml(currentMarginMode === 'new' ? 'Ny marginmodel' : 'Klassisk marginmodel') + '</div>';
                html += '</div>';
                html += '</div>';
                html += '<div class="report-pill-row">';
                html += '<span class="report-pill ok"><strong>' + formatNumber(revenue) + ' DKK</strong> omsætning</span>';
                html += '<span class="report-pill"><strong>' + formatNumber(cost) + ' DKK</strong> kost</span>';
                html += '<span class="report-pill"><strong>' + formatNumber(marginAmount) + ' DKK</strong> margin</span>';
                html += '<span class="report-pill warn"><strong>' + formatCount(exceptionGroupCount) + '</strong> spor/advarsler</span>';
                html += '</div>';
                html += '</div>';

                html += '<div class="order-report-grid">';
                html += '<div class="order-report-card"><div class="label">Samlet omsætning</div><div class="value">' + formatNumber(revenue) + ' DKK</div></div>';
                html += '<div class="order-report-card"><div class="label">Samlet kost</div><div class="value">' + formatNumber(cost) + ' DKK</div></div>';
                html += '<div class="order-report-card"><div class="label">Margin</div><div class="value">' + formatNumber(marginAmount) + ' DKK (' + marginPct + '%)</div></div>';
                html += '<div class="order-report-card"><div class="label">Produktionsordrer</div><div class="value">' + formatNumber(productionOrders.length) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Planlagte minutter</div><div class="value">' + formatNumber(totalPlannedMinutes) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Brugte minutter</div><div class="value">' + formatNumber(totalUsedMinutes) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Operation kost</div><div class="value">' + formatNumber(totalOperationCost) + ' DKK</div></div>';
                html += '<div class="order-report-card"><div class="label">Laser / materiale kost</div><div class="value">' + formatNumber(totalLaserCost) + ' DKK</div></div>';
                html += '</div>';

                html += '<table class="order-report-table">';
                html += '<tr><th>Produktionsordre</th><th>Produkt</th><th>Planlagt min.</th><th>Brugt min.</th><th>Afvigelse</th><th>Operation kost</th><th>Laser / materiale kost</th><th>Samlet kost</th></tr>';
                for (const row of rows) {
                    html += '<tr>';
                    html += '<td>' + row.ordNo + '</td>';
                    html += '<td>' + escapeHtml(row.productLabel || '-') + '</td>';
                    html += '<td>' + formatNumber(row.plannedMinutes || 0) + '</td>';
                    html += '<td>' + formatNumber(row.usedMinutes || 0) + '</td>';
                    html += '<td>' + formatNumber(row.deltaMinutes || 0) + '</td>';
                    html += '<td>' + formatNumber(row.operationCost || 0) + ' DKK</td>';
                    html += '<td>' + formatNumber(row.laserCost || 0) + ' DKK</td>';
                    html += '<td><strong>' + formatNumber(row.totalCost || 0) + ' DKK</strong></td>';
                    html += '</tr>';
                }
                if (rows.length === 0) {
                    html += '<tr><td colspan="8">Ingen produktionsordrer fundet.</td></tr>';
                }
                html += '<tr class="summary-row"><td colspan="2">Samlet</td><td>' + formatNumber(totalPlannedMinutes || 0) + '</td><td>' + formatNumber(totalUsedMinutes || 0) + '</td><td>' + formatNumber(totalUsedMinutes - totalPlannedMinutes) + '</td><td>' + formatNumber(totalOperationCost || 0) + ' DKK</td><td>' + formatNumber(totalLaserCost || 0) + ' DKK</td><td><strong>' + formatNumber(cost || 0) + ' DKK</strong></td></tr>';
                html += '</table>';

                html += '<div class="order-report-grid" style="margin-top:12px;">';
                html += '<div class="order-report-card"><div class="label">Salgsordrer</div><div class="value">' + formatNumber(salesOrderLines.length) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Ekstra salgslinjer</div><div class="value">' + formatNumber(salesLines.length) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Kost til dato</div><div class="value">' + formatNumber(costToDateFromProduction || 0) + ' DKK</div></div>';
                html += '<div class="order-report-card"><div class="label">Marginprocent</div><div class="value">' + marginPct + '%</div></div>';
                html += '</div>';

                html += '<div class="order-report-card" style="margin-top:12px;">';
                html += '<div class="label">Exceptioner og spor</div>';
                html += '<div class="value" style="font-size:14px; font-weight:600; margin-bottom:8px;">Advarsler: ' + formatCount(exceptionGroupCount) + ' grupper • ' + formatCount(exceptionOccurrenceCount) + ' forekomster</div>';
                if (exceptionCompactRows.length > 0) {
                    html += '<table class="order-report-table" style="margin-top:0;">';
                    html += '<tr><th>Prod.ordre</th><th>Type</th><th>Linjer/Ref</th><th>Beskrivelse</th><th>Antal</th></tr>';
                    for (const ex of exceptionCompactRows) {
                        const refsPreview = ex.refs.length <= 4
                            ? ex.refs.join(', ')
                            : (ex.refs.slice(0, 4).join(', ') + ' +' + (ex.refs.length - 4));
                        const refsTitle = ex.refs.join(', ');
                        html += '<tr>';
                        html += '<td>' + escapeHtml(ex.prodOrdNo || '-') + '</td>';
                        html += '<td>' + escapeHtml(ex.types.join(' + ') || '-') + '</td>';
                        html += '<td title="' + escapeHtml(refsTitle) + '">' + escapeHtml(refsPreview || '-') + '</td>';
                        html += '<td>' + escapeHtml(ex.message || '-') + '</td>';
                        html += '<td>' + formatCount(ex.count || 0) + '</td>';
                        html += '</tr>';
                    }
                    html += '</table>';
                } else {
                    html += '<div style="color:#4b5563; font-size:13px;">Ingen ekceptioner fundet i den aktuelle ordre.</div>';
                }
                html += '</div>';
                html += '</div>';
                return html;
            }

            function toggleOrderDetailReport() {
                const report = document.getElementById('orderDetailReport');
                if (!report) return;
                const isClosed = report.style.display === 'none';
                report.style.display = isClosed ? '' : 'none';
            }

            let currentPrintPreviewMode = null;
            let reportPrintRestoreState = null;

            function isOrderDetailReportViewActive() {
                const bodyEl = document.getElementById('orderDetailModalBody');
                return Boolean(bodyEl && bodyEl.querySelector('#orderDetailReport'));
            }

            function updateOrderDetailModalBackButton() {
                const btn = document.getElementById('orderDetailModalBackBtn');
                if (!btn) return;
                const show = Boolean(reportOriginState && isOrderDetailReportViewActive());
                btn.style.display = show ? '' : 'none';
            }

            function restoreOrderDetailFromReport() {
                if (!reportOriginState) return false;
                const snapshot = reportOriginState;
                reportOriginState = null;
                openOrderDetailModal(snapshot.html, snapshot.title, snapshot.subtitle);
                return true;
            }

            function goBackFromReportToOrder() {
                restoreOrderDetailFromReport();
            }

            function openOrderDetailModal(html, titleText, subtitleText) {
                const overlay = document.getElementById('orderDetailModal');
                const titleEl = document.getElementById('orderDetailModalTitle');
                const subtitleEl = document.getElementById('orderDetailModalSubtitle');
                const bodyEl = document.getElementById('orderDetailModalBody');
                if (!overlay || !titleEl || !subtitleEl || !bodyEl) return;
                titleEl.textContent = titleText || 'Ordre-rapport';
                subtitleEl.textContent = subtitleText || 'Manager-oversigt med produktion, cost og sporbarhed';
                bodyEl.innerHTML = html || '';
                applyMicroTablePolish(bodyEl);
                overlay.style.display = 'flex';
                document.body.classList.add('report-modal-open');
                updateOrderDetailModalBackButton();
            }

            function closeOrderDetailModal(event) {
                if (event && event.target && event.target.id !== 'orderDetailModal') return;
                if (isOrderDetailReportViewActive() && reportOriginState) {
                    restoreOrderDetailFromReport();
                    return;
                }
                const overlay = document.getElementById('orderDetailModal');
                const bodyEl = document.getElementById('orderDetailModalBody');
                if (overlay) overlay.style.display = 'none';
                if (bodyEl) bodyEl.innerHTML = '';
                document.body.classList.remove('report-modal-open');
                reportOriginState = null;
                updateOrderDetailModalBackButton();
                goBackToList();
            }

            function buildStandaloneReportPrintCss() {
                return [
                    '@page { size: A4 portrait; margin: 12mm; }',
                    'body { margin: 0; background: #fff; color: #10253f; font-family: Arial, sans-serif; }',
                    '.order-detail-report { display:block !important; }',
                    '.order-report-toolbar,',
                    '.order-report-actions,',
                    '.list-toggle-btn,',
                    'button { display:none !important; }',
                    '.section { break-inside: avoid; page-break-inside: avoid; margin-bottom: 14px; }',
                    '.order-report-grid { grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; }',
                    '.order-report-card { break-inside: avoid; page-break-inside: avoid; }',
                    '.order-report-table { width:100%; border-collapse:collapse; }',
                    '.order-report-table th, .order-report-table td { border-bottom:1px solid #dfe8f3; padding:8px 10px; }',
                    '.order-report-table th { background:#eef5ff; }',
                    '.summary-box { break-inside: avoid; page-break-inside: avoid; }'
                ].join('\\n');
            }

            function buildStandaloneListPrintCss() {
                return [
                    '@page { size: A4 portrait; margin: 7mm; }',
                    'body { margin:0; font-family: Arial, sans-serif; color:#1f2937; font-size:10px; line-height:1.2; }',
                    '.order-list-summary-actions, .order-report-actions, .search-box, .header-banner-wrapper { display:none !important; }',
                    '.order-list-section { margin:0 !important; box-shadow:none !important; border:none !important; padding:0 !important; }',
                    '.order-list-summary { margin:0 0 8px 0 !important; padding:7px 9px !important; font-size:10px !important; }',
                    '.order-list-section h3 { margin:0 0 7px 0 !important; padding:0 0 5px 0 !important; font-size:12px !important; }',
                    '.order-list-table { width:100%; font-size:9.3px; border-collapse:collapse; table-layout:auto; }',
                    '.order-list-table th, .order-list-table td { padding:3px 4px; line-height:1.15; }',
                    '.order-list-table th:nth-child(8), .order-list-table td:nth-child(8), .order-list-table th:nth-child(9), .order-list-table td:nth-child(9) { display:none; }',
                    '.order-list-table tr { break-inside: avoid; page-break-inside: avoid; }'
                ].join('\\n');
            }

            function printStandaloneHtml(title, html, cssText) {
                const iframe = document.createElement('iframe');
                iframe.setAttribute('aria-hidden', 'true');
                iframe.style.position = 'fixed';
                iframe.style.right = '0';
                iframe.style.bottom = '0';
                iframe.style.width = '0';
                iframe.style.height = '0';
                iframe.style.border = '0';
                iframe.style.opacity = '0';
                document.body.appendChild(iframe);

                const styles = Array.from(document.querySelectorAll('style, link[rel="stylesheet"]'))
                    .map(el => el.outerHTML)
                    .join('\\n');
                const safeTitle = escapeHtml(title || 'Udskrift');
                const doc = iframe.contentWindow.document;
                doc.open();
                doc.write('<!doctype html><html><head><meta charset="utf-8"><title>' + safeTitle + '</title>' + styles + '<style>' + (cssText || '') + '</style></head><body>' + (html || '') + '</body></html>');
                doc.close();

                setTimeout(() => {
                    try {
                        iframe.contentWindow.focus();
                        iframe.contentWindow.print();
                    } finally {
                        setTimeout(() => iframe.remove(), 1500);
                    }
                }, 250);
            }

            function printOrderDetailReport() {
                const bodyEl = document.getElementById('orderDetailModalBody');
                if (!bodyEl) return;
                const titleEl = document.getElementById('orderDetailModalTitle');
                const reportTitle = titleEl ? titleEl.textContent : 'Ordre-rapport';
                printStandaloneHtml(reportTitle, bodyEl.innerHTML, buildStandaloneReportPrintCss());
            }

            function renderPrintPreview(title, html, mode) {
                const overlay = document.getElementById('printPreviewOverlay');
                const titleEl = document.getElementById('printPreviewTitle');
                const bodyEl = document.getElementById('printPreviewBody');
                if (!overlay || !titleEl || !bodyEl) return;
                currentPrintPreviewMode = mode;
                titleEl.textContent = title;
                bodyEl.innerHTML = html;
                overlay.style.display = 'flex';
                document.body.classList.add('print-preview-lock');
                document.body.classList.add('print-preview-mode');
                const dialog = overlay.querySelector('.print-preview-dialog');
                if (dialog && dialog.focus) dialog.focus();
            }

            function closePrintPreview(event) {
                if (event && event.target && event.target.id !== 'printPreviewOverlay') return;
                const overlay = document.getElementById('printPreviewOverlay');
                if (overlay) overlay.style.display = 'none';
                const bodyEl = document.getElementById('printPreviewBody');
                if (bodyEl) bodyEl.innerHTML = '';
                document.body.classList.remove('print-preview-lock');
                document.body.classList.remove('print-preview-mode');
                currentPrintPreviewMode = null;
            }

            function confirmPrintFromPreview() {
                const bodyEl = document.getElementById('printPreviewBody');
                if (!bodyEl) return;
                const titleEl = document.getElementById('printPreviewTitle');
                const title = titleEl ? titleEl.textContent : 'Forhåndsvisning';
                const cssText = currentPrintPreviewMode === 'list'
                    ? buildStandaloneListPrintCss()
                    : buildStandaloneReportPrintCss();
                printStandaloneHtml(title, bodyEl.innerHTML, cssText);
            }

            document.addEventListener('keydown', function(event) {
                if (event.key === 'Escape' && document.body.classList.contains('print-preview-mode')) {
                    closePrintPreview();
                }
            });

            function openOrderListPrintPreview() {
                if (!orderListVisible) {
                    toggleOrderList();
                }
                const listEl = document.getElementById('orderList');
                if (!listEl) return;
                renderPrintPreview('Forhåndsvisning - ordreliste', listEl.innerHTML, 'list');
            }

            function openOrderDetailPrintPreview() {
                const report = document.getElementById('orderDetailReport');
                if (report) {
                    renderPrintPreview('Forhåndsvisning - rapport', report.outerHTML, 'report');
                    return;
                }
                if (lastOrderReportHtml) {
                    renderPrintPreview(lastOrderReportTitle || 'Forhåndsvisning - rapport', lastOrderReportHtml, 'report');
                }
            }

            function openLatestOrderReportPreview() {
                if (!lastOrderReportHtml) {
                    alert('Rapporten er ikke klar endnu. Søg efter en ordre først.');
                    return;
                }

                const overlay = document.getElementById('orderDetailModal');
                const bodyEl = document.getElementById('orderDetailModalBody');
                const titleEl = document.getElementById('orderDetailModalTitle');
                const subtitleEl = document.getElementById('orderDetailModalSubtitle');
                const modalOpen = overlay && getComputedStyle(overlay).display === 'flex';
                if (modalOpen && bodyEl && !isOrderDetailReportViewActive() && !reportOriginState) {
                    reportOriginState = {
                        html: bodyEl.innerHTML,
                        title: titleEl ? titleEl.textContent : 'Ordre-rapport',
                        subtitle: subtitleEl ? subtitleEl.textContent : 'Manager-oversigt med produktion, cost og sporbarhed'
                    };
                }

                openOrderDetailModal(
                    lastOrderReportHtml,
                    lastOrderReportTitle || 'Rapport 2.0',
                    'Manager-oversigt med produktion, cost og sporbarhed'
                );
            }

            function updateReportOpenButtonState(isReady, ordNo) {
                const btn = document.getElementById('openReportBtn');
                if (!btn) return;
                const ready = Boolean(isReady && lastOrderReportHtml);
                btn.disabled = !ready;
                if (ready) {
                    btn.title = 'Åbn seneste rapport i separat visning';
                    btn.textContent = ordNo ? ('Rapport ' + ordNo) : 'Rapport 2.0';
                } else {
                    btn.title = 'Søg efter en ordre for at aktivere rapport';
                    btn.textContent = 'Rapport 2.0';
                }
            }

            window.addEventListener('afterprint', function() {
                document.body.classList.remove('print-report-mode');
                document.body.classList.remove('print-list-mode');
                document.body.classList.remove('print-preview-mode');
                const report = document.getElementById('orderDetailReport');
                if (report && reportPrintRestoreState !== null) {
                    report.style.display = reportPrintRestoreState;
                    reportPrintRestoreState = null;
                }
                const overlay = document.getElementById('printPreviewOverlay');
                if (overlay) overlay.style.display = 'none';
                const bodyEl = document.getElementById('printPreviewBody');
                if (bodyEl) bodyEl.innerHTML = '';
                currentPrintPreviewMode = null;
            });

            function updateOrderListSummaryPanel() {
                const summaryEl = document.getElementById('orderListSummary');
                if (!summaryEl || !orderListVisible) return;
                const orders = getFilteredOrders();
                summaryEl.innerHTML = buildOrderListSummaryHtml(orders);
            }

            async function searchOrder() {
                const ordNo = document.getElementById('orderInput').value;
                if (!ordNo) {
                    alert('Indtast et ordrenummer');
                    return;
                }

                const requestId = ++activeSearchRequestId;

                // Keep customer list visible during direct order search.
                const result = document.getElementById('result');
                result.innerHTML = '<div class="loading">Indlæser...</div>';
                openOrderDetailModal(
                    '<div class="section"><h3>Ordre ' + escapeHtml(String(ordNo)) + '</h3><div class="loading">Henter ordredata...</div></div>',
                    'Ordre ' + escapeHtml(String(ordNo)) + ' - indlæser...',
                    'Forbereder produktion, cost og sporbarhed...'
                );
                
                try {
                    const data = await requestAftercalcData(ordNo);
                    if (requestId !== activeSearchRequestId) return;
                    
                    if (data.error) {
                        openOrderDetailModal(
                            '<div class="section"><div class="error">Fejl: ' + escapeHtml(String(data.error)) + '</div></div>',
                            'Ordre ' + escapeHtml(String(ordNo)) + ' - fejl',
                            'Kunne ikke hente data for ordren'
                        );
                        result.innerHTML = '<div class="error">Fejl: ' + data.error + '</div>';
                        return;
                    }

                    // NOTE: Gr4 is order type (e.g., Multiordre). Do not change Gr4 logic here.
                    currentSalesOrderGr4 = Number((data.orderHeader && data.orderHeader.Gr4) || 0);
                    currentSearchOrderData = data;
                    const orderMarginPercent = calculateOrderMarginPercent(data.summary.totalRevenue, data.summary.totalCost).toFixed(2);
                    const _invoAm = Number(data.orderHeader.InvoAm || 0);
                    const _dInvoIF = Number(data.orderHeader.DInvoIF || 0);
                    let invoiceStatusBadge, invoiceStatusSub = '';
                    if (_invoAm === 0) {
                        invoiceStatusBadge = '<span class="invoice-status-badge status-in-production">🔧 I produktion</span>';
                    } else if (_dInvoIF <= 0) {
                        invoiceStatusBadge = '<span class="invoice-status-badge status-fully-invoiced">✅ Komplet faktureret</span>';
                    } else {
                        invoiceStatusBadge = '<span class="invoice-status-badge status-partial-invoiced">⏳ Delvist faktureret</span>';
                        invoiceStatusSub = '<div style="font-size:12px; opacity:0.85; margin-top:5px;">Faktureret: ' + formatNumber(_invoAm) + ' | Mangler: ' + formatNumber(_dInvoIF) + '</div>';
                    }
                    const productionOrderByOrdNo = new Map((Array.isArray(data.productionOrders) ? data.productionOrders : []).map(order => [Number(order.ordNo || 0), order]));
                    const costToDateFromProduction = (Array.isArray(data.productionOrders) ? data.productionOrders : [])
                        .reduce((sum, order) => sum + Number((order && order.totalCost) || 0), 0);
                    const getSalesLineCostBreakdown = (purcNo) => {
                        const prodOrder = productionOrderByOrdNo.get(Number(purcNo || 0));
                        const lines = Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : [];
                        let operationTotal = 0;
                        let laserTotal = 0;
                        let materialTotal = 0;

                        for (const line of lines) {
                            const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                            const lnNo = Number((line && line.LnNo) || 0);
                            if (lnNo === 1 || key === '0' || key === '3' || key === '5') continue;
                            const effectiveCost = Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                            if (key === '1') {
                                operationTotal += effectiveCost;
                            } else if (key === '2') {
                                laserTotal += effectiveCost;
                                if (!isLaserLProdNo(line && line.ProdNo)) {
                                    materialTotal += effectiveCost;
                                }
                            }
                        }

                        return {
                            operationTotal,
                            laserTotal
                        };
                    };
                    
                    let html = '<div class="order-header">';
                    html += '<h2>Salgsordre: ' + data.orderHeader.OrdNo + ' - ' + (data.orderHeader.CustomerName || '-') + '</h2>';
                    const _noteOrdNo = Number(data.orderHeader.OrdNo);
                    const _existingNote = orderNotesCache[String(_noteOrdNo)];
                    const _noteIcons = { ok: '✅', error: '❌', check: '⚠️', credit: '🧾' };
                    const _noteIcon = _existingNote && _existingNote.isCreditNote
                        ? _noteIcons.credit
                        : (_existingNote && _existingNote.status ? (_noteIcons[_existingNote.status] || '📝') : '📝');
                    const _noteCls = _existingNote && _existingNote.isCreditNote
                        ? 'credit'
                        : (_existingNote && _existingNote.status ? _existingNote.status : 'text');
                    const _noteDisplay = _existingNote && (_existingNote.status || _existingNote.text || _existingNote.isCreditNote) ? 'flex' : 'none';
                    html += '<div id="order-note-banner-' + _noteOrdNo + '" class="order-note-banner ' + _noteCls + '" style="display:' + _noteDisplay + ';" onclick="openNotePopup(' + _noteOrdNo + ',true)">';
                    if (_existingNote && (_existingNote.status || _existingNote.text || _existingNote.isCreditNote)) {
                        html += '<span class="note-icon">' + _noteIcon + '</span><div class="note-body"><strong>' + (_existingNote.isCreditNote ? 'Kreditnota' : (_existingNote.status === 'ok' ? 'OK' : _existingNote.status === 'error' ? 'Fejl' : _existingNote.status === 'check' ? 'Tjek' : 'Note')) + '</strong>' + (_existingNote.text ? ': ' + escapeHtml(_existingNote.text) : '') + '</div><span style="font-size:11px;opacity:0.7;margin-left:auto;">✏️ Rediger</span>';
                    }
                    html += '</div>';
                    html += '<button onclick="openNotePopup(' + _noteOrdNo + ',true)" style="border:none;background:transparent;cursor:pointer;font-size:12px;color:#888;padding:0 0 8px 0;">📝 ' + (_existingNote && (_existingNote.status || _existingNote.text || _existingNote.isCreditNote) ? 'Rediger note' : 'Tilføj note') + '</button>';
                    html += '<div class="order-report-actions" style="display:flex; gap:8px; flex-wrap:wrap; margin:6px 0 12px 0;">';
                    html += '<button class="list-toggle-btn" onclick="openLatestOrderReportPreview()" title="Åbn rapporten i separat visning">Rapport 2.0</button>';
                    html += '</div>';
                    loadOrderNote(_noteOrdNo).catch(() => {});
                    html += '<div class="order-header-row">';
                    if (_invoAm === 0) {
                        // I Produktion: show cost to date + projected margin if DInvoIF available
                        html += '<div class="order-header-item"><div class="order-header-label">Kost til dato (estimat)</div><div class="order-header-value">' + formatNumber(costToDateFromProduction) + ' DKK</div></div>';
                        if (_dInvoIF > 0) {
                            const projectedMargin = _dInvoIF - costToDateFromProduction;
                            const projectedMarginPct = costToDateFromProduction > 0 ? calculateOrderMarginPercent(_dInvoIF, costToDateFromProduction).toFixed(2) : '0.00';
                            html += '<div class="order-header-item"><div class="order-header-label">Forventet salgsbeløb</div><div class="order-header-value">' + formatNumber(_dInvoIF) + ' DKK</div></div>';
                            html += '<div class="order-header-item"><div class="order-header-label">Forventet margin (prognose)</div><div class="order-header-value">' + getMarginBadge(projectedMarginPct) + '<div style="font-size:13px; opacity:0.85; margin-top:4px;">' + formatNumber(projectedMargin) + ' DKK</div></div></div>';
                        } else {
                            html += '<div class="order-header-item"><div class="order-header-label">Forventet salgsbeløb</div><div class="order-header-value" style="opacity:0.6; font-size:16px;">— (ukendt)</div></div>';
                            html += '<div class="order-header-item"><div class="order-header-label">Margin</div><div class="order-header-value"><span style="background:rgba(255,255,255,0.15); color:#fff; font-weight:bold; padding:2px 8px; border-radius:4px; font-size:14px;">— Ingen data</span></div></div>';
                        }
                    } else {
                        html += '<div class="order-header-item"><div class="order-header-label">Faktureret beløb</div><div class="order-header-value">' + formatNumber(data.summary.totalRevenue) + ' DKK</div></div>';
                        html += '<div class="order-header-item"><div class="order-header-label">Kostpris</div><div class="order-header-value">' + formatNumber(data.summary.totalCost) + ' DKK</div></div>';
                        html += '<div class="order-header-item"><div class="order-header-label">Margin (' + getMarginModeLabel() + ')</div><div class="order-header-value">' + getMarginBadge(orderMarginPercent) + '</div></div>';
                    }
                    html += '<div class="order-header-item"><div class="order-header-label">Fakturastatus</div><div class="order-header-value">' + invoiceStatusBadge + invoiceStatusSub + '</div></div>';
                    html += '</div></div>';

                    html += '<div class="section oversigt-launcher-section">';
                    html += '<h3>Produktionsoversigter</h3>';
                    html += '<div class="oversigt-launcher-grid">';
                    html += '<article class="oversigt-launcher-card">';
                    html += '<h4>Laseroversigt (L-linjer)</h4>';
                    html += '<div class="desc">Nesting, vægtafvigelser og kost på tværs af ruter.</div>';
                    html += '<div id="laserOversigtSummaryTeaser" class="oversigt-launcher-kpi"><span class="loading">Indlæser laser-KPI...</span></div>';
                    html += '<button class="list-toggle-btn" onclick="openOversigtModal(\\\'laser\\\')" title="Åbn detaljeret laseroversigt">Åbn laseroversigt</button>';
                    html += '</article>';
                    html += '<article class="oversigt-launcher-card">';
                    html += '<h4>Operation Oversigt</h4>';
                    html += '<div class="desc">Operationstid, afvigelser og omkostninger i én driftsvisning.</div>';
                    html += '<div id="operationOversigtSummaryTeaser" class="oversigt-launcher-kpi"><span class="loading">Indlæser operations-KPI...</span></div>';
                    html += '<button class="list-toggle-btn" onclick="openOversigtModal(\\\'operation\\\')" title="Åbn detaljeret operationsoversigt">Åbn operationer</button>';
                    html += '</article>';
                    html += '</div>';
                    html += '</div>';

                    html += '<div id="laserOrderSummaryPanel" style="display:none;" aria-hidden="true">';
                    html += '<div id="laserOrderSummaryTotals" class="summary-box"><div class="loading">Indlæser totaler...</div></div>';
                    html += '<div class="laser-summary-layout">';
                    html += '<div id="laserOrderSummaryBody" class="loading">Indlæser laserdata...</div>';
                    html += '<aside id="laserImagePanel" class="laser-image-panel hidden"></aside>';
                    html += '</div>';
                    html += '</div>';

                    html += '<div id="operationOrderSummaryPanel" style="display:none;" aria-hidden="true">';
                    html += '<div id="operationOrderSummaryTotals" class="summary-box"><div class="loading">Indlæser totaler...</div></div>';
                    html += '<div id="operationOrderSummaryBody" class="loading">Indlæser operationsdata...</div>';
                    html += '</div>';

                    lastOrderReportHtml = buildOrderDetailReportHtml(data, orderMarginPercent, costToDateFromProduction);
                    lastOrderReportTitle = 'Rapport 2.0 - ordre ' + String(data.orderHeader.OrdNo || '-');
                    updateReportOpenButtonState(true, String(data.orderHeader.OrdNo || ''));

                    // Sezione linee ORDINE DI VENDITA complete
                    if (data.salesOrderLines && data.salesOrderLines.length > 0) {
                        const hasSalesOrderDrawing = data.salesOrderLines.some(line => !!line.DrawingWebPg);
                        const salesOrderColspan = hasSalesOrderDrawing ? 11 : 10;
                        html += '<div class="section"><h3>Salgsordrelinjer</h3>';
                        html += '<table><tr><th>Linje</th><th>Produkt</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Kostpris</th><th>Samlet kost</th><th>Salgspris/enhed</th><th>Salgspris</th><th>Margin (%)</th><th>Prod.ordre</th>' + (hasSalesOrderDrawing ? '<th>Vis tegning</th>' : '') + '</tr>';

                        for (const line of data.salesOrderLines) {
                            const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
                            const lineCost = line.EffectiveLineCost || 0;
                            const lineProdNo = String(line.ProdNo || '').trim();
                            const includeForMargin = lineProdNo.startsWith('1') || lineProdNo.startsWith('3');
                            const lineMarginValue = calculateLineMarginPercent(lineSalesPrice, lineCost);
                            const isExactlyHundred = Math.abs(lineMarginValue - 100) < 0.0001;
                            const lineMarginPercent = lineMarginValue.toFixed(2);
                            const hasProductionOrder = Boolean(line.PurcNo && line.PurcNo !== 0);
                            const breakdownRowId = 'sales-line-breakdown-' + String(data.orderHeader.OrdNo || '0') + '-' + String(line.LnNo || 0);
                            const breakdownInfo = hasProductionOrder
                                ? getSalesLineCostBreakdown(line.PurcNo)
                                : { operationTotal: 0, laserTotal: 0 };
                            // Rabat-badge fjernet: linjer med salgspris=0 (underlinjer af hovedprodukt) vises som N/A.
                            const lineMarginBadge = (!includeForMargin || lineSalesPrice === 0 || isExactlyHundred)
                                ? '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>'
                                : getMarginBadge(lineMarginPercent);
                            html += '<tr>';
                            html += '<td>' + (hasProductionOrder
                                ? ('<button type="button" onclick="toggleSalesLineBreakdown(\\'' + breakdownRowId + '\\', this)" title="Vis kost-opdeling" style="margin-right:6px; width:22px; height:22px; border:1px solid #90caf9; background:#e3f2fd; color:#0d47a1; border-radius:4px; cursor:pointer; font-weight:700;">+</button>')
                                : '') + (line.LnNo || 0) + '</td>';

                            const salesWarningFlag = getWarningFlagHtml(line, 'Tilknyttet produktionsordre har en advarsel.');
                            if (line.PurcNo && line.PurcNo !== 0) {
                                html += '<td><span class="prod-link" onclick="openProduction(' + line.PurcNo + ')">' + (line.ProdNo || '-') + '</span>' + salesWarningFlag + '</td>';
                            } else {
                                html += '<td>' + (line.ProdNo || '-') + salesWarningFlag + '</td>';
                            }

                            const displaySalesQty = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                                ? line.DisplayQuantity
                                : (line.NoFin || 0);
                            html += '<td>' + (line.Descr || '') + '</td>';
                            html += '<td>' + formatNumber(displaySalesQty) + '</td>';
                            const productionTotalCost = Number(line.ProductionOrderTotalCost || 0);
                            const lineQty = Number(line.NoFin || 0);
                            const displayKostpris = (line.PurcNo && line.PurcNo !== 0)
                                ? (lineQty > 0 ? (productionTotalCost / lineQty) : productionTotalCost)
                                : (line.CCstPr || 0);
                            html += '<td>' + formatNumber(displayKostpris) + '</td>';
                            html += '<td><strong>' + formatNumber(lineCost) + '</strong></td>';
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                            html += '<td>' + formatNumber(lineSalesPrice) + '</td>';
                            html += '<td>' + lineMarginBadge + '</td>';
                            html += '<td>' + ((line.PurcNo && line.PurcNo !== 0) ? line.PurcNo : '-') + '</td>';
                            if (hasSalesOrderDrawing) {
                                if (line.DrawingWebPg) {
                                    html += '<td><button class="list-toggle-btn drawing-open-btn" data-drawing-path="' + escapeHtml(String(line.DrawingWebPg || '')) + '" data-prod-no="' + escapeHtml(String(line.ProdNo || '')) + '" data-ord-no="' + escapeHtml(String(line.PurcNo || data.orderHeader.OrdNo || '')) + '" style="padding:4px 8px; margin-left:0;">Vis tegning</button></td>';
                                } else {
                                    html += '<td></td>';
                                }
                            }
                            html += '</tr>';
                            if (hasProductionOrder) {
                                html += '<tr id="' + breakdownRowId + '" style="display:none; background:#f8fbff;">';
                                html += '<td colspan="' + salesOrderColspan + '" style="padding:10px 16px; border-top:none;">';
                                html += '<div style="display:grid; gap:6px; color:#1f2937;">';
                                html += '<div><strong>Operation:</strong> ' + formatNumber(breakdownInfo.operationTotal || 0) + ' DKK</div>';
                                html += '<div><strong>Laser / materiale:</strong> ' + formatNumber(breakdownInfo.laserAndMaterialTotal || 0) + ' DKK</div>';
                                if ((breakdownInfo.materialTotal || 0) > 0) {
                                    html += '<div><strong>Materiale (ikke L):</strong> ' + formatNumber(breakdownInfo.materialTotal || 0) + ' DKK</div>';
                                }
                                html += '</div>';
                                html += '</td>';
                                html += '</tr>';
                            }
                        }

                        html += '</table></div>';
                    }
                    
                    // Sezione linee di vendita
                    if (data.salesLines.length > 0) {
                        const hasSalesLinesDrawing = data.salesLines.some(line => !!line.DrawingWebPg);
                        html += '<div class="section"><h3>Salgslinjer (Ekstra produkter)</h3>';
                        html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Salgspris</th><th>Kostpris/enhed</th><th>Samlet kost</th>' + (hasSalesLinesDrawing ? '<th>Vis tegning</th>' : '') + '</tr>';
                        
                        for (const line of data.salesLines) {
                            const salesExtraWarningFlag = getWarningFlagHtml(line, 'Inkonsekvens på salgslinje.');
                            const displaySalesExtraQty = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                                ? line.DisplayQuantity
                                : (line.NoFin || 0);
                            html += '<tr>';
                            html += '<td>' + (line.ProdNo || '-') + salesExtraWarningFlag + '</td>';
                            html += '<td>' + (line.Descr || '') + '</td>';
                            html += '<td>' + formatNumber(displaySalesExtraQty) + '</td>';
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                            html += '<td>' + formatNumber(line.CCstPr || 0) + '</td>';
                            html += '<td><strong>' + formatNumber(line.EffectiveLineCost || 0) + '</strong></td>';
                            if (hasSalesLinesDrawing) {
                                if (line.DrawingWebPg) {
                                    html += '<td><button class="list-toggle-btn drawing-open-btn" data-drawing-path="' + escapeHtml(String(line.DrawingWebPg || '')) + '" data-prod-no="' + escapeHtml(String(line.ProdNo || '')) + '" data-ord-no="' + escapeHtml(String(line.PurcNo || data.orderHeader.OrdNo || '')) + '" style="padding:4px 8px; margin-left:0;">Vis tegning</button></td>';
                                } else {
                                    html += '<td></td>';
                                }
                            }
                            html += '</tr>';
                        }
                        
                        html += '<tr class="summary-row"><td colspan="5">Total salgslinjer:</td><td>' + formatNumber(data.salesLinesTotalCost) + ' DKK</td>' + (hasSalesLinesDrawing ? '<td></td>' : '') + '</tr>';
                        html += '</table></div>';
                    }
                    
                    // Sezione ordini di produzione
                    if (data.productionOrders.length > 0) {
                        html += '<div class="section"><h3>Produktionsordrer</h3>';
                        const prodTp4Labels = {
                            '1': 'Operation',
                            '2': 'Materiale Laser',
                            '4': 'Produkt dele',
                            '5': 'Rute',
                            '6': 'Ydelse',
                            '7': 'Underleverandor',
                            '8': 'Materiale fast antal',
                            '9': 'Indkøbt dele',
                            'NA': 'Ikke sat'
                        };
                        
                        for (const prodOrder of data.productionOrders) {
                            const mainProductLine = prodOrder.lines.find(line => line.ProdTp4 === 0) || prodOrder.lines.find(line => line.LnNo === 1);
                            const mainProductText = mainProductLine
                                ? ((mainProductLine.ProdNo || '-') + ' - ' + (mainProductLine.Descr || ''))
                                : '-';

                            html += '<div id="po-' + prodOrder.ordNo + '" data-order="' + prodOrder.ordNo + '" style="margin-bottom: 20px; border: 1px solid #ddd; padding: 15px; border-radius: 4px;">';
                            const prodOrderTimeFlagHtml = getTimeAdjustmentFlagHtml({
                                hasEstimatedOperationTime: !!prodOrder.hasEstimatedOperationTime,
                                EstimatedTimeText: 'Mindst én operation er genberegnet ud fra Stykliste Minutter, fordi Færdigmeldt var 0.'
                            });
                            html += '<h4>Produktionsordre: ' + prodOrder.ordNo + prodOrderTimeFlagHtml + getWarningFlagHtml({ HasWarning: !!prodOrder.hasWarnings, WarningText: prodOrder.warningText || '' }, 'Denne produktionsordre indeholder mindst en advarselslinje.') + '</h4>';
                            html += '<div class="main-product-box">';
                            html += '<div class="value">' + mainProductText + '</div>';
                            html += '</div>';
                            html += '<div class="prodtp4-hint">Klik paa en linje for at aabne/lukke detaljer.</div>';

                            const groupedLines = {};
                            const operationMergeMap = new Map();
                            const pendingNoOrgFromTp3 = new Map();
                            for (const line of prodOrder.lines) {
                                const rawKey = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                                if (rawKey === '0' || rawKey === '5') continue;

                                // Merge operation rows where ProdTp4 is 1 or 3 and ProdNo is the same.
                                const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                                const normalizedKey = getDisplayProdTp4Key(rawKey, prodNoKey, line.PurcNo);


                                // R1090/R8200 must be fully excluded from Operations: no row and no cost contribution.
                                if (normalizedKey === '1' && isExcludedOperationProdNo(prodNoKey)) {
                                    continue;
                                }

                                // R-products under Produkt dele must never be shown or counted.
                                if (normalizedKey === '4' && prodNoKey.startsWith('R')) {
                                    continue;
                                }

                                if (normalizedKey === '1') {
                                    if (prodNoKey) {
                                        const mergeKey = normalizedKey + '|' + prodNoKey;
                                        if (rawKey === '3') {
                                            const extraNoOrg = Number(line.NoOrg || 0);
                                            if (operationMergeMap.has(mergeKey)) {
                                                const mergedLine = operationMergeMap.get(mergeKey);
                                                mergedLine.NoOrg = Number(mergedLine.NoOrg || 0) + extraNoOrg;
                                            } else {
                                                pendingNoOrgFromTp3.set(mergeKey, Number(pendingNoOrgFromTp3.get(mergeKey) || 0) + extraNoOrg);
                                            }
                                            continue;
                                        }

                                        if (!operationMergeMap.has(mergeKey)) {
                                            const extraNoOrg = Number(pendingNoOrgFromTp3.get(mergeKey) || 0);
                                            const mergedLine = {
                                                ...line,
                                                ProdTp4: 1,
                                                NoOrg: Number(line.NoOrg || 0) + extraNoOrg,
                                                NoFin: Number(line.NoFin || 0),
                                                LineCost: Number(line.LineCost || 0),
                                                EffectiveLineCost: Number(line.EffectiveLineCost || 0)
                                            };
                                            operationMergeMap.set(mergeKey, mergedLine);
                                            if (!groupedLines[normalizedKey]) groupedLines[normalizedKey] = [];
                                            groupedLines[normalizedKey].push(mergedLine);
                                        } else {
                                            const mergedLine = operationMergeMap.get(mergeKey);
                                            mergedLine.NoOrg = Number(mergedLine.NoOrg || 0) + Number(line.NoOrg || 0);
                                            mergedLine.NoFin = Number(mergedLine.NoFin || 0) + Number(line.NoFin || 0);
                                            mergedLine.LineCost = Number(mergedLine.LineCost || 0) + Number(line.LineCost || 0);
                                            mergedLine.EffectiveLineCost = Number(mergedLine.EffectiveLineCost || 0) + Number(line.EffectiveLineCost || 0);
                                            if ((!mergedLine.Descr || mergedLine.Descr === '-') && line.Descr) {
                                                mergedLine.Descr = line.Descr;
                                            }
                                        }
                                        continue;
                                    }
                                }

                                if (!groupedLines[normalizedKey]) groupedLines[normalizedKey] = [];
                                groupedLines[normalizedKey].push({ ...line, ProdTp4: normalizedKey === '1' ? 1 : line.ProdTp4 });
                            }

                            const groupKeys = Object.keys(groupedLines).sort((a, b) => {
                                if (a === 'NA') return 1;
                                if (b === 'NA') return -1;
                                return Number(a) - Number(b);
                            });

                            let orderVisibleTotal = 0;

                            for (let i = 0; i < groupKeys.length; i++) {
                                const key = groupKeys[i];
                                const lines = groupedLines[key];
                                const subtotal = key === '2'
                                    ? lines.filter(line => line.LnNo !== 1).reduce((sum, line) => {
                                        if (!isLaserLProdNo(line.ProdNo)) {
                                            return sum + (line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null ? (line.EffectiveLineCost || 0) : (line.LineCost || 0));
                                        }
                                        if (line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null) {
                                            return sum + (line.EffectiveLineCost || 0);
                                        }
                                        const hasNestingCost = Number(line.NestingCost || 0) > 0;
                                        return sum + (hasNestingCost
                                            ? ((line.NestingCost || 0) * (line.NoFin || 0))
                                            : (line.LineCost || 0));
                                    }, 0)
                                    : lines.filter(line => line.LnNo !== 1).reduce((sum, line) => {
                                        const pn = String(line.ProdNo || '').toUpperCase();
                                        if (pn === 'R6200' && String(key) === '1') {
                                            return sum + ((line.NoOrg || 0) * (line.CCstPr || 0));
                                        }
                                        return sum + (line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null ? (line.EffectiveLineCost || 0) : (line.LineCost || 0));
                                    }, 0);
                                const isOpenByDefault = false;
                                orderVisibleTotal += subtotal;
                                const groupWarningFlagHtml = getWarningFlagHtml(lines, 'Denne gruppe indeholder mindst en advarselslinje.');

                                html += '<div class="prodtp4-group">';
                                html += '<div class="prodtp4-header" onclick="toggleProdTp4Group(' + prodOrder.ordNo + ', &quot;' + key + '&quot;)">';
                                html += '<span class="prodtp4-label"><span id="po-' + prodOrder.ordNo + '-group-' + key + '-icon">' + (isOpenByDefault ? '▾' : '▸') + '</span> ' + key + ' - ' + (prodTp4Labels[key] || 'Altro') + groupWarningFlagHtml + '</span>';
                                html += '<span class="prodtp4-subtotal">Delsum: ' + formatNumber(subtotal) + ' DKK</span>';
                                html += '</div>';

                                html += '<div id="po-' + prodOrder.ordNo + '-group-' + key + '" class="prodtp4-body" style="display:' + (isOpenByDefault ? '' : 'none') + ';">';
                                if (key === '2') {
                                    const laserCostHeader = currentSalesOrderGr4 === 3 ? 'NestMultiPris' : 'Kostpris nesting';
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>' + laserCostHeader + '</th><th>Samlet kost</th></tr>';
                                } else if (key === '1') {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Stykliste Minutter</th><th>Færdigmeldt minutter</th><th>Kostpris/enhed</th><th>Samlet kost</th></tr>';
                                } else if (key === '6' || key === '9') {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Pris/enhed</th><th>Samlet kost</th></tr>';
                                } else {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Kostpris/enhed</th><th>Samlet kost</th></tr>';
                                }

                                for (const line of lines) {
                                    html += '<tr>';
                                    const warningFlagHtml = getWarningFlagHtml(line);
                                    const invoiceStatusFlagHtml = getInvoiceStatusFlagHtml(line);
                                    const timeAdjustFlagHtml = getTimeAdjustmentFlagHtml(line);
                                    const laserAllocationFlagHtml = getLaserAllocationFlagHtml(line);
                                    const hasChildProductionOrder = Number(line.PurcNo || 0) > 0;
                                    if (String(key) === '1' && line.ProdNo) {
                                        const safeProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeProdLabel = escapeHtml(getResourceDisplayLabel(line.ProdNo, line.Descr));
                                        const trInf2Value = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : prodOrder.ordNo);
                                        const safeTrInf2 = trInf2Value.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeTrInf4 = String(line.TrInf4 || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        html += '<td><span class="prod-no-link" data-prodno="' + safeProdNo + '" data-ordno="' + prodOrder.ordNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="' + key + '" data-trinf2="' + safeTrInf2 + '" data-trinf4="' + safeTrInf4 + '">' + safeProdLabel + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (String(key) === '2' && line.ProdNo && isLaserLProdNo(line.ProdNo)) {
                                        const safeProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeProdLabel = escapeHtml(getResourceDisplayLabel(line.ProdNo, line.Descr));
                                        const trInf2Value = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : prodOrder.ordNo);
                                        const safeTrInf2 = trInf2Value.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeTrInf4 = String(line.TrInf4 || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const linkNoFin = Number(line.NoFin || 0);
                                        const linkHasNestingCost = Number(line.NestingCost || 0) > 0;
                                        const linkHasEffectiveLaserCost = linkNoFin > 0
                                            && line.EffectiveLineCost !== undefined
                                            && line.EffectiveLineCost !== null;
                                        const linkDisplayLaserUnitCost = linkHasEffectiveLaserCost
                                            ? ((line.EffectiveLineCost || 0) / linkNoFin)
                                            : (linkHasNestingCost ? (line.NestingCost || 0) : (line.CCstPr || 0));
                                        html += '<td><span class="prod-no-link" data-prodno="' + safeProdNo + '" data-ordno="' + prodOrder.ordNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="' + key + '" data-trinf2="' + safeTrInf2 + '" data-trinf4="' + safeTrInf4 + '" data-showallroutes="1" data-nofin="' + linkNoFin + '" data-nestingcost="' + Number(linkDisplayLaserUnitCost || 0) + '">' + safeProdLabel + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (hasChildProductionOrder) {
                                        const safeChildProdNoForSummary = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const childSummaryArgs = shouldFilterChildSummary(key, line.ProdNo, line.PurcNo)
                                            ? (Number(line.PurcNo || 0) + ', &quot;' + safeChildProdNoForSummary + '&quot;, true')
                                            : Number(line.PurcNo || 0);
                                        html += '<td><span class="inline-link" onclick="showChildProductionSummary(' + childSummaryArgs + ')">' + (line.ProdNo || '-') + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (line.ProdNo) {
                                        html += '<td>' + (line.ProdNo || '-') + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else {
                                        html += '<td>-' + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    }
                                    html += '<td>' + (line.Descr || '') + '</td>';
                                    if (key === '1') {
                                        const effectiveNoFin = (line.EffectiveOperationMinutes !== undefined && line.EffectiveOperationMinutes !== null)
                                            ? (line.EffectiveOperationMinutes || 0)
                                            : (line.UsesEstimatedOperationTime ? (line.NoOrg || 0) : (line.NoFin || 0));
                                        html += '<td>' + formatNumber(line.NoOrg || 0) + '</td>';
                                        html += '<td>' + formatNumber(effectiveNoFin) + '</td>';
                                        const displayUnitCost1 = (line.CCstPr || 0);
                                        const displayTotalCost1 = (line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null)
                                            ? (line.EffectiveLineCost || 0)
                                            : (effectiveNoFin * (line.CCstPr || 0));
                                        html += '<td>' + formatNumber(displayUnitCost1) + '</td>';
                                        html += '<td><strong>' + formatNumber(displayTotalCost1) + '</strong></td>';
                                    } else {
                                        const displayQty = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                                            ? line.DisplayQuantity
                                            : (line.NoFin || 0);
                                        html += '<td>' + formatNumber(displayQty) + '</td>';
                                    }
                                    if (key === '2') {
                                        const isLaserLine = isLaserLProdNo(line.ProdNo);
                                        const hasNestingCost = Number(line.NestingCost || 0) > 0;
                                        const hasEffectiveLaserCost = Number(line.NoFin || 0) > 0
                                            && line.EffectiveLineCost !== undefined
                                            && line.EffectiveLineCost !== null;
                                        const nestingUnitCost = isLaserLine
                                            ? (hasEffectiveLaserCost
                                                ? ((line.EffectiveLineCost || 0) / (line.NoFin || 0))
                                                : (hasNestingCost ? (line.NestingCost || 0) : (line.CCstPr || 0)))
                                            : (line.CCstPr || 0);
                                        const nestingSamlet = isLaserLine
                                            ? (hasEffectiveLaserCost
                                                ? (line.EffectiveLineCost || 0)
                                                : (hasNestingCost
                                                    ? ((line.NestingCost || 0) * (line.NoFin || 0))
                                                    : (line.LineCost || 0)))
                                            : (line.LineCost || 0);
                                        html += '<td>' + formatNumber(nestingUnitCost) + '</td>';
                                        html += '<td><strong>' + formatNumber(nestingSamlet) + '</strong></td>';
                                    } else if (key !== '1') {
                                        const displayQtyNonOperation = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                                            ? Number(line.DisplayQuantity || 0)
                                            : Number(line.NoFin || 0);
                                        const displayUnitCost = (displayQtyNonOperation > 0 && line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null)
                                            ? ((line.EffectiveLineCost || 0) / displayQtyNonOperation)
                                            : ((line.DisplayUnitCost !== undefined && line.DisplayUnitCost !== null)
                                                ? line.DisplayUnitCost
                                                : (line.CCstPr || line.DPrice || 0));
                                        const displayTotalCost = line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null
                                            ? (line.EffectiveLineCost || 0)
                                            : (line.LineCost || 0);
                                        html += '<td>' + formatNumber(displayUnitCost) + '</td>';
                                        html += '<td><strong>' + formatNumber(displayTotalCost) + '</strong></td>';
                                    }
                                    html += '</tr>';
                                }

                                html += '</table>';
                                html += '</div>';
                                html += '</div>';
                            }
                            
                            html += '<div class="po-total-row">Total ordre: <span id="po-total-' + prodOrder.ordNo + '">' + formatNumber(orderVisibleTotal) + ' DKK</span></div>';
                            html += '</div>';
                        }
                        
                        html += '</div>';
                    }
                    
                    reportOriginState = null;
                    openOrderDetailModal(
                        html,
                        'Ordre ' + data.orderHeader.OrdNo + ' - ' + (data.orderHeader.CustomerName || '-'),
                        'Produktion, cost og sporbarhed i en separat rapportvisning'
                    );
                    result.innerHTML = '';
                    loadSalesOrderLaserSummary(data);
                    loadSalesOrderOperationSummary(data);
                } catch (err) {
                    if (requestId !== activeSearchRequestId) return;
                    openOrderDetailModal(
                        '<div class="section"><div class="error">Fejl: ' + escapeHtml(String(err.message || err)) + '</div></div>',
                        'Ordre ' + escapeHtml(String(ordNo)) + ' - fejl',
                        'Der opstod en fejl under indlæsning'
                    );
                    result.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                }
            }

            function loadSalesOrderOperationSummary(orderData) {
                const body = document.getElementById('operationOrderSummaryBody');
                const totals = document.getElementById('operationOrderSummaryTotals');
                const teaser = document.getElementById('operationOversigtSummaryTeaser');
                if (!body || !totals) return;

                try {
                    const productionOrders = Array.isArray(orderData && orderData.productionOrders) ? orderData.productionOrders : [];
                    const groupedRows = new Map();
                    let totalOperationCost = 0;
                    let totalStyklisteMinutes = 0;
                    let totalFinishedMinutes = 0;

                    for (const prodOrder of productionOrders) {
                        const lines = Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : [];

                        for (const line of lines) {
                            const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                            const lnNo = Number((line && line.LnNo) || 0);
                            if (lnNo === 1 || key !== '1') continue;

                            const prodNo = String((line && line.ProdNo) || '').trim();
                            if (!prodNo) continue;
                            const totalCost = Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                            const qty = Number((line && (line.DisplayQuantity ?? line.NoFin)) || 0);
                            const styklisteMinutes = Number((line && line.NoOrg) || 0);
                            const effectiveMinutes = Number((line && (line.EffectiveOperationMinutes ?? line.NoFin)) || 0);

                            totalOperationCost += totalCost;
                            totalStyklisteMinutes += styklisteMinutes;
                            totalFinishedMinutes += effectiveMinutes;
                            if (!groupedRows.has(prodNo)) {
                                groupedRows.set(prodNo, {
                                    prodNo,
                                    descr: String((line && line.Descr) || '').trim(),
                                    styklisteQty: 0,
                                    qty: 0,
                                    totalCost: 0,
                                    occurrences: 0
                                });
                            }

                            const group = groupedRows.get(prodNo);
                            group.styklisteQty += Number((line && line.NoOrg) || 0);
                            group.qty += qty;
                            group.totalCost += totalCost;
                            group.occurrences += 1;
                            if ((!group.descr || group.descr === '-') && line && line.Descr) {
                                group.descr = String(line.Descr).trim();
                            }
                        }
                    }

                    const rows = Array.from(groupedRows.values()).sort((a, b) => a.prodNo.localeCompare(b.prodNo));

                    if (rows.length === 0) {
                        body.innerHTML = '<div>Ingen operationer fundet for denne salgsordre.</div>';
                        totals.innerHTML = '<div><strong>Samlet Operation kost:</strong> 0,00 DKK</div><div><strong>Ordre stykliste minutter:</strong> 0,00</div><div><strong>Ordre færdigmeldt minutter:</strong> 0,00</div><div><strong>Afvigelse minutter:</strong> 0,00</div><div><strong>Samlet afvigelse %:</strong> NULL</div>';
                        if (teaser) teaser.innerHTML = '<strong>Ingen operationer fundet</strong>';
                        if (currentOversigtModalType === 'operation') buildOversigtModalView('operation');
                        return;
                    }

                    let html = '<table class="oversigt-table-operation">';
                    html += '<tr><th>Operation</th><th>Beskrivelse</th><th>Linjer</th><th>Stykliste</th><th>Færdig</th><th>Kost/enh.</th><th>Samlet</th></tr>';
                    for (const row of rows) {
                        const unitCost = row.qty > 0 ? (row.totalCost / row.qty) : 0;
                        html += '<tr>';
                        html += '<td>' + (row.prodNo || '-') + '</td>';
                        html += '<td>' + (row.descr || '-') + '</td>';
                        html += '<td>' + formatNumber(row.occurrences || 0) + '</td>';
                        html += '<td>' + formatNumber(row.styklisteQty || 0) + '</td>';
                        html += '<td>' + formatNumber(row.qty || 0) + '</td>';
                        html += '<td>' + formatNumber(unitCost || 0) + '</td>';
                        html += '<td><strong>' + formatNumber(row.totalCost || 0) + '</strong></td>';
                        html += '</tr>';
                    }
                    html += '</table>';
                    body.innerHTML = html;
                    const deltaMinutes = totalFinishedMinutes - totalStyklisteMinutes;
                    const deltaPct = totalStyklisteMinutes > 0
                        ? ((deltaMinutes / totalStyklisteMinutes) * 100)
                        : null;
                    totals.innerHTML = ''
                        + '<div><strong>Samlet Operation kost:</strong> ' + formatNumber(totalOperationCost) + ' DKK</div>'
                        + '<div><strong>Ordre stykliste minutter:</strong> ' + formatNumber(totalStyklisteMinutes) + '</div>'
                        + '<div><strong>Ordre færdigmeldt minutter:</strong> ' + formatNumber(totalFinishedMinutes) + '</div>'
                        + '<div><strong>Afvigelse minutter:</strong> ' + formatNumber(deltaMinutes) + '</div>'
                        + '<div><strong>Samlet afvigelse %:</strong> ' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + '</div>';
                    if (teaser) {
                        teaser.innerHTML = ''
                            + '<div><strong>' + formatNumber(totalOperationCost) + ' DKK</strong> samlet operation kost</div>'
                            + '<div>Afvigelse: <strong>' + formatNumber(deltaMinutes) + ' min</strong> (' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + ')</div>';
                    }
                    if (currentOversigtModalType === 'operation') buildOversigtModalView('operation');
                } catch (err) {
                    body.innerHTML = '<div class="error">Fejl operationsoversigt: ' + err.message + '</div>';
                    totals.innerHTML = '<div class="error">Fejl i samlet operationsoversigt: ' + err.message + '</div>';
                    if (teaser) teaser.innerHTML = '<strong>Fejl i operation KPI:</strong> ' + err.message;
                }
            }

            async function loadSalesOrderLaserSummary(orderData) {
                const body = document.getElementById('laserOrderSummaryBody');
                const totals = document.getElementById('laserOrderSummaryTotals');
                const teaser = document.getElementById('laserOversigtSummaryTeaser');
                if (!body || !totals) return;
                // NOTE: Preserve existing Gr4 branching exactly; Gr4=3 indicates Multiordre type.
                const orderGr4 = Number((orderData && orderData.orderHeader && orderData.orderHeader.Gr4) || currentSalesOrderGr4 || 0);

                try {
                    const requests = [];
                    const targetDedupe = new Set();
                    const visitedProdOrders = new Set();
                    const productionOrderLinesByOrdNo = new Map();
                    const productionOrders = Array.isArray(orderData.productionOrders) ? orderData.productionOrders : [];
                    const laserTargets = [];

                    function getOperationCostFromLines(lines) {
                        let total = 0;
                        for (const line of (Array.isArray(lines) ? lines : [])) {
                            const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                            const lnNo = Number((line && line.LnNo) || 0);
                            if (lnNo === 1 || key === '0' || key === '3' || key === '5') continue;
                            if (key !== '1') continue;
                            total += Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                        }
                        return total;
                    }

                    function addLaserTarget(targetOrdNo, prodNo, nestingCost) {
                        const cleanedOrdNo = Number(targetOrdNo || 0);
                        const cleanedProdNo = String(prodNo || '').trim();
                        const cleanedNestingCost = Number(nestingCost || 0);
                        if (!cleanedOrdNo || !cleanedProdNo) return;

                        if (cleanedNestingCost > 0) {
                            setLaserNestCostHint(cleanedOrdNo, cleanedProdNo, cleanedNestingCost);
                        }

                        const key = cleanedOrdNo + '|' + cleanedProdNo;
                        if (targetDedupe.has(key)) return;
                        targetDedupe.add(key);
                        laserTargets.push({ ordNo: cleanedOrdNo, prodNo: cleanedProdNo, nestingCost: cleanedNestingCost > 0 ? cleanedNestingCost : null });
                    }

                    function collectLaserTargetsFromLines(sourceOrdNo, lines) {
                        const childOrdNos = [];
                        for (const line of (Array.isArray(lines) ? lines : [])) {
                            const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                            const prodNo = String(line.ProdNo || '').trim();
                            if (key === '2' && isLaserLProdNo(prodNo)) {
                                addLaserTarget(sourceOrdNo, prodNo, line.NestingCost);
                            }
                            if (key === '4' && Number(line.PurcNo || 0) > 0) {
                                childOrdNos.push(Number(line.PurcNo || 0));
                            }
                        }
                        return childOrdNos;
                    }

                    async function fetchProductionSummarySafe(childOrdNo) {
                        try {
                            const response = await fetch('/production-summary/' + childOrdNo + (orderGr4 === 3 ? '?gr4=3' : ''));
                            const data = await response.json();
                            if (!response.ok || !data || data.error) return null;
                            return data;
                        } catch (_) {
                            return null;
                        }
                    }

                    const pendingChildOrdNos = [];

                    for (const prodOrder of productionOrders) {
                        const currentOrdNo = Number(prodOrder && prodOrder.ordNo || 0);
                        if (!currentOrdNo || visitedProdOrders.has(currentOrdNo)) continue;
                        visitedProdOrders.add(currentOrdNo);
                        productionOrderLinesByOrdNo.set(currentOrdNo, Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : []);
                        const discoveredChildOrdNos = collectLaserTargetsFromLines(currentOrdNo, prodOrder.lines);
                        for (const childOrdNo of discoveredChildOrdNos) {
                            if (!visitedProdOrders.has(childOrdNo)) {
                                pendingChildOrdNos.push(childOrdNo);
                            }
                        }
                    }

                    while (pendingChildOrdNos.length > 0) {
                        const childOrdNo = Number(pendingChildOrdNos.shift() || 0);
                        if (!childOrdNo || visitedProdOrders.has(childOrdNo)) continue;
                        visitedProdOrders.add(childOrdNo);

                        const childSummary = await fetchProductionSummarySafe(childOrdNo);
                        if (!childSummary) continue;

                        productionOrderLinesByOrdNo.set(childOrdNo, Array.isArray(childSummary.lines) ? childSummary.lines : []);

                        const discoveredChildOrdNos = collectLaserTargetsFromLines(childOrdNo, childSummary.lines);
                        for (const nestedOrdNo of discoveredChildOrdNos) {
                            if (!visitedProdOrders.has(nestedOrdNo)) {
                                pendingChildOrdNos.push(nestedOrdNo);
                            }
                        }
                    }

                    for (const target of laserTargets) {
                        const endpoint = '/laser-route-metrics?ordine=' + encodeURIComponent(target.ordNo)
                            + '&prodNo=' + encodeURIComponent(target.prodNo)
                            + '&showAllRoutes=1'
                            + (orderGr4 === 3 ? '&gr4=3' : '');

                        requests.push(
                            fetch(endpoint)
                                .then(r => r.json().then(data => ({ ok: r.ok, data })))
                                .then(({ ok, data }) => ({ ok, data, prodOrderNo: target.ordNo, requestedProdNo: target.prodNo, requestedRoute: null, requestedNestingCost: target.nestingCost }))
                                .catch(() => null)
                        );
                    }

                    if (requests.length === 0) {
                        body.innerHTML = '<div>Ingen L-linjer fundet for denne salgsordre.</div>';
                        totals.innerHTML = '<div><strong>Samlet L-kost (NestKost):</strong> 0,00 DKK</div><div><strong>Ordre stykliste kg:</strong> 0,00 kg</div><div><strong>Ordre forbrugt kg:</strong> 0,00 kg</div><div><strong>Afvigelse kg:</strong> 0,00 kg</div><div><strong>Samlet afvigelse %:</strong> NULL</div>';
                        if (teaser) teaser.innerHTML = '<strong>Ingen laserlinjer fundet</strong>';
                        if (currentOversigtModalType === 'laser') buildOversigtModalView('laser');
                        return;
                    }

                    const results = await Promise.all(requests);
                    const rows = [];

                    for (const item of results) {
                        if (!item || !item.ok || !item.data || item.data.error) continue;
                        const products = Array.isArray(item.data.products) ? item.data.products : [];
                        for (const p of products) {
                            const expected = p.NWgtU_medio;
                            const effective = p.KgPerPezzoEffettivo;
                            const hintedNestCost = getLaserNestCostHint(item.prodOrderNo, p.ProdNo);
                            const hasHintedNestCost = hintedNestCost !== null && hintedNestCost !== undefined && Number(hintedNestCost) > 0;
                            const routeSpecificCostPerPiece = hasHintedNestCost
                                ? hintedNestCost
                                : ((p.CostoPerPezzo !== null && p.CostoPerPezzo !== undefined)
                                    ? p.CostoPerPezzo
                                    : null);
                            const extraPct = (expected !== null && expected !== undefined && expected > 0 && effective !== null && effective !== undefined)
                                ? (((effective - expected) / expected) * 100)
                                : null;

                            rows.push({
                                prodOrderNo: item.prodOrderNo,
                                nestingOrdNo: p.NestingOrdNo || item.data.nestingOrdNo,
                                prodNo: p.ProdNo,
                                route: p.Route || item.data.route || item.requestedRoute,
                                noFin: p.QtaPezzi,
                                oldNWgtU_medio: p.OldNWgtU_medio,
                                expected,
                                effective,
                                costPerPiece: routeSpecificCostPerPiece,
                                quotaCost: p.QuotaCosto,
                                extraPct,
                                imageItems: Array.isArray(p.ImageItems) ? p.ImageItems : []
                            });
                        }
                    }

                    if (rows.length === 0) {
                        body.innerHTML = '<div>Ingen laserberegninger tilgaengelige for denne salgsordre.</div>';
                        totals.innerHTML = '<div><strong>Samlet L-kost (NestKost):</strong> 0,00 DKK</div><div><strong>Ordre stykliste kg:</strong> 0,00 kg</div><div><strong>Ordre forbrugt kg:</strong> 0,00 kg</div><div><strong>Afvigelse kg:</strong> 0,00 kg</div><div><strong>Samlet afvigelse %:</strong> NULL</div>';
                        if (teaser) teaser.innerHTML = '<strong>Ingen laserdata klar</strong>';
                        if (currentOversigtModalType === 'laser') buildOversigtModalView('laser');
                        return;
                    }

                    const multiNestHeader = orderGr4 === 3 ? 'NestMultiPris' : 'NestKost pr. stk';
                    let html = '<table class="oversigt-table-laser">';
                    html += '<tr><th>Prod.ordre</th><th>Nesting</th><th>Produkt</th><th>Rute</th><th>Færdig</th><th>Icon kg</th><th>Stykl. kg</th><th>Forbrug kg</th><th>' + (orderGr4 === 3 ? 'Multi/stk' : 'Nest/stk') + '</th><th>Samlet</th><th>Afvig. %</th><th>Vis</th></tr>';
                    let totalKgUtilizzati = 0;
                    let totalKgPrevisti = 0;
                    let totalKgIcon = 0;
                    let totalLaserCost = 0;
                    for (const r of rows) {
                        const rowNoFin = Number(r.noFin || 0);
                        const rowIcon = Number(r.oldNWgtU_medio || 0);
                        const rowExpected = Number(r.expected || 0);
                        const rowEffective = Number(r.effective || 0);
                        const rowCostPerPiece = Number(r.costPerPiece || 0);
                        const rowTotalCost = (r.costPerPiece !== null && r.costPerPiece !== undefined && rowNoFin > 0)
                            ? (rowNoFin * rowCostPerPiece)
                            : ((r.costPerPiece === null || r.costPerPiece === undefined || r.noFin === null || r.noFin === undefined)
                                ? null
                                : (rowNoFin * rowCostPerPiece));
                        totalKgIcon += rowNoFin * rowIcon;
                        totalKgPrevisti += rowNoFin * rowExpected;
                        totalKgUtilizzati += rowNoFin * rowEffective;
                        totalLaserCost += rowTotalCost || 0;

                        html += '<tr>';
                        html += '<td>' + (r.prodOrderNo || '-') + '</td>';
                        html += '<td>' + (r.nestingOrdNo || '-') + '</td>';
                        html += '<td>' + (r.prodNo || '-') + '</td>';
                        html += '<td>' + (r.route || '-') + '</td>';
                        html += '<td>' + (r.noFin === null || r.noFin === undefined ? 'NULL' : formatNumber(r.noFin)) + '</td>';
                        html += '<td>' + (r.oldNWgtU_medio === null || r.oldNWgtU_medio === undefined ? 'NULL' : formatNumber(r.oldNWgtU_medio)) + '</td>';
                        html += '<td>' + (r.expected === null || r.expected === undefined ? 'NULL' : formatNumber(r.expected)) + '</td>';
                        html += '<td>' + (r.effective === null || r.effective === undefined ? 'NULL' : formatNumber(r.effective)) + '</td>';
                        html += '<td>' + (r.costPerPiece === null || r.costPerPiece === undefined ? 'NULL' : formatNumber(r.costPerPiece)) + '</td>';
                        html += '<td>' + (rowTotalCost === null ? 'NULL' : formatNumber(rowTotalCost)) + '</td>';
                        html += '<td>' + (r.extraPct === null || r.extraPct === undefined ? 'NULL' : (formatNumber(r.extraPct) + '%')) + '</td>';
                        if (Array.isArray(r.imageItems) && r.imageItems.length > 0) {
                            const imageKey = registerSummaryImageData('Billeder for ' + (r.prodNo || 'produkt') + ' / rute ' + (r.route || '-'), r.imageItems);
                            html += '<td><button class="image-preview-btn" data-image-mode="compact" data-image-key="' + imageKey + '">Vis</button></td>';
                        } else {
                            html += '<td>-</td>';
                        }
                        html += '</tr>';
                    }
                    const deltaKg = totalKgUtilizzati - totalKgPrevisti;
                    const deltaPct = totalKgPrevisti > 0
                        ? ((deltaKg / totalKgPrevisti) * 100)
                        : null;
                    html += '</table>';
                    body.innerHTML = html;
                    applyMicroTablePolish(body);
                    totals.innerHTML = ''
                        + '<div><strong>Samlet L-kost (' + (orderGr4 === 3 ? 'NestMultiPris' : 'NestKost') + '):</strong> ' + formatNumber(totalLaserCost) + ' DKK</div>'
                        + '<div><strong>Ordre icon kg:</strong> ' + formatNumber(totalKgIcon) + ' kg</div>'
                        + '<div><strong>Ordre stykliste kg:</strong> ' + formatNumber(totalKgPrevisti) + ' kg</div>'
                        + '<div><strong>Ordre forbrugt kg:</strong> ' + formatNumber(totalKgUtilizzati) + ' kg</div>'
                        + '<div><strong>Afvigelse kg:</strong> ' + formatNumber(deltaKg) + ' kg</div>'
                        + '<div><strong>Samlet afvigelse %:</strong> ' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + '</div>';
                    if (teaser) {
                        teaser.innerHTML = ''
                            + '<div><strong>' + formatNumber(totalLaserCost) + ' DKK</strong> samlet L-kost</div>'
                            + '<div>Afvigelse: <strong>' + formatNumber(deltaKg) + ' kg</strong> (' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + ')</div>';
                    }
                    if (currentOversigtModalType === 'laser') buildOversigtModalView('laser');
                } catch (err) {
                    body.innerHTML = '<div class="error">Fejl laseroversigt: ' + err.message + '</div>';
                    totals.innerHTML = '<div class="error">Fejl i samlet laseroversigt: ' + err.message + '</div>';
                    if (teaser) teaser.innerHTML = '<strong>Fejl i laser KPI:</strong> ' + err.message;
                }
            }

            async function onProductClick(prodNo, ordNo, lnNo, prodTp4, trInf2, trInf4, showAllRoutes, clickedNoFin, clickedNestingCost) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');

                const modalWasOpen = modal.style.display === 'flex';
                if (modalWasOpen) {
                    pushSummaryModalState();
                } else {
                    summaryModalHistory = [];
                    updateSummaryModalBackBtn();
                }

                closeSummaryImagePanel();

                title.textContent = 'Produkt: ' + prodNo;
                modal.style.display = 'flex';

                if (String(prodTp4) === '1') {
                    body.innerHTML = '<div class="modal-loading">Indlæser transaktioner...</div>';
                    try {
                        const response = await fetch('/prodtr/' + ordNo + '/' + lnNo);
                        const rows = await response.json();
                        if (!response.ok || rows.error) {
                            body.innerHTML = '<div class="error">Fejl: ' + (rows.error || 'Uventet fejl') + '</div>';
                            return;
                        }
                        if (!rows.length) {
                            body.innerHTML = '<div>Ingen ProdTr-linjer fundet.</div>';
                            return;
                        }
                        let html = '<table>';
                        html += '<tr><th>Færdigmeldingsdato</th><th>Færdigmeldingstid</th><th>Minutter</th><th>Hvem</th></tr>';
                        for (const r of rows) {
                            const rawFinDt = String(r.FinDt || '').trim();
                            const compactFinDt = rawFinDt.split('T')[0].replace(/-/g, '');
                            let finDt = '-';
                            if (/^\\d{8}$/.test(compactFinDt)) {
                                finDt = compactFinDt.slice(6, 8) + '-' + compactFinDt.slice(4, 6) + '-' + compactFinDt.slice(0, 4);
                            } else if (rawFinDt) {
                                finDt = rawFinDt;
                            }
                            const rawFinTm = r.FinTm != null ? String(r.FinTm).trim() : '';
                            const finTm = rawFinTm
                                ? rawFinTm.padStart(4, '0').replace(/^(\\d{2})(\\d{2})$/, '$1:$2')
                                : '-';
                            html += '<tr>';
                            html += '<td>' + finDt + '</td>';
                            html += '<td>' + finTm + '</td>';
                            html += '<td>' + formatNumber(r.NoInvoAb || 0) + '</td>';
                            html += '<td>' + (r.HvemNm || '-') + '</td>';
                            html += '</tr>';
                        }
                        html += '</table>';
                        body.innerHTML = html;
                        applyMicroTablePolish(body);
                    } catch (err) {
                        body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                    }
                } else if (String(prodTp4) === '2') {
                    body.innerHTML = '<div class="modal-loading">Indlaeser ruteberegning...</div>';
                    try {
                        const effectiveOrdine = String(ordNo || trInf2 || '').trim();
                        let effectiveRoute = String(trInf4 || '').trim();

                        if (!effectiveOrdine) {
                            body.innerHTML = '<div class="error">Fejl: OrdNo/TrInf2 mangler paa den valgte linje.</div>';
                            return;
                        }

                        if (!showAllRoutes && !effectiveRoute) {
                            const encProdNo = encodeURIComponent(prodNo || '');
                            const fallbackResponse = await fetch('/nesting-detail/' + encodeURIComponent(effectiveOrdine) + '/' + encProdNo);
                            const fallbackRows = await fallbackResponse.json();
                            if (fallbackResponse.ok && Array.isArray(fallbackRows) && fallbackRows.length > 0) {
                                effectiveRoute = String(fallbackRows[0].TrInf4 || '').trim();
                            }
                        }

                        if (!showAllRoutes && !effectiveRoute) {
                            body.innerHTML = '<div class="error">Fejl: TrInf4 (route) mangler paa den valgte linje.</div>';
                            return;
                        }

                        const endpoint = buildLaserRouteMetricsEndpoint(effectiveOrdine, effectiveRoute, prodNo, showAllRoutes);
                        const data = await requestRouteMetricsData(endpoint);

                        const finalData = data;
                        const usedProdFilter = Boolean(prodNo);

                        const s = finalData.summary || {};
                        const products = Array.isArray(finalData.products) ? finalData.products : [];
                        const formatNullable = (value, suffix = '') => {
                            return value === null || value === undefined
                                ? 'NULL'
                                : (formatNumber(value) + suffix);
                        };

                        if (!products.length) {
                            body.innerHTML = usedProdFilter
                                ? '<div>Ingen faerdigvarer (TrTp=7) fundet for valgt produkt/route.</div>'
                                : '<div>Ingen faerdigvarer (TrTp=7) fundet for valgt rute.</div>';
                            return;
                        }

                        const multiNestHeader = currentSalesOrderGr4 === 3 ? 'NestMultiPris' : 'NestKost pr. stk';
                        let html = '<table>';
                        html += '<tr><th>Nestingordre</th><th>Produkt</th><th>Rute</th><th>Færdigmeldt</th><th>Icon vægt (kg/stk)</th><th>Stykliste vaegt (kg/stk)</th><th>Forbrugt (kg/stk)</th><th>' + multiNestHeader + '</th><th>Samlet kost</th><th>Afvigelse (%)</th><th>Billeder</th></tr>';
                        let totalKgPrevisti = 0;
                        let totalKgUtilizzati = 0;
                        let totalKgIcon = 0;
                        let totalLaserCost = 0;
                        const clickedNoFinNum = Number(clickedNoFin || 0);
                        const clickedNestingCostNum = Number(clickedNestingCost || 0);

                        for (const rowProduct of products) {
                            const oldExpected = rowProduct ? rowProduct.OldNWgtU_medio : null;
                            const expected = rowProduct ? rowProduct.NWgtU_medio : null;
                            const effective = rowProduct ? rowProduct.KgPerPezzoEffettivo : null;
                            const routeNoFin = rowProduct ? rowProduct.QtaPezzi : null;
                            const prodNoForCost = rowProduct ? (rowProduct.ProdNo || prodNo) : prodNo;
                            const isClickedProd = String(prodNoForCost || '').trim().toUpperCase() === String(prodNo || '').trim().toUpperCase();
                            const hasClickedNestCost = !showAllRoutes && isClickedProd && clickedNestingCostNum > 0;
                            const noFin = (!showAllRoutes && hasClickedNestCost && clickedNoFinNum > 0) ? clickedNoFinNum : routeNoFin;
                            const hintedNestCost = getLaserNestCostHint(effectiveOrdine, prodNoForCost);
                            const costPerPiece = hasClickedNestCost
                                ? clickedNestingCostNum
                                : ((rowProduct && rowProduct.CostoPerPezzo !== null && rowProduct.CostoPerPezzo !== undefined)
                                    ? rowProduct.CostoPerPezzo
                                    : hintedNestCost);
                            const noFinNum = Number(noFin || 0);
                            const expectedNum = Number(expected || 0);
                            const effectiveNum = Number(effective || 0);
                            const baseTotalCost = hasClickedNestCost
                                ? (noFinNum > 0 ? (noFinNum * Number(costPerPiece || 0)) : null)
                                : ((rowProduct && rowProduct.QuotaCosto !== null && rowProduct.QuotaCosto !== undefined)
                                    ? rowProduct.QuotaCosto
                                : ((costPerPiece === null || costPerPiece === undefined || noFin === null || noFin === undefined)
                                    ? null
                                    : (noFinNum * Number(costPerPiece || 0))));
                            let totalCost = baseTotalCost;
                            let displayCostPerPiece = costPerPiece;
                            totalKgIcon += noFinNum * Number(oldExpected || 0);
                            totalKgPrevisti += noFinNum * expectedNum;
                            totalKgUtilizzati += noFinNum * effectiveNum;
                            totalLaserCost += totalCost || 0;
                            const extraPct = (expected !== null && expected !== undefined && expected > 0 && effective !== null && effective !== undefined)
                                ? (((effective - expected) / expected) * 100)
                                : null;

                            html += '<tr>';
                            html += '<td>' + ((rowProduct && rowProduct.NestingOrdNo) || finalData.nestingOrdNo || '-') + '</td>';
                            html += '<td>' + (rowProduct ? (rowProduct.ProdNo || '-') : '-') + '</td>';
                            html += '<td>' + (rowProduct ? (rowProduct.Route || '-') : (finalData.route || '-')) + '</td>';
                            html += '<td>' + formatNullable(noFin) + '</td>';
                            html += '<td>' + formatNullable(oldExpected) + '</td>';
                            html += '<td>' + formatNullable(expected) + '</td>';
                            html += '<td>' + formatNullable(effective) + '</td>';
                            html += '<td>' + formatNullable(displayCostPerPiece) + '</td>';
                            html += '<td>' + formatNullable(totalCost) + '</td>';
                            html += '<td>' + (extraPct === null ? 'NULL' : (formatNumber(extraPct) + '%')) + '</td>';
                            if (Array.isArray(rowProduct.ImageItems) && rowProduct.ImageItems.length > 0) {
                                const imageKey = registerSummaryImageData('Billeder for ' + (rowProduct.ProdNo || 'produkt') + ' / rute ' + (rowProduct.Route || '-'), rowProduct.ImageItems);
                                html += '<td><button class="image-preview-btn" data-image-mode="compact" data-image-key="' + imageKey + '">Vis</button></td>';
                            } else {
                                html += '<td>-</td>';
                            }
                            html += '</tr>';
                        }
                        html += '</table>';
                        const displayedLaserTotal = (showAllRoutes && clickedNoFinNum > 0 && clickedNestingCostNum > 0)
                            ? (clickedNoFinNum * clickedNestingCostNum)
                            : totalLaserCost;
                        html += '<div class="summary-box" style="margin-top:12px;">'
                            + '<div><strong>Samlet L-kost (NestKost):</strong> ' + formatNumber(displayedLaserTotal) + ' DKK</div>'
                            + '<div><strong>Ordre icon kg:</strong> ' + formatNumber(totalKgIcon) + ' kg</div>'
                            + '<div><strong>Ordre stykliste kg:</strong> ' + formatNumber(totalKgPrevisti) + ' kg</div>'
                            + '<div><strong>Ordre forbrugt kg:</strong> ' + formatNumber(totalKgUtilizzati) + ' kg</div>'
                            + '</div>';
                        body.innerHTML = html;
                        applyMicroTablePolish(body);
                    } catch (err) {
                        body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                    }
                }
            }

            function handleProdNoClick(e) {
                const span = e.target.closest('.prod-no-link');
                if (!span) return;
                const prodNo = span.dataset.prodno;
                const ordNo = span.dataset.ordno;
                const lnNo = span.dataset.lnno;
                const prodTp4 = span.dataset.prodtp4;
                const trInf2 = span.dataset.trinf2;
                const trInf4 = span.dataset.trinf4;
                const showAllRoutes = span.dataset.showallroutes === '1';
                const noFin = span.dataset.nofin;
                const nestingCost = span.dataset.nestingcost;
                if (prodNo) onProductClick(prodNo, ordNo, lnNo, prodTp4, trInf2, trInf4, showAllRoutes, noFin, nestingCost);
            }

            function handleProdNoHover(e) {
                const span = e.target.closest('.prod-no-link');
                if (!span) return;
                if (span.dataset.prodtp4 !== '2') return;
                if (span.dataset.routePrefetchStarted === '1') return;
                span.dataset.routePrefetchStarted = '1';
                prefetchRouteMetricsForProduct(
                    span.dataset.prodno,
                    span.dataset.ordno,
                    span.dataset.trinf2,
                    span.dataset.trinf4,
                    span.dataset.showallroutes === '1'
                );
            }

            function handleImagePreviewClick(e) {
                const btn = e.target.closest('.image-preview-btn');
                if (!btn) return;
                const imageKey = btn.dataset.imageKey;
                if (imageKey) {
                    if (btn.dataset.imageMode === 'compact') {
                        openCompactImageModal(imageKey);
                        return;
                    }
                    const inLaserPanel = !!e.target.closest('#laserOrderSummaryPanel');
                    const preferredPanelId = inLaserPanel ? 'laserImagePanel' : 'summaryImagePanel';
                    openSummaryImagePanel(imageKey, preferredPanelId);
                }
            }

            function handleDrawingOpenClick(e) {
                const btn = e.target.closest('.drawing-open-btn');
                if (!btn) return;
                openDrawingPdf({
                    path: btn.dataset.drawingPath || '',
                    prodNo: btn.dataset.prodNo || '',
                    ordNo: btn.dataset.ordNo || ''
                });
            }

            function handlePreviewImageZoom(e) {
                const image = e.target.closest('.image-preview-zoomable');
                if (!image) return;
                openImageLightbox(
                    image.dataset.fullsrc || image.getAttribute('src') || '',
                    image.dataset.title || image.getAttribute('alt') || 'Billede',
                    image.dataset.path || ''
                );
            }

            // Outside modal content.
            document.addEventListener('click', handleProdNoClick);
            document.addEventListener('mouseover', handleProdNoHover);
            document.addEventListener('focusin', handleProdNoHover);
            document.addEventListener('click', handleImagePreviewClick);
            document.addEventListener('click', handleDrawingOpenClick);
            document.addEventListener('click', handlePreviewImageZoom);
            document.addEventListener('keydown', function(event) {
                if (event.key === 'Escape') {
                    closeSideMenu();
                    closeImageLightbox();
                    closeCompactImageModal();
                    closeOversigtModal();
                }
            });
            // Inside modal content (document listener is blocked by modal stopPropagation).
            const summaryModalBodyEl = document.getElementById('summaryModalBody');
            if (summaryModalBodyEl) {
                summaryModalBodyEl.addEventListener('click', handleProdNoClick);
                summaryModalBodyEl.addEventListener('mouseover', handleProdNoHover);
                summaryModalBodyEl.addEventListener('focusin', handleProdNoHover);
                summaryModalBodyEl.addEventListener('click', handleImagePreviewClick);
                summaryModalBodyEl.addEventListener('click', handlePreviewImageZoom);
            }
            const summaryImagePanelEl = document.getElementById('summaryImagePanel');
            if (summaryImagePanelEl) {
                summaryImagePanelEl.addEventListener('click', handlePreviewImageZoom);
            }
            window.addEventListener('resize', updateSummaryImagePanelLayout);
            const oversigtModalBodyEl = document.getElementById('oversigtModalBody');
            if (oversigtModalBodyEl) {
                oversigtModalBodyEl.addEventListener('click', handleProdNoClick);
                oversigtModalBodyEl.addEventListener('click', handleImagePreviewClick);
                oversigtModalBodyEl.addEventListener('click', handleDrawingOpenClick);
                oversigtModalBodyEl.addEventListener('click', handlePreviewImageZoom);
            }
            const orderDetailModalBodyEl = document.getElementById('orderDetailModalBody');
            if (orderDetailModalBodyEl) {
                // Clicks inside the order detail modal do not bubble to document because the shell stops propagation.
                orderDetailModalBodyEl.addEventListener('click', handleProdNoClick);
                orderDetailModalBodyEl.addEventListener('click', handleImagePreviewClick);
                orderDetailModalBodyEl.addEventListener('click', handleDrawingOpenClick);
                orderDetailModalBodyEl.addEventListener('click', handlePreviewImageZoom);
            }
            const laserImagePanelEl = document.getElementById('laserImagePanel');
            if (laserImagePanelEl) {
                laserImagePanelEl.addEventListener('click', handlePreviewImageZoom);
            }

            function toggleSalesLineBreakdown(rowId, buttonEl) {
                const row = document.getElementById(rowId);
                if (!row) return;
                const isClosed = row.style.display === 'none';
                row.style.display = isClosed ? 'table-row' : 'none';
                if (buttonEl) buttonEl.textContent = isClosed ? '−' : '+';
            }

            function toggleProdTp4Group(orderNo, prodTp4Key) {
                const el = document.getElementById('po-' + orderNo + '-group-' + prodTp4Key);
                const icon = document.getElementById('po-' + orderNo + '-group-' + prodTp4Key + '-icon');
                if (!el) return;
                const isClosed = el.style.display === 'none';
                el.style.display = isClosed ? '' : 'none';
                if (icon) icon.textContent = isClosed ? '▾' : '▸';
            }

            async function showChildProductionSummary(childOrdNo, targetProdNo, forceInvoiceStatus) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');

                const modalWasOpen = modal.style.display === 'flex';
                if (modalWasOpen) {
                    pushSummaryModalState();
                } else {
                    summaryModalHistory = [];
                    updateSummaryModalBackBtn();
                }
                closeSummaryImagePanel();
                const normalizedTargetProdNo = String(targetProdNo || '').trim();
                title.textContent = normalizedTargetProdNo
                    ? ('Produktoversigt for ordre ' + childOrdNo + ' - ' + normalizedTargetProdNo)
                    : ('Produktoversigt for ordre ' + childOrdNo);
                body.innerHTML = '<div class="modal-loading">Indlaeser...</div>';
                modal.style.display = 'flex';

                try {
                    const response = await fetch('/production-summary/' + childOrdNo + (currentSalesOrderGr4 === 3 ? '?gr4=3' : ''));
                    const data = await response.json();

                    if (!response.ok || data.error) {
                        body.innerHTML = '<div class="error">Fejl: ' + (data.error || 'Uventet fejl') + '</div>';
                        return;
                    }

                    if (!data.lines || data.lines.length === 0) {
                        body.innerHTML = '<div>Ingen linjer fundet for denne produktionsordre.</div>';
                        return;
                    }

                    const filteredLines = normalizedTargetProdNo
                        ? data.lines.filter(line => String(line && line.ProdNo || '').trim().toUpperCase() === normalizedTargetProdNo.toUpperCase())
                        : data.lines;

                    if (!filteredLines || filteredLines.length === 0) {
                        body.innerHTML = '<div>Det valgte produkt blev ikke fundet i denne produktionsordre.</div>';
                        return;
                    }

                    const baseTitleText = normalizedTargetProdNo
                        ? ('Produktoversigt for ordre ' + childOrdNo + ' - ' + normalizedTargetProdNo)
                        : ('Produktoversigt for ordre ' + childOrdNo);
                    const titleFlags = [
                        data.hasEstimatedOperationTime ? '🕒' : '',
                        data.hasWarnings ? '⚠️' : ''
                    ].filter(Boolean).join(' ');
                    title.textContent = titleFlags
                        ? (baseTitleText + ' ' + titleFlags)
                        : baseTitleText;

                    const isYdelseFilteredView = !!normalizedTargetProdNo;
                    const modalTotalCost = normalizedTargetProdNo
                        ? filteredLines.reduce((sum, line) => sum + Number(line && line.EffectiveLineCost || 0), 0)
                        : Number(data.totalCost || 0);

                    const shouldShowInvoiceStatus = Boolean(forceInvoiceStatus) || filteredLines.some(line => line && (line.IsInvoiceTracked || line.isInvoiceTracked || isInvoiceTrackedProdNo(line.ProdNo)));

                    let html = '';
                    html += getInvoiceStatusSummaryHtml(filteredLines, shouldShowInvoiceStatus);
                    html += isYdelseFilteredView
                        ? '<table><tr><th>Linje</th><th>ProdTp4</th><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Pris/enhed</th><th>Samlet kost (beregnet)</th></tr>'
                        : '<table><tr><th>Linje</th><th>ProdTp4</th><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Salgspris</th><th>Kostpris/enhed</th><th>Nesting/enhed</th><th>Samlet kost (beregnet)</th></tr>';
                    for (const line of filteredLines) {
                        const lineExcludedFromTotal = !isYdelseFilteredView && isProductionSummaryExcludedLine(line);
                        const displayLineCost = lineExcludedFromTotal
                            ? null
                            : Number(line.EffectiveLineCost || 0);
                        const warningFlagHtml = getWarningFlagHtml(line);
                        const invoiceStatusFlagHtml = getInvoiceStatusFlagHtml(line, shouldShowInvoiceStatus);
                        const timeAdjustmentFlagHtml = getTimeAdjustmentFlagHtml(line);
                        const laserAllocationFlagHtml = getLaserAllocationFlagHtml(line);
                        html += '<tr>';
                        html += '<td>' + (line.LnNo || 0) + '</td>';
                        const displayProdTp4 = getDisplayProdTp4Key(line.ProdTp4, line.ProdNo, line.PurcNo);
                        html += '<td>' + (displayProdTp4 === 'NA' ? '-' : displayProdTp4) + '</td>';
                        const childHasPurcNo = Number(line.PurcNo || 0) > 0;
                        if (String(line.ProdTp4 || '') === '1' && line.ProdNo) {
                            const safeChildProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const trInf2FromLine = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : childOrdNo);
                            const trInf4FromLine = String(line.TrInf4 || '');
                            const safeChildTrInf2 = trInf2FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const safeChildTrInf4 = trInf4FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            html += '<td><span class="prod-no-link" data-prodno="' + safeChildProdNo + '" data-ordno="' + childOrdNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="1" data-trinf2="' + safeChildTrInf2 + '" data-trinf4="' + safeChildTrInf4 + '">' + safeChildProdNo + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else if (line.ProdNo && String(line.ProdNo).trim().toUpperCase().endsWith('L')) {
                            const safeChildProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const trInf2FromLine = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : childOrdNo);
                            const trInf4FromLine = String(line.TrInf4 || '');
                            const safeChildTrInf2 = trInf2FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const safeChildTrInf4 = trInf4FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            html += '<td><span class="prod-no-link" data-prodno="' + safeChildProdNo + '" data-ordno="' + childOrdNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="2" data-trinf2="' + safeChildTrInf2 + '" data-trinf4="' + safeChildTrInf4 + '" data-showallroutes="1" data-nofin="' + Number(line.NoFin || 0) + '" data-nestingcost="' + Number(line.NestingCost || 0) + '">' + safeChildProdNo + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else if (childHasPurcNo) {
                            const safeChildProdNoForSummary = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const childSummaryArgs = shouldFilterChildSummary(displayProdTp4, line.ProdNo, line.PurcNo)
                                ? (Number(line.PurcNo || 0) + ', &quot;' + safeChildProdNoForSummary + '&quot;, true')
                                : Number(line.PurcNo || 0);
                            html += '<td><span class="inline-link" onclick="showChildProductionSummary(' + childSummaryArgs + ')">' + (line.ProdNo || '-') + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else {
                            html += '<td>' + (line.ProdNo || '-') + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        }
                        const displayQty = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                            ? line.DisplayQuantity
                            : (line.NoFin || 0);
                        const displayUnitCost = (Number(displayQty || 0) > 0 && displayLineCost !== undefined && displayLineCost !== null)
                            ? ((displayLineCost || 0) / displayQty)
                            : ((line.DisplayUnitCost !== undefined && line.DisplayUnitCost !== null)
                                ? line.DisplayUnitCost
                                : (line.CCstPr || 0));
                        const isLaserProdLine = isLaserLProdNo(line.ProdNo);
                        html += '<td>' + (line.Descr || '') + '</td>';
                        html += '<td>' + formatNumber(displayQty) + '</td>';
                        if (isYdelseFilteredView) {
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                        } else {
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                            html += '<td>' + (isLaserProdLine ? '-' : formatNumber(displayUnitCost)) + '</td>';
                            const hasLaserAllocationSpread = Boolean(line.UsesLaserAllocationSpread || line.usesLaserAllocationSpread);
                            const allocationTitle = hasLaserAllocationSpread
                                ? ' title="Nesting-fordeling bruger et andet antal end ordrelinjen; pris pr. stk kan afvige."'
                                : '';
                            const allocationHint = hasLaserAllocationSpread ? ' <span style="color:#b26a00; font-weight:700;">*</span>' : '';
                            html += '<td' + allocationTitle + '>' + formatNumber(line.NestingCost || 0) + allocationHint + '</td>';
                        }
                        html += '<td><strong>' + (displayLineCost === null ? '-' : formatNumber(displayLineCost)) + '</strong></td>';
                        html += '</tr>';
                    }
                    html += isYdelseFilteredView
                        ? '<tr class="summary-row"><td colspan="6">Total beregnet kost:</td><td><strong>' + formatNumber(modalTotalCost || 0) + ' DKK</strong></td></tr>'
                        : '<tr class="summary-row"><td colspan="8">Total beregnet kost:</td><td><strong>' + formatNumber(modalTotalCost || 0) + ' DKK</strong></td></tr>';
                    html += '</table>';
                    body.innerHTML = html;
                    applyMicroTablePolish(body);
                } catch (err) {
                    body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                }
            }

            function closeSummaryModal(event) {
                if (event && event.target && event.target.id !== 'summaryModal') return;
                const modal = document.getElementById('summaryModal');
                modal.style.display = 'none';
                summaryModalHistory = [];
                closeSummaryImagePanel();
                updateSummaryModalBackBtn();
            }

            function scrollToElementWithStickyOffset(el) {
                if (!el) return;
                const header = document.querySelector('.header-banner-wrapper');
                const searchBox = document.getElementById('searchBox');
                const headerH = header ? header.offsetHeight : 0;
                const searchH = searchBox ? searchBox.offsetHeight : 0;
                const extraGap = 14;
                const targetTop = window.pageYOffset + el.getBoundingClientRect().top - headerH - searchH - extraGap;
                window.scrollTo({ top: Math.max(targetTop, 0), behavior: 'auto' });
            }

            function syncStickyOffsets() {
                const root = document.documentElement;
                const header = document.querySelector('.header-banner-wrapper');
                const searchBox = document.getElementById('searchBox');

                const headerH = header ? Math.ceil(header.getBoundingClientRect().height || 0) : 0;
                const searchVisible = !!searchBox && getComputedStyle(searchBox).display !== 'none';
                const searchH = searchVisible ? Math.ceil(searchBox.getBoundingClientRect().height || 0) : 0;

                // Header height can increase on smaller viewports; keep sticky controls below it.
                const searchTop = Math.max(headerH + 8, 58);
                const tableTop = Math.max(headerH + (searchVisible ? searchH : 0) + 8, 0);

                root.style.setProperty('--search-sticky-top', String(searchTop) + 'px');
                root.style.setProperty('--table-sticky-top', String(tableTop) + 'px');
            }

            function openProduction(ordNo) {
                const el = document.getElementById('po-' + ordNo);
                if (!el) {
                    alert('Produktionsordre ' + ordNo + ' blev ikke fundet i de indlaeste resultater.');
                    return;
                }

                const modal = document.getElementById('orderDetailModal');
                const modalBody = document.getElementById('orderDetailModalBody');
                const modalIsOpen = modal && getComputedStyle(modal).display === 'flex';
                if (modalIsOpen && modalBody && modalBody.contains(el)) {
                    // When browsing inside the report modal, scroll the modal body instead of the page.
                    const top = Math.max(el.offsetTop - 16, 0);
                    modalBody.scrollTo({ top, behavior: 'auto' });
                } else {
                    scrollToElementWithStickyOffset(el);
                }
                el.classList.add('po-highlight');
                setTimeout(() => el.classList.remove('po-highlight'), 1800);
            }
            
            function renderOrderList() {
                const el = document.getElementById('orderList');
                const toggleBtn = document.getElementById('listToggleBtn');

                if (!orderListVisible) {
                    if (toggleBtn) toggleBtn.textContent = 'Vis kundeliste';
                    el.innerHTML = '';
                    return;
                }

                if (toggleBtn) toggleBtn.textContent = 'Skjul kundeliste';
                if (!orderListData || orderListData.length === 0) {
                    el.innerHTML = '<div class="loading">Indlaeser ordreliste...</div>';
                    return;
                }

                const orders = getFilteredOrders();
                if (orders.length === 0) {
                    el.innerHTML = '<div class="order-list-section"><h3>Ingen kunder fundet</h3><div>Prøv en anden søgning.</div></div>';
                    return;
                }

                let html = '<div class="order-list-section">';
                html += '<div id="orderListSummary" class="order-list-summary">';
                html += buildOrderListSummaryHtml(orders);
                html += '</div>';
                html += '<h3>Seneste fakturerede ordrer (' + ORDER_LIST_DAYS_BACK_CLIENT + ' dage) &mdash; ' + orders.length + ' af ' + orderListData.length + ' ordrer</h3>';
                const sortMark = (field) => {
                    if (orderListSortField !== field) return ' <span style="opacity:0.4;">^v</span>';
                    return orderListSortDir === 'asc'
                        ? ' <span style="color:#1976d2;">^</span>'
                        : ' <span style="color:#1976d2;">v</span>';
                };
                html += '<table class="order-list-table"><tr>';
                html += '<th class="order-sortable-header" data-sort-field="bruger" style="cursor:pointer; user-select:none;">Bruger' + sortMark('bruger') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="ordno" style="cursor:pointer; user-select:none;">Ordrenr.' + sortMark('ordno') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="kunde" style="cursor:pointer; user-select:none;">Kunde' + sortMark('kunde') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="date" style="cursor:pointer; user-select:none;">Fakturadato' + sortMark('date') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="belob" style="cursor:pointer; user-select:none;">Fakturabelob' + sortMark('belob') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="margin" style="cursor:pointer; user-select:none;">Margin' + sortMark('margin') + '</th>';
                html += '<th>Faktura</th>';
                html += '<th>Note</th>';
                html += '<th>Opdater</th>';
                html += '</tr>';
                for (const o of orders) {
                    const marginHtml = getOrderMarginHtml(o.OrdNo);

                    const d = String(o.LstInvDt || '');
                    const invDate = d.length === 8 ? d.slice(0,4) + '-' + d.slice(4,6) + '-' + d.slice(6,8) : (d || '-');
                    const orderWarningFlag = getWarningFlagHtml(o, 'Ordren indeholder mindst én advarsel.');
                    // NOTE: Visual badge only. No logic change: Gr4=3 => Multiordre order type.
                    const gr4TypeBadge = Number(o.Gr4 || 0) === 3
                        ? '<span title="Ordretype: Multiordre (Gr4=3)" aria-label="Ordretype: Multiordre" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:#1565c0;color:#fff;font-size:11px;font-weight:700;margin-left:6px;vertical-align:middle;">M</span>'
                        : '';
                    html += '<tr data-ordno="' + o.OrdNo + '" class="order-list-row">'
                    html += '<td>' + (o.SellerUsr || '-') + '</td>';
                    html += '<td><strong>' + o.OrdNo + '</strong>' + gr4TypeBadge + orderWarningFlag + '</td>';
                    html += '<td>' + (o.CustomerName || '-') + '</td>';
                    html += '<td>' + invDate + '</td>';
                    html += '<td>' + formatNumber(o.InvoAm || 0) + ' DKK</td>';
                    html += '<td class="order-margin-cell" data-ordno="' + o.OrdNo + '">' + marginHtml + '</td>';
                    html += '<td class="order-invoice-cell" data-ordno="' + o.OrdNo + '">' + getOrderInvoiceStatusHtml(o.OrdNo) + '</td>';
                    html += '<td class="order-note-cell" data-ordno="' + o.OrdNo + '" onclick="event.stopPropagation();openNotePopup(' + o.OrdNo + ')">' + getOrderNoteHtml(o.OrdNo) + '</td>';
                    html += '<td class="order-refresh-cell"><button class="list-toggle-btn order-refresh-one-btn" data-ordno="' + o.OrdNo + '" style="padding:4px 8px; margin-left:0; background:#00695c !important;" title="Opdater cache for denne ordre">Opdater</button></td>';
                    html += '</tr>';
                }
                html += '</table>';

                html += '</div>';
                el.innerHTML = html;

                // Carica i margini in coda per tutti gli ordini visibili.
                const queuedOrders = orders.slice(0, MARGIN_PREFETCH_ROWS);
                queueMarginLoad(queuedOrders.map(o => o.OrdNo));
                updateSystemStatusFromOrders(queuedOrders);
            }

            async function loadOrderList(forceRefresh = false) {
                const el = document.getElementById('orderList');
                if (!el) return;

                const showOrderListError = (message) => {
                    el.innerHTML = '<div class="order-list-section"><h3>Ordreliste kunne ikke indlæses</h3><div>' + escapeHtml(message) + '</div><div style="margin-top:8px;"><button class="list-toggle-btn" onclick="refreshOrderList()">Prøv igen</button></div></div>';
                };

                if (orderListLoading && !forceRefresh) return;
                if (!forceRefresh && orderListData && orderListData.length > 0) {
                    renderOrderList();
                    return;
                }

                orderListLoading = true;
                const previousHtml = el.innerHTML;
                setSystemStatus('System loading...', '#fff3cd', '#8a6d3b');
                if (!orderListData || orderListData.length === 0) {
                    el.innerHTML = '<div class="loading">Indlaeser ordreliste...</div>';
                }
                try {
                    const endpoint = forceRefresh
                        ? '/order-list?force=1&t=' + Date.now()
                        : '/order-list';
                    const response = await fetch(endpoint);
                    if (!response.ok) {
                        setSystemStatus('System error', '#fdecea', '#b71c1c');
                        if (previousHtml) {
                            el.innerHTML = previousHtml;
                        } else {
                            showOrderListError('Serveren svarede med fejl (HTTP ' + response.status + ').');
                        }
                        return;
                    }
                    const orders = await response.json();
                    if (!orders || orders.error) {
                        setSystemStatus('System error', '#fdecea', '#b71c1c');
                        if (previousHtml) {
                            el.innerHTML = previousHtml;
                        } else {
                            showOrderListError((orders && orders.error) ? String(orders.error) : 'Ugyldigt svar fra serveren.');
                        }
                        return;
                    }
                    orderListData = orders;
                    hydrateMarginStateFromOrderList(orders);
                    populateBrugerFilterOptions();
                    loadAllNotes().then(() => {
                        if (orderListVisible) renderOrderList();
                    }).catch(() => {});
                    renderOrderList();
                    checkOrderListFreshness();
                } catch (err) {
                    console.error('Fejl i loadOrderList:', err);
                    setSystemStatus('System error', '#fdecea', '#b71c1c');
                    if (previousHtml) {
                        el.innerHTML = previousHtml;
                    } else {
                        showOrderListError(err && err.message ? err.message : 'Ukendt fejl.');
                    }
                } finally {
                    orderListLoading = false;
                }
            }

            function startOrderListAutoRefresh() {
                if (orderListAutoRefreshTimer) return;
                orderListAutoRefreshTimer = setInterval(() => {
                    if (document.hidden) return;
                    checkOrderListFreshness();
                }, ORDER_LIST_AUTO_REFRESH_MS);
            }

            async function refreshOrderList() {
                const btn = document.getElementById('refreshListBtn');
                if (btn) {
                    btn.disabled = true;
                    btn.textContent = 'Opdaterer...';
                }

                try {
                    await loadOrderList(true);
                } finally {
                    if (btn) {
                        btn.disabled = false;
                        btn.textContent = 'Opdater liste';
                    }
                }
            }

            async function refreshSingleOrderCache() {
                const ordNo = String(document.getElementById('orderInput').value || '').trim();
                if (!ordNo) {
                    alert('Indtast et ordrenummer foerst.');
                    return;
                }

                return refreshSingleOrderCacheByOrdNo(ordNo, true);
            }

            async function refreshSingleOrderCacheByOrdNo(ordNo, openAfter = false, clickedBtn = null) {
                const normalizedOrdNo = String(ordNo || '').trim();
                if (!normalizedOrdNo) return;
                const ordNoNum = Number(normalizedOrdNo);
                aftercalcClientCache.delete(normalizedOrdNo);
                let refreshSucceeded = false;

                const btn = document.getElementById('refreshSingleOrderBtn');
                if (btn) {
                    btn.disabled = true;
                    btn.textContent = 'Opdaterer ordre...';
                }
                if (clickedBtn) {
                    clickedBtn.disabled = true;
                    clickedBtn.textContent = '...';
                }

                try {
                    const r = await fetch('/cache-refresh-order/' + encodeURIComponent(normalizedOrdNo), { method: 'POST' });
                    const d = await r.json();
                    if (!r.ok || d.error) throw new Error((d && d.error) ? d.error : ('HTTP ' + r.status));

                    const startedAt = Date.now();
                    while (true) {
                        await new Promise(resolve => setTimeout(resolve, 1000));
                        const sr = await fetch('/cache-refresh-order-status/' + encodeURIComponent(normalizedOrdNo));
                        const sd = await sr.json();
                        if (sd && sd.status === 'done') {
                            break;
                        }
                        if (sd && sd.status === 'error') {
                            throw new Error(sd.error || 'Order refresh failed');
                        }
                        if (Date.now() - startedAt > 120000) {
                            throw new Error('Timeout waiting for order refresh');
                        }
                    }

                    await loadOrderList(true);
                    refreshSucceeded = true;
                    if (openAfter && Number.isFinite(ordNoNum)) {
                        await searchOrder();
                    } else {
                        const currentInputOrdNo = String((document.getElementById('orderInput') || {}).value || '').trim();
                        const detailModal = document.getElementById('orderDetailModal');
                        const detailOpen = detailModal && detailModal.style.display === 'flex';
                        if (detailOpen && currentInputOrdNo === normalizedOrdNo) {
                            await searchOrder();
                        }
                    }
                    if (clickedBtn) {
                        alert('Ordre ' + normalizedOrdNo + ' er opdateret fra kilden.');
                    }
                } catch (e) {
                    alert('Fejl ved ordre-cache opdatering: ' + e.message);
                } finally {
                    if (btn) {
                        btn.disabled = false;
                        btn.textContent = 'Opdater ordre-cache';
                    }
                    if (clickedBtn) {
                        clickedBtn.disabled = false;
                        clickedBtn.textContent = refreshSucceeded ? 'Opdateret' : 'Opdater';
                        if (refreshSucceeded) {
                            setTimeout(() => {
                                if (clickedBtn && clickedBtn.isConnected) clickedBtn.textContent = 'Opdater';
                            }, 1400);
                        }
                    }
                }
            }

            async function clearAppCache() {
                const confirmed = confirm('Er du sikker? Dette vil slette alt cache og tage lang tid at genindlæse data.');
                if (!confirmed) return;
                
                const btn = document.getElementById('clearCacheBtn');
                const dashBtn = document.getElementById('dashboardClearCacheBtn');
                if (btn) { btn.disabled = true; btn.textContent = 'Rydder...'; }
                if (dashBtn) { dashBtn.disabled = true; dashBtn.textContent = 'Rydder cache...'; }
                try {
                    showDashboardWarmupNotice = true;
                    warmupCombinedReady = false;
                    warmupCombinedPct = 0;
                    warmupCombinedDone = 0;
                    warmupCombinedTotal = 0;
                    const r = await fetch('/cache-clear', { method: 'POST' });
                    if (!r.ok) throw new Error('HTTP ' + r.status);
                    const d = await r.json();
                    alert('Cache ryddet: ' + (d.deleted || 0) + ' filer slettet.');
                } catch (e) {
                    alert('Fejl ved cache-rydning: ' + e.message);
                } finally {
                    if (btn) { btn.disabled = false; btn.textContent = 'Ryd cache'; }
                    if (dashBtn) { dashBtn.disabled = false; dashBtn.textContent = 'Ryd Efterkalk cache'; }
                }
            }

            async function checkDesktopUpdateNow() {
                const btn = document.getElementById('checkUpdateBtn');
                const dashBtn = document.getElementById('dashboardUpdateCheckBtn');
                if (btn) { btn.disabled = true; btn.textContent = 'Tjekker...'; }
                if (dashBtn) { dashBtn.disabled = true; dashBtn.textContent = 'Tjekker...'; }

                try {
                    const r = await fetch('/desktop-update-check', { method: 'POST' });
                    const d = await r.json();
                    if (!r.ok) throw new Error((d && d.message) ? d.message : ('HTTP ' + r.status));

                    // Show status only in dashboard update section (no popup alerts).
                    applyDashboardUpdateNotice({
                        status: d && d.status ? d.status : 'checking',
                        latestVersion: d && (d.latestVersion || d.version) ? (d.latestVersion || d.version) : undefined,
                        currentVersion: d && d.currentVersion ? d.currentVersion : undefined,
                        downloaded: d && d.downloaded === true,
                        canInstallNow: d && d.canInstallNow === true,
                        message: d && d.message ? d.message : 'Opdateringskontrol sendt.'
                    });
                } catch (e) {
                    applyDashboardUpdateNotice({ status: 'error', message: 'Fejl ved opdateringskontrol: ' + e.message });
                } finally {
                    if (btn) { btn.disabled = false; btn.textContent = 'Tjek opdatering nu'; }
                    if (dashBtn) { dashBtn.disabled = false; dashBtn.textContent = 'Tjek nu'; }
                    refreshDashboardUpdateNotice().catch(() => {});
                }
            }

            function applyDashboardUpdateNotice(state) {
                const wrap = document.getElementById('dashboardUpdateNotice');
                const titleEl = document.getElementById('dashboardUpdateTitle');
                const textEl = document.getElementById('dashboardUpdateText');
                const installBtn = document.getElementById('dashboardUpdateInstallBtn');
                if (!wrap || !titleEl || !textEl || !installBtn) return;

                const safe = state && typeof state === 'object' ? state : {};
                const status = String(safe.status || 'unavailable');
                const latest = safe.latestVersion ? String(safe.latestVersion) : null;
                const current = safe.currentVersion ? String(safe.currentVersion) : null;
                const canInstall = safe.canInstallNow === true || safe.downloaded === true || status === 'downloaded';

                let title = 'Programopdatering';
                let line = safe.message ? String(safe.message) : 'Ingen status endnu.';

                if (status === 'downloaded') {
                    title = 'Ny version klar';
                    line = 'Version ' + (latest || '?') + ' er hentet og klar til installation.';
                } else if (status === 'available') {
                    title = 'Ny version fundet';
                    line = 'Version ' + (latest || '?') + ' er fundet og hentes i baggrunden.';
                } else if (status === 'up-to-date') {
                    title = 'Programmet er opdateret';
                    line = current ? ('Aktuel version: ' + current + '.') : 'Du har den nyeste version.';
                } else if (status === 'checking' || status === 'busy') {
                    title = 'Søger efter opdateringer';
                    line = safe.message ? String(safe.message) : 'Tjekker...';
                } else if (status === 'unsupported' || status === 'unavailable') {
                    title = 'Opdatering ikke tilgængelig';
                } else if (status === 'installing') {
                    title = 'Installerer opdatering';
                } else if (status === 'error') {
                    title = 'Opdateringsfejl';
                }

                titleEl.textContent = title;
                textEl.textContent = line;
                wrap.classList.add('active');
                installBtn.style.display = canInstall ? 'inline-block' : 'none';
                installBtn.disabled = !canInstall;
            }

            async function refreshDashboardUpdateNotice() {
                const r = await fetch('/desktop-update-status');
                const d = await r.json();
                if (!r.ok) {
                    applyDashboardUpdateNotice({
                        status: 'error',
                        message: (d && d.message) ? d.message : ('HTTP ' + r.status)
                    });
                    return;
                }
                applyDashboardUpdateNotice(d || { status: 'unavailable', message: 'Ingen status.' });
            }

            function startDashboardUpdatePolling() {
                if (dashboardUpdatePollTimer) return;
                refreshDashboardUpdateNotice().catch(() => {});
                dashboardUpdatePollTimer = setInterval(() => {
                    if (document.hidden) return;
                    refreshDashboardUpdateNotice().catch(() => {});
                }, 20000);
            }

            async function installDesktopUpdateNow() {
                const installBtn = document.getElementById('dashboardUpdateInstallBtn');
                if (installBtn) {
                    installBtn.disabled = true;
                    installBtn.textContent = 'Installerer...';
                }

                try {
                    const r = await fetch('/desktop-update-install', { method: 'POST' });
                    const d = await r.json();
                    if (!r.ok) throw new Error((d && d.message) ? d.message : ('HTTP ' + r.status));
                    alert((d && d.message) ? d.message : 'Installering starter.');
                } catch (e) {
                    alert('Fejl ved installering: ' + e.message);
                } finally {
                    if (installBtn) {
                        installBtn.disabled = false;
                        installBtn.textContent = 'Installer nu';
                    }
                    refreshDashboardUpdateNotice().catch(() => {});
                }
            }

            function selectOrder(ordNo) {
                document.getElementById('orderInput').value = ordNo;
                prefetchAftercalcData(ordNo);
                orderListVisible = false;
                renderOrderList();
                searchOrder();
                window.scrollTo({ top: 0, behavior: 'auto' });
            }

            function goBackToList() {
                document.getElementById('result').innerHTML = '';
                orderListVisible = true;
                renderOrderList();
                setTimeout(() => {
                    const listEl = document.getElementById('orderList');
                    if (listEl) scrollToElementWithStickyOffset(listEl);
                }, 50);
            }

            function toggleSearchBox() {
                const searchBox = document.getElementById('searchBox');
                const collapseToggleBtn = document.getElementById('collapseToggleBtn');
                const collapseExpandBtn = document.getElementById('collapseExpandBtn');
                
                searchBox.classList.toggle('collapsed');
                if (searchBox.classList.contains('collapsed')) {
                    collapseToggleBtn.style.display = 'inline-block';
                    collapseExpandBtn.style.display = 'none';
                    collapseToggleBtn.textContent = '↗ Søg';
                } else {
                    collapseToggleBtn.style.display = 'none';
                    collapseExpandBtn.style.display = 'inline-block';
                }
                setTimeout(syncStickyOffsets, 0);
            }

            let uiBootstrapped = false;
            function bootstrapUiAfterLoad() {
                if (uiBootstrapped) return;
                uiBootstrapped = true;
                try {
                    const storedName = localStorage.getItem('afterkalk_logged_user_name');
                    if (storedName) {
                        loggedUserDisplayName = sanitizeDisplayName(storedName);
                    }
                } catch {}
                updateHeaderGreeting();
                showAccessGate();
                syncStickyOffsets();
                window.addEventListener('resize', syncStickyOffsets);
                const orderInput = document.getElementById('orderInput');
                const accessGateBtn = document.getElementById('accessGateBtn');
                if (orderInput) {
                    orderInput.addEventListener('keydown', function(event) {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            if (!accessGranted) {
                                submitAccessCode();
                                return;
                            }
                            searchOrder();
                        }
                    });
                    orderInput.addEventListener('input', function() {
                        clearTimeout(prefetchOrderDebounceTimer);
                        const ordNo = String(orderInput.value || '').trim();
                        if (!ordNo || ordNo.length < 4) return;
                        prefetchOrderDebounceTimer = setTimeout(() => {
                            prefetchAftercalcData(ordNo);
                        }, 260);
                    });
                }
                if (accessGateBtn) {
                    accessGateBtn.addEventListener('click', function(event) {
                        event.preventDefault();
                        submitAccessCode();
                    });
                }
                const accessGateInput = document.getElementById('accessGateInput');
                if (accessGateInput) {
                    accessGateInput.addEventListener('keydown', function(event) {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            submitAccessCode();
                        }
                    });
                }
                const accessGateUserInput = document.getElementById('accessGateUserInput');
                if (accessGateUserInput) {
                    accessGateUserInput.addEventListener('keydown', function(event) {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            const codeInput = document.getElementById('accessGateInput');
                            if (codeInput) {
                                codeInput.focus();
                                codeInput.select();
                            }
                        }
                    });
                }
                const sideMenuLoginInput = document.getElementById('sideMenuLoginInput');
                if (sideMenuLoginInput) {
                    sideMenuLoginInput.addEventListener('keydown', function(event) {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            submitAccessCodeFromSideMenu();
                        }
                    });
                }
                const sideMenuUserInput = document.getElementById('sideMenuUserInput');
                if (sideMenuUserInput) {
                    sideMenuUserInput.value = sanitizeDisplayName(loggedUserDisplayName);
                }
                updateReportOpenButtonState(Boolean(lastOrderReportHtml));
                refreshSideMenuAuthState();
                const orderListEl = document.getElementById('orderList');
                if (orderListEl) {
                    orderListEl.addEventListener('pointerdown', function(e) {
                        const sortHeader = e.target.closest('.order-sortable-header');
                        if (!sortHeader) return;
                        e.preventDefault();
                        const field = sortHeader.getAttribute('data-sort-field');
                        if (field) setOrderListSort(field);
                    });

                    orderListEl.addEventListener('click', function(e) {
                        const sortHeader = e.target.closest('.order-sortable-header');
                        if (sortHeader) {
                            return;
                        }
                        const refreshBtn = e.target.closest('.order-refresh-one-btn');
                        if (refreshBtn) {
                            e.preventDefault();
                            e.stopPropagation();
                            const ordNo = refreshBtn.getAttribute('data-ordno');
                            refreshSingleOrderCacheByOrdNo(ordNo, false, refreshBtn);
                            return;
                        }
                        if (e.target.closest('.order-refresh-cell')) {
                            e.preventDefault();
                            e.stopPropagation();
                            return;
                        }
                        const tr = e.target.closest('tr[data-ordno]');
                        if (tr) selectOrder(Number(tr.dataset.ordno));
                    });

                    orderListEl.addEventListener('mouseover', function(e) {
                        const tr = e.target.closest('tr[data-ordno]');
                        if (!tr) return;
                        const ordNo = String(tr.dataset.ordno || '').trim();
                        if (!ordNo) return;
                        prefetchAftercalcData(ordNo);
                    });
                }
            }

            // Soeg ved indlaesning hvis ordrenummer er i query string
            window.addEventListener('load', bootstrapUiAfterLoad, { once: true });
            document.addEventListener('DOMContentLoaded', bootstrapUiAfterLoad, { once: true });
            if (document.readyState === 'complete' || document.readyState === 'interactive') {
                setTimeout(bootstrapUiAfterLoad, 0);
            }
        </script>
    </body>
    </html>
    `);
});

const PORT = Number(process.env.PORT || 3000);
let startedServerPromise = null;
let scheduledRefreshTimer = null;

function startScheduledRefresh() {
    if (scheduledRefreshTimer) return;
    scheduledRefreshTimer = setInterval(() => {
        refreshOrderListCache(true)
            .then(() => {
                logEvent('Scheduled refresh completed (8h)');
            })
            .catch(err => {
                logEvent('ERROR scheduled refresh: ' + err.message);
            });
    }, BACKGROUND_WARM_INTERVAL_MS);
}

function ensureServerStarted() {
    if (startedServerPromise) return startedServerPromise;

    startedServerPromise = new Promise((resolve, reject) => {
        const server = app.listen(PORT, async () => {
            try {
                console.log('Server in ascolto su http://localhost:' + PORT);
                logEvent('Server started - smart preload phase beginning');

                // Try to load from persistent cache first for faster startup
                const cachedList = tryLoadOrderListFromCache();
                if (cachedList && cachedList.length > 0) {
                    orderListCache.data = cachedList;
                    orderListCache.loadedAt = Date.now();
                    logEvent('Cache primed from disk: ' + cachedList.length + ' orders ready');
                    
                    // Preload margins AND aftercalc details from disk (instant load)
                    const preloadOrdNos = cachedList.slice(0, STARTUP_MARGIN_WARM_COUNT).map(r => r.OrdNo);
                    preloadMarginsAndDetailsFromCache(preloadOrdNos);
                } else {
                    // No cache: defer DB fetch until dashboard/open-after-access flow.
                    logEvent('No disk cache found: DB warmup deferred until dashboard access');
                }

                logEvent('Cache primed: order list loaded and ready');
                startScheduledRefresh();
                resolve(server);
            } catch (err) {
                logEvent('WARNING cache warmup error (non-fatal): ' + err.message);
                resolve(server); // server is running — warmup error is not fatal
            }
        });

        server.on('error', reject);
    });

    return startedServerPromise;
}

if (require.main === module) {
    ensureServerStarted().catch(err => {
        logEvent('FATAL server startup error: ' + err.message);
        process.exit(1);
    });
}

module.exports = {
    app,
    ensureServerStarted
};
