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

const CACHE_TTL_AFTERCALC_MS        = 30 * 60 * 1000;  // 30 min
const CACHE_TTL_PRODUCTION_SUMMARY_MS = 30 * 60 * 1000;  // 30 min
const CACHE_TTL_LASER_METRICS_MS    = 60 * 60 * 1000;  // 60 min
const CACHE_TTL_ORDER_MARGIN_MS     = 30 * 60 * 1000;  // 30 min
const AFTERCALC_CACHE_KEY_PREFIX = 'aftercalc_v21_';
const ORDER_MARGIN_CACHE_KEY_PREFIX = 'order_margin_v21_';
const LEGACY_AFTERCALC_CACHE_KEY_PREFIXES = ['aftercalc_v20_', 'aftercalc_v19_', 'aftercalc_v18_', 'aftercalc_v17_', 'aftercalc_'];

const app = express();
app.use(express.json({ limit: '256kb' }));
app.use('/assets', express.static(path.join(__dirname, 'assets')));

// Read version from package.json
let pkgVersion = '1.0.0';
try {
    const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
    pkgVersion = pkg.version || '1.0.0';
} catch (e) {
    console.warn('Could not read package.json version');
}
const APP_VERSION = 'Gantech Efterkalkulation - v' + pkgVersion;

const { logEvent } = createLogger(APP_VERSION);
const ORDER_LIST_CACHE_TTL_MS = 10 * 60 * 1000;
const ORDER_LIST_MAX_ROWS = 150;
const ORDER_LIST_DAYS_BACK = 30;
const STARTUP_MARGIN_WARM_COUNT = ORDER_LIST_MAX_ROWS;
const BACKGROUND_WARM_INTERVAL_MS = 10 * 60 * 1000;
const BACKGROUND_AFTERCALC_WARM_COUNT = 60;  // Startup gate: warm top orders quickly, remaining orders prefetch on demand/background
const BACKGROUND_WARM_DELAY_MS = 10;  // Small stagger between queue submissions to keep DB stable
const MAX_DB_CALC_CONCURRENCY = 2;  // Controlled parallel DB calculations for faster startup warmup
const AFTERCALC_QUERY_TIMEOUT_MS = 45 * 1000;

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
        <title>Efterkalkulation</title>
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
            .header-brand { display:flex; align-items:center; gap:10px; min-width:0; flex:1; }
            .header-brand-logo { width:38px; height:38px; border-radius:8px; background:transparent; padding:5px; object-fit:contain; border:1px solid rgba(255,255,255,0.22); box-shadow:0 6px 16px rgba(4,16,30,0.25); filter:brightness(0) invert(1) contrast(1.08); flex-shrink:0; }
            .header-brand-text { min-width:0; font-size:20px; font-weight:800; letter-spacing:0.01em; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
            .header-status-badge { display: inline-block; font-size: 12px; font-weight: 700; color: #8a6d3b; background: #fff3cd; border: 1px solid #fff3cd; border-radius: 999px; padding: 4px 10px; white-space: nowrap; }
            #warmupBarWrap { display:none; align-items:center; gap:8px; background:rgba(0,0,0,0.15); border-radius:8px; padding:4px 10px; font-size:12px; color:#fff; white-space:nowrap; }
            #warmupBarWrap.active { display:flex; }
            #warmupBarBg { background:rgba(255,255,255,0.25); border-radius:999px; height:6px; width:110px; overflow:hidden; flex-shrink:0; }
            #warmupBarFill { background:#fff; height:100%; border-radius:999px; width:0%; transition:width 0.35s ease; }
            .search-box { background: linear-gradient(180deg, #ffffff 0%, #f4f9ff 100%); padding: 14px; margin-bottom: 18px; border-radius: 12px; box-shadow: 0 10px 24px rgba(15,53,96,0.10); border: 1px solid #d9e8f9; position: sticky; top: 58px; z-index: 1100; display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
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
            .filter-select { width: 180px; padding: 8px 10px; border: 1px solid #ddd; border-radius: 3px; background: #fff; }
            .section { background: white; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); padding: 20px; }
            .order-header { background: linear-gradient(135deg, #1976D2 0%, #1565C0 100%); color: white; padding: 25px; border-radius: 6px; margin-bottom: 25px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
            .order-header h2 { margin: 0 0 20px 0; font-size: 28px; font-weight: 700; }
            .order-header-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; }
            .order-header-item { display: flex; flex-direction: column; }
            .order-header-label { font-size: 12px; font-weight: 600; opacity: 0.9; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
            .order-header-value { font-size: 22px; font-weight: 700; color: #fff; }
            .invoice-status-badge { display: inline-block; font-size: 14px; font-weight: 700; padding: 4px 10px; border-radius: 6px; white-space: nowrap; }
            .status-in-production { background: rgba(255,255,255,0.15); color: #fff; border: 2px solid rgba(255,255,255,0.4); }
            .status-partial-invoiced { background: #e65100; color: #fff; }
            .status-fully-invoiced { background: #2e7d32; color: #fff; }
            h3 { color: var(--ink-900); margin-bottom: 14px; border-bottom: 2px solid #7eb1e6; padding-bottom: 10px; font-size: clamp(15px, 1.5vw, 19px); letter-spacing: 0.01em; }
            table { width: 100%; border-collapse: separate; border-spacing: 0; margin-top: 10px; background: #fff; border: 1px solid var(--line-soft); border-radius: var(--radius-m); overflow: hidden; box-shadow: 0 6px 16px rgba(15,53,96,0.05); }
            th, td { padding: 10px 11px; text-align: left; border-bottom: 1px solid #e8eff8; font-size: 13px; color: var(--text-900); }
            th { position: sticky; top: 0; z-index: 1; background: linear-gradient(180deg, #f4f9ff 0%, #eaf2ff 100%); font-weight: 700; color: var(--ink-900); text-transform: uppercase; font-size: 11px; letter-spacing: 0.04em; }
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
            .prodtp4-group { border: 1px solid #e5e5e5; border-radius: 4px; margin-bottom: 10px; overflow: hidden; }
            .prodtp4-header { background: linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%); padding: 10px 12px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; font-weight: 700; }
            .prodtp4-header:hover { background: linear-gradient(180deg, #eff6ff 0%, #e5efff 100%); }
            .prodtp4-label { color: #2b2b2b; }
            .prodtp4-subtotal { color: #1976D2; font-weight: 700; }
            .prodtp4-body { padding: 8px 12px 12px; }
            .po-total-row { margin-top: 10px; padding: 10px 12px; border-top: 1px solid #d8e5f7; font-weight: 800; text-align: right; background: linear-gradient(180deg, #fbfdff 0%, #f2f7ff 100%); color: var(--ink-900); }
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
            @media (max-width: 900px) {
                .modal-box { width: 99vw; max-height: 93vh; padding: 12px; }
                .modal-box th, .modal-box td { padding: 8px 6px; font-size: 13px; }
                .dashboard-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
                .omsaetning-filters { grid-template-columns:1fr 1fr; }
                .omsaetning-field.omsaetning-accounts-field,
                .omsaetning-field.omsaetning-threshold-field,
                .omsaetning-actions { grid-column:span 2; }
                .omsaetning-kpis { grid-template-columns:1fr; }
                .omsaetning-charts { grid-template-columns:1fr; }
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
                .dashboard-grid { grid-template-columns:1fr; }
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
            }
            .order-list-section { background: #fff; padding: 16px 20px; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            .order-list-section h3 { color: #333; margin-bottom: 12px; border-bottom: 2px solid #2196F3; padding-bottom: 8px; }
            .order-list-table { width: 100%; border-collapse: collapse; font-size: 13px; }
            .order-list-table th { background: #1565C0; color: #fff; padding: 8px 10px; text-align: left; }
            .order-list-table td { padding: 8px 10px; border-bottom: 1px solid #e0e0e0; cursor: pointer; }
            .order-list-table tr:hover td { background: #e3f2fd; }
            .note-badge { display:inline-flex; align-items:center; gap:3px; font-size:11px; font-weight:700; padding:2px 7px; border-radius:10px; border:1px solid; cursor:pointer; white-space:nowrap; }
            .note-badge.ok  { background:#e8f5e9; color:#1b5e20; border-color:#a5d6a7; }
            .note-badge.error { background:#ffebee; color:#b71c1c; border-color:#ef9a9a; }
            .note-badge.check { background:#fff8e1; color:#f57f17; border-color:#ffe082; }
            .note-badge.text { background:#f3e5f5; color:#4a148c; border-color:#ce93d8; }
            .note-badge.credit { background:#e3f2fd; color:#0d47a1; border-color:#90caf9; }
            .note-popup-overlay { position:fixed; inset:0; background:rgba(0,0,0,0.45); z-index:14000; display:flex; align-items:center; justify-content:center; }
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
            .access-gate-overlay { position: fixed; inset: 0; background: rgba(20, 26, 36, 0.72); display: none; align-items: center; justify-content: center; z-index: 12000; }
            .access-gate-box { width: min(430px, 92vw); background: #ffffff; border-radius: 10px; padding: 22px; box-shadow: 0 18px 42px rgba(0,0,0,0.28); }
            .access-gate-box h3 { margin: 0 0 10px 0; border: none; padding: 0; color: #1f2937; }
            .access-gate-box p { margin: 0 0 14px 0; color: #4b5563; }
            .access-gate-row { display: flex; gap: 8px; }
            .access-gate-row input { flex: 1; padding: 9px 10px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 16px; }
            .access-gate-row button { border: none; border-radius: 6px; background: #1565c0; color: #fff; font-weight: 700; padding: 9px 14px; cursor: pointer; }
            .access-gate-error { margin-top: 10px; min-height: 18px; color: #b71c1c; font-weight: 600; font-size: 13px; }
            .main-dashboard { display:none; margin-bottom:16px; }
            .dashboard-shell { position:relative; overflow:hidden; background:radial-gradient(1100px 360px at 8% -12%, rgba(22,101,192,0.19) 0%, rgba(22,101,192,0.03) 40%, transparent 70%), linear-gradient(160deg, #ffffff 0%, #f3f8ff 62%, #edf4ff 100%); border:1px solid #d7e6fb; border-radius:16px; box-shadow:0 14px 30px rgba(15,53,96,0.10); padding:16px; }
            .dashboard-shell::before { content:''; position:absolute; width:240px; height:240px; right:-95px; top:-105px; border-radius:50%; background:radial-gradient(circle at 30% 30%, rgba(86,164,255,0.24), rgba(86,164,255,0)); pointer-events:none; }
            .dashboard-shell::after { content:''; position:absolute; width:360px; height:110px; left:-90px; bottom:-60px; transform:rotate(-8deg); background:linear-gradient(90deg, rgba(15,53,96,0.00), rgba(15,53,96,0.08), rgba(15,53,96,0.00)); pointer-events:none; }
            .dashboard-head h2 { margin:0; color:#0f3560; font-size:26px; letter-spacing:0.01em; }
            .dashboard-head p { margin:6px 0 0 0; color:#4d6680; font-size:13px; }
            .dashboard-grid { margin-top:14px; display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; }
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
            .omsaetning-years { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:10px; }
            .omsaetning-year-btn { border:1px solid #c9ddf8; background:linear-gradient(180deg,#fff 0%,#edf4ff 100%); color:#0f3560; border-radius:999px; padding:6px 10px; font-size:12px; font-weight:700; cursor:pointer; }
            .omsaetning-year-btn.active { background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); color:#fff; border-color:#0f3560; }
            .omsaetning-filters { display:grid; grid-template-columns:160px 160px 1fr 220px auto; gap:10px; align-items:end; }
            .omsaetning-field label { display:block; font-size:12px; color:#3f5875; font-weight:700; margin-bottom:4px; }
            .omsaetning-field input, .omsaetning-field select { width:100%; border:1px solid #cfe0f7; border-radius:8px; padding:8px 10px; font-size:13px; }
            .omsaetning-threshold-grid { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
            .omsaetning-accounts-panel { border:1px solid #d3e4f8; border-radius:10px; background:#f9fcff; padding:8px; max-height:220px; overflow:auto; }
            .omsaetning-accounts-toolbar { display:flex; gap:6px; margin-bottom:8px; }
            .omsaetning-accounts-toolbar button { border:1px solid #c6dcf8; border-radius:999px; background:#fff; color:#0f3560; padding:4px 9px; cursor:pointer; font-size:11px; font-weight:700; }
            .omsaetning-accounts-list { display:flex; flex-direction:column; gap:4px; }
            .omsaetning-account-item { display:flex; align-items:center; gap:6px; padding:4px 6px; border-radius:6px; }
            .omsaetning-account-item:hover { background:#ecf4ff; }
            .omsaetning-account-item input { width:15px; height:15px; }
            .omsaetning-account-item span { font-size:12px; color:#244766; }
            .omsaetning-actions button { border:none; border-radius:999px; padding:8px 14px; font-weight:700; cursor:pointer; color:#fff; background:linear-gradient(180deg,#1565c0 0%,#0f3560 100%); }
            .omsaetning-kpis { margin-top:12px; display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:10px; }
            .omsaetning-kpi { border:1px solid #d6e7fb; border-radius:10px; background:#f8fbff; padding:10px; }
            .omsaetning-kpi .lbl { font-size:11px; font-weight:700; color:#4f6d8c; text-transform:uppercase; letter-spacing:0.03em; }
            .omsaetning-kpi .val { margin-top:4px; font-size:20px; font-weight:800; color:#0f3560; }
            .omsaetning-charts { margin-top:12px; display:grid; grid-template-columns:1fr 1fr; gap:10px; }
            .omsaetning-chart-card { border:1px solid #dbe8f9; border-radius:10px; background:linear-gradient(180deg,#ffffff 0%,#f6faff 100%); overflow:hidden; }
            .omsaetning-chart-head { display:flex; justify-content:space-between; align-items:center; gap:8px; padding:8px 10px; border-bottom:1px solid #dbe8f9; background:#f4f9ff; }
            .omsaetning-chart-title { font-size:12px; font-weight:800; color:#2f5475; }
            .omsaetning-chart-sub { font-size:11px; color:#5f7892; }
            .omsaetning-chart-body { padding:8px; overflow:auto; }
            .omsaetning-chart-svg { width:100%; min-width:680px; height:260px; display:block; }
            .omsaetning-legend { display:flex; flex-wrap:wrap; gap:6px 10px; margin-top:8px; }
            .omsaetning-legend-item { display:inline-flex; align-items:center; gap:6px; font-size:11px; color:#355675; }
            .omsaetning-legend-swatch { width:10px; height:10px; border-radius:3px; border:1px solid rgba(0,0,0,0.16); }
            .omsaetning-table-card { margin-top:12px; border:1px solid #dbe8f9; border-radius:10px; background:#fff; overflow:hidden; }
            .omsaetning-table-title { padding:8px 10px; font-size:12px; font-weight:700; color:#2f5475; background:#f4f9ff; border-bottom:1px solid #dbe8f9; }
            .omsaetning-table-wrap { margin-top:12px; overflow:auto; border:1px solid #dbe8f9; border-radius:10px; }
            .omsaetning-table { width:100%; border-collapse:collapse; min-width:760px; font-size:12px; }
            .omsaetning-table th { background:#1565c0; color:#fff; text-align:left; padding:8px 10px; position:sticky; top:0; z-index:1; }
            .omsaetning-table td { padding:7px 10px; border-bottom:1px solid #e6eef9; }
            .omsaetning-cell-right { text-align:right; }
            .omsaetning-status { display:inline-flex; align-items:center; border-radius:999px; padding:2px 8px; font-size:11px; font-weight:700; }
            .omsaetning-status.good { color:#1b5e20; background:#e8f5e9; border:1px solid #a5d6a7; }
            .omsaetning-status.mid { color:#8d6e00; background:#fff8e1; border:1px solid #ffe082; }
            .omsaetning-status.low { color:#b71c1c; background:#ffebee; border:1px solid #ef9a9a; }
            .omsaetning-empty { margin-top:10px; padding:10px; border:1px dashed #c7daef; border-radius:8px; color:#4f6d8c; background:#f8fbff; }
            #mainWorkspace { display:none; }
            .warning-flag { display:inline-flex; align-items:center; justify-content:center; margin-left:6px; font-size:14px; line-height:1; cursor:help; vertical-align:middle; }
            .allocation-flag { display:inline-flex; align-items:center; justify-content:center; margin-left:4px; color:#b26a00; font-size:16px; font-weight:700; line-height:1; cursor:help; vertical-align:middle; }
            .invoice-status-banner { margin: 0 0 10px 0; padding: 8px 10px; border-radius: 6px; font-size: 13px; font-weight: 600; }
            .invoice-status-banner.ok { background: #e8f5e9; color: #1b5e20; border: 1px solid #c8e6c9; }
            .invoice-status-banner.warn { background: #fff8e1; color: #8d6e00; border: 1px solid #ffe082; }
        </style>
    </head>
    <body>
        <div id="accessGateOverlay" class="access-gate-overlay" style="display:flex;">
            <div class="access-gate-box">
                <h3>Adgangskode</h3>
                <p>Indtast kode for at se ordreliste og detaljer.</p>
                <div class="access-gate-row">
                    <input id="accessGateInput" type="password" placeholder="Kode" autocomplete="off" />
                    <button id="accessGateBtn" type="button" onclick="submitAccessCode()">Åbn</button>
                </div>
                <div id="accessGateError" class="access-gate-error"></div>
            </div>
        </div>
        <div class="header-banner-wrapper">
            <button id="homeBtn" onclick="goToDashboard()" title="Tilbage til dashboard" style="background:rgba(255,255,255,0.18); border:none; border-radius:5px; color:#fff; font-size:20px; width:38px; height:38px; cursor:pointer; display:flex; align-items:center; justify-content:center; flex-shrink:0;">🏠</button>
            <div class="header-brand">
                <img class="header-brand-logo" src="/assets/brand/logo-gantech.png" alt="Gantech logo" />
                <span class="header-brand-text">${APP_VERSION}</span>
            </div>
            <div id="warmupBarWrap" title="Forberegner ordredata i baggrunden">
                <div id="warmupBarBg"><div id="warmupBarFill"></div></div>
                <span id="warmupBarText">Forberegner...</span>
            </div>
            <span class="header-status-badge" id="systemStatusBadge">System indlæser...</span>
        </div>
        <div class="container main-dashboard" id="mainDashboard">
            <section class="dashboard-shell">
                <div class="dashboard-head">
                    <h2>Gantech Operations Hub</h2>
                    <p>Vælg modul for at gå videre. Efterkalk er aktiv nu, de øvrige klargøres til næste fase.</p>
                </div>
                <div class="dashboard-grid">
                    <article class="dash-card">
                        <span class="dash-chip">Aktiv</span>
                        <h4>Efterkalkulation</h4>
                        <p>Ordreliste, kost, margin, produktion og rapportvisning.</p>
                        <button onclick="openModule('efterkalk')">Åbn Efterkalk</button>
                    </article>
                    <article class="dash-card">
                        <span class="dash-chip">Planlagt</span>
                        <h4>Belastning</h4>
                        <p>Kapacitetsbelastning, ressourcer, ordreflyt og planlægningsudsving.</p>
                        <button onclick="openModule('belastning')" disabled>Kommer snart</button>
                    </article>
                    <article class="dash-card">
                        <span class="dash-chip">Planlagt</span>
                        <h4>Omsætning</h4>
                        <p>Total omsætning, KPI-overblik og udvikling pr. periode/kunde.</p>
                        <button onclick="openModule('omsaetning')">Åbn Omsætning</button>
                    </article>
                    <article class="dash-card">
                        <span class="dash-chip">Planlagt</span>
                        <h4>Faktura</h4>
                        <p>Fakturaflow, status, opfølgning og administration af økonomidata.</p>
                        <button onclick="openModule('faktura')" disabled>Kommer snart</button>
                    </article>
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
                </div>
                <div id="omsaetningYears" class="omsaetning-years"></div>
                <div class="omsaetning-filters">
                    <div class="omsaetning-field">
                        <label for="omsaetningFraMonth">Fra måned</label>
                        <input id="omsaetningFraMonth" type="month" />
                    </div>
                    <div class="omsaetning-field">
                        <label for="omsaetningTilMonth">Til måned</label>
                        <input id="omsaetningTilMonth" type="month" />
                    </div>
                    <div class="omsaetning-field omsaetning-accounts-field">
                        <label for="omsaetningAccountSearch">Kontoer (multi)</label>
                        <input id="omsaetningAccountSearch" type="text" placeholder="Søg konto/navn..." oninput="filterOmsaetningAccounts()" />
                        <div class="omsaetning-accounts-panel">
                            <div class="omsaetning-accounts-toolbar">
                                <button type="button" onclick="setAllOmsaetningAccounts(true)">Alle</button>
                                <button type="button" onclick="setAllOmsaetningAccounts(false)">Ingen</button>
                            </div>
                            <div id="omsaetningAccountsList" class="omsaetning-accounts-list"></div>
                        </div>
                    </div>
                    <div class="omsaetning-field omsaetning-threshold-field">
                        <label>Soglie (Mio)</label>
                        <div class="omsaetning-threshold-grid">
                            <input id="omsaetningWarnThreshold" type="number" step="0.1" min="0" value="3.0" title="Under denne værdi markeres lav" />
                            <input id="omsaetningGoodThreshold" type="number" step="0.1" min="0" value="5.0" title="Over denne værdi markeres god" />
                        </div>
                    </div>
                    <div class="omsaetning-actions">
                        <button id="omsaetningLoadBtn" onclick="loadOmsaetningSummary()">Opdater</button>
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
                            <span class="omsaetning-chart-title">Omsætning pr. måned (stacked pr. konto)</span>
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
                    <div class="omsaetning-table-title">Månedstabel med soglier</div>
                    <div id="omsaetningThresholdTable" class="omsaetning-table-wrap" style="margin-top:0;border:none;border-radius:0;"></div>
                </div>
                <div id="omsaetningTableWrap" class="omsaetning-table-wrap" style="display:none;"></div>
                <div id="omsaetningEmpty" class="omsaetning-empty">Vælg perioder og konti, og tryk Opdater.</div>
            </section>
        </div>

        <div class="container" id="mainWorkspace">
            <div class="search-box" id="searchBox">
                <button id="collapseToggleBtn" onclick="toggleSearchBox()" style="display:none;" title="Åbn søgefelt og filtre">▼ Søg</button>
                <input type="number" id="orderInput" placeholder="Indtast ordrenummer..." style="display:none;" />
                <button onclick="searchOrder()" title="Aabn detaljer for ordrenummeret" style="display:none;">Søg</button>
                <select id="updateActionSelect" class="filter-select" onchange="handleUpdateActionSelection()" title="Vaelg hvad du vil opdatere">
                    <option value="">Opdater...</option>
                    <option value="order-cache">Ordre cache</option>
                    <option value="list">Liste</option>
                    <option value="program">Program</option>
                </select>
                <button class="mode-btn" onclick="toggleMarginMode()" title="Skift hvordan margin beregnes i visningen">Skift marginberegning</button>
                <button class="mode-btn" onclick="openOrderListPrintPreview()" title="Vis forhåndsvisning af den filtrerede ordreliste">Liste / PDF</button>
                <button id="listToggleBtn" class="list-toggle-btn" onclick="toggleOrderList()" title="Vis eller skjul kundelisten">Skjul kundeliste</button>
                <button id="clearCacheBtn" class="list-toggle-btn" onclick="clearAppCache()" style="background:#b71c1c !important;" title="DET TAGER LANG TID!!! Slet disk-cache og genindlaes data">Ryd cache</button>
                <select id="brugerFilterSelect" class="filter-select" onchange="setBrugerFilter()">
                    <option value="">Alle brugere</option>
                </select>
                <input type="text" id="customerFilterInput" class="filter-input" placeholder="Søg kunde i listen..." oninput="setOrderListFilter()" />
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
            const MARGIN_MAX_CONCURRENT = 2;
            const MARGIN_QUEUE_DELAY_MS = 120;
            const MARGIN_FETCH_TIMEOUT_MS = 20000;
            const MARGIN_PREFETCH_ROWS = 150;
            const ORDER_LIST_AUTO_REFRESH_MS = 2 * 60 * 1000;
            let lastOrderListCheckTime = 0;
            let lastOrderListRemoteTime = 0;
            let updateActionRunning = false;
            let omsaetningInitialized = false;
            let omsaetningAccounts = [];
            let omsaetningSelectedAccounts = new Set();
            let omsaetningActiveYear = null;

            function formatMio(value) {
                const numeric = Number(value || 0);
                return numeric.toLocaleString('da-DK', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
            }

            function formatMonthDa(dateValue) {
                if (!dateValue) return '-';
                const dt = new Date(dateValue);
                if (Number.isNaN(dt.getTime())) return String(dateValue);
                return dt.toLocaleDateString('da-DK', { month: 'short', year: 'numeric' });
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
                return {
                    fra: String(fromMeta.period),
                    til: String(exclusiveToDate.getFullYear() * 100 + (exclusiveToDate.getMonth() + 1))
                };
            }

            function renderOmsaetningYearChips(centerYear) {
                const wrap = document.getElementById('omsaetningYears');
                const fromEl = document.getElementById('omsaetningFraMonth');
                const toEl = document.getElementById('omsaetningTilMonth');
                if (!wrap || !fromEl || !toEl) return;

                const currentYear = Number(centerYear || new Date().getFullYear());
                const years = [currentYear - 3, currentYear - 2, currentYear - 1, currentYear, currentYear + 1];
                wrap.innerHTML = years.map(year => {
                    const activeCls = (omsaetningActiveYear === year) ? ' active' : '';
                    return '<button type="button" class="omsaetning-year-btn' + activeCls + '" onclick="setOmsaetningYear(' + year + ')">' +
                        (year === currentYear ? 'Nu ' : '') + escapeHtmlFE(String(year)) +
                        '</button>';
                }).join('');
            }

            function setOmsaetningYear(year) {
                const y = Number(year);
                if (!Number.isFinite(y)) return;
                const fromEl = document.getElementById('omsaetningFraMonth');
                const toEl = document.getElementById('omsaetningTilMonth');
                if (fromEl) fromEl.value = String(y) + '-01';
                if (toEl) toEl.value = String(y) + '-12';
                omsaetningActiveYear = y;
                renderOmsaetningYearChips(new Date().getFullYear());
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
            }

            function toggleOmsaetningAccount(inputEl) {
                if (!inputEl) return;
                const value = String(inputEl.getAttribute('data-accno') || '').trim();
                if (!value) return;
                if (inputEl.checked) omsaetningSelectedAccounts.add(value);
                else omsaetningSelectedAccounts.delete(value);
            }

            function getOmsaetningStatusClass(valueMio, warnThreshold, goodThreshold) {
                const n = Number(valueMio || 0);
                if (n >= goodThreshold) return 'good';
                if (n >= warnThreshold) return 'mid';
                return 'low';
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

            function renderOmsaetningCharts(rows) {
                const chartsWrap = document.getElementById('omsaetningChartsWrap');
                const stackedSvg = document.getElementById('omsaetningStackedChart');
                const trendSvg = document.getElementById('omsaetningTrendChart');
                const legend = document.getElementById('omsaetningLegend');
                if (!chartsWrap || !stackedSvg || !trendSvg || !legend) return;

                if (!Array.isArray(rows) || rows.length === 0) {
                    chartsWrap.style.display = 'none';
                    stackedSvg.innerHTML = '';
                    trendSvg.innerHTML = '';
                    legend.innerHTML = '';
                    return;
                }

                const monthMap = new Map();
                const accountOrder = [];
                const seenAccounts = new Set();
                for (const row of rows) {
                    const monthKey = String(row.date || '');
                    if (!monthMap.has(monthKey)) monthMap.set(monthKey, new Map());
                    const accountKey = String(row.acNo || '');
                    if (!seenAccounts.has(accountKey)) {
                        seenAccounts.add(accountKey);
                        accountOrder.push({ acNo: accountKey, name: String(row.name || '') });
                    }
                    const monthAcc = monthMap.get(monthKey);
                    monthAcc.set(accountKey, (monthAcc.get(accountKey) || 0) + Number(row.revenueMio || 0));
                }

                const monthKeys = Array.from(monthMap.keys()).sort((a, b) => String(a).localeCompare(String(b)));
                const monthlyTotals = monthKeys.map(key => {
                    const m = monthMap.get(key);
                    let t = 0;
                    for (const value of m.values()) t += Number(value || 0);
                    return t;
                });
                const maxTotal = Math.max(0.1, ...monthlyTotals);

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

                let stackedSvgHtml = '<g>';
                for (let i = 0; i <= 4; i++) {
                    const y = topPad + (innerHeight * i / 4);
                    const val = maxTotal * (1 - i / 4);
                    stackedSvgHtml += '<line x1="' + leftPad + '" y1="' + y + '" x2="' + (leftPad + innerWidth) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
                    stackedSvgHtml += '<text x="' + (leftPad - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="#5f7892">' + escapeHtmlFE(formatMio(val)) + '</text>';
                }

                monthKeys.forEach((monthKey, monthIndex) => {
                    const x = leftPad + monthIndex * (barWidth + barGap);
                    let stackedTop = topPad + innerHeight;
                    const values = monthMap.get(monthKey);
                    accountOrder.forEach((acc, accIndex) => {
                        const value = Number(values.get(acc.acNo) || 0);
                        if (value === 0) return;
                        const h = Math.max(1, (value / maxTotal) * innerHeight);
                        stackedTop -= h;
                        stackedSvgHtml += '<rect x="' + x + '" y="' + stackedTop + '" width="' + barWidth + '" height="' + h + '" fill="' + getOmsaetningColor(accIndex) + '" rx="2" />';
                    });
                    stackedSvgHtml += '<text x="' + (x + barWidth / 2) + '" y="' + (topPad + innerHeight + 14) + '" text-anchor="middle" font-size="10" fill="#47617c">' + escapeHtmlFE(formatMonthDa(monthKey)) + '</text>';
                });
                stackedSvgHtml += '</g>';

                stackedSvg.setAttribute('viewBox', '0 0 ' + viewWidth + ' ' + viewHeight);
                stackedSvg.innerHTML = stackedSvgHtml;

                legend.innerHTML = accountOrder.map((acc, idx) =>
                    '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:' + getOmsaetningColor(idx) + ';"></span>' +
                    escapeHtmlFE(String(acc.acNo)) + ' ' + escapeHtmlFE(acc.name || '') + '</span>'
                ).join('');

                const trendLeftPad = 42;
                const trendTopPad = 16;
                const trendBottomPad = 28;
                const trendHeight = 190;
                const trendInnerHeight = trendHeight - trendTopPad - trendBottomPad;
                const trendInnerWidth = Math.max(560, monthKeys.length * 54);
                const trendViewWidth = trendLeftPad + trendInnerWidth + 16;
                const trendViewHeight = trendHeight;
                const maxTrend = Math.max(0.1, ...monthlyTotals);

                let trendSvgHtml = '<g>';
                for (let i = 0; i <= 4; i++) {
                    const y = trendTopPad + (trendInnerHeight * i / 4);
                    trendSvgHtml += '<line x1="' + trendLeftPad + '" y1="' + y + '" x2="' + (trendLeftPad + trendInnerWidth) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
                }

                const points = monthKeys.map((monthKey, idx) => {
                    const x = trendLeftPad + (trendInnerWidth * (monthKeys.length === 1 ? 0.5 : (idx / (monthKeys.length - 1))));
                    const y = trendTopPad + trendInnerHeight - ((monthlyTotals[idx] / maxTrend) * trendInnerHeight);
                    return { x, y, monthKey, total: monthlyTotals[idx] };
                });
                const linePath = points.map((p, idx) => (idx === 0 ? 'M' : 'L') + p.x + ' ' + p.y).join(' ');
                const areaPath = linePath + ' L ' + points[points.length - 1].x + ' ' + (trendTopPad + trendInnerHeight) + ' L ' + points[0].x + ' ' + (trendTopPad + trendInnerHeight) + ' Z';
                trendSvgHtml += '<path d="' + areaPath + '" fill="rgba(21,101,192,0.12)" />';
                trendSvgHtml += '<path d="' + linePath + '" fill="none" stroke="#1565c0" stroke-width="3" />';
                points.forEach(p => {
                    trendSvgHtml += '<circle cx="' + p.x + '" cy="' + p.y + '" r="3.5" fill="#0f3560" />';
                });
                trendSvgHtml += '</g>';

                trendSvg.setAttribute('viewBox', '0 0 ' + trendViewWidth + ' ' + trendViewHeight);
                trendSvg.innerHTML = trendSvgHtml;
                chartsWrap.style.display = 'grid';
            }

            function showAccessGate() {
                const overlay = document.getElementById('accessGateOverlay');
                const input = document.getElementById('accessGateInput');
                const err = document.getElementById('accessGateError');
                if (!overlay) return;
                if (err) err.textContent = '';
                overlay.style.display = 'flex';
                setTimeout(() => {
                    if (input) input.focus();
                }, 30);
            }

            function hideAccessGate() {
                const overlay = document.getElementById('accessGateOverlay');
                if (!overlay) return;
                overlay.style.display = 'none';
            }

            function submitAccessCode() {
                const input = document.getElementById('accessGateInput');
                const err = document.getElementById('accessGateError');
                const btn = document.getElementById('accessGateBtn');
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
                        accessGranted = true;
                        hideAccessGate();
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
                loadOrderList(false);
                setTimeout(() => {
                    if (!orderListData || orderListData.length === 0) {
                        loadOrderList(true);
                    }
                }, 2500);
                startOrderListAutoRefresh();

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

                if (moduleKey === 'efterkalk') {
                    if (dashboard) dashboard.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'none';
                    if (workspace) workspace.style.display = 'block';
                    goBackToList();
                    return;
                }

                if (moduleKey === 'omsaetning') {
                    if (dashboard) dashboard.style.display = 'none';
                    if (workspace) workspace.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'block';
                    initializeOmsaetningIfNeeded();
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
                if (workspace) workspace.style.display = 'none';
                if (omsaetning) omsaetning.style.display = 'none';
                if (dashboard) dashboard.style.display = 'block';
                const detailModal = document.getElementById('orderDetailModal');
                const detailBody = document.getElementById('orderDetailModalBody');
                if (detailModal) detailModal.style.display = 'none';
                if (detailBody) detailBody.innerHTML = '';
                document.body.classList.remove('report-modal-open');
            }

            async function initializeOmsaetningIfNeeded() {
                if (omsaetningInitialized) return;
                omsaetningInitialized = true;

                const fraInput = document.getElementById('omsaetningFraMonth');
                const tilInput = document.getElementById('omsaetningTilMonth');
                const nowYear = new Date().getFullYear();
                if (fraInput && !fraInput.value) fraInput.value = String(nowYear - 1) + '-01';
                if (tilInput && !tilInput.value) tilInput.value = String(nowYear) + '-12';
                omsaetningActiveYear = nowYear;
                renderOmsaetningYearChips(nowYear);

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
                    omsaetningSelectedAccounts = new Set(accounts.map(acc => String(acc.acNo || '').trim()).filter(Boolean));
                    renderOmsaetningAccountsList();
                } catch (err) {
                    list.innerHTML = '<div class="omsaetning-account-item"><span>Fejl ved konti</span></div>';
                    console.error('loadOmsaetningAccounts failed:', err);
                }
            }

            async function loadOmsaetningSummary() {
                const loadBtn = document.getElementById('omsaetningLoadBtn');
                const empty = document.getElementById('omsaetningEmpty');
                const tableWrap = document.getElementById('omsaetningTableWrap');
                const thresholdWrap = document.getElementById('omsaetningThresholdWrap');
                const thresholdTable = document.getElementById('omsaetningThresholdTable');
                const chartsWrap = document.getElementById('omsaetningChartsWrap');
                const totalEl = document.getElementById('omsaetningTotalMio');
                const rowsEl = document.getElementById('omsaetningRowsCount');
                const periodsEl = document.getElementById('omsaetningPeriodsCount');
                const warnThresholdInput = document.getElementById('omsaetningWarnThreshold');
                const goodThresholdInput = document.getElementById('omsaetningGoodThreshold');

                const periodRange = buildOmsaetningPeriodRange();
                if (!periodRange) {
                    alert('Vælg gyldig periode (Fra måned skal være før eller lig Til måned).');
                    return;
                }

                const fra = periodRange.fra;
                const til = periodRange.til;

                const selected = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean);
                if (selected.length === 0) {
                    alert('Vælg mindst én konto.');
                    return;
                }

                const warnThreshold = Math.max(0, Number((warnThresholdInput && warnThresholdInput.value) || 3));
                const goodThreshold = Math.max(warnThreshold, Number((goodThresholdInput && goodThresholdInput.value) || 5));

                if (loadBtn) {
                    loadBtn.disabled = true;
                    loadBtn.textContent = 'Indlæser...';
                }

                try {
                    const query = new URLSearchParams({
                        fra,
                        til,
                        accounts: selected.join(',')
                    });

                    const response = await fetch('/omsaetning/summary?' + query.toString());
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    const rows = Array.isArray(payload.rows) ? payload.rows : [];
                    const uniquePeriods = new Set(rows.map(r => String(r.date || '')));

                    const monthTotals = new Map();
                    for (const row of rows) {
                        const monthKey = String(row.date || '');
                        const prev = monthTotals.get(monthKey) || 0;
                        monthTotals.set(monthKey, prev + Number(row.revenueMio || 0));
                    }

                    if (totalEl) totalEl.textContent = formatMio(payload.totalRevenueMio || 0);
                    if (rowsEl) rowsEl.textContent = formatCount(rows.length);
                    if (periodsEl) periodsEl.textContent = formatCount(uniquePeriods.size);

                    if (rows.length === 0) {
                        if (tableWrap) {
                            tableWrap.style.display = 'none';
                            tableWrap.innerHTML = '';
                        }
                        if (thresholdWrap) thresholdWrap.style.display = 'none';
                        if (thresholdTable) thresholdTable.innerHTML = '';
                        if (chartsWrap) chartsWrap.style.display = 'none';
                        if (empty) {
                            empty.style.display = 'block';
                            empty.textContent = 'Ingen data fundet for valgte filtre.';
                        }
                        return;
                    }

                    const sortedMonths = Array.from(monthTotals.entries()).sort((a, b) => String(a[0]).localeCompare(String(b[0])));
                    let thresholdHtml = '<table class="omsaetning-table"><thead><tr>' +
                        '<th>Måned</th><th class="omsaetning-cell-right">Omsætning (Mio)</th><th>Soglia</th>' +
                        '</tr></thead><tbody>';
                    for (const [monthKey, amountMio] of sortedMonths) {
                        const statusClass = getOmsaetningStatusClass(amountMio, warnThreshold, goodThreshold);
                        thresholdHtml += '<tr>' +
                            '<td>' + escapeHtmlFE(formatMonthDa(monthKey)) + '</td>' +
                            '<td class="omsaetning-cell-right">' + escapeHtmlFE(formatMio(amountMio)) + '</td>' +
                            '<td><span class="omsaetning-status ' + statusClass + '">' + escapeHtmlFE(getOmsaetningStatusLabel(statusClass)) + '</span></td>' +
                            '</tr>';
                    }
                    thresholdHtml += '</tbody></table>';

                    let html = '<table class="omsaetning-table"><thead><tr>' +
                        '<th>Måned</th><th>Konto</th><th>Navn</th><th style="text-align:right;">Omsætning (Mio)</th>' +
                        '</tr></thead><tbody>';

                    for (const row of rows) {
                        html += '<tr>' +
                            '<td>' + escapeHtmlFE(formatMonthDa(row.date)) + '</td>' +
                            '<td>' + escapeHtmlFE(String(row.acNo || '')) + '</td>' +
                            '<td>' + escapeHtmlFE(String(row.name || '')) + '</td>' +
                            '<td style="text-align:right;">' + escapeHtmlFE(formatMio(row.revenueMio || 0)) + '</td>' +
                            '</tr>';
                    }

                    html += '</tbody></table>';
                    if (thresholdTable) thresholdTable.innerHTML = thresholdHtml;
                    if (thresholdWrap) thresholdWrap.style.display = 'block';
                    if (tableWrap) {
                        tableWrap.innerHTML = html;
                        tableWrap.style.display = 'block';
                    }
                    renderOmsaetningCharts(rows);
                    if (empty) empty.style.display = 'none';
                } catch (err) {
                    if (tableWrap) {
                        tableWrap.style.display = 'none';
                        tableWrap.innerHTML = '';
                    }
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
                const badge = document.getElementById('systemStatusBadge');
                if (!badge) return;
                badge.textContent = text;
                badge.style.background = bgColor;
                badge.style.color = textColor;
                badge.style.borderColor = bgColor;
            }

            // Warmup progress bar polling
            let warmupPollTimer = null;
            function startWarmupPolling() {
                const wrap = document.getElementById('warmupBarWrap');
                const fill = document.getElementById('warmupBarFill');
                const txt  = document.getElementById('warmupBarText');
                if (!wrap) return;

                warmupPollTimer = setInterval(async () => {
                    try {
                        const r = await fetch('/warmup-status');
                        if (!r.ok) return;
                        const d = await r.json();

                        if (d.total === 0) {
                            wrap.classList.remove('active');
                            clearInterval(warmupPollTimer);
                            return;
                        }

                        wrap.classList.add('active');
                        fill.style.width = d.pct + '%';

                        if (d.running) {
                            txt.textContent = 'Forberegner ' + d.done + '/' + d.total + ' ordrer...';
                        } else {
                            txt.textContent = 'Klar! ' + d.loaded + ' nye + ' + d.cached + ' fra cache';
                            fill.style.width = '100%';
                            setTimeout(() => {
                                wrap.classList.remove('active');
                                clearInterval(warmupPollTimer);
                                warmupPollTimer = null;
                            }, 3000);
                        }
                    } catch(e) {
                        // ignore polling errors silently
                    }
                }, 800);
            }
            startWarmupPolling();

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
                const ordNo = document.getElementById('orderInput').value;
                if (ordNo) searchOrder();
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
                    return matchesText && matchesBruger;
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
                                        html += '<td><span class="prod-no-link" data-prodno="' + safeProdNo + '" data-ordno="' + prodOrder.ordNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="' + key + '" data-trinf2="' + safeTrInf2 + '" data-trinf4="' + safeTrInf4 + '" data-showallroutes="1" data-nofin="' + Number(line.NoFin || 0) + '" data-nestingcost="' + Number(line.NestingCost || 0) + '">' + safeProdLabel + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
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

                        if (!effectiveRoute) {
                            const encProdNo = encodeURIComponent(prodNo || '');
                            const fallbackResponse = await fetch('/nesting-detail/' + encodeURIComponent(effectiveOrdine) + '/' + encProdNo);
                            const fallbackRows = await fallbackResponse.json();
                            if (fallbackResponse.ok && Array.isArray(fallbackRows) && fallbackRows.length > 0) {
                                effectiveRoute = String(fallbackRows[0].TrInf4 || '').trim();
                            }
                        }

                        if (!effectiveRoute) {
                            body.innerHTML = '<div class="error">Fejl: TrInf4 (route) mangler paa den valgte linje.</div>';
                            return;
                        }

                        const endpoint = '/laser-route-metrics?ordine=' + encodeURIComponent(effectiveOrdine)
                            + '&route=' + encodeURIComponent(effectiveRoute)
                            + '&prodNo=' + encodeURIComponent(prodNo || '')
                            + '&showAllRoutes=' + (showAllRoutes ? '1' : '0')
                            + (currentSalesOrderGr4 === 3 ? '&gr4=3' : '');
                        const response = await fetch(endpoint);
                        const data = await response.json();
                        if (!response.ok || data.error) {
                            body.innerHTML = '<div class="error">Fejl: ' + (data.error || 'Uventet fejl') + '</div>';
                            return;
                        }

                        let finalData = data;
                        let usedProdFilter = Boolean(prodNo);

                        if (usedProdFilter && Array.isArray(data.products) && data.products.length === 0) {
                            const fallbackEndpoint = '/laser-route-metrics?ordine=' + encodeURIComponent(effectiveOrdine)
                                + '&route=' + encodeURIComponent(effectiveRoute)
                                + (currentSalesOrderGr4 === 3 ? '&gr4=3' : '');
                            const fallbackResponse = await fetch(fallbackEndpoint);
                            const fallbackData = await fallbackResponse.json();
                            if (fallbackResponse.ok && !fallbackData.error && Array.isArray(fallbackData.products) && fallbackData.products.length > 0) {
                                finalData = fallbackData;
                                usedProdFilter = false;
                            }
                        }

                        const s = finalData.summary || {};
                        const products = Array.isArray(finalData.products) ? finalData.products : [];
                        const formatNullable = (value, suffix = '') => {
                            return value === null || value === undefined
                                ? 'NULL'
                                : (formatNumber(value) + suffix);
                        };

                        if (!products.length) {
                            body.innerHTML = '<div>Ingen faerdigvarer (TrTp=7) fundet for valgt rute.</div>';
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
                            // In showAllRoutes mode use per-route API data; clickedNestCost/clickedNoFin are order-level totals and must not override per-route values.
                            const hasClickedNestCost = !showAllRoutes && isClickedProd && clickedNestingCostNum > 0;
                            const noFin = (hasClickedNestCost && clickedNoFinNum > 0) ? clickedNoFinNum : routeNoFin;
                            const hintedNestCost = getLaserNestCostHint(effectiveOrdine, prodNoForCost);
                            const costPerPiece = hasClickedNestCost
                                ? clickedNestingCostNum
                                : ((rowProduct && rowProduct.CostoPerPezzo !== null && rowProduct.CostoPerPezzo !== undefined)
                                    ? rowProduct.CostoPerPezzo
                                    : hintedNestCost);
                            const noFinNum = Number(noFin || 0);
                            const expectedNum = Number(expected || 0);
                            const effectiveNum = Number(effective || 0);
                            const totalCost = hasClickedNestCost
                                ? (noFinNum > 0 ? (noFinNum * Number(costPerPiece || 0)) : null)
                                : ((rowProduct && rowProduct.QuotaCosto !== null && rowProduct.QuotaCosto !== undefined)
                                    ? rowProduct.QuotaCosto
                                : ((costPerPiece === null || costPerPiece === undefined || noFin === null || noFin === undefined)
                                    ? null
                                    : (noFinNum * Number(costPerPiece || 0))));
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
                            html += '<td>' + formatNullable(costPerPiece) + '</td>';
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
                        html += '<div class="summary-box" style="margin-top:12px;">'
                            + '<div><strong>Samlet L-kost (NestKost):</strong> ' + formatNumber(totalLaserCost) + ' DKK</div>'
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
            document.addEventListener('click', handleImagePreviewClick);
            document.addEventListener('click', handleDrawingOpenClick);
            document.addEventListener('click', handlePreviewImageZoom);
            document.addEventListener('keydown', function(event) {
                if (event.key === 'Escape') {
                    closeImageLightbox();
                    closeCompactImageModal();
                    closeOversigtModal();
                }
            });
            // Inside modal content (document listener is blocked by modal stopPropagation).
            const summaryModalBodyEl = document.getElementById('summaryModalBody');
            if (summaryModalBodyEl) {
                summaryModalBodyEl.addEventListener('click', handleProdNoClick);
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
                            html += '<td>' + formatNumber(line.NestingCost || 0) + '</td>';
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

                    await loadOrderList(false);
                    if (openAfter && Number.isFinite(ordNoNum)) {
                        await searchOrder();
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
                        clickedBtn.textContent = 'Opdater';
                    }
                }
            }

            async function clearAppCache() {
                const confirmed = confirm('Er du sikker? Dette vil slette alt cache og tage lang tid at genindlæse data.');
                if (!confirmed) return;
                
                const btn = document.getElementById('clearCacheBtn');
                if (btn) { btn.disabled = true; btn.textContent = 'Rydder...'; }
                try {
                    const r = await fetch('/cache-clear', { method: 'POST' });
                    if (!r.ok) throw new Error('HTTP ' + r.status);
                    const d = await r.json();
                    alert('Cache ryddet: ' + (d.deleted || 0) + ' filer slettet.');
                } catch (e) {
                    alert('Fejl ved cache-rydning: ' + e.message);
                } finally {
                    if (btn) { btn.disabled = false; btn.textContent = 'Ryd cache'; }
                }
            }

            async function checkDesktopUpdateNow() {
                const btn = document.getElementById('checkUpdateBtn');
                if (btn) { btn.disabled = true; btn.textContent = 'Tjekker...'; }

                try {
                    const r = await fetch('/desktop-update-check', { method: 'POST' });
                    const d = await r.json();
                    if (!r.ok) throw new Error((d && d.message) ? d.message : ('HTTP ' + r.status));

                    if (d.status === 'available' && d.version) {
                        alert('Ny version fundet: ' + d.version + '. Den downloades i baggrunden.');
                    } else if (d.status === 'up-to-date') {
                        alert('Du har allerede den nyeste version.');
                    } else if (d.status === 'busy') {
                        alert('Opdateringskontrol kører allerede. Prøv igen om lidt.');
                    } else if (d.status === 'checking') {
                        alert('Opdateringskontrol startet. Vent lidt og prøv igen.');
                    } else {
                        alert((d && d.message) ? d.message : 'Opdateringskontrol sendt.');
                    }
                } catch (e) {
                    alert('Fejl ved opdateringskontrol: ' + e.message);
                } finally {
                    if (btn) { btn.disabled = false; btn.textContent = 'Tjek opdatering nu'; }
                }
            }

            async function handleUpdateActionSelection() {
                const select = document.getElementById('updateActionSelect');
                if (!select) return;

                const action = String(select.value || '');
                if (!action) return;

                if (updateActionRunning) {
                    alert('En opdatering koerer allerede. Vent venligst.');
                    select.value = '';
                    return;
                }

                updateActionRunning = true;
                select.disabled = true;
                try {
                    if (action === 'order-cache') {
                        await refreshSingleOrderCache();
                    } else if (action === 'list') {
                        await refreshOrderList();
                    } else if (action === 'program') {
                        await checkDesktopUpdateNow();
                    }
                } finally {
                    select.disabled = false;
                    select.value = '';
                    updateActionRunning = false;
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
            }

            let uiBootstrapped = false;
            function bootstrapUiAfterLoad() {
                if (uiBootstrapped) return;
                uiBootstrapped = true;
                showAccessGate();
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
                updateReportOpenButtonState(Boolean(lastOrderReportHtml));
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
                logEvent('Scheduled refresh completed (10 min)');
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

                    // Warm up margins in background (will check disk first, then refresh if needed)
                    warmMarginsInBackground(preloadOrdNos);
                    
                    // Refresh from DB in background (don't block startup)
                    refreshOrderListCache(true).catch(err => {
                        logEvent('WARNING: background DB refresh failed: ' + err.message);
                    });
                } else {
                    // No cache: load from DB (fresh startup)
                    await refreshOrderListCache(true);
                    logEvent('Cache primed from database (first startup)');
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
