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
const AFTERCALC_CACHE_KEY_PREFIX = 'aftercalc_v14_';
const ORDER_MARGIN_CACHE_KEY_PREFIX = 'order_margin_v14_';

const app = express();
app.use(express.json({ limit: '256kb' }));

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
const STARTUP_MARGIN_WARM_COUNT = 150;
const BACKGROUND_WARM_INTERVAL_MS = 10 * 60 * 1000;
const BACKGROUND_AFTERCALC_WARM_COUNT = 150;
const BACKGROUND_WARM_DELAY_MS = 50;  // 50ms delay tra calcoli warmup (faster for startup)
const MAX_DB_CALC_CONCURRENCY = 2;

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

async function getOrComputeAftercalc(ordNo, options = {}) {
    const priority = options.priority || 'normal';
    const key = Number(ordNo);
    if (!Number.isFinite(key)) {
        throw new Error('Ordrenummer ugyldigt');
    }

    const cacheKey = AFTERCALC_CACHE_KEY_PREFIX + key;
    const cached = diskCache.get(cacheKey);
    if (cached) {
        return cached;
    }

    let computePromise = afterCalcInFlight.get(key);
    if (!computePromise) {
        computePromise = runWithDbCalcLimit(async () => {
            const data = await getAfterCalc(key);
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

    if (!forceRefresh && orderMarginInFlight.has(key)) {
        return orderMarginInFlight.get(key);
    }

    const computePromise = runWithDbCalcLimit(async () => {
        const data = await getAfterCalc(key);
        if (data.error) {
            throw new Error(data.error);
        }

        const marginInfo = {
            ordNo: key,
            totalRevenue: Number(data.summary.totalRevenue || 0),
            totalCost: Number(data.summary.totalCost || 0),
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
    logEvent('WARM-MARGIN: queueing ' + ordNos.length + ' orders');
    for (const ordNo of ordNos) {
        const numericOrdNo = Number(ordNo);
        if (!Number.isFinite(numericOrdNo)) continue;
        getOrComputeOrderMargin(numericOrdNo).catch(() => {});
    }
}

async function warmAftercalcInBackground(ordNos, sourceLabel, maxDelayMs = BACKGROUND_WARM_DELAY_MS) {
    if (backgroundAftercalcWarmRunning) {
        logEvent('WARM-AFTERCALC: skipped (' + sourceLabel + ') because previous run is still active');
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

    try {
        for (const ordNo of ordNos) {
            const numericOrdNo = Number(ordNo);
            if (!Number.isFinite(numericOrdNo)) continue;
            total += 1;

            if (diskCache.get(AFTERCALC_CACHE_KEY_PREFIX + numericOrdNo)) {
                alreadyCached += 1;
                warmupProgress.cached += 1;
                warmupProgress.current = numericOrdNo;
                continue;
            }

            warmupProgress.current = numericOrdNo;
            try {
                await getOrComputeAftercalc(numericOrdNo, { priority: 'normal' });
                warmed += 1;
                warmupProgress.loaded += 1;
            } catch (err) {
                failed += 1;
                warmupProgress.failed += 1;
                logEvent('WARM-AFTERCALC ERROR ordNo=' + numericOrdNo + ': ' + err.message);
            }

            if (maxDelayMs > 0) {
                await new Promise(resolve => setTimeout(resolve, maxDelayMs));
            }
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
            const detailsCached = diskCache.get(AFTERCALC_CACHE_KEY_PREFIX + key);
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
            const warmAftercalcOrdNos = rows.slice(0, BACKGROUND_AFTERCALC_WARM_COUNT).map(r => r.OrdNo);
            warmAftercalcInBackground(warmAftercalcOrdNos, force ? 'refresh-force' : 'refresh-auto', 50);
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
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { font-family: Arial, sans-serif; background: #f5f5f5; padding: 20px; }
            .container { max-width: 1200px; margin: 0 auto; }
            .header-banner-wrapper { background: #c0392b; color: #fff; font-weight: 800; font-size: 25px; padding: 10px 12px; border-radius: 6px; margin-bottom: 20px; letter-spacing: 0.2px; width: 100%; position: sticky; top: 0; z-index: 1200; display: flex; align-items: center; justify-content: space-between; gap: 12px; }
            .header-status-badge { display: inline-block; font-size: 12px; font-weight: 700; color: #8a6d3b; background: #fff3cd; border: 1px solid #fff3cd; border-radius: 999px; padding: 4px 10px; white-space: nowrap; }
            #warmupBarWrap { display:none; align-items:center; gap:8px; background:rgba(0,0,0,0.15); border-radius:8px; padding:4px 10px; font-size:12px; color:#fff; white-space:nowrap; }
            #warmupBarWrap.active { display:flex; }
            #warmupBarBg { background:rgba(255,255,255,0.25); border-radius:999px; height:6px; width:110px; overflow:hidden; flex-shrink:0; }
            #warmupBarFill { background:#fff; height:100%; border-radius:999px; width:0%; transition:width 0.35s ease; }
            .search-box { background: #fff; padding: 20px; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); position: sticky; top: 58px; z-index: 1100; display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
            .search-box.collapsed { padding: 8px 12px; height: 36px; }
            .search-box.collapsed > * { display: none; }
            .search-box.collapsed > #collapseToggleBtn { display: inline-block; }
            #collapseToggleBtn { background: #1976d2; color: #fff; border: none; padding: 8px 12px; border-radius: 3px; cursor: pointer; font-weight: 600; font-size: 12px; }
            .build-badge { display: inline-block; font-size: 12px; color: #444; background: #f1f1f1; border: 1px solid #ddd; border-radius: 4px; padding: 4px 8px; }
            .build-banner { display: none; }
            .search-box input { padding: 8px 12px; font-size: 14px; width: 200px; border: 1px solid #ddd; border-radius: 3px; }
            .search-box button { padding: 8px 16px; background: #4CAF50; color: white; border: none; border-radius: 3px; cursor: pointer; margin-left: 10px; }
            .mode-btn { background: #0d47a1 !important; }
            .list-toggle-btn { background: #455a64 !important; color: #fff !important; }
            .filter-input { width: 260px !important; margin-left: 10px; }
            .filter-select { width: 180px; padding: 8px 10px; border: 1px solid #ddd; border-radius: 3px; background: #fff; }
            .section { background: white; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); padding: 20px; }
            .order-header { background: linear-gradient(135deg, #1976D2 0%, #1565C0 100%); color: white; padding: 25px; border-radius: 6px; margin-bottom: 25px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
            .order-header h2 { margin: 0 0 20px 0; font-size: 28px; font-weight: 700; }
            .order-header-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; }
            .order-header-item { display: flex; flex-direction: column; }
            .order-header-label { font-size: 12px; font-weight: 600; opacity: 0.9; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
            .order-header-value { font-size: 22px; font-weight: 700; color: #fff; }
            h3 { color: #333; margin-bottom: 15px; border-bottom: 2px solid #2196F3; padding-bottom: 10px; }
            table { width: 100%; border-collapse: collapse; margin-top: 10px; }
            th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
            th { background: #f0f0f0; font-weight: bold; }
            tr:hover { background: #fafafa; }
            .summary-row { font-weight: bold; background: #f9f9f9; }
            .summary-box { background: #e8f5e9; padding: 15px; border-radius: 4px; margin-top: 15px; }
            .summary-box div { margin: 8px 0; font-size: 14px; }
            .summary-box .total { font-size: 18px; color: #2196F3; font-weight: bold; }
            .margin-positive { color: green; }
            .margin-negative { color: red; }
            .error { color: red; padding: 20px; background: #ffebee; border-radius: 4px; }
            .loading { color: #666; padding: 20px; }
            .prod-link { color: #1976D2; text-decoration: underline; cursor: pointer; }
            .prod-link:hover { color: #0D47A1; }
            .po-highlight { box-shadow: 0 0 0 3px #90CAF9; }
            .prodtp4-group { border: 1px solid #e5e5e5; border-radius: 4px; margin-bottom: 10px; overflow: hidden; }
            .prodtp4-header { background: #f7f9fc; padding: 10px 12px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; font-weight: 600; }
            .prodtp4-header:hover { background: #eef4fb; }
            .prodtp4-label { color: #2b2b2b; }
            .prodtp4-subtotal { color: #1976D2; font-weight: 700; }
            .prodtp4-body { padding: 8px 12px 12px; }
            .po-total-row { margin-top: 10px; padding: 10px 12px; border-top: 1px solid #ddd; font-weight: 700; text-align: right; background: #fafafa; }
            .prodtp4-hint { color: #555; margin: 6px 0 10px; font-size: 13px; }
            .main-product-box { background: #eef6ff; border: 2px solid #90caf9; border-radius: 6px; padding: 10px 12px; margin: 8px 0 12px; }
            .main-product-box .value { font-size: 20px; font-weight: 800; color: #0d47a1; margin-top: 3px; }
            .inline-link { color: #1565c0; text-decoration: underline; cursor: pointer; }
            .inline-link:hover { color: #0d47a1; }
            .prod-no-link { color: #1565c0; text-decoration: underline; cursor: pointer; }
            .prod-no-link:hover { color: #0d47a1; }
            .modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.35); display: none; align-items: center; justify-content: center; z-index: 9999; }
            .modal-box { width: min(1280px, 96vw); max-height: 88vh; overflow: auto; background: #fff; border-radius: 8px; box-shadow: 0 10px 30px rgba(0,0,0,0.2); padding: 16px; }
            .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
            .modal-header-left { display: flex; align-items: center; gap: 8px; }
            .modal-content-wrap { display: flex; gap: 16px; align-items: flex-start; }
            #summaryModalBody { flex: 1; min-width: 0; }
            .modal-back { border: none; background: #efefef; border-radius: 4px; padding: 6px 10px; cursor: pointer; font-weight: 700; }
            .modal-back.hidden { display: none; }
            .modal-close { border: none; background: #efefef; border-radius: 4px; padding: 6px 10px; cursor: pointer; }
            .modal-loading { color: #666; padding: 8px 0; }
            .summary-image-panel { width: min(380px, 32vw); min-width: 320px; max-height: 76vh; overflow: auto; border-left: 1px solid #e0e0e0; padding-left: 16px; position: sticky; top: 0; background: #fff; }
            .summary-image-panel.hidden { display: none; }
            .laser-summary-layout { display: flex; gap: 12px; align-items: flex-start; }
            .laser-image-panel { width: min(290px, 28vw); min-width: 220px; max-height: 68vh; overflow: auto; border: 1px solid #e0e0e0; border-radius: 6px; padding: 10px; position: sticky; top: 12px; background: #fff; }
            .laser-image-panel.hidden { display: none; }
            .summary-image-panel-header { display: flex; justify-content: space-between; align-items: center; gap: 10px; margin-bottom: 12px; }
            .summary-image-panel-title { font-size: 16px; font-weight: 700; color: #1f2937; }
            .summary-image-close { border: none; background: #efefef; border-radius: 4px; padding: 6px 10px; cursor: pointer; }
            .image-preview-btn { padding: 6px 10px; border: none; border-radius: 4px; background: #1565c0; color: #fff; cursor: pointer; font-size: 12px; }
            .image-preview-btn:hover { background: #0d47a1; }
            .image-preview-gallery { display: grid; gap: 12px; }
            .image-preview-card { border: 1px solid #e5e7eb; border-radius: 6px; padding: 10px; background: #fafafa; }
            .image-preview-card img { display: block; width: 100%; max-height: 240px; object-fit: contain; background: #fff; border-radius: 4px; border: 1px solid #e5e7eb; cursor: zoom-in; }
            .image-preview-label { font-size: 12px; font-weight: 700; color: #374151; margin-bottom: 8px; }
            .image-preview-path { font-size: 11px; color: #6b7280; word-break: break-all; margin-top: 8px; }
            .image-preview-empty { font-size: 13px; color: #6b7280; padding: 8px 0; }
            .image-lightbox { position: fixed; inset: 0; background: rgba(17, 24, 39, 0.88); display: flex; align-items: center; justify-content: center; padding: 24px; z-index: 11000; }
            .image-lightbox.hidden { display: none; }
            .image-lightbox-dialog { width: min(1200px, 96vw); max-height: 92vh; background: #111827; color: #f9fafb; border-radius: 10px; box-shadow: 0 18px 40px rgba(0,0,0,0.35); padding: 16px; display: flex; flex-direction: column; gap: 12px; }
            .image-lightbox-header { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
            .image-lightbox-title { font-size: 15px; font-weight: 700; color: #f9fafb; }
            .image-lightbox-close { border: none; background: rgba(255,255,255,0.12); color: #fff; border-radius: 4px; padding: 6px 10px; cursor: pointer; }
            .image-lightbox-close:hover { background: rgba(255,255,255,0.2); }
            .image-lightbox-body { display: flex; align-items: center; justify-content: center; min-height: 0; overflow: auto; }
            .image-lightbox-body img { display: block; max-width: 100%; max-height: calc(92vh - 110px); object-fit: contain; border-radius: 6px; background: #fff; }
            .image-lightbox-path { font-size: 12px; color: #d1d5db; word-break: break-all; }
            @media (max-width: 900px) {
                .modal-box { width: 98vw; max-height: 92vh; padding: 12px; }
                .modal-box th, .modal-box td { padding: 8px 6px; font-size: 13px; }
                .modal-content-wrap { flex-direction: column; }
                .summary-image-panel { width: 100%; min-width: 0; max-height: none; border-left: none; border-top: 1px solid #e0e0e0; padding-left: 0; padding-top: 12px; }
                .laser-summary-layout { flex-direction: column; }
                .laser-image-panel { width: 100%; min-width: 0; max-height: none; position: static; }
                .image-lightbox { padding: 12px; }
                .image-lightbox-dialog { width: 100%; max-height: 96vh; padding: 12px; }
                .image-lightbox-body img { max-height: calc(96vh - 110px); }
            }
            .order-list-section { background: #fff; padding: 16px 20px; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            .order-list-section h3 { color: #333; margin-bottom: 12px; border-bottom: 2px solid #2196F3; padding-bottom: 8px; }
            .order-list-table { width: 100%; border-collapse: collapse; font-size: 13px; }
            .order-list-table th { background: #1565C0; color: #fff; padding: 8px 10px; text-align: left; }
            .order-list-table td { padding: 8px 10px; border-bottom: 1px solid #e0e0e0; cursor: pointer; }
            .order-list-table tr:hover td { background: #e3f2fd; }
            .access-gate-overlay { position: fixed; inset: 0; background: rgba(20, 26, 36, 0.72); display: none; align-items: center; justify-content: center; z-index: 12000; }
            .access-gate-box { width: min(430px, 92vw); background: #ffffff; border-radius: 10px; padding: 22px; box-shadow: 0 18px 42px rgba(0,0,0,0.28); }
            .access-gate-box h3 { margin: 0 0 10px 0; border: none; padding: 0; color: #1f2937; }
            .access-gate-box p { margin: 0 0 14px 0; color: #4b5563; }
            .access-gate-row { display: flex; gap: 8px; }
            .access-gate-row input { flex: 1; padding: 9px 10px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 16px; }
            .access-gate-row button { border: none; border-radius: 6px; background: #1565c0; color: #fff; font-weight: 700; padding: 9px 14px; cursor: pointer; }
            .access-gate-error { margin-top: 10px; min-height: 18px; color: #b71c1c; font-weight: 600; font-size: 13px; }
            .warning-flag { display:inline-flex; align-items:center; justify-content:center; margin-left:6px; font-size:14px; line-height:1; cursor:help; vertical-align:middle; }
        </style>
    </head>
    <body>
        <div id="accessGateOverlay" class="access-gate-overlay">
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
            <button id="homeBtn" onclick="goBackToList()" title="Tilbage til ordreliste" style="background:rgba(255,255,255,0.18); border:none; border-radius:5px; color:#fff; font-size:20px; width:38px; height:38px; cursor:pointer; display:flex; align-items:center; justify-content:center; flex-shrink:0;">🏠</button>
            <span style="flex:1;">🔷 ${APP_VERSION}</span>
            <div id="warmupBarWrap" title="Forberegner ordredata i baggrunden">
                <div id="warmupBarBg"><div id="warmupBarFill"></div></div>
                <span id="warmupBarText">Forberegner...</span>
            </div>
            <span class="header-status-badge" id="systemStatusBadge">System indlæser...</span>
        </div>
        <div class="container">
            <div class="search-box" id="searchBox">
                <button id="collapseToggleBtn" onclick="toggleSearchBox()" style="display:none;" title="Åbn søgefelt og filtre">▼ Søg</button>
                <input type="number" id="orderInput" placeholder="Indtast ordrenummer..." />
                <button onclick="searchOrder()" title="Aabn detaljer for ordrenummeret">Søg</button>
                <select id="updateActionSelect" class="filter-select" onchange="handleUpdateActionSelection()" title="Vaelg hvad du vil opdatere">
                    <option value="">Opdater...</option>
                    <option value="order-cache">Ordre cache</option>
                    <option value="list">Liste</option>
                    <option value="program">Program</option>
                </select>
                <button class="mode-btn" onclick="toggleMarginMode()" title="Skift hvordan margin beregnes i visningen">Skift marginberegning</button>
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

            function openDrawingPdf(pathValue) {
                const value = String(pathValue || '').trim();
                if (!value) return;
                fetch('/open-drawing', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ path: value })
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

            let currentMarginMode = 'classic';
            let orderListData = [];
            let orderListVisible = true;
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
            let summaryModalHistory = [];
            let summaryImageRegistry = {};
            let summaryImageRegistryCounter = 0;
            const ACCESS_CODE = '12345';
            let accessGranted = false;
            const MARGIN_MAX_CONCURRENT = 2;
            const MARGIN_QUEUE_DELAY_MS = 120;
            const MARGIN_FETCH_TIMEOUT_MS = 20000;
            const MARGIN_PREFETCH_ROWS = ${ORDER_LIST_MAX_ROWS};
            const ORDER_LIST_AUTO_REFRESH_MS = 2 * 60 * 1000;
            let lastOrderListCheckTime = 0;
            let lastOrderListRemoteTime = 0;
            let updateActionRunning = false;

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
                const value = input ? String(input.value || '').trim() : '';
                if (value !== ACCESS_CODE) {
                    if (err) err.textContent = 'Forkert kode.';
                    if (input) {
                        input.select();
                        input.focus();
                    }
                    return;
                }

                accessGranted = true;
                hideAccessGate();
                initializeAfterAccess();
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
                    searchOrder();
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
            }

            function openSummaryImagePanel(imageKey, preferredPanelId) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const laserPanelWrap = document.getElementById('laserOrderSummaryPanel');
                const laserPanel = document.getElementById('laserImagePanel');
                const summaryPanel = document.getElementById('summaryImagePanel');
                const isLaserVisible = laserPanelWrap && laserPanelWrap.style.display !== 'none';
                let panel = null;
                if (preferredPanelId === 'laserImagePanel') {
                    panel = laserPanel;
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

            function getOrderMarginHtml(ordNo) {
                const marginState = getMarginState(ordNo);
                let marginHtml = '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>';
                if (marginState && marginState.status === 'loading') {
                    marginHtml = '<span style="background:#546e7a; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">...</span>';
                } else if (marginState && marginState.status === 'success') {
                    const margin = calculateOrderMarginPercent(marginState.totalRevenue || 0, marginState.totalCost || 0).toFixed(2);
                    marginHtml = getMarginBadge(margin);
                }
                return marginHtml;
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
                        totalCost: Number(data.totalCost || 0)
                    };
                    updateOrderMarginCell(ordNo);
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

            async function searchOrder() {
                const ordNo = document.getElementById('orderInput').value;
                if (!ordNo) {
                    alert('Indtast et ordrenummer');
                    return;
                }

                // Keep customer list visible during direct order search.
                const result = document.getElementById('result');
                result.innerHTML = '<div class="loading">Indlæser...</div>';
                
                try {
                    const response = await fetch('/aftercalc/' + ordNo);
                    const data = await response.json();
                    
                    if (data.error) {
                        result.innerHTML = '<div class="error">Fejl: ' + data.error + '</div>';
                        return;
                    }

                    currentSalesOrderGr4 = Number((data.orderHeader && data.orderHeader.Gr4) || 0);
                    const orderMarginPercent = calculateOrderMarginPercent(data.summary.totalRevenue, data.summary.totalCost).toFixed(2);
                    const productionOrderByOrdNo = new Map((Array.isArray(data.productionOrders) ? data.productionOrders : []).map(order => [Number(order.ordNo || 0), order]));
                    const getSalesLineCostBreakdown = (purcNo) => {
                        const prodOrder = productionOrderByOrdNo.get(Number(purcNo || 0));
                        const lines = Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : [];
                        let operationTotal = 0;
                        let laserTotal = 0;

                        for (const line of lines) {
                            const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                            const lnNo = Number((line && line.LnNo) || 0);
                            if (lnNo === 1 || key === '0' || key === '3' || key === '5') continue;
                            const effectiveCost = Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                            if (key === '1') {
                                operationTotal += effectiveCost;
                            } else if (key === '2' && isLaserLProdNo(line && line.ProdNo)) {
                                laserTotal += effectiveCost;
                            }
                        }

                        return {
                            operationTotal,
                            laserTotal
                        };
                    };
                    
                    let html = '<div class="order-header">';
                    html += '<h2>Salgsordre: ' + data.orderHeader.OrdNo + ' - ' + (data.orderHeader.CustomerName || '-') + '</h2>';
                    html += '<div class="order-header-row">';
                    html += '<div class="order-header-item"><div class="order-header-label">Faktureret beløb</div><div class="order-header-value">' + formatNumber(data.summary.totalRevenue) + ' DKK</div></div>';
                    html += '<div class="order-header-item"><div class="order-header-label">Kostpris</div><div class="order-header-value">' + formatNumber(data.summary.totalCost) + ' DKK</div></div>';
                    html += '<div class="order-header-item"><div class="order-header-label">Margin (' + getMarginModeLabel() + ')</div><div class="order-header-value">' + getMarginBadge(orderMarginPercent) + '</div></div>';
                    html += '</div></div>';

                    html += '<div class="section">';
                    html += '<h3 style="display:flex; align-items:center; justify-content:space-between; gap:10px;">';
                    html += '<span>Laseroversigt (L-linjer)</span>';
                    html += '<button id="laserOrderSummaryToggleBtn" class="list-toggle-btn" style="padding:6px 10px;" onclick="toggleLaserOrderSummary()" title="Vis eller skjul detaljeret laseroversigt">Vis laseroversigt</button>';
                    html += '</h3>';
                    html += '<div id="laserOrderSummaryTotals" class="summary-box"><div class="loading">Indlæser totaler...</div></div>';
                    html += '<div id="laserOrderSummaryPanel" style="display:none;">';
                    html += '<div class="laser-summary-layout">';
                    html += '<div id="laserOrderSummaryBody" class="loading">Indlæser laserdata...</div>';
                    html += '<aside id="laserImagePanel" class="laser-image-panel hidden"></aside>';
                    html += '</div>';
                    html += '</div>';
                    html += '</div>';

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
                            const lineMarginBadge = !includeForMargin
                                ? '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>'
                                : lineSalesPrice === 0
                                ? '<span style="background:#757575; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">Rabatt</span>'
                                : (isExactlyHundred
                                    ? '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>'
                                    : getMarginBadge(lineMarginPercent));
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
                                    html += '<td><button class="list-toggle-btn drawing-open-btn" data-drawing-path="' + escapeHtml(String(line.DrawingWebPg || '')) + '" style="padding:4px 8px; margin-left:0;">Vis tegning</button></td>';
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
                                html += '<div><strong>L:</strong> ' + formatNumber(breakdownInfo.laserTotal || 0) + ' DKK</div>';
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
                                    html += '<td><button class="list-toggle-btn drawing-open-btn" data-drawing-path="' + escapeHtml(String(line.DrawingWebPg || '')) + '" style="padding:4px 8px; margin-left:0;">Vis tegning</button></td>';
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
                                const normalizedKey = (rawKey === '3') ? '1' : rawKey;
                                const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();

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
                                            return sum + (line.EffectiveLineCost || line.LineCost || 0);
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
                                        return sum + (line.EffectiveLineCost || line.LineCost || 0);
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
                                } else if (key === '6') {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Ydelse pris/enhed</th><th>Samlet kost</th></tr>';
                                } else {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Kostpris/enhed</th><th>Samlet kost</th></tr>';
                                }

                                for (const line of lines) {
                                    html += '<tr>';
                                    const warningFlagHtml = getWarningFlagHtml(line);
                                    const timeAdjustFlagHtml = getTimeAdjustmentFlagHtml(line);
                                    const hasChildProductionOrder = Number(line.PurcNo || 0) > 0;
                                    if (String(key) === '1' && line.ProdNo) {
                                        const safeProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const trInf2Value = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : prodOrder.ordNo);
                                        const safeTrInf2 = trInf2Value.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeTrInf4 = String(line.TrInf4 || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        html += '<td><span class="prod-no-link" data-prodno="' + safeProdNo + '" data-ordno="' + prodOrder.ordNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="' + key + '" data-trinf2="' + safeTrInf2 + '" data-trinf4="' + safeTrInf4 + '">' + safeProdNo + '</span>' + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (String(key) === '2' && line.ProdNo && isLaserLProdNo(line.ProdNo)) {
                                        const safeProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const trInf2Value = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : prodOrder.ordNo);
                                        const safeTrInf2 = trInf2Value.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeTrInf4 = String(line.TrInf4 || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        html += '<td><span class="prod-no-link" data-prodno="' + safeProdNo + '" data-ordno="' + prodOrder.ordNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="' + key + '" data-trinf2="' + safeTrInf2 + '" data-trinf4="' + safeTrInf4 + '" data-showallroutes="1">' + safeProdNo + '</span>' + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (hasChildProductionOrder) {
                                        const safeChildProdNoForSummary = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const childSummaryArgs = String(key) === '6'
                                            ? (Number(line.PurcNo || 0) + ', &quot;' + safeChildProdNoForSummary + '&quot;')
                                            : Number(line.PurcNo || 0);
                                        html += '<td><span class="inline-link" onclick="showChildProductionSummary(' + childSummaryArgs + ')">' + (line.ProdNo || '-') + '</span>' + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (line.ProdNo) {
                                        html += '<td>' + (line.ProdNo || '-') + timeAdjustFlagHtml + warningFlagHtml + '</td>';
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
                    
                    // Riepilogo
                    html += '<div class="section"><h3>Ordresammendrag</h3>';
                    html += '<div class="summary-box">';
                    html += '<div><strong>Samlet faktureret beløb:</strong> ' + formatNumber(data.summary.totalRevenue) + ' DKK</div>';
                    html += '<div><strong>Samlet kost:</strong> ' + formatNumber(data.summary.totalCost) + ' DKK</div>';
                    let marginClass = data.summary.margin >= 0 ? 'margin-positive' : 'margin-negative';
                    html += '<div class="total"><span class="' + marginClass + '">Margin: ' + formatNumber(data.summary.margin) + ' DKK (' + orderMarginPercent + '%)</span></div>';
                    html += '</div></div>';
                    
                    result.innerHTML = html;
                    loadSalesOrderLaserSummary(data);
                } catch (err) {
                    result.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                }
            }

            async function loadSalesOrderLaserSummary(orderData) {
                const body = document.getElementById('laserOrderSummaryBody');
                const totals = document.getElementById('laserOrderSummaryTotals');
                if (!body || !totals) return;
                const orderGr4 = Number((orderData && orderData.orderHeader && orderData.orderHeader.Gr4) || currentSalesOrderGr4 || 0);

                try {
                    const requests = [];
                    const targetDedupe = new Set();
                    const visitedProdOrders = new Set();
                    const productionOrders = Array.isArray(orderData.productionOrders) ? orderData.productionOrders : [];
                    const laserTargets = [];

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
                            const routeSpecificCostPerPiece = (p.CostoPerPezzo !== null && p.CostoPerPezzo !== undefined)
                                ? p.CostoPerPezzo
                                : hintedNestCost;
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
                        return;
                    }

                    const multiNestHeader = orderGr4 === 3 ? 'NestMultiPris' : 'NestKost pr. stk';
                    let html = '<table>';
                    html += '<tr><th>Prod.ordre</th><th>Nestingordre</th><th>Produkt</th><th>Rute</th><th>Færdigmeldt</th><th>Icon vægt (kg/stk)</th><th>Stykliste vaegt (kg/stk)</th><th>Forbrugt (kg/stk)</th><th>' + multiNestHeader + '</th><th>Samlet kost</th><th>Afvigelse (%)</th><th>Billeder</th></tr>';
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
                        const rowTotalCost = (r.quotaCost !== null && r.quotaCost !== undefined)
                            ? Number(r.quotaCost || 0)
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
                            html += '<td><button class="image-preview-btn" data-image-key="' + imageKey + '">Vis</button></td>';
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
                    totals.innerHTML = ''
                        + '<div><strong>Samlet L-kost (' + (orderGr4 === 3 ? 'NestMultiPris' : 'NestKost') + '):</strong> ' + formatNumber(totalLaserCost) + ' DKK</div>'
                        + '<div><strong>Ordre icon kg:</strong> ' + formatNumber(totalKgIcon) + ' kg</div>'
                        + '<div><strong>Ordre stykliste kg:</strong> ' + formatNumber(totalKgPrevisti) + ' kg</div>'
                        + '<div><strong>Ordre forbrugt kg:</strong> ' + formatNumber(totalKgUtilizzati) + ' kg</div>'
                        + '<div><strong>Afvigelse kg:</strong> ' + formatNumber(deltaKg) + ' kg</div>'
                        + '<div><strong>Samlet afvigelse %:</strong> ' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + '</div>';
                } catch (err) {
                    body.innerHTML = '<div class="error">Fejl laseroversigt: ' + err.message + '</div>';
                    totals.innerHTML = '<div class="error">Fejl i samlet laseroversigt: ' + err.message + '</div>';
                }
            }

            async function onProductClick(prodNo, ordNo, lnNo, prodTp4, trInf2, trInf4, showAllRoutes) {
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
                        for (const rowProduct of products) {
                            const oldExpected = rowProduct ? rowProduct.OldNWgtU_medio : null;
                            const expected = rowProduct ? rowProduct.NWgtU_medio : null;
                            const effective = rowProduct ? rowProduct.KgPerPezzoEffettivo : null;
                            const noFin = rowProduct ? rowProduct.QtaPezzi : null;
                            const prodNoForCost = rowProduct ? (rowProduct.ProdNo || prodNo) : prodNo;
                            const hintedNestCost = getLaserNestCostHint(effectiveOrdine, prodNoForCost);
                            const costPerPiece = (rowProduct && rowProduct.CostoPerPezzo !== null && rowProduct.CostoPerPezzo !== undefined)
                                ? rowProduct.CostoPerPezzo
                                : hintedNestCost;
                            const noFinNum = Number(noFin || 0);
                            const expectedNum = Number(expected || 0);
                            const effectiveNum = Number(effective || 0);
                            const totalCost = (rowProduct && rowProduct.QuotaCosto !== null && rowProduct.QuotaCosto !== undefined)
                                ? rowProduct.QuotaCosto
                                : ((costPerPiece === null || costPerPiece === undefined || noFin === null || noFin === undefined)
                                    ? null
                                    : (noFinNum * Number(costPerPiece || 0)));
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
                                html += '<td><button class="image-preview-btn" data-image-key="' + imageKey + '">Vis</button></td>';
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
                if (prodNo) onProductClick(prodNo, ordNo, lnNo, prodTp4, trInf2, trInf4, showAllRoutes);
            }

            function handleImagePreviewClick(e) {
                const btn = e.target.closest('.image-preview-btn');
                if (!btn) return;
                const imageKey = btn.dataset.imageKey;
                if (imageKey) {
                    const inLaserPanel = !!e.target.closest('#laserOrderSummaryPanel');
                    const preferredPanelId = inLaserPanel ? 'laserImagePanel' : 'summaryImagePanel';
                    openSummaryImagePanel(imageKey, preferredPanelId);
                }
            }

            function handleDrawingOpenClick(e) {
                const btn = e.target.closest('.drawing-open-btn');
                if (!btn) return;
                const pathValue = btn.dataset.drawingPath || '';
                if (!pathValue) return;
                openDrawingPdf(pathValue);
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
                if (event.key === 'Escape') closeImageLightbox();
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

            async function showChildProductionSummary(childOrdNo, targetProdNo) {
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

                    let html = '';
                    html += isYdelseFilteredView
                        ? '<table><tr><th>Linje</th><th>ProdTp4</th><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Ydelse pris/enhed</th><th>Samlet kost (beregnet)</th></tr>'
                        : '<table><tr><th>Linje</th><th>ProdTp4</th><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Salgspris</th><th>Kostpris/enhed</th><th>Nesting/enhed</th><th>Samlet kost (beregnet)</th></tr>';
                    for (const line of filteredLines) {
                        const displayLineCost = Number(line.EffectiveLineCost || 0);
                        const warningFlagHtml = getWarningFlagHtml(line);
                        const timeAdjustmentFlagHtml = getTimeAdjustmentFlagHtml(line);
                        html += '<tr>';
                        html += '<td>' + (line.LnNo || 0) + '</td>';
                        html += '<td>' + (line.ProdTp4 === null || line.ProdTp4 === undefined ? '-' : line.ProdTp4) + '</td>';
                        const childHasPurcNo = Number(line.PurcNo || 0) > 0;
                        if (String(line.ProdTp4 || '') === '1' && line.ProdNo) {
                            const safeChildProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const trInf2FromLine = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : childOrdNo);
                            const trInf4FromLine = String(line.TrInf4 || '');
                            const safeChildTrInf2 = trInf2FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const safeChildTrInf4 = trInf4FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            html += '<td><span class="prod-no-link" data-prodno="' + safeChildProdNo + '" data-ordno="' + childOrdNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="1" data-trinf2="' + safeChildTrInf2 + '" data-trinf4="' + safeChildTrInf4 + '">' + safeChildProdNo + '</span>' + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else if (line.ProdNo && String(line.ProdNo).trim().toUpperCase().endsWith('L')) {
                            const safeChildProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const trInf2FromLine = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : childOrdNo);
                            const trInf4FromLine = String(line.TrInf4 || '');
                            const safeChildTrInf2 = trInf2FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const safeChildTrInf4 = trInf4FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            html += '<td><span class="prod-no-link" data-prodno="' + safeChildProdNo + '" data-ordno="' + childOrdNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="2" data-trinf2="' + safeChildTrInf2 + '" data-trinf4="' + safeChildTrInf4 + '" data-showallroutes="1">' + safeChildProdNo + '</span>' + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else if (childHasPurcNo) {
                            const safeChildProdNoForSummary = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const childSummaryArgs = String(line.ProdTp4 || '') === '6'
                                ? (Number(line.PurcNo || 0) + ', &quot;' + safeChildProdNoForSummary + '&quot;')
                                : Number(line.PurcNo || 0);
                            html += '<td><span class="inline-link" onclick="showChildProductionSummary(' + childSummaryArgs + ')">' + (line.ProdNo || '-') + '</span>' + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else {
                            html += '<td>' + (line.ProdNo || '-') + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
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
                        html += '<td><strong>' + formatNumber(displayLineCost) + '</strong></td>';
                        html += '</tr>';
                    }
                    html += isYdelseFilteredView
                        ? '<tr class="summary-row"><td colspan="6">Total beregnet kost:</td><td><strong>' + formatNumber(modalTotalCost || 0) + ' DKK</strong></td></tr>'
                        : '<tr class="summary-row"><td colspan="8">Total beregnet kost:</td><td><strong>' + formatNumber(modalTotalCost || 0) + ' DKK</strong></td></tr>';
                    html += '</table>';
                    body.innerHTML = html;
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

                scrollToElementWithStickyOffset(el);
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
                html += '<h3>Seneste fakturerede ordrer (${ORDER_LIST_DAYS_BACK} dage) &mdash; ' + orders.length + ' af ' + orderListData.length + ' ordrer</h3>';
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
                html += '<th>Opdater</th>';
                html += '</tr>';
                for (const o of orders) {
                    const marginHtml = getOrderMarginHtml(o.OrdNo);

                    const d = String(o.LstInvDt || '');
                    const invDate = d.length === 8 ? d.slice(0,4) + '-' + d.slice(4,6) + '-' + d.slice(6,8) : (d || '-');
                    const orderWarningFlag = getWarningFlagHtml(o, 'Ordren indeholder mindst én advarsel.');
                    const gr4ManualBadge = Number(o.Gr4 || 0) === 3
                        ? '<span title="MultiOrdre" aria-label="MultiOrdre" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:#1565c0;color:#fff;font-size:11px;font-weight:700;margin-left:6px;vertical-align:middle;">M</span>'
                        : '';
                    html += '<tr data-ordno="' + o.OrdNo + '" class="order-list-row">'
                    html += '<td>' + (o.SellerUsr || '-') + '</td>';
                    html += '<td><strong>' + o.OrdNo + '</strong>' + gr4ManualBadge + orderWarningFlag + '</td>';
                    html += '<td>' + (o.CustomerName || '-') + '</td>';
                    html += '<td>' + invDate + '</td>';
                    html += '<td>' + formatNumber(o.InvoAm || 0) + ' DKK</td>';
                    html += '<td class="order-margin-cell" data-ordno="' + o.OrdNo + '">' + marginHtml + '</td>';
                    html += '<td class="order-refresh-cell"><button class="list-toggle-btn order-refresh-one-btn" data-ordno="' + o.OrdNo + '" style="padding:4px 8px; margin-left:0; background:#00695c !important;" title="Opdater cache for denne ordre">Opdater</button></td>';
                    html += '</tr>';
                }
                html += '</table></div>';
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
                    collapseToggleBtn.textContent = '▼ Søg';
                } else {
                    collapseToggleBtn.style.display = 'none';
                    collapseExpandBtn.style.display = 'inline-block';
                }
            }

            // Soeg ved indlaesning hvis ordrenummer er i query string
            window.onload = function() {
                showAccessGate();
                const orderInput = document.getElementById('orderInput');
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
                const orderListEl = document.getElementById('orderList');
                if (orderListEl) {
                    orderListEl.addEventListener('pointerdown', function(e) {
                        const sortHeader = e.target.closest('.order-sortable-header');
                        if (!sortHeader) return;
                        e.preventDefault();
                        const field = sortHeader.getAttribute('data-sort-field');
                        if (field) setOrderListSort(field);
                    });

                    // Prefetch on hover: start loading order data before the user clicks
                    const prefetchInFlight = new Set();
                    orderListEl.addEventListener('mouseover', function(e) {
                        const tr = e.target.closest('tr[data-ordno]');
                        if (!tr) return;
                        const ordNo = Number(tr.dataset.ordno);
                        if (!ordNo || prefetchInFlight.has(ordNo)) return;
                        prefetchInFlight.add(ordNo);
                        fetch('/aftercalc/' + ordNo).catch(() => {});
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
                }
            };
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
                    const warmAftercalcOrdNos = cachedList.slice(0, BACKGROUND_AFTERCALC_WARM_COUNT).map(r => r.OrdNo);
                    warmAftercalcInBackground(warmAftercalcOrdNos, 'startup-cached-list', 25);
                    
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
