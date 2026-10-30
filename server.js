const express = require('express');
const sql = require('mssql/msnodesqlv8');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const getConnection = require('./db');
const localDiskCache = require('./diskCache');
const gohCache = require('./services/gohCacheService');
const gohData = require('./services/gohDataService');
// Mirror trasparente: ogni scrittura/cancellazione locale viene replicata su GOH
// (fail-soft, fire-and-forget). Le letture restano locali e sincrone.
const diskCache = {
    get: localDiskCache.get,
    getStale: localDiskCache.getStale,
    list: localDiskCache.list,
    // "Ryd cache" deve svuotare anche la cache condivisa, altrimenti la sync
    // riporta subito indietro i dati vecchi da GOH.
    clearAll() {
        const deleted = localDiskCache.clearAll();
        gohCache.clearAll().catch(() => {});
        return deleted;
    },
    set(key, data, ttlMs) {
        localDiskCache.set(key, data, ttlMs);
        gohCache.set(key, data, ttlMs).catch(() => {});
    },
    del(key) {
        localDiskCache.del(key);
        gohCache.del(key).catch(() => {});
    }
};
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
const aftercalcCostExclusionsService = require('./services/aftercalcCostExclusionsService');
const { calculateAdjustedCost } = require('./assets/js/aftercalc-cost-exclusions');
const { fetchSalgordreViaRows } = require('./services/viaService');
const { createApiRouter } = require('./routes/apiRoutes');

const CACHE_TTL_AFTERCALC_MS        = 8 * 60 * 60 * 1000;  // 8 hours - match background refresh cadence
const CACHE_TTL_PRODUCTION_SUMMARY_MS = 30 * 60 * 1000;  // 30 min
const CACHE_TTL_LASER_METRICS_MS    = 60 * 60 * 1000;  // 60 min
const CACHE_TTL_ORDER_MARGIN_MS     = 30 * 60 * 1000;  // 30 min
const AFTERCALC_CACHE_KEY_PREFIX = 'aftercalc_v29_';
const ORDER_MARGIN_CACHE_KEY_PREFIX = 'order_margin_v26_';
const ORDER_COST_BREAKDOWN_CACHE_KEY_PREFIX = 'order_cost_breakdown_v1_';
const LEGACY_AFTERCALC_CACHE_KEY_PREFIXES = ['aftercalc_v28_', 'aftercalc_v27_', 'aftercalc_v21_', 'aftercalc_v20_', 'aftercalc_v19_', 'aftercalc_v18_', 'aftercalc_v17_', 'aftercalc_'];

const app = express();
// Parser JSON globale 256kb, ma /bom/analyze-file ha il proprio parser 40mb a livello di route
const jsonBodyParser = express.json({ limit: '256kb' });
const snapshotJsonBodyParser = express.json({ limit: '15mb' });
app.use((req, res, next) => {
    if (req.path === '/bom/analyze-file') return next();
    if (req.path === '/lagerliste/snapshots' && req.method === 'POST') {
        return snapshotJsonBodyParser(req, res, next);
    }
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
gohCache.configure({ logEvent });
gohData.configure({ logEvent });
const ORDER_LIST_CACHE_TTL_MS = 8 * 60 * 60 * 1000;
const ORDER_LIST_MAX_ROWS = 150;
const ORDER_LIST_DAYS_BACK = 30;
const STARTUP_MARGIN_WARM_COUNT = ORDER_LIST_MAX_ROWS;
const BACKGROUND_WARM_INTERVAL_MS = 8 * 60 * 60 * 1000;
const BACKGROUND_AFTERCALC_WARM_COUNT = ORDER_LIST_MAX_ROWS;  // Warm full aftercalc for ALL orders in the list so clicking is instant
const BACKGROUND_WARM_DELAY_MS = 10;  // Small stagger between queue submissions to keep DB stable
const MAX_DB_CALC_CONCURRENCY = 3;  // Controlled parallel DB calculations for faster startup warmup
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
const orderCostBreakdownCache = new Map();
const orderCostBreakdownInFlight = new Map();
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

async function enrichAftercalcInvoiceNo(cached, ordNo) {
    if (!cached || !cached.orderHeader || Object.prototype.hasOwnProperty.call(cached.orderHeader, 'InvoNo')) return cached;
    try {
        const pool = await getConnection();
        const result = await pool.request()
            .input('ordNo', sql.Numeric, Number(ordNo))
            .query('SELECT InvoNo FROM Ord WHERE OrdNo=@ordNo');
        const row = result.recordset && result.recordset[0];
        cached.orderHeader.InvoNo = row ? row.InvoNo : null;
        diskCache.set(AFTERCALC_CACHE_KEY_PREFIX + Number(ordNo), cached, CACHE_TTL_AFTERCALC_MS);
    } catch (err) {
        logEvent('AFTERCALC INVOICE HEADER ERROR: ordNo=' + ordNo + ', ' + err.message);
    }
    return cached;
}

async function getOrComputeAftercalc(ordNo, options = {}) {
    const priority = options.priority || 'normal';
    const forceRefresh = Boolean(options.forceRefresh);
    const key = Number(ordNo);
    if (!Number.isFinite(key)) {
        throw new Error('Ordrenummer ugyldigt');
    }

    const cacheKey = AFTERCALC_CACHE_KEY_PREFIX + key;
    if (forceRefresh) {
        diskCache.del(cacheKey);
    }
    const cached = getAftercalcCacheWithFallback(key, true);
    if (cached && !forceRefresh) {
        logEvent('AFTERCALC CACHE HIT: ordNo=' + key);
        return options.includeInvoiceNo ? enrichAftercalcInvoiceNo(cached, key) : cached;
    }

    // Solo lettura cache (locale → GOH): nessun calcolo Visma, null se assente
    if (options.cachedOnly) {
        if (forceRefresh) return null;
        const shared = await gohCache.get(cacheKey);
        if (shared) {
            logEvent('AFTERCALC GOH-CACHE HIT (cachedOnly): ordNo=' + key);
            localDiskCache.set(cacheKey, shared, CACHE_TTL_AFTERCALC_MS);
            return shared;
        }
        return null;
    }

    let computePromise = afterCalcInFlight.get(key);
    if (computePromise && !forceRefresh) {
        logEvent('AFTERCALC IN-FLIGHT REUSE: ordNo=' + key);
    }
    if (forceRefresh) computePromise = null;
    if (!computePromise) {
        computePromise = (async () => {
            // Prova la cache condivisa GOH prima di occupare uno slot di calcolo Visma
            if (!forceRefresh) {
                const shared = await gohCache.get(cacheKey);
                if (shared) {
                    logEvent('AFTERCALC GOH-CACHE HIT: ordNo=' + key);
                    localDiskCache.set(cacheKey, shared, CACHE_TTL_AFTERCALC_MS);
                    return shared;
                }
            }
            logEvent('AFTERCALC FRESH COMPUTE: ordNo=' + key + ', priority=' + priority);
            return runWithDbCalcLimit(async () => {
                const data = await withTimeout(
                    getAfterCalc(key),
                    AFTERCALC_QUERY_TIMEOUT_MS,
                    'Aftercalc timeout for ordNo=' + key
                );
                if (!data.error) {
                    diskCache.set(cacheKey, data, CACHE_TTL_AFTERCALC_MS);
                    if (data.summary) {
                        gohData.recordAftercalcSnapshot(key, data.summary).catch(() => {});
                    }
                }
                return data;
            }, priority);
        })().finally(() => {
            afterCalcInFlight.delete(key);
        });
        afterCalcInFlight.set(key, computePromise);
    }

    const computed = await computePromise;
    return options.includeInvoiceNo ? enrichAftercalcInvoiceNo(computed, key) : computed;
}


async function getOrComputeOrderMargin(ordNo, options = {}) {
    const forceRefresh = Boolean(options.forceRefresh);
    const key = Number(ordNo);
    if (!Number.isFinite(key)) {
        throw new Error('Ordrenummer ugyldigt');
    }
    await aftercalcCostExclusionsService.ensureHydrated();
    const costExclusionFingerprint = aftercalcCostExclusionsService.getFingerprintSync(key);
    const cachedAftercalc = forceRefresh ? null : getAftercalcCacheWithFallback(key, false);
    const currentSummary = cachedAftercalc && !cachedAftercalc.error && cachedAftercalc.summary;
    const currentCost = currentSummary ? calculateAdjustedCost(
        Number(currentSummary.totalCost || 0),
        cachedAftercalc.salesOrderLines,
        aftercalcCostExclusionsService.getExcludedLineKeysSync(key)
    ).adjustedCost : null;
    const matchesCurrentCalculation = margin => !currentSummary || (
        Number(margin.totalRevenue || 0) === Number(currentSummary.totalRevenue || 0)
        && Number(margin.totalCost) === currentCost
        && Number(margin.styklisteFallbackCost || 0) === Number(currentSummary.styklisteFallbackCost || 0)
        && Boolean(margin.hasInvoiceWarning) === Boolean(currentSummary.hasInvoiceWarning)
    );

    if (!forceRefresh && orderMarginCache.has(key)) {
        const memoryMargin = orderMarginCache.get(key);
        if (memoryMargin && memoryMargin.costExclusionFingerprint === costExclusionFingerprint && matchesCurrentCalculation(memoryMargin)) {
            return memoryMargin;
        }
        orderMarginCache.delete(key);
    }

    if (!forceRefresh) {
        const diskMargin = diskCache.get(ORDER_MARGIN_CACHE_KEY_PREFIX + key)
            || diskCache.getStale(ORDER_MARGIN_CACHE_KEY_PREFIX + key)
            || diskCache.getStale('order_margin_v6_' + key);
        const diskFingerprint = diskMargin && typeof diskMargin.costExclusionFingerprint === 'string'
            ? diskMargin.costExclusionFingerprint
            : '';
        if (diskMargin
            && diskMargin.totalCost !== null
            && diskMargin.totalCost !== undefined
            && diskFingerprint === costExclusionFingerprint
            && matchesCurrentCalculation(diskMargin)) {
            const marginInfo = {
                ordNo: key,
                totalRevenue: Number(diskMargin.totalRevenue || 0),
                totalCost: Number(diskMargin.totalCost || 0),
                styklisteFallbackCost: Number(diskMargin.styklisteFallbackCost || 0),
                hasInvoiceWarning: diskMargin.hasInvoiceWarning === true,
                costExclusionFingerprint,
                computedAt: Date.now()
            };
            orderMarginCache.set(key, marginInfo);
            return marginInfo;
        }
    }

    if (!forceRefresh && orderMarginInFlight.has(key)) {
        return orderMarginInFlight.get(key);
    }

    // Deriva il margine dalla cache/calcolo aftercalc condiviso: evita il doppio
    // calcolo DB per ordine durante il warmup (aftercalc + margin sono lo stesso dato).
    const computePromise = (async () => {
        const data = currentSummary ? cachedAftercalc : await getOrComputeAftercalc(key, {
            priority: options.priority || 'normal',
            forceRefresh
        });
        if (data.error) {
            throw new Error(data.error);
        }
        const marginInfo = {
            ordNo: key,
            totalRevenue: Number(data.summary.totalRevenue || 0),
            totalCost: calculateAdjustedCost(
                Number(data.summary.totalCost || 0),
                data.salesOrderLines,
                aftercalcCostExclusionsService.getExcludedLineKeysSync(key)
            ).adjustedCost,
            // Campo separato: sommato SOLO dal client Efterkalkulation, mai da altri moduli
            styklisteFallbackCost: Number(data.summary.styklisteFallbackCost || 0),
            hasInvoiceWarning: Boolean(data.summary.hasInvoiceWarning),
            costExclusionFingerprint,
            computedAt: Date.now()
        };

        orderMarginCache.set(key, marginInfo);
        // Also save to persistent disk cache (24 hours) for faster startup next time
        diskCache.set(ORDER_MARGIN_CACHE_KEY_PREFIX + key, marginInfo, 24 * 60 * 60 * 1000);
        return marginInfo;
    })().finally(() => {
        orderMarginInFlight.delete(key);
    });

    orderMarginInFlight.set(key, computePromise);
    return computePromise;
}

// Material/stang/tid/indkøbte dele/underleverandør pr. ordre, fra samme ProdTp4/Gr6-baserede
// SQL som "Salgsordre VIA" (services/viaService.js), men uafhængig cache: ordrer her kan
// være allerede fakturerede (lukkede), ikke kun åbne som i VIA-backloggen.
async function getOrComputeOrderCostBreakdown(ordNo, options = {}) {
    const forceRefresh = Boolean(options.forceRefresh);
    const key = Number(ordNo);
    if (!Number.isFinite(key)) {
        throw new Error('Ordrenummer ugyldigt');
    }

    if (!forceRefresh && orderCostBreakdownCache.has(key)) {
        return orderCostBreakdownCache.get(key);
    }

    if (!forceRefresh) {
        const diskBreakdown = diskCache.get(ORDER_COST_BREAKDOWN_CACHE_KEY_PREFIX + key);
        if (diskBreakdown) {
            orderCostBreakdownCache.set(key, diskBreakdown);
            return diskBreakdown;
        }
    }

    if (!forceRefresh && orderCostBreakdownInFlight.has(key)) {
        return orderCostBreakdownInFlight.get(key);
    }

    const computePromise = (async () => {
        const rows = await fetchSalgordreViaRows({ getConnection, sql, orderNos: [key] });
        const row = rows[0] || null;
        const costFields = ['MaterialCost', 'StangCost', 'TimeCost', 'PurchasedPartCost', 'UnderleverandorCost'];
        const costBreakdownAvailable = !!row && costFields.every(field => row[field] != null && Number.isFinite(Number(row[field])));
        const breakdown = {
            ordNo: key,
            materialCost: costBreakdownAvailable ? Number(row.MaterialCost) : null,
            stangCost: costBreakdownAvailable ? Number(row.StangCost) : null,
            timeCost: costBreakdownAvailable ? Number(row.TimeCost) : null,
            purchasedPartCost: costBreakdownAvailable ? Number(row.PurchasedPartCost) : null,
            underleverandorCost: costBreakdownAvailable ? Number(row.UnderleverandorCost) : null,
            costBreakdownAvailable,
            computedAt: Date.now()
        };

        orderCostBreakdownCache.set(key, breakdown);
        diskCache.set(ORDER_COST_BREAKDOWN_CACHE_KEY_PREFIX + key, breakdown, 24 * 60 * 60 * 1000);
        return breakdown;
    })().finally(() => {
        orderCostBreakdownInFlight.delete(key);
    });

    orderCostBreakdownInFlight.set(key, computePromise);
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
        // Prefetch bulk dalla cache condivisa GOH: gli ordini già calcolati da
        // un'altra macchina diventano hit locali senza toccare Visma.
        const missingKeys = ordNos
            .map(o => Number(o))
            .filter(n => Number.isFinite(n) && !getAftercalcCacheWithFallback(n, false))
            .map(n => AFTERCALC_CACHE_KEY_PREFIX + n);
        if (missingKeys.length > 0 && gohCache.isEnabled()) {
            const shared = await gohCache.getMany(missingKeys);
            for (const [sharedKey, entry] of shared) {
                const remainingTtl = Math.max(60 * 1000, entry.ttlMs - (Date.now() - entry.cachedAtMs));
                localDiskCache.set(sharedKey, entry.data, remainingTtl);
            }
            if (shared.size > 0) {
                logEvent('WARM-AFTERCALC (' + sourceLabel + '): ' + shared.size + '/' + missingKeys.length + ' ordini precaricati da GOH-cache');
            }
        }

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

// Sync bidirezionale con la cache condivisa GOH: scarica in locale le entry
// più fresche calcolate da altre macchine, pubblica quelle locali mancanti.
// Da qui in poi ogni scrittura runtime è già replicata dal mirror diskCache.
async function syncCacheWithGoh() {
    if (!gohCache.isEnabled()) return;
    try {
        logEvent('GOH-SYNC: start (ping ' + gohCache.serverLabel + ')');
        // Timeout duro: msnodesqlv8 può restare appeso oltre il connectionTimeout
        const pingOk = await withTimeout(gohCache.ping(), 15000, 'GOH ping timeout (15s)')
            .catch(err => { logEvent('GOH-SYNC: ' + err.message); return false; });
        if (!pingOk) return;
        gohCache.purgeExpired().then(n => { if (n > 0) logEvent('GOH-CACHE PURGE: ' + n + ' udløbne entries slettet'); }).catch(() => {});
        const localEntries = localDiskCache.list().filter(e => e.key && e.fresh);
        const localByKey = new Map(localEntries.map(e => [e.key, e]));
        const remote = await gohCache.getAllFresh();
        if (!gohCache.isEnabled()) return;

        let downloaded = 0;
        for (const [key, entry] of remote) {
            const local = localByKey.get(key);
            const localMs = local ? (Date.parse(local.cachedAt) || 0) : 0;
            if (entry.cachedAtMs <= localMs) continue;
            const remainingTtl = Math.max(60 * 1000, entry.ttlMs - (Date.now() - entry.cachedAtMs));
            localDiskCache.set(key, entry.data, remainingTtl);
            downloaded += 1;
        }

        let uploaded = 0;
        for (const entry of localEntries) {
            const remoteEntry = remote.get(entry.key);
            const localCachedAtMs = Date.parse(entry.cachedAt) || 0;
            if (remoteEntry && Number(remoteEntry.cachedAtMs) >= localCachedAtMs) continue;
            const data = localDiskCache.get(entry.key);
            if (!data) continue;
            const remainingTtlMs = Math.max(60 * 1000, (entry.ttlMinutes - entry.ageMinutes) * 60 * 1000);
            if (await gohCache.set(entry.key, data, remainingTtlMs)) uploaded += 1;
            if (!gohCache.isEnabled()) return;
        }
        logEvent('GOH-SYNC: scaricati ' + downloaded + ', pubblicati ' + uploaded + ' (remoto=' + remote.size + ', locale=' + localEntries.length + ')');
    } catch (err) {
        logEvent('GOH-SYNC ERROR: ' + err.message);
    }
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
    gohData,
    gohCache,
    logEvent,
    getOrComputeAftercalc,
    getOrComputeOrderMargin,
    getOrComputeOrderCostBreakdown,
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

// Storico efterkalk (GantechOperationHub): serie temporale omsætning/kost per ordine
app.get('/order-trend/:ordno', async (req, res) => {
    const ordNo = parseInt(req.params.ordno, 10);
    if (!Number.isFinite(ordNo)) {
        return res.status(400).json({ ok: false, error: 'Ugyldigt ordrenummer' });
    }
    try {
        const readings = await gohData.getOrderTrend(ordNo);
        res.json({ ok: true, ordNo, available: gohData.isEnabled(), readings });
    } catch (err) {
        res.status(500).json({ ok: false, error: err.message });
    }
});

// Endpoint per HTML — markup/CSS estratto verbatim in views/index.html (nessun placeholder
// dinamico salvo pkgVersion/APP_VERSION, sostituiti qui sotto e messi in cache una volta sola).
let cachedIndexHtml = null;
app.get("/", (req, res) => {
    logEvent("HTTP GET /");
    if (!cachedIndexHtml) {
        cachedIndexHtml = fs.readFileSync(path.join(__dirname, "views", "index.html"), "utf8")
            .split("__PKG_VERSION__").join(pkgVersion)
            .split("__APP_VERSION__").join(APP_VERSION);
    }
    res.send(cachedIndexHtml);
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
        syncCacheWithGoh().catch(() => {});
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
                // Connessione eager a GOH + sync bidirezionale della cache (fail-soft)
                syncCacheWithGoh().catch(() => {});
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
