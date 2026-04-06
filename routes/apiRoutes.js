const express = require('express');

function createApiRouter({
    getConnection,
    sql,
    fs,
    spawn,
    diskCache,
    logEvent,
    getOrComputeAftercalc,
    getOrComputeOrderMargin,
    getProductionSummary,
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
}) {
    const router = express.Router();

    router.get('/aftercalc/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            logEvent('SEARCH: OrdNo=' + ordNo);
            const cached = diskCache.get('aftercalc_' + ordNo);
            if (cached) {
                logEvent('  -> Cache hit: OrdNo=' + ordNo);
                return res.json(cached);
            }

            const data = await getOrComputeAftercalc(ordNo, { priority: 'high' });
            if (!data.error) {
                logEvent('  -> Found: Revenue=' + data.summary.totalRevenue + ', Margin=' + data.summary.marginPercentage + '%');
            }
            res.json(data);
        } catch (err) {
            console.error('Errore API:', err);
            logEvent('ERROR: ' + err.message);
            res.status(500).json({ error: err.message });
        }
    });

    router.get('/order-margin/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            if (Number.isNaN(ordNo)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }
            const cacheKey = 'order_margin_' + ordNo;
            const cached = diskCache.get(cacheKey);
            if (cached) return res.json({ ...cached, cached: true });

            const marginInfo = await getOrComputeOrderMargin(ordNo);
            const result = {
                ordNo: marginInfo.ordNo,
                totalRevenue: marginInfo.totalRevenue,
                totalCost: marginInfo.totalCost,
                cached: true
            };
            diskCache.set(cacheKey, result, CACHE_TTL_ORDER_MARGIN_MS);
            return res.json(result);
        } catch (err) {
            logEvent('ERROR order-margin: ' + err.message);
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/production-summary/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            if (Number.isNaN(ordNo)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }

            const result = await getProductionSummary(ordNo);
            return res.json(result);
        } catch (err) {
            console.error('Errore production-summary:', err);
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/nesting-detail/:ordno/:prodno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            const prodNo = req.params.prodno;
            if (Number.isNaN(ordNo) || !prodNo) {
                return res.status(400).json({ error: 'Ugyldige parametre' });
            }
            const pool = await getConnection();
            const result = await pool.request()
                .input('ordNo', sql.Numeric, ordNo)
                .input('prodNo', sql.VarChar, prodNo)
                .query(`
                    SELECT OrdNo, TrInf4, CstPr, NoFin, Descr
                    FROM OrdLn
                    WHERE TrInf2 = CAST(@ordNo AS VARCHAR(20))
                      AND ProdNo = @prodNo
                    ORDER BY OrdNo, LnNo
                `);
            res.json(result.recordset);
        } catch (err) {
            res.status(500).json({ error: err.message });
        }
    });

    router.get('/image-file', async (req, res) => {
        try {
            const rawPath = String(req.query.path || '').trim();
            if (!rawPath) {
                return res.status(400).json({ error: 'Billedsti mangler' });
            }

            if (isHttpUrl(rawPath)) {
                return res.redirect(rawPath);
            }

            const normalizedPath = normalizeWindowsPath(rawPath);
            if (!isAbsoluteWindowsPath(normalizedPath)) {
                return res.status(400).json({ error: 'Kun absolutte billedstier er tilladt' });
            }

            if (!isSupportedImagePath(normalizedPath)) {
                return res.status(400).json({ error: 'Filtypen understoettes ikke som billede' });
            }

            if (!fs.existsSync(normalizedPath)) {
                return res.status(404).json({ error: 'Billedfilen blev ikke fundet' });
            }

            return res.sendFile(normalizedPath, err => {
                if (err && !res.headersSent) {
                    res.status(err.statusCode || 500).json({ error: err.message });
                }
            });
        } catch (err) {
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/laser-route-metrics', async (req, res) => {
        try {
            const ordine = String(req.query.ordine || '').trim();
            const route = String(req.query.route || '').trim();
            const prodNoFilter = String(req.query.prodNo || '').trim();
            const normalizedProdNoFilter = prodNoFilter.toUpperCase();
            const showAllRoutes = req.query.showAllRoutes === '1';

            if (!ordine) {
                return res.status(400).json({ error: 'Ugyldige parametre: ordine er paakraevet' });
            }

            const laserCacheKey = 'laser_' + ordine + '_' + (route || 'all') + '_' + (prodNoFilter || 'all') + '_' + (showAllRoutes ? '1' : '0');
            const cachedLaser = diskCache.get(laserCacheKey);
            if (cachedLaser) return res.json(cachedLaser);

            const pool = await getConnection();

            const candidateResult = await pool.request()
                .input('ordine', sql.VarChar, ordine)
                .query(`
                    SELECT OrdNo, TrInf4, ProdNo
                    FROM OrdLn
                    WHERE TrInf2 = @ordine
                      AND TrTp = 7
                `);

            const candidates = candidateResult.recordset || [];
            const normalizedRoute = route ? route.toUpperCase() : '';

            const pickCandidate = () => {
                const withProd = prodNoFilter
                    ? candidates.filter(c => String(c.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                    : candidates;

                if (showAllRoutes && withProd.length > 0) return withProd[0];

                const withRoute = normalizedRoute
                    ? withProd.filter(c => String(c.TrInf4 || '').trim().toUpperCase() === normalizedRoute)
                    : withProd;

                if (withRoute.length > 0) return withRoute[0];
                if (withProd.length > 0) return withProd[0];

                const onlyRoute = normalizedRoute
                    ? candidates.filter(c => String(c.TrInf4 || '').trim().toUpperCase() === normalizedRoute)
                    : [];
                if (onlyRoute.length > 0) return onlyRoute[0];

                return candidates[0] || null;
            };

            const selectedCandidate = pickCandidate();
            if (!selectedCandidate) {
                return res.json({
                    ordine,
                    route: route || null,
                    nestingOrdNo: null,
                    prodNo: prodNoFilter || null,
                    summary: {
                        KgConsumati: null,
                        CostoLastre: null,
                        KgFiniti: null,
                        SfridoKg: null,
                        SfridoPct: null
                    },
                    products: []
                });
            }

            const nestingOrdNo = String(selectedCandidate.OrdNo || '').trim();
            const effectiveRoute = String(selectedCandidate.TrInf4 || '').trim();

            if (!nestingOrdNo || !effectiveRoute) {
                return res.json({
                    ordine,
                    route: effectiveRoute || null,
                    nestingOrdNo: nestingOrdNo || null,
                    prodNo: prodNoFilter || null,
                    summary: {
                        KgConsumati: null,
                        CostoLastre: null,
                        KgFiniti: null,
                        SfridoKg: null,
                        SfridoPct: null
                    },
                    products: []
                });
            }

            const result = await pool.request()
                .input('nestingOrdNo', sql.VarChar, nestingOrdNo)
                .query(`
                    SELECT LnNo, OrdNo, TrInf2, TrInf4, ProdNo, TrTp, NoFin, Free3, IncCst, WebPg, PictFNm
                    FROM OrdLn
                    WHERE OrdNo = @nestingOrdNo
                      AND TrTp IN (5, 7)
                `);

            const rows = result.recordset || [];
            const toNumber = (v) => {
                if (v === null || v === undefined || v === '') return 0;
                if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
                const raw = String(v).trim();
                if (!raw) return 0;
                let normalized = raw;
                if (normalized.includes(',') && normalized.includes('.')) {
                    normalized = normalized.replace(/\./g, '').replace(',', '.');
                } else if (normalized.includes(',')) {
                    normalized = normalized.replace(',', '.');
                }
                const parsed = Number(normalized);
                return Number.isFinite(parsed) ? parsed : 0;
            };
            const normalizeExpectedWeight = (v) => {
                const parsed = toNumber(v);
                return Math.abs(parsed) >= 1000 ? (parsed / 1000) : parsed;
            };
            const round = (v) => Number.isFinite(v) ? parseFloat(Number(v).toFixed(6)) : null;

            const inScopeRows = showAllRoutes
                ? rows
                : rows.filter(r => String(r.TrInf4 || '').trim() === effectiveRoute);

            const sheetRows = inScopeRows.filter(r => Number(r.TrTp) === 5);
            const finishedRows = inScopeRows.filter(r => Number(r.TrTp) === 7);

            const kgConsumati = sheetRows.reduce((s, r) => s + toNumber(r.NoFin), 0);
            const costoLastre = sheetRows.reduce((s, r) => s + toNumber(r.IncCst), 0);
            const kgFiniti = finishedRows.reduce((s, r) => s + (normalizeExpectedWeight(r.Free3) * toNumber(r.NoFin)), 0);
            const sfridoKg = kgConsumati - kgFiniti;
            const sfridoPct = kgConsumati > 0 ? (sfridoKg / kgConsumati) : null;

            const filteredFinishedRows = (prodNoFilter
                ? finishedRows.filter(r => String(r.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                : finishedRows)
                .sort((a, b) => toNumber(a.LnNo) - toNumber(b.LnNo));

            const filteredNestingRows = (prodNoFilter
                ? inScopeRows.filter(r => String(r.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                : finishedRows)
                .sort((a, b) => toNumber(a.LnNo) - toNumber(b.LnNo));

            const routeStats = new Map();
            for (const row of inScopeRows) {
                const routeKey = String(row.TrInf4 || '').trim();
                if (!routeStats.has(routeKey)) {
                    routeStats.set(routeKey, {
                        kgConsumati: 0,
                        costoLastre: 0,
                        kgFiniti: 0
                    });
                }

                const stats = routeStats.get(routeKey);
                if (Number(row.TrTp) === 5) {
                    stats.kgConsumati += toNumber(row.NoFin);
                    stats.costoLastre += toNumber(row.IncCst);
                } else if (Number(row.TrTp) === 7) {
                    stats.kgFiniti += normalizeExpectedWeight(row.Free3) * toNumber(row.NoFin);
                }
            }

            const structMap = new Map();
            try {
                const uniqueProdNos = Array.from(new Set(
                    filteredNestingRows.map(r => String(r.ProdNo || '').trim().toUpperCase())
                ));
                if (uniqueProdNos.length > 0) {
                    const placeholders = uniqueProdNos.map((_, i) => `@p${i}`).join(', ');
                    const request = pool.request();
                    uniqueProdNos.forEach((prodNo, i) => {
                        request.input(`p${i}`, sql.VarChar, prodNo);
                    });
                    const structResult = await request.query(`
                        SELECT ProdNo, NoPerStr
                        FROM Struct
                        WHERE ProdNo IN (${placeholders})
                          AND SubProd LIKE '3%'
                    `);
                    const structRows = structResult.recordset || [];
                    for (const sr of structRows) {
                        const prodKey = String(sr.ProdNo || '').trim().toUpperCase();
                        const noPerStr = toNumber(sr.NoPerStr);
                        if (!structMap.has(prodKey)) {
                            structMap.set(prodKey, noPerStr);
                            if (noPerStr > 0) {
                                logEvent(`DEBUG Struct: ProdNo=${prodKey}, NoPerStr=${noPerStr}`);
                            }
                        }
                    }
                }
            } catch (err) {
                logEvent(`Errore lettura Struct: ${err.message}`);
            }

            const products = filteredNestingRows.map(r => {
                const routeKey = String(r.TrInf4 || '').trim();
                const refFinished = Number(r.TrTp) === 7
                    ? r
                    : filteredFinishedRows.find(fr => String(fr.TrInf4 || '').trim() === routeKey);

                const qtaPezzi = refFinished ? toNumber(refFinished.NoFin) : null;
                const prodKey = String(r.ProdNo || '').trim().toUpperCase();
                const structNoPerStr = structMap.get(prodKey) || null;
                const oldExpectedUnitWeight = refFinished ? normalizeExpectedWeight(refFinished.Free3) : null;
                const expectedUnitWeight = (structNoPerStr !== null && structNoPerStr > 0)
                    ? structNoPerStr
                    : oldExpectedUnitWeight;
                const kgProdotto = (qtaPezzi !== null && expectedUnitWeight !== null)
                    ? (expectedUnitWeight * qtaPezzi)
                    : null;
                const stats = routeStats.get(routeKey) || { kgConsumati: 0, costoLastre: 0, kgFiniti: 0 };
                const nWgtUMedio = (qtaPezzi !== null && qtaPezzi > 0 && kgProdotto !== null) ? (kgProdotto / qtaPezzi) : null;
                const oldKgProdotto = (qtaPezzi !== null && oldExpectedUnitWeight !== null)
                    ? (oldExpectedUnitWeight * qtaPezzi)
                    : null;
                const oldNWgtUMedio = (qtaPezzi !== null && qtaPezzi > 0 && oldKgProdotto !== null) ? (oldKgProdotto / qtaPezzi) : null;
                const kgUtilizzatiEffettivi = (oldKgProdotto !== null && stats.kgFiniti > 0)
                    ? ((oldKgProdotto / stats.kgFiniti) * stats.kgConsumati)
                    : null;
                const kgPerPezzoEffettivo = (kgUtilizzatiEffettivi !== null && qtaPezzi !== null && qtaPezzi > 0)
                    ? (kgUtilizzatiEffettivi / qtaPezzi)
                    : null;
                const quotaCosto = (oldKgProdotto !== null && stats.kgFiniti > 0)
                    ? ((oldKgProdotto / stats.kgFiniti) * stats.costoLastre)
                    : null;
                const costoPerPezzo = (quotaCosto !== null && qtaPezzi !== null && qtaPezzi > 0) ? (quotaCosto / qtaPezzi) : null;
                const euroPerKgFinito = (costoPerPezzo !== null && nWgtUMedio !== null && nWgtUMedio > 0)
                    ? (costoPerPezzo / nWgtUMedio)
                    : null;
                const imageRow = refFinished || r;
                const imageItems = buildImageItems(imageRow ? imageRow.WebPg : null, imageRow ? imageRow.PictFNm : null);

                return {
                    LnNo: toNumber(r.LnNo),
                    ProdNo: String(r.ProdNo || '').trim(),
                    Route: routeKey,
                    TrTp: toNumber(r.TrTp),
                    QtaPezzi: qtaPezzi === null ? null : round(qtaPezzi),
                    KgProdotto: kgProdotto === null ? null : round(kgProdotto),
                    OldNWgtU_medio: oldNWgtUMedio === null ? null : round(oldNWgtUMedio),
                    NWgtU_medio: nWgtUMedio === null ? null : round(nWgtUMedio),
                    KgUtilizzatiEffettivi: kgUtilizzatiEffettivi === null ? null : round(kgUtilizzatiEffettivi),
                    KgPerPezzoEffettivo: kgPerPezzoEffettivo === null ? null : round(kgPerPezzoEffettivo),
                    QuotaCosto: quotaCosto === null ? null : round(quotaCosto),
                    CostoPerPezzo: costoPerPezzo === null ? null : round(costoPerPezzo),
                    EuroPerKgFinito: euroPerKgFinito === null ? null : round(euroPerKgFinito),
                    ImageItems: imageItems,
                    WebPg: imageRow ? String(imageRow.WebPg || '').trim() : '',
                    PictFNm: imageRow ? String(imageRow.PictFNm || '').trim() : ''
                };
            });

            const laserResult = {
                ordine,
                route: showAllRoutes ? null : effectiveRoute,
                nestingOrdNo,
                prodNo: prodNoFilter || null,
                showAllRoutes,
                summary: {
                    KgConsumati: round(kgConsumati),
                    CostoLastre: round(costoLastre),
                    KgFiniti: round(kgFiniti),
                    SfridoKg: round(sfridoKg),
                    SfridoPct: sfridoPct === null ? null : round(sfridoPct)
                },
                products
            };
            diskCache.set(laserCacheKey, laserResult, CACHE_TTL_LASER_METRICS_MS);
            return res.json(laserResult);
        } catch (err) {
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/cache-status', (req, res) => {
        const entries = diskCache.list();
        res.json({ count: entries.length, entries });
    });

    router.get('/health', (req, res) => {
        res.json({ ok: true, version: pkgVersion });
    });

    router.get('/warmup-status', (req, res) => {
        const done = warmupProgress.cached + warmupProgress.loaded + warmupProgress.failed;
        const pct = warmupProgress.total > 0 ? Math.round((done / warmupProgress.total) * 100) : 100;
        res.json({
            running: warmupProgress.running,
            total: warmupProgress.total,
            cached: warmupProgress.cached,
            loaded: warmupProgress.loaded,
            failed: warmupProgress.failed,
            done,
            pct,
            current: warmupProgress.current
        });
    });

    router.post('/cache-refresh-order/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            if (Number.isNaN(ordNo)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }

            if (!orderRefreshInFlight.has(ordNo)) {
                const refreshPromise = (async () => {
                    logEvent('CACHE REFRESH ORDER: ordNo=' + ordNo + ' start');
                    orderRefreshStatus.set(ordNo, { status: 'running', startedAt: Date.now() });

                    diskCache.del('aftercalc_' + ordNo);
                    diskCache.del('prod_summary_' + ordNo);
                    orderMarginInFlight.delete(ordNo);
                    afterCalcInFlight.delete(ordNo);

                    const aftercalc = await getOrComputeAftercalc(ordNo, { priority: 'high' });
                    if (aftercalc && !aftercalc.error) {
                        const marginInfo = {
                            ordNo,
                            totalRevenue: Number(aftercalc.summary && aftercalc.summary.totalRevenue || 0),
                            totalCost: Number(aftercalc.summary && aftercalc.summary.totalCost || 0),
                            computedAt: Date.now()
                        };
                        orderMarginCache.set(ordNo, marginInfo);
                        const marginResult = {
                            ordNo,
                            totalRevenue: marginInfo.totalRevenue,
                            totalCost: marginInfo.totalCost,
                            cached: true
                        };
                        diskCache.set('order_margin_' + ordNo, marginResult, CACHE_TTL_ORDER_MARGIN_MS);
                        logEvent('CACHE REFRESH ORDER: ordNo=' + ordNo + ' margin updated');
                        orderRefreshStatus.set(ordNo, { status: 'done', startedAt: Date.now(), finishedAt: Date.now() });
                    } else {
                        const errMsg = (aftercalc && aftercalc.error) ? aftercalc.error : 'unknown error';
                        orderRefreshStatus.set(ordNo, { status: 'error', error: errMsg, startedAt: Date.now(), finishedAt: Date.now() });
                    }

                    logEvent('CACHE REFRESH ORDER: ordNo=' + ordNo + ' done');
                })()
                    .catch(err => {
                        orderRefreshStatus.set(ordNo, { status: 'error', error: err.message, startedAt: Date.now(), finishedAt: Date.now() });
                        logEvent('ERROR cache-refresh-order worker ordNo=' + ordNo + ': ' + err.message);
                    })
                    .finally(() => {
                        orderRefreshInFlight.delete(ordNo);
                    });

                orderRefreshInFlight.set(ordNo, refreshPromise);
            } else {
                logEvent('CACHE REFRESH ORDER: ordNo=' + ordNo + ' already running');
            }

            return res.json({ ok: true, ordNo, started: true });
        } catch (err) {
            logEvent('ERROR cache-refresh-order: ' + err.message);
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/cache-refresh-order-status/:ordno', (req, res) => {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) {
            return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
        }
        const state = orderRefreshStatus.get(ordNo);
        if (!state) {
            return res.json({ ordNo, status: 'idle' });
        }
        return res.json({ ordNo, ...state });
    });

    router.post('/cache-clear', (req, res) => {
        const deleted = diskCache.clearAll();
        orderMarginCache.clear();
        orderMarginInFlight.clear();
        afterCalcInFlight.clear();
        orderListCache.data = [];
        orderListCache.loadedAt = 0;
        orderListCache.lastError = null;
        logEvent('CACHE CLEARED: ' + deleted + ' files deleted, in-memory caches reset');
        res.json({ ok: true, deleted });
    });

    router.post('/desktop-update-check', async (req, res) => {
        try {
            const checkFn = global.__desktopManualUpdateCheck;
            if (typeof checkFn !== 'function') {
                return res.status(503).json({ ok: false, status: 'unavailable', message: 'Opdateringskontrol er ikke tilgaengelig i denne mode.' });
            }

            const result = await checkFn();
            logEvent('MANUAL-UPDATE-CHECK: status=' + String(result && result.status || 'unknown') + ', ok=' + String(!!(result && result.ok)));
            return res.json(result || { ok: false, status: 'error', message: 'Tomt svar fra updater.' });
        } catch (err) {
            logEvent('MANUAL-UPDATE-CHECK ERROR: ' + err.message);
            return res.status(500).json({ ok: false, status: 'error', message: err.message });
        }
    });

    router.post('/open-drawing', (req, res) => {
        try {
            const rawPath = String((req.body && req.body.path) || '').trim();
            if (!rawPath) {
                return res.status(400).json({ ok: false, message: 'Path mangler.' });
            }

            const lower = rawPath.toLowerCase();
            if (lower.indexOf('.pdf') === -1) {
                return res.status(400).json({ ok: false, message: 'Kun PDF er tilladt.' });
            }

            const child = spawn('cmd', ['/c', 'start', '', rawPath], {
                windowsHide: true,
                detached: true,
                stdio: 'ignore'
            });
            child.unref();

            logEvent('OPEN-DRAWING: ' + rawPath);
            return res.json({ ok: true });
        } catch (err) {
            logEvent('OPEN-DRAWING ERROR: ' + err.message);
            return res.status(500).json({ ok: false, message: err.message });
        }
    });

    router.get('/prodtr/:ordno/:lnno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            const lnNo = parseInt(req.params.lnno);
            if (Number.isNaN(ordNo) || Number.isNaN(lnNo)) {
                return res.status(400).json({ error: 'Ugyldige parametre' });
            }
            const pool = await getConnection();
            const result = await pool.request()
                .input('ordNo', sql.Numeric, ordNo)
                .input('lnNo', sql.Numeric, lnNo)
                .query(`
                    SELECT
                        P.FinDt,
                        P.FinTm,
                        P.NoInvoAb,
                        A.Nm AS HvemNm
                    FROM ProdTr P
                    LEFT JOIN Actor A ON A.EmpNo = P.EmpNo
                    WHERE P.OrdNo = @ordNo AND P.OrdLnNo = @lnNo
                    ORDER BY P.FinDt DESC, P.FinTm DESC
                `);
            res.json(result.recordset);
        } catch (err) {
            res.status(500).json({ error: err.message });
        }
    });

    router.get('/order-list-check-time', async (req, res) => {
        try {
            const pool = await getConnection();
            const result = await pool.request().query(`
                SELECT MAX(CAST(LstInvDt AS INT)) as maxInvDate
                FROM Ord
                WHERE CAST(CAST(LstInvDt AS CHAR(8)) AS INT)
                    >= CONVERT(INT, FORMAT(DATEADD(DAY, -${ORDER_LIST_DAYS_BACK}, GETDATE()), 'yyyyMMdd'))
            `);

            const maxDate = result.recordset[0]?.maxInvDate || 0;
            const serverTime = Date.now();

            res.json({
                lastModifiedDate: maxDate,
                serverTime: serverTime,
                cacheLastModified: orderListCache.lastModifiedTime
            });
        } catch (err) {
            logEvent('ERROR order-list-check-time: ' + err.message);
            res.status(500).json({ error: err.message });
        }
    });

    router.get('/order-list', async (req, res) => {
        try {
            const forceRefresh = req.query.force === '1';
            logEvent('ORDER-LIST: force=' + (forceRefresh ? '1' : '0'));

            if (forceRefresh) {
                await refreshOrderListCache(true);
            } else if (!isOrderListCacheFresh()) {
                await refreshOrderListCache();
            }

            let marginFromMemory = 0;
            let marginFromDisk = 0;
            let marginFromDiskStale = 0;
            let marginMissing = 0;

            const data = orderListCache.data.map(row => {
                const ordNoNum = Number(row.OrdNo);
                let marginInfo = orderMarginCache.get(ordNoNum);
                if (marginInfo) {
                    marginFromMemory += 1;
                }

                if (!marginInfo) {
                    const cachedMargin = diskCache.get('order_margin_' + ordNoNum);
                    if (cachedMargin && cachedMargin.totalCost !== null && cachedMargin.totalCost !== undefined) {
                        marginInfo = {
                            ordNo: ordNoNum,
                            totalRevenue: Number(cachedMargin.totalRevenue || row.InvoAm || 0),
                            totalCost: Number(cachedMargin.totalCost || 0),
                            computedAt: Date.now()
                        };
                        orderMarginCache.set(ordNoNum, marginInfo);
                        marginFromDisk += 1;
                    }
                }

                if (!marginInfo) {
                    const staleMargin = diskCache.getStale('order_margin_' + ordNoNum);
                    if (staleMargin && staleMargin.totalCost !== null && staleMargin.totalCost !== undefined) {
                        marginInfo = {
                            ordNo: ordNoNum,
                            totalRevenue: Number(staleMargin.totalRevenue || row.InvoAm || 0),
                            totalCost: Number(staleMargin.totalCost || 0),
                            computedAt: Date.now()
                        };
                        orderMarginCache.set(ordNoNum, marginInfo);
                        marginFromDiskStale += 1;
                    }
                }

                if (!marginInfo) {
                    marginMissing += 1;
                }

                return {
                    ...row,
                    TotalCost: marginInfo ? marginInfo.totalCost : null
                };
            });

            orderListCache.lastModifiedTime = Date.now();

            if (!isOrderListCacheFresh() && !orderListCache.loading) {
                refreshOrderListCache(true).catch(err => {
                    logEvent('ERROR order-list refresh: ' + err.message);
                });
            }

            logEvent('ORDER-LIST: returned ' + data.length + ' rows (margin memory=' + marginFromMemory + ', disk=' + marginFromDisk + ', stale=' + marginFromDiskStale + ', missing=' + marginMissing + ')');
            res.json(data);
        } catch (err) {
            logEvent('ERROR order-list: ' + err.message);
            res.status(500).json({ error: err.message });
        }
    });

    return router;
}

module.exports = {
    createApiRouter
};
