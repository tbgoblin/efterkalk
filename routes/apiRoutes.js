const express = require('express');
const orderNotesService = require('../services/orderNotesService');
const omsaetningThresholdsService = require('../services/omsaetningThresholdsService');
const { createOmsaetningService } = require('../services/omsaetningService');
const { createOrdreindgangService } = require('../services/ordreindgangService');

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
}) {
    const router = express.Router();
    const omsaetningService = createOmsaetningService({ getConnection, sql });
    const ordreindgangService = createOrdreindgangService({ getConnection, sql });

    router.get('/aftercalc/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            logEvent('SEARCH: OrdNo=' + ordNo);
            const data = await getOrComputeAftercalc(ordNo, { priority: 'high' });
            if (!data || data.error) {
                return res.json(data);
            }

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
            const cacheKey = ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo;
            const cached = diskCache.get(cacheKey);
            if (cached) return res.json({ ...cached, cached: true });

            const marginInfo = await getOrComputeOrderMargin(ordNo);
            const result = {
                ordNo: marginInfo.ordNo,
                totalRevenue: marginInfo.totalRevenue,
                totalCost: marginInfo.totalCost,
                hasInvoiceWarning: Boolean(marginInfo.hasInvoiceWarning),
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

            const orderGr4 = Number(req.query.gr4 || 0);
            const result = await getProductionSummary(ordNo, new Set(), { orderGr4 });
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
            const orderGr4 = Number(req.query.gr4 || 0);
            const useSpecialLaserCost = orderGr4 === 3;

            if (!ordine) {
                return res.status(400).json({ error: 'Ugyldige parametre: ordine er paakraevet' });
            }

            const laserCacheKey = 'laser_v4_' + ordine + '_' + (route || 'all') + '_' + (prodNoFilter || 'all') + '_' + (showAllRoutes ? '1' : '0') + '_gr4_' + (useSpecialLaserCost ? '3' : '0');
            const cachedLaser = diskCache.get(laserCacheKey);
            if (cachedLaser) return res.json(cachedLaser);

            const pool = await getConnection();

            const candidateResult = await pool.request()
                .input('ordine', sql.VarChar, ordine)
                .query(`
                    SELECT OrdNo, TrInf4, ProdNo, NoFin
                    FROM OrdLn
                    WHERE TrInf2 = @ordine
                      AND TrTp = 7
                `);

            const candidates = candidateResult.recordset || [];
            const normalizedRoute = route ? route.toUpperCase() : '';
            const routeMatches = candidate => String(candidate.TrInf4 || '').trim().toUpperCase() === normalizedRoute;

            const withProd = prodNoFilter
                ? candidates.filter(c => String(c.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                : candidates;
            const withRoute = normalizedRoute
                ? withProd.filter(routeMatches)
                : withProd;
            const routeOnly = normalizedRoute
                ? candidates.filter(routeMatches)
                : [];

            const selectedCandidates = showAllRoutes
                ? (withProd.length > 0 ? withProd : (routeOnly.length > 0 ? routeOnly : candidates))
                : (withRoute.length > 0
                    ? [withRoute[0]]
                    : (withProd.length > 0
                        ? [withProd[0]]
                        : (routeOnly.length > 0 ? [routeOnly[0]] : (candidates[0] ? [candidates[0]] : []))));

            if (selectedCandidates.length === 0) {
                return res.json({
                    ordine,
                    route: route || null,
                    nestingOrdNo: null,
                    nestingOrdNos: [],
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

            const nestingOrdNos = Array.from(new Set(
                selectedCandidates
                    .map(candidate => String(candidate.OrdNo || '').trim())
                    .filter(Boolean)
            ));
            const effectiveRoute = showAllRoutes ? '' : String(selectedCandidates[0].TrInf4 || '').trim();

            // Mappa OrdNo_TrInf4_ProdNo → NoFin dalla produzione: contiene il vero Færdigmeldt per rotta.
            // I nesting order rows hanno spesso NoFin=totale (es. 40 per tutte le rotte),
            // mentre questi record (TrInf2=produzione) hanno il valore corretto per singola rotta.
            const candidateNoFinMap = new Map();
            for (const c of candidates) {
                const k = String(c.OrdNo || '').trim() + '_' + String(c.TrInf4 || '').trim() + '_' + String(c.ProdNo || '').trim().toUpperCase();
                if (!candidateNoFinMap.has(k)) candidateNoFinMap.set(k, Number(c.NoFin || 0));
            }

            if (nestingOrdNos.length === 0 || (!showAllRoutes && !effectiveRoute)) {
                return res.json({
                    ordine,
                    route: effectiveRoute || null,
                    nestingOrdNo: nestingOrdNos[0] || null,
                    nestingOrdNos,
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

            const nestingRowsRequest = pool.request();
            nestingOrdNos.forEach((ordValue, index) => {
                nestingRowsRequest.input(`nestingOrdNo${index}`, sql.VarChar, ordValue);
            });
            const nestingPlaceholders = nestingOrdNos.map((_, index) => `@nestingOrdNo${index}`).join(', ');
            const result = await nestingRowsRequest.query(`
                    SELECT LnNo, OrdNo, TrInf2, TrInf4, ProdNo, TrTp, NoFin, Free3, IncCst, CstPr, WebPg, PictFNm
                    FROM OrdLn
                    WHERE OrdNo IN (${nestingPlaceholders})
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

            const filteredFinishedRows = (prodNoFilter
                ? finishedRows.filter(r => String(r.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                : finishedRows)
                .sort((a, b) => toNumber(a.LnNo) - toNumber(b.LnNo));

            const filteredNestingRows = (prodNoFilter
                ? finishedRows.filter(r => String(r.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                : finishedRows)
                .sort((a, b) => toNumber(a.LnNo) - toNumber(b.LnNo));

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

            const getExpectedUnitWeight = (row) => {
                const free3Weight = normalizeExpectedWeight(row.Free3);
                if (free3Weight > 0) return free3Weight;
                const prodKey = String(row.ProdNo || '').trim().toUpperCase();
                return normalizeExpectedWeight(structMap.get(prodKey));
            };

            const kgFiniti = finishedRows.reduce((s, r) => s + (getExpectedUnitWeight(r) * toNumber(r.NoFin)), 0);
            const sfridoKg = kgConsumati - kgFiniti;
            const sfridoPct = kgConsumati > 0 ? (sfridoKg / kgConsumati) : null;

            const routeStats = new Map();
            const routeStatsKey = (row) => String(row.OrdNo || '').trim() + '|' + String(row.TrInf4 || '').trim();
            for (const row of inScopeRows) {
                const statsKey = routeStatsKey(row);
                if (!routeStats.has(statsKey)) {
                    routeStats.set(statsKey, {
                        kgConsumati: 0,
                        costoLastre: 0,
                        kgFiniti: 0,
                        cstPrSum: 0,
                        cstPrCount: 0
                    });
                }

                const stats = routeStats.get(statsKey);
                if (Number(row.TrTp) === 5) {
                    stats.kgConsumati += toNumber(row.NoFin);
                    stats.costoLastre += toNumber(row.IncCst);
                    const rowCstPr = toNumber(row.CstPr);
                    if (rowCstPr > 0) {
                        stats.cstPrSum += rowCstPr;
                        stats.cstPrCount += 1;
                    }
                } else if (Number(row.TrTp) === 7) {
                    stats.kgFiniti += getExpectedUnitWeight(row) * toNumber(row.NoFin);
                }
            }

            const products = filteredNestingRows.map(r => {
                const routeKey = String(r.TrInf4 || '').trim();
                const refFinished = Number(r.TrTp) === 7
                    ? r
                    : filteredFinishedRows.find(fr => String(fr.TrInf4 || '').trim() === routeKey && String(fr.OrdNo || '').trim() === String(r.OrdNo || '').trim());

                const prodKey = String(r.ProdNo || '').trim().toUpperCase();
                const candidateLookupKey = String(r.OrdNo || '').trim() + '_' + routeKey + '_' + prodKey;
                const candidateNoFin = candidateNoFinMap.has(candidateLookupKey) ? candidateNoFinMap.get(candidateLookupKey) : null;
                const rowNoFin = refFinished ? toNumber(refFinished.NoFin) : null;
                // Multiordre can have multiple finished rows with same ProdNo/Route but different quantities.
                // Keep row-level NoFin when present; only fallback to candidate map if row quantity is missing.
                const qtaPezzi = rowNoFin !== null && rowNoFin > 0
                    ? rowNoFin
                    : (candidateNoFin !== null && candidateNoFin > 0 ? candidateNoFin : null);
                const structNoPerStr = structMap.get(prodKey) || null;
                const oldExpectedUnitWeight = refFinished ? getExpectedUnitWeight(refFinished) : null;
                const expectedUnitWeight = (structNoPerStr !== null && structNoPerStr > 0)
                    ? structNoPerStr
                    : oldExpectedUnitWeight;
                const kgProdotto = (qtaPezzi !== null && expectedUnitWeight !== null)
                    ? (expectedUnitWeight * qtaPezzi)
                    : null;
                const stats = routeStats.get(String(r.OrdNo || '').trim() + '|' + routeKey) || { kgConsumati: 0, costoLastre: 0, kgFiniti: 0, cstPrSum: 0, cstPrCount: 0 };
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
                const avgSheetCstPr = stats.cstPrCount > 0 ? (stats.cstPrSum / stats.cstPrCount) : null;
                const quotaCosto = useSpecialLaserCost
                    ? ((kgUtilizzatiEffettivi !== null && avgSheetCstPr !== null) ? (kgUtilizzatiEffettivi * avgSheetCstPr) : null)
                    : ((oldKgProdotto !== null && stats.kgFiniti > 0)
                        ? ((oldKgProdotto / stats.kgFiniti) * stats.costoLastre)
                        : null);
                const costoPerPezzo = (quotaCosto !== null && qtaPezzi !== null && qtaPezzi > 0) ? (quotaCosto / qtaPezzi) : null;
                const euroPerKgFinito = (costoPerPezzo !== null && nWgtUMedio !== null && nWgtUMedio > 0)
                    ? (costoPerPezzo / nWgtUMedio)
                    : null;
                const imageRow = refFinished || r;
                const imageItems = buildImageItems(imageRow ? imageRow.WebPg : null, imageRow ? imageRow.PictFNm : null);

                return {
                    LnNo: toNumber(r.LnNo),
                    NestingOrdNo: toNumber(r.OrdNo),
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

            const _debugLookups = filteredNestingRows.map(r => {
                const rk = String(r.TrInf4 || '').trim();
                const pk = String(r.ProdNo || '').trim().toUpperCase();
                const lk = String(r.OrdNo || '').trim() + '_' + rk + '_' + pk;
                return { OrdNo: r.OrdNo, TrInf4: r.TrInf4, ProdNo: r.ProdNo, TrTp: r.TrTp, NoFin_row: r.NoFin, lookupKey: lk, found: candidateNoFinMap.has(lk), candidateNoFin: candidateNoFinMap.get(lk) };
            });
            logEvent('LASER_DEBUG ordine=' + ordine + ' candidates=' + JSON.stringify(candidates.map(c => ({ OrdNo: c.OrdNo, TrInf4: c.TrInf4, ProdNo: c.ProdNo, NoFin: c.NoFin }))));
            logEvent('LASER_DEBUG mapEntries=' + JSON.stringify(Array.from(candidateNoFinMap.entries()).map(([k, v]) => ({ key: k, noFin: v }))));
            logEvent('LASER_DEBUG lookups=' + JSON.stringify(_debugLookups));
            const laserResult = {
                ordine,
                route: showAllRoutes ? null : effectiveRoute,
                nestingOrdNo: showAllRoutes
                    ? (nestingOrdNos.length === 1 ? nestingOrdNos[0] : null)
                    : (nestingOrdNos[0] || null),
                nestingOrdNos,
                prodNo: prodNoFilter || null,
                showAllRoutes,
                summary: {
                    KgConsumati: round(kgConsumati),
                    CostoLastre: round(costoLastre),
                    KgFiniti: round(kgFiniti),
                    SfridoKg: round(sfridoKg),
                    SfridoPct: sfridoPct === null ? null : round(sfridoPct)
                },
                products,
                _debug: {
                    candidates: candidates.map(c => ({ OrdNo: c.OrdNo, TrInf4: c.TrInf4, ProdNo: c.ProdNo, NoFin: c.NoFin })),
                    candidateNoFinMapEntries: Array.from(candidateNoFinMap.entries()).map(([k, v]) => ({ key: k, noFin: v })),
                    lookups: _debugLookups
                }
            };
            diskCache.set(laserCacheKey, laserResult, CACHE_TTL_LASER_METRICS_MS);
            return res.json(laserResult);
        } catch (err) {
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/omsaetning/accounts', async (req, res) => {
        try {
            const accounts = await omsaetningService.getAccounts();

            return res.json({
                ok: true,
                accounts
            });
        } catch (err) {
            logEvent('ERROR omsaetning/accounts: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/omsaetning/customers', async (req, res) => {
        try {
            const queryText = String(req.query.q || '').trim();
            const limit = Number(req.query.limit || 20);
            const customers = await omsaetningService.searchCustomers({ queryText, limit });

            return res.json({
                ok: true,
                customers
            });
        } catch (err) {
            logEvent('ERROR omsaetning/customers: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/omsaetning/summary', async (req, res) => {
        try {
            const fra = String(req.query.fra || '').trim();
            const til = String(req.query.til || '').trim();
            const accountCsv = String(req.query.accounts || '').trim();
            const customerCsv = String(req.query.customers || '').trim();
            const summary = await omsaetningService.getSummary({ fra, til, accountCsv, customerCsv });

            return res.json({
                ok: true,
                ...summary
            });
        } catch (err) {
            if (err && err.statusCode) {
                return res.status(err.statusCode).json({ ok: false, error: err.message || 'Ugyldig forespørgsel' });
            }
            logEvent('ERROR omsaetning/summary: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/ordreindgang/summary', async (req, res) => {
        try {
            const fraWeek = String(req.query.fraWeek || '').trim();
            const tilWeek = String(req.query.tilWeek || '').trim();
            const summary = await ordreindgangService.getSummary({ fraWeek, tilWeek });

            return res.json({
                ok: true,
                ...summary
            });
        } catch (err) {
            if (err && err.statusCode) {
                return res.status(err.statusCode).json({ ok: false, error: err.message || 'Ugyldig forespørgsel' });
            }
            logEvent('ERROR ordreindgang/summary: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/omsaetning/customer-threshold/:custno', (req, res) => {
        const custNo = String(req.params.custno || '').trim();
        if (!/^\d{1,20}$/.test(custNo)) {
            return res.status(400).json({ ok: false, error: 'Ugyldigt kundenummer' });
        }

        const threshold = omsaetningThresholdsService.getThreshold(custNo);
        const meta = omsaetningThresholdsService.getStorageMeta();
        if (!threshold) {
            return res.json({
                ok: true,
                custNo,
                warnThreshold: meta.defaultWarnThreshold,
                goodThreshold: meta.defaultGoodThreshold,
                updatedAt: null,
                exists: false,
                storageFile: meta.filePath
            });
        }

        return res.json({
            ok: true,
            custNo,
            warnThreshold: threshold.warnThreshold,
            goodThreshold: threshold.goodThreshold,
            updatedAt: threshold.updatedAt,
            exists: true,
            storageFile: meta.filePath
        });
    });

    router.post('/omsaetning/customer-threshold/:custno', express.json(), (req, res) => {
        const custNo = String(req.params.custno || '').trim();
        if (!/^\d{1,20}$/.test(custNo)) {
            return res.status(400).json({ ok: false, error: 'Ugyldigt kundenummer' });
        }

        const { warnThreshold, goodThreshold } = req.body || {};
        const saved = omsaetningThresholdsService.setThreshold(custNo, { warnThreshold, goodThreshold });
        const meta = omsaetningThresholdsService.getStorageMeta();
        if (!saved) {
            return res.status(400).json({ ok: false, error: 'Ugyldige tærskelværdier' });
        }

        return res.json({
            ok: true,
            custNo,
            warnThreshold: saved.warnThreshold,
            goodThreshold: saved.goodThreshold,
            updatedAt: saved.updatedAt,
            storageFile: meta.filePath
        });
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

        const marginOrdNos = (Array.isArray(orderListCache.data) ? orderListCache.data : [])
            .map(row => Number(row && row.OrdNo))
            .filter(ordNo => Number.isFinite(ordNo));
        let marginDone = 0;
        for (const ordNo of marginOrdNos) {
            if (orderMarginCache.has(ordNo)) {
                marginDone += 1;
                continue;
            }
            const cachedMargin = diskCache.get(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo)
                || diskCache.getStale(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo)
                || diskCache.getStale('order_margin_v6_' + ordNo);
            if (cachedMargin && cachedMargin.totalCost !== null && cachedMargin.totalCost !== undefined) {
                marginDone += 1;
            }
        }
        const marginTotal = marginOrdNos.length;

        const combinedTotal = (warmupProgress.total || 0) + marginTotal;
        const combinedDone = done + marginDone;
        const combinedPct = combinedTotal > 0 ? Math.round((combinedDone / combinedTotal) * 100) : 100;
        const ready = !orderListCache.loading
            && !warmupProgress.running
            && (warmupProgress.total === 0 || done >= warmupProgress.total)
            && (marginTotal === 0 || marginDone >= marginTotal);

        res.json({
            running: warmupProgress.running,
            total: warmupProgress.total,
            cached: warmupProgress.cached,
            loaded: warmupProgress.loaded,
            failed: warmupProgress.failed,
            done,
            pct,
            current: warmupProgress.current,
            marginDone,
            marginTotal,
            combinedDone,
            combinedTotal,
            combinedPct,
            ready
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

                    diskCache.del(AFTERCALC_CACHE_KEY_PREFIX + ordNo);
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
                        diskCache.set(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo, marginResult, CACHE_TTL_ORDER_MARGIN_MS);
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

    router.get('/desktop-update-status', (req, res) => {
        try {
            const statusFn = global.__desktopManualUpdateStatus;
            if (typeof statusFn !== 'function') {
                return res.status(503).json({
                    ok: false,
                    status: 'unavailable',
                    message: 'Opdateringsstatus er ikke tilgaengelig i denne mode.'
                });
            }

            const result = statusFn();
            return res.json(result || {
                ok: false,
                status: 'error',
                message: 'Tomt svar fra updater-status.'
            });
        } catch (err) {
            logEvent('DESKTOP-UPDATE-STATUS ERROR: ' + err.message);
            return res.status(500).json({ ok: false, status: 'error', message: err.message });
        }
    });

    router.post('/desktop-update-install', (req, res) => {
        try {
            const installFn = global.__desktopManualUpdateInstall;
            if (typeof installFn !== 'function') {
                return res.status(503).json({
                    ok: false,
                    status: 'unavailable',
                    message: 'Installering er ikke tilgaengelig i denne mode.'
                });
            }

            const result = installFn();
            logEvent('DESKTOP-UPDATE-INSTALL: status=' + String(result && result.status || 'unknown') + ', ok=' + String(!!(result && result.ok)));
            return res.json(result || {
                ok: false,
                status: 'error',
                message: 'Tomt svar fra install-funktion.'
            });
        } catch (err) {
            logEvent('DESKTOP-UPDATE-INSTALL ERROR: ' + err.message);
            return res.status(500).json({ ok: false, status: 'error', message: err.message });
        }
    });

    router.post('/open-drawing', (req, res) => {
        (async () => {
            try {
                const rawPath = String((req.body && req.body.path) || '').trim();
                const prodNo = String((req.body && req.body.prodNo) || '').trim();
                let candidatePath = rawPath;

                if (!candidatePath && prodNo) {
                    const pool = await getConnection();
                    const drawingRow = await pool.request()
                        .input('prodNo', sql.VarChar(100), prodNo)
                        .query(`
                            SELECT TOP 1 LTRIM(RTRIM(CONVERT(VARCHAR(1000), WebPg))) AS WebPg
                            FROM FreeInf2
                            WHERE LTRIM(RTRIM(CONVERT(VARCHAR(100), ProdNo))) = @prodNo
                              AND WebPg IS NOT NULL
                              AND LTRIM(RTRIM(CONVERT(VARCHAR(1000), WebPg))) <> ''
                            ORDER BY LTRIM(RTRIM(CONVERT(VARCHAR(1000), WebPg))) DESC
                        `);

                    const webPg = String((drawingRow.recordset && drawingRow.recordset[0] && drawingRow.recordset[0].WebPg) || '').trim();
                    if (webPg) {
                        candidatePath = webPg;
                    }
                }

                if (!candidatePath) {
                    return res.status(400).json({ ok: false, message: 'Path mangler.' });
                }

                const lower = candidatePath.toLowerCase();
                if (lower.indexOf('.pdf') === -1) {
                    return res.status(400).json({ ok: false, message: 'Kun PDF er tilladt.' });
                }

                const child = spawn('cmd', ['/c', 'start', '', candidatePath], {
                    windowsHide: true,
                    detached: true,
                    stdio: 'ignore'
                });
                child.unref();

                logEvent('OPEN-DRAWING: ' + candidatePath + (prodNo ? (' [prodNo=' + prodNo + ']') : ''));
                return res.json({ ok: true });
            } catch (err) {
                logEvent('OPEN-DRAWING ERROR: ' + err.message);
                return res.status(500).json({ ok: false, message: err.message });
            }
        })();
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

            const rowHasWarnings = (aftercalc) => {
                if (!aftercalc || typeof aftercalc !== 'object') return false;
                if (aftercalc.hasWarnings === true) return true;

                const lineHasWarning = (lines) => Array.isArray(lines) && lines.some(line => {
                    if (!line || typeof line !== 'object') return false;
                    if (line.HasWarning) return true;
                    const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                    const noFinValue = Number(line.NoFin || 0);
                    const noOrgValue = Number(line.NoOrg || 0);
                    return prodNoKey.startsWith('3') && noFinValue === 0 && noOrgValue > 0;
                });

                const prodOrderHasWarning = Array.isArray(aftercalc.productionOrders)
                    && aftercalc.productionOrders.some(order => order && (order.hasWarnings || lineHasWarning(order.lines)));

                return lineHasWarning(aftercalc.salesOrderLines)
                    || lineHasWarning(aftercalc.salesLines)
                    || prodOrderHasWarning;
            };

            const data = orderListCache.data.map(row => {
                const ordNoNum = Number(row.OrdNo);
                let marginInfo = orderMarginCache.get(ordNoNum);
                let warningSource = diskCache.get(AFTERCALC_CACHE_KEY_PREFIX + ordNoNum)
                    || diskCache.getStale(AFTERCALC_CACHE_KEY_PREFIX + ordNoNum)
                    || diskCache.get('aftercalc_' + ordNoNum)
                    || diskCache.getStale('aftercalc_' + ordNoNum);
                const hasWarning = rowHasWarnings(warningSource);
                if (marginInfo) {
                    marginFromMemory += 1;
                }

                if (!marginInfo) {
                    const cachedMargin = diskCache.get(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNoNum);
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
                    const staleMargin = diskCache.getStale(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNoNum)
                        || diskCache.getStale('order_margin_v6_' + ordNoNum);
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
                    HasWarning: hasWarning,
                    WarningText: hasWarning ? 'Ordren indeholder mindst én advarsel.' : '',
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

    // ── ORDER NOTES ─────────────────────────────────────────────────────────
    router.get('/order-note/:ordno', (req, res) => {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) return res.status(400).json({ error: 'Ugyldigt ordrenummer' });
        const note = orderNotesService.getNote(ordNo);
        res.json(note || { status: '', text: '', isCreditNote: false, updatedAt: null });
    });

    router.get('/order-notes-all', (req, res) => {
        res.json(orderNotesService.getAllNotes());
    });

    router.post('/order-note/:ordno', express.json(), (req, res) => {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) return res.status(400).json({ error: 'Ugyldigt ordrenummer' });
        const { status = '', text = '', isCreditNote = false } = req.body || {};
        const note = orderNotesService.setNote(ordNo, { status, text, isCreditNote });
        res.json(note || { status: '', text: '', isCreditNote: false, updatedAt: null });
    });

    router.delete('/order-note/:ordno', (req, res) => {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) return res.status(400).json({ error: 'Ugyldigt ordrenummer' });
        orderNotesService.deleteNote(ordNo);
        res.json({ ok: true });
    });

    return router;
}

module.exports = {
    createApiRouter
};
