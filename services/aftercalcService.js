function createAftercalcService({
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
    orderListMaxRows,
    orderListDaysBack,
    cacheTtlProductionSummaryMs
}) {
    const PRODUCTION_SUMMARY_CACHE_SCHEMA_VERSION = 23;
    const laserRoutePricingCache = new Map();

    function buildLineWarnings(line, extraWarnings = []) {
        const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
        const prodNoKey = String((line && line.ProdNo) || '').trim().toUpperCase();
        const noFinValue = Number((line && line.NoFin) || 0);
        const noOrgValue = Number((line && line.NoOrg) || 0);
        const purcNoValue = Number((line && line.PurcNo) || 0);
        const noInvoValue = Number((line && line.NoInvo) || 0);
        const noInvoAbValue = Number((line && line.NoInvoAb) || 0);
        const isYdelseLikeLine = isInvoiceTrackedLine(line);
        const warnings = [];

        if (prodNoKey.startsWith('3') && noFinValue === 0 && noOrgValue > 0) {
            warnings.push(key === '2'
                ? 'Inkonsekvens: materiale/rør med NoFin=0 men NoOrg>0.'
                : 'Inkonsekvens på salgsordre: produkt/rør med NoFin=0 men NoOrg>0.');
        }

        if (isYdelseLikeLine && noInvoValue === 0 && noFinValue > 0) {
            warnings.push('Mangler faktura: NoInvo er 0, bruger NoFin til kostberegning.');
        }

        if (isYdelseLikeLine && purcNoValue > 0 && noInvoAbValue > noInvoValue) {
            warnings.push('manglede indkøbsfaktura');
        }

        for (const warning of extraWarnings || []) {
            const text = String(warning || '').trim();
            if (text && !warnings.includes(text)) warnings.push(text);
        }

        return warnings;
    }

    function joinWarningMessages(messages = []) {
        const unique = [];
        for (const message of messages || []) {
            const chunks = String(message || '').split('|');
            for (const chunk of chunks) {
                const text = String(chunk || '').trim();
                if (text && !unique.includes(text)) unique.push(text);
            }
        }
        return unique.join(' | ');
    }

    function getInconsistentTubeFallbackCost(line) {
        const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
        const prodNoKey = String((line && line.ProdNo) || '').trim().toUpperCase();
        const noFinValue = Number((line && line.NoFin) || 0);
        const noOrgValue = Number((line && line.NoOrg) || 0);
        const unitCost = Number((line && line.CCstPr) || 0);

        if (key === '2' && prodNoKey.startsWith('3') && noFinValue === 0 && noOrgValue > 0 && unitCost > 0) {
            return Number(noOrgValue * unitCost);
        }

        return null;
    }

    function getOperationTimeInfo(line) {
        const effectiveMinutes = Number(getEffectiveOperationMinutes(line) || 0);
        const usesEstimatedMinutes = Boolean(isEstimatedOperationMinutesFallback(line));
        return {
            effectiveMinutes,
            usesEstimatedMinutes,
            infoText: usesEstimatedMinutes
                ? 'Færdigmeldt minutter var 0; beregnet ud fra Stykliste Minutter.'
                : ''
        };
    }

    function getYdelseCostInfo(line, invoiceSourceLine = null) {
        const sourceLine = invoiceSourceLine || line || {};
        const noInvoValue = Number(sourceLine.NoInvo || 0);
        const noFinValue = Number(sourceLine.NoFin || 0);
        const usesNoFinFallback = noInvoValue === 0 && noFinValue > 0;
        const effectiveQuantity = noInvoValue > 0 ? noInvoValue : noFinValue;
        const hasInvoice = noInvoValue > 0;
        const statusText = hasInvoice
            ? ('Faktura registreret: NoInvo = ' + noInvoValue + '.')
            : (usesNoFinFallback
                ? 'Mangler faktura: NoInvo er 0, bruger NoFin til kostberegning.'
                : 'Ingen fakturainfo fundet.');

        return {
            effectiveQuantity,
            usesNoFinFallback,
            hasInvoice,
            statusText,
            infoText: usesNoFinFallback
                ? 'Mangler faktura: NoInvo er 0, bruger NoFin til kostberegning.'
                : ''
        };
    }

    function isUProduct(prodNo) {
        return String(prodNo || '').trim().toUpperCase().startsWith('U');
    }

    function isPurchasedPartLine(line) {
        if (!line) return false;
        const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
        const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
        const purcNoValue = Number(line.PurcNo || 0);
        return key === '2' && purcNoValue > 0 && prodNoKey && !isLaserLProduct(prodNoKey);
    }

    function isInvoiceTrackedLine(line) {
        if (!line) return false;
        const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
        return key === '6' || isPurchasedPartLine(line) || isUProduct(line.ProdNo);
    }

    function findMatchingChildYdelseLine(lines, prodNo) {
        if (!Array.isArray(lines)) return null;
        const prodKey = String(prodNo || '').trim().toUpperCase();
        if (!prodKey) return lines.find(line => Number(line.LnNo || 0) === 1) || null;

        return lines.find(line => String(line.ProdNo || '').trim().toUpperCase() === prodKey)
            || lines.find(line => Number(line.LnNo || 0) === 1)
            || null;
    }

    function normalizeExpectedWeight(value) {
        const parsed = Number(value || 0);
        if (!Number.isFinite(parsed) || parsed <= 0) return 0;
        return Math.abs(parsed) >= 1000 ? (parsed / 1000) : parsed;
    }

    async function loadLaserRoutePricingData(pool, sourceOrdNo) {
        const numericOrdNo = Number(sourceOrdNo || 0);
        if (!Number.isFinite(numericOrdNo) || numericOrdNo <= 0) {
            return {
                avgSheetCstPrByRoute: new Map(),
                kgPerPieceByProdRoute: new Map(),
                unitCostByProd: new Map(),
                totalCostByProd: new Map(),
                qtyByProd: new Map()
            };
        }

        if (laserRoutePricingCache.has(numericOrdNo)) {
            return laserRoutePricingCache.get(numericOrdNo);
        }

        const pricingPromise = (async () => {
            const linkedFinishedRowsResult = await pool.request()
                .input('ordNo', sql.Numeric, numericOrdNo)
                .query(`
                    SELECT OrdNo, TrInf4, ProdNo, TrTp, NoFin, Free3, CstPr, IncCst
                    FROM OrdLn
                    WHERE TrInf2 = CAST(@ordNo AS VARCHAR(20))
                      AND TrTp = 7
                `);

            const linkedFinishedRows = linkedFinishedRowsResult.recordset || [];
            const nestingOrdNos = Array.from(new Set(
                linkedFinishedRows
                    .map(row => Number(row.OrdNo || 0))
                    .filter(value => Number.isFinite(value) && value > 0)
            ));

            if (nestingOrdNos.length === 0) {
                return {
                    avgSheetCstPrByRoute: new Map(),
                    kgPerPieceByProdRoute: new Map(),
                    unitCostByProd: new Map(),
                    totalCostByProd: new Map(),
                    qtyByProd: new Map()
                };
            }

            const ordPlaceholders = nestingOrdNos.map((_, i) => `@nest${i}`).join(', ');
            const routeRowsRequest = pool.request();
            nestingOrdNos.forEach((ordNoValue, i) => {
                routeRowsRequest.input(`nest${i}`, sql.Numeric, ordNoValue);
            });
            const routeRowsResult = await routeRowsRequest.query(`
                SELECT OrdNo, TrInf4, ProdNo, TrTp, NoFin, Free3, CstPr, IncCst
                FROM OrdLn
                WHERE OrdNo IN (${ordPlaceholders})
                  AND TrTp IN (5, 7)
            `);

            const rows = routeRowsResult.recordset || [];
            const finishedProdNos = Array.from(new Set(
                rows
                    .filter(row => Number(row.TrTp) === 7)
                    .map(row => String(row.ProdNo || '').trim().toUpperCase())
                    .filter(Boolean)
            ));

            const structMap = new Map();
            if (finishedProdNos.length > 0) {
                const placeholders = finishedProdNos.map((_, i) => `@prod${i}`).join(', ');
                const structRequest = pool.request();
                finishedProdNos.forEach((prodNo, i) => {
                    structRequest.input(`prod${i}`, sql.VarChar, prodNo);
                });
                const structResult = await structRequest.query(`
                    SELECT ProdNo, NoPerStr
                    FROM Struct
                    WHERE ProdNo IN (${placeholders})
                      AND SubProd LIKE '3%'
                `);
                for (const row of (structResult.recordset || [])) {
                    const prodKey = String(row.ProdNo || '').trim().toUpperCase();
                    const unitWeight = Number(row.NoPerStr || 0);
                    if (prodKey && unitWeight > 0 && !structMap.has(prodKey)) {
                        structMap.set(prodKey, unitWeight);
                    }
                }
            }

            // Per-product nestingOrdNos: una stampa di laser può condividere
            // route (es. "04") con prodotti diversi su nesting orders diversi.
            // Il popup /laser-route-metrics scopa il calcolo ai soli nesting
            // orders che contengono il prodotto target → riproduciamo lo stesso
            // qui, altrimenti la riga riga raddoppia/triplica i kg di lastre
            // condivise con altri prodotti.
            const nestingOrdNosByProd = new Map();
            for (const row of rows) {
                if (Number(row.TrTp) !== 7) continue;
                const prodKey = String(row.ProdNo || '').trim().toUpperCase();
                const ordKey = String(row.OrdNo || '').trim();
                if (!prodKey || !ordKey) continue;
                if (!nestingOrdNosByProd.has(prodKey)) nestingOrdNosByProd.set(prodKey, new Set());
                nestingOrdNosByProd.get(prodKey).add(ordKey);
            }

            // Stat per (OrdNo, route) — una entry per ogni combinazione, così
            // sheet e finished sono allineati esattamente come fa /laser-route-metrics.
            const routeStatsByOrdRoute = new Map();
            const finishedRowsByProd = new Map();

            for (const row of rows) {
                const routeKey = String(row.TrInf4 || '').trim().toUpperCase();
                const ordKey = String(row.OrdNo || '').trim();
                if (!routeKey || !ordKey) continue;
                const statsKey = ordKey + '|' + routeKey;

                if (!routeStatsByOrdRoute.has(statsKey)) {
                    routeStatsByOrdRoute.set(statsKey, { kgConsumati: 0, kgFiniti: 0, costoLastre: 0, cstPrSum: 0, cstPrCount: 0 });
                }
                const stats = routeStatsByOrdRoute.get(statsKey);

                if (Number(row.TrTp) === 5) {
                    const cstPr = Number(row.CstPr || 0);
                    if (cstPr > 0) {
                        stats.cstPrSum += cstPr;
                        stats.cstPrCount += 1;
                    }
                    stats.kgConsumati += Number(row.NoFin || 0);
                    stats.costoLastre += Number(row.IncCst || 0);
                    continue;
                }

                if (Number(row.TrTp) === 7) {
                    const prodKey = String(row.ProdNo || '').trim().toUpperCase();
                    const noFin = Number(row.NoFin || 0);
                    const free3ExpectedUnitWeight = normalizeExpectedWeight(row.Free3);
                    const structExpectedUnitWeight = prodKey ? normalizeExpectedWeight(structMap.get(prodKey)) : 0;
                    const oldExpectedUnitWeight = free3ExpectedUnitWeight > 0
                        ? free3ExpectedUnitWeight
                        : structExpectedUnitWeight;
                    const oldKgProdotto = oldExpectedUnitWeight > 0 && noFin > 0
                        ? (oldExpectedUnitWeight * noFin)
                        : 0;
                    stats.kgFiniti += oldKgProdotto;

                    if (prodKey && oldKgProdotto > 0 && noFin > 0) {
                        if (!finishedRowsByProd.has(prodKey)) finishedRowsByProd.set(prodKey, []);
                        finishedRowsByProd.get(prodKey).push({ statsKey, routeKey, ordKey, noFin, oldKgProdotto });
                    }
                }
            }

            const kgPerPieceByProdRoute = new Map();
            const unitCostByProdRoute = new Map();
            const totalCostByProdRoute = new Map();
            const qtyByProdRoute = new Map();
            const kgStatsByProdRoute = new Map();
            const cstPrStatsByRoute = new Map();
            const unitCostByProd = new Map();
            const totalCostByProd = new Map();
            const qtyByProd = new Map();

            for (const [prodKey, finishedRows] of finishedRowsByProd.entries()) {
                let totalCost = 0;
                let totalQty = 0;
                for (const fr of finishedRows) {
                    const stats = routeStatsByOrdRoute.get(fr.statsKey);
                    if (!stats || stats.kgFiniti <= 0) continue;
                    if (stats.costoLastre <= 0) continue;
                    const quotaCosto = (fr.oldKgProdotto / stats.kgFiniti) * stats.costoLastre;
                    const kgUtilizzatiEffettivi = (fr.oldKgProdotto / stats.kgFiniti) * stats.kgConsumati;

                    totalCost += quotaCosto;
                    totalQty += fr.noFin;

                    const routeProdKey = fr.routeKey + '|' + prodKey;
                    totalCostByProdRoute.set(routeProdKey, (totalCostByProdRoute.get(routeProdKey) || 0) + quotaCosto);
                    qtyByProdRoute.set(routeProdKey, (qtyByProdRoute.get(routeProdKey) || 0) + fr.noFin);
                    const kgAcc = kgStatsByProdRoute.get(routeProdKey) || { kgUtilizzati: 0, qty: 0 };
                    kgAcc.kgUtilizzati += kgUtilizzatiEffettivi;
                    kgAcc.qty += fr.noFin;
                    kgStatsByProdRoute.set(routeProdKey, kgAcc);

                    if (stats.cstPrCount > 0) {
                        const cs = cstPrStatsByRoute.get(fr.routeKey) || { sum: 0, count: 0 };
                        cs.sum += stats.cstPrSum;
                        cs.count += stats.cstPrCount;
                        cstPrStatsByRoute.set(fr.routeKey, cs);
                    }
                }
                if (totalQty > 0) {
                    totalCostByProd.set(prodKey, totalCost);
                    qtyByProd.set(prodKey, totalQty);
                    unitCostByProd.set(prodKey, totalCost / totalQty);
                }
            }

            for (const [routeProdKey, totalCost] of totalCostByProdRoute.entries()) {
                const qty = qtyByProdRoute.get(routeProdKey) || 0;
                if (qty > 0) unitCostByProdRoute.set(routeProdKey, totalCost / qty);
            }
            for (const [routeProdKey, kgAcc] of kgStatsByProdRoute.entries()) {
                if (kgAcc.qty > 0) kgPerPieceByProdRoute.set(routeProdKey, kgAcc.kgUtilizzati / kgAcc.qty);
            }
            const avgSheetCstPrByRoute = new Map();
            for (const [routeKey, cs] of cstPrStatsByRoute.entries()) {
                if (cs.count > 0) avgSheetCstPrByRoute.set(routeKey, cs.sum / cs.count);
            }

            return {
                avgSheetCstPrByRoute,
                kgPerPieceByProdRoute,
                unitCostByProdRoute,
                totalCostByProdRoute,
                unitCostByProd,
                totalCostByProd,
                qtyByProd
            };
        })().catch(err => {
            laserRoutePricingCache.delete(numericOrdNo);
            throw err;
        });

        laserRoutePricingCache.set(numericOrdNo, pricingPromise);
        return pricingPromise;
    }

    function getSpecialGr4LaserCostInfo(pricingData, line) {
        if (!pricingData || !line) return null;
        const routeKey = String(line.TrInf4 || '').trim().toUpperCase();
        const prodKey = String(line.ProdNo || '').trim().toUpperCase();
        const qty = Number(line.NoFin || 0);
        if (!prodKey) return null;

        const allocatedQty = Number((pricingData.qtyByProd && pricingData.qtyByProd.get(prodKey)) || 0);
        const qtyMismatch = qty > 0 && allocatedQty > 0 && Math.abs(allocatedQty - qty) > 0.0001;
        const allocationInfoText = qtyMismatch
            ? (allocatedQty > qty
                ? ('Der er registreret flere stk i nestingdata (' + allocatedQty + ') end på denne ordrelinje (' + qty + '), så laserkosten bliver fordelt og pris pr. stk kan afvige.')
                : ('Der er registreret færre stk i nestingdata (' + allocatedQty + ') end på denne ordrelinje (' + qty + '), så laserkosten bliver fordelt og pris pr. stk kan afvige.'))
            : '';

        if (routeKey) {
            const routeProdKey = routeKey + '|' + prodKey;
            const routeAllocatedUnitCost = Number((pricingData.unitCostByProdRoute && pricingData.unitCostByProdRoute.get(routeProdKey)) || 0);
            if (routeAllocatedUnitCost > 0) {
                return {
                    unitCost: parseFloat(Number(routeAllocatedUnitCost).toFixed(6)),
                    totalCost: qty > 0 ? parseFloat(Number(routeAllocatedUnitCost * qty).toFixed(2)) : null,
                    allocatedQty,
                    usesAllocationSpread: qtyMismatch,
                    infoText: allocationInfoText
                };
            }
            const avgSheetCstPr = Number(pricingData.avgSheetCstPrByRoute.get(routeKey) || 0);
            const kgPerPiece = Number(pricingData.kgPerPieceByProdRoute.get(routeProdKey) || 0);
            if (avgSheetCstPr > 0 && kgPerPiece > 0) {
                const unitCost = parseFloat(Number(avgSheetCstPr * kgPerPiece).toFixed(6));
                return {
                    unitCost,
                    totalCost: qty > 0 ? parseFloat(Number(unitCost * qty).toFixed(2)) : null,
                    allocatedQty,
                    usesAllocationSpread: qtyMismatch,
                    infoText: allocationInfoText
                };
            }
        }

        const aggregatedTotalCost = Number(pricingData.totalCostByProd.get(prodKey) || 0);
        if (aggregatedTotalCost > 0) {
            const allocationBaseQty = allocatedQty > 0 ? allocatedQty : qty;
            const fallbackUnitCost = allocationBaseQty > 0
                ? (aggregatedTotalCost / allocationBaseQty)
                : Number(pricingData.unitCostByProd.get(prodKey) || 0);
            const allocatedTotalCost = (allocationBaseQty > 0 && qty > 0)
                ? (aggregatedTotalCost * (qty / allocationBaseQty))
                : aggregatedTotalCost;
            return {
                unitCost: parseFloat(Number(fallbackUnitCost).toFixed(6)),
                totalCost: parseFloat(Number(allocatedTotalCost).toFixed(2)),
                allocatedQty,
                usesAllocationSpread: qtyMismatch,
                infoText: allocationInfoText
            };
        }

        const aggregatedUnitCost = Number(pricingData.unitCostByProd.get(prodKey) || 0);
        if (aggregatedUnitCost > 0) {
            return {
                unitCost: parseFloat(Number(aggregatedUnitCost).toFixed(6)),
                totalCost: qty > 0 ? parseFloat(Number(aggregatedUnitCost * qty).toFixed(2)) : null,
                allocatedQty,
                usesAllocationSpread: qtyMismatch,
                infoText: allocationInfoText
            };
        }

        return null;
    }

    async function getAfterCalc(ordNo) {
        const pool = await getConnection();

        try {
            const [orderResult, salesOrderLinesResult, salesLinesResult, productionLinesResult] = await Promise.all([
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                        SELECT O.OrdNo, O.TrTp, O.InvoAm, O.DInvoIF, O.Gr4, A.Nm as CustomerName
                        FROM Ord O
                        LEFT JOIN Actor A ON O.CustNo = A.CustNo
                        WHERE O.OrdNo = @ordNo
                    `),
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                        SELECT
                            OrdNo,
                            LnNo,
                            ProdNo,
                            Descr,
                            DPrice,
                            NoOrg,
                            NoFin,
                            NoInvo,
                            NoInvoAb,
                            ProdTp4,
                            CCstPr,
                            PurcNo,
                            CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost
                        FROM OrdLn
                        WHERE OrdNo = @ordNo
                        ORDER BY LnNo
                    `),
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                        SELECT 
                            OrdNo, 
                            LnNo, 
                            ProdNo, 
                            Descr, 
                            DPrice,
                            NoOrg,
                            NoFin,
                            NoInvo,
                            NoInvoAb,
                            ProdTp4,
                            CCstPr,
                            CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost
                        FROM OrdLn
                        WHERE OrdNo = @ordNo AND PurcNo IS NULL
                        ORDER BY LnNo
                    `),
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                        SELECT DISTINCT PurcNo FROM OrdLn
                        WHERE OrdNo = @ordNo AND PurcNo IS NOT NULL
                    `)
            ]);

            if (orderResult.recordset.length === 0) {
                return { error: 'Ordre ikke fundet' };
            }

            const orderHeader = orderResult.recordset[0];
            const useRouteSpecificLaserCost = true;

            const prodNosForDrawings = [
                ...salesOrderLinesResult.recordset.map(r => r.ProdNo),
                ...salesLinesResult.recordset.map(r => r.ProdNo)
            ];
            const drawingByProdNo = await getLatestDrawingByProdNo(pool, prodNosForDrawings, logEvent);

            const salesLines = salesLinesResult.recordset
                .filter(line => !isGloballyExcludedProdNo(line.ProdNo))
                .map(line => {
                    const tubeFallbackCost = getInconsistentTubeFallbackCost(line);
                    // Rabat/IsDiscountLine-logikken er fjernet: linjer med salgspris=0 (typisk underlinjer
                    // af et hovedprodukt) skal stadig medregne deres kostpris i ordrens totalkost.
                    const effectiveLineCost = tubeFallbackCost !== null ? tubeFallbackCost : (line.LineCost || 0);
                    const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                    const warningMessages = buildLineWarnings(line);
                    const displayQuantity = Number(line.NoFin || 0) === 0 && Number(line.NoOrg || 0) > 0
                        ? Number(line.NoOrg || 0)
                        : Number(line.NoFin || 0);
                    return {
                        ...line,
                        IsDiscountLine: false,
                        DisplayQuantity: parseFloat(Number(displayQuantity).toFixed(2)),
                        HasWarning: warningMessages.length > 0,
                        WarningText: joinWarningMessages(warningMessages),
                        EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2)),
                        DrawingWebPg: drawingByProdNo.get(prodNoKey) || null
                    };
                });
            const salesLinesTotalCost = salesLines.reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);

            const salesOrderLines = salesOrderLinesResult.recordset
                .filter(line => !isGloballyExcludedProdNo(line.ProdNo))
                .map(line => {
                    const tubeFallbackCost = getInconsistentTubeFallbackCost(line);
                    // Rabat/IsDiscountLine-logikken er fjernet: linjer med salgspris=0 (typisk underlinjer
                    // af et hovedprodukt) skal stadig medregne deres kostpris i ordrens totalkost.
                    const effectiveLineCost = tubeFallbackCost !== null ? tubeFallbackCost : (line.LineCost || 0);
                    const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                    const warningMessages = buildLineWarnings(line);
                    const displayQuantity = Number(line.NoFin || 0) === 0 && Number(line.NoOrg || 0) > 0
                        ? Number(line.NoOrg || 0)
                        : Number(line.NoFin || 0);
                    return {
                        ...line,
                        IsDiscountLine: false,
                        DisplayQuantity: parseFloat(Number(displayQuantity).toFixed(2)),
                        HasWarning: warningMessages.length > 0,
                        WarningText: joinWarningMessages(warningMessages),
                        EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2)),
                        DrawingWebPg: drawingByProdNo.get(prodNoKey) || null
                    };
                });

            const productionOrderDetailsCache = new Map();

            // Cache dei prezzi DPrice dagli ordini di acquisto (TrTp=6).
            // Serve per usare il prezzo reale pagato al fornitore invece del CCstPr standard.
            const purchaseOrderDPriceCache = new Map();

            async function getPurchaseOrderDPriceMap(purcNo) {
                const cacheKey = String(purcNo);
                if (purchaseOrderDPriceCache.has(cacheKey)) {
                    return purchaseOrderDPriceCache.get(cacheKey);
                }
                const promise = pool.request()
                    .input('poNo', sql.Numeric, purcNo)
                    .query(`
                        SELECT L.ProdNo, L.DPrice
                        FROM OrdLn L
                        INNER JOIN Ord O ON O.OrdNo = L.OrdNo AND O.TrTp = 6
                        WHERE L.OrdNo = @poNo AND L.DPrice > 0
                    `)
                    .then(res => {
                        const map = new Map();
                        for (const row of res.recordset) {
                            const k = String(row.ProdNo || '').trim().toUpperCase();
                            if (!map.has(k)) map.set(k, Number(row.DPrice));
                        }
                        return map;
                    })
                    .catch(() => new Map());
                purchaseOrderDPriceCache.set(cacheKey, promise);
                return promise;
            }

            async function loadProductionOrderDetails(prodOrdNo, visited = new Set(), options = {}) {
                const numericProdOrdNo = Number(prodOrdNo);
                if (!Number.isFinite(numericProdOrdNo)) {
                    return { lines: [], totalCost: 0 };
                }

                const useSpecialLaserCost = Boolean(options.useSpecialLaserCost);
                const detailsCacheKey = String(numericProdOrdNo) + '|' + (useSpecialLaserCost ? 'gr4-3' : 'default');

                if (productionOrderDetailsCache.has(detailsCacheKey)) {
                    return productionOrderDetailsCache.get(detailsCacheKey);
                }

                if (visited.has(numericProdOrdNo)) {
                    return { lines: [], totalCost: 0 };
                }

                const nextVisited = new Set(visited);
                nextVisited.add(numericProdOrdNo);

                const detailsPromise = (async () => {
                    const prodLinesResult = await pool.request()
                        .input('purcNo', sql.Numeric, numericProdOrdNo)
                        .query(`
                            SELECT 
                                OrdLn.OrdNo, 
                                OrdLn.LnNo, 
                                OrdLn.ProdNo, 
                                OrdLn.Descr, 
                                OrdLn.DPrice,
                                OrdLn.NoOrg,
                                OrdLn.NoFin,
                                OrdLn.NoInvo,
                                OrdLn.NoInvoAb,
                                OrdLn.CCstPr,
                                OrdLn.PurcNo,
                                OrdLn.TrInf2,
                                OrdLn.TrInf4,
                                OrdLn.ProdTp4,
                                (
                                    SELECT TOP 1 A.Nm
                                    FROM ProdTr P
                                    LEFT JOIN Actor A ON A.EmpNo = P.EmpNo
                                    WHERE P.OrdNo = @purcNo
                                      AND P.OrdLnNo = OrdLn.LnNo
                                    ORDER BY P.FinDt DESC, P.FinTm DESC
                                ) AS HvemNm,
                                CAST(OrdLn.NoFin * OrdLn.CCstPr AS DECIMAL(10,2)) AS LineCost,
                                (
                                    SELECT SUM(CAST(n.CstPr AS DECIMAL(18,6)) * CAST(n.NoFin AS DECIMAL(18,6)))
                                         / NULLIF(SUM(CAST(n.NoFin AS DECIMAL(18,6))), 0)
                                    FROM OrdLn n
                                    WHERE n.TrInf2 = CAST(@purcNo AS VARCHAR(20))
                                      AND n.ProdNo = OrdLn.ProdNo
                                ) AS NestingCost
                            FROM OrdLn
                            WHERE OrdLn.OrdNo = @purcNo
                            ORDER BY LnNo
                        `);

                    const hasLaserLLines = (prodLinesResult.recordset || []).some(row => {
                        const key = (row.ProdTp4 === null || row.ProdTp4 === undefined) ? 'NA' : String(row.ProdTp4);
                        return key === '2' && isLaserLProduct(row.ProdNo);
                    });
                    const specialLaserPricingData = hasLaserLLines
                        ? await loadLaserRoutePricingData(pool, numericProdOrdNo).catch(() => null)
                        : null;

                    const lines = [];
                    let total = 0;
                    let hasWarnings = false;
                    let hasEstimatedOperationTime = false;

                    for (const rawLine of prodLinesResult.recordset) {
                        if (isGloballyExcludedProdNo(rawLine.ProdNo)) continue;

                        const line = adjustOperationLinePricing({ ...rawLine });
                        const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                        const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                        const excludeRProductsInSubOrders = Boolean(options.excludeRProductsInSubOrders);
                        if (excludeRProductsInSubOrders && prodNoKey.startsWith('R')) continue;
                        if (key === '1' && isExcludedOperationProdNo(line.ProdNo)) continue;
                        if (key === '4' && prodNoKey.startsWith('R')) continue;

                        const noFinValue = Number(line.NoFin || 0);
                        const noOrgValue = Number(line.NoOrg || 0);
                        const isTubeMaterialLine = key === '2' && prodNoKey.startsWith('3');
                        const isInvoiceTracked = isInvoiceTrackedLine(line);
                        const operationTimeInfo = key === '1'
                            ? getOperationTimeInfo(line)
                            : { effectiveMinutes: noFinValue, usesEstimatedMinutes: false, infoText: '' };
                        let ydelseInvoiceSourceLine = null;
                        if (isInvoiceTracked && line.PurcNo && Number(line.PurcNo) !== 0 && !nextVisited.has(Number(line.PurcNo))) {
                            const childYdelseDetails = await loadProductionOrderDetails(Number(line.PurcNo), nextVisited, {
                                ...options,
                                excludeRProductsInSubOrders: true
                            });
                            ydelseInvoiceSourceLine = findMatchingChildYdelseLine(childYdelseDetails && childYdelseDetails.lines, line.ProdNo);
                        }
                        const ydelseCostInfo = isInvoiceTracked
                            ? getYdelseCostInfo(line, ydelseInvoiceSourceLine)
                            : { effectiveQuantity: noFinValue, usesNoFinFallback: false, hasInvoice: false, statusText: '', infoText: '' };
                        const ydelseSourceQuantity = Number((ydelseInvoiceSourceLine && (ydelseInvoiceSourceLine.DisplayQuantity ?? ydelseInvoiceSourceLine.NoFin)) || 0);
                        // Per "purchased part" (ProdTp4=2 con PurcNo → ordine acquisto):
                        // usa DPrice dall'ordine acquisto (prezzo reale) invece di CCstPr (standard con overhead ~42%).
                        // Per ydelse/servizi: mantiene la logica basata su fattura (NoInvo × CCstPr fonte).
                        const ydelseSourceUnitCost = isInvoiceTracked
                            ? (isPurchasedPartLine(line) && ydelseInvoiceSourceLine && Number(ydelseInvoiceSourceLine.DPrice || 0) > 0
                                ? Number(ydelseInvoiceSourceLine.DPrice)
                                : ((ydelseSourceQuantity > 0 && ydelseInvoiceSourceLine && ydelseInvoiceSourceLine.EffectiveLineCost !== undefined && ydelseInvoiceSourceLine.EffectiveLineCost !== null)
                                    ? (Number(ydelseInvoiceSourceLine.EffectiveLineCost || 0) / ydelseSourceQuantity)
                                    : Number((ydelseInvoiceSourceLine && (ydelseInvoiceSourceLine.CCstPr ?? ydelseInvoiceSourceLine.DPrice ?? ydelseInvoiceSourceLine.DisplayUnitCost)) ?? (line.CCstPr ?? line.DPrice ?? 0))))
                            : Number(line.CCstPr || 0);
                        const displayQuantity = key === '1'
                            ? operationTimeInfo.effectiveMinutes
                            : (isPurchasedPartLine(line)
                                ? (noFinValue || noOrgValue)
                                : (isInvoiceTracked
                                    ? ydelseCostInfo.effectiveQuantity
                                    : ((isTubeMaterialLine && noFinValue === 0 && noOrgValue > 0)
                                        ? noOrgValue
                                        : noFinValue)));
                        let effectiveLineCost = Number(line.LineCost || 0);
                        let childProductionTotalCost = null;
                        let specialLaserCostInfo = null;
                        const warningMessages = buildLineWarnings(line, ydelseCostInfo.usesNoFinFallback ? [ydelseCostInfo.infoText] : []);

                        if (key === '1') {
                            effectiveLineCost = Number(operationTimeInfo.effectiveMinutes * (line.CCstPr || 0));
                        } else if (isInvoiceTracked) {
                            const costQty = isPurchasedPartLine(line) ? (noFinValue || noOrgValue) : ydelseCostInfo.effectiveQuantity;
                            effectiveLineCost = Number(costQty * ydelseSourceUnitCost);
                        } else if (key === '2' && isLaserLProduct(line.ProdNo)) {
                            specialLaserCostInfo = specialLaserPricingData
                                ? getSpecialGr4LaserCostInfo(specialLaserPricingData, line)
                                : null;
                            if (specialLaserCostInfo && specialLaserCostInfo.unitCost !== null) {
                                line.NestingCost = specialLaserCostInfo.unitCost;
                            }
                            const hasNestingCost = Number(line.NestingCost || 0) > 0;
                            effectiveLineCost = (specialLaserCostInfo && specialLaserCostInfo.totalCost !== null)
                                ? Number(specialLaserCostInfo.totalCost)
                                : (hasNestingCost
                                    ? Number((line.NestingCost || 0) * (line.NoFin || 0))
                                    : Number(line.LineCost || 0));
                        } else if (isTubeMaterialLine && noFinValue === 0 && noOrgValue > 0) {
                            const tubeFallbackCost = getInconsistentTubeFallbackCost(line);
                            effectiveLineCost = tubeFallbackCost !== null ? tubeFallbackCost : Number(line.LineCost || 0);
                        } else if (key === '4' && line.PurcNo && Number(line.PurcNo) !== 0) {
                            const childDetails = await loadProductionOrderDetails(Number(line.PurcNo), nextVisited, {
                                ...options,
                                excludeRProductsInSubOrders: true
                            });
                            childProductionTotalCost = Number(childDetails.totalCost || 0);
                            effectiveLineCost = childProductionTotalCost;
                            if (childDetails.hasWarnings) {
                                warningMessages.push(childDetails.warningText || 'Underliggende produktionsordre har en advarsel.');
                            }
                        } else if (line.PurcNo && Number(line.PurcNo) !== 0) {
                            // Riga materiale allocata a un ordine di acquisto (TrTp=6):
                            // usa il prezzo effettivo DPrice dall'ordine allocato invece del CCstPr standard.
                            // Funziona ricorsivamente: ogni sotto-produzione che ha materiali allocati
                            // usa anch'essa questo percorso.
                            const priceMap = await getPurchaseOrderDPriceMap(Number(line.PurcNo));
                            const actualDPrice = priceMap.get(prodNoKey);
                            if (actualDPrice !== undefined && actualDPrice > 0) {
                                const qty = noFinValue > 0 ? noFinValue : noOrgValue;
                                effectiveLineCost = parseFloat(Number(actualDPrice * qty).toFixed(2));
                            }
                            // Se DPrice non disponibile (ordine non è acquisto, o prodotto non trovato):
                            // rimane l'effectiveLineCost = LineCost (CCstPr × NoFin) come fallback
                        }

                        line.ChildProductionTotalCost = childProductionTotalCost === null
                            ? null
                            : parseFloat(Number(childProductionTotalCost).toFixed(2));
                        line.DisplayQuantity = parseFloat(Number(displayQuantity).toFixed(2));
                        line.EffectiveOperationMinutes = key === '1'
                            ? parseFloat(Number(operationTimeInfo.effectiveMinutes).toFixed(2))
                            : null;
                        line.UsesEstimatedOperationTime = Boolean(operationTimeInfo.usesEstimatedMinutes);
                        line.EstimatedTimeText = operationTimeInfo.infoText;
                        line.IsInvoiceTracked = Boolean(isInvoiceTracked);
                        line.HasInvoice = isInvoiceTracked ? Boolean(ydelseCostInfo.hasInvoice) : null;
                        line.InvoiceStatusText = isInvoiceTracked ? String(ydelseCostInfo.statusText || '') : '';
                        line.UsesMissingInvoiceFallback = Boolean(isInvoiceTracked && ydelseCostInfo.usesNoFinFallback);
                        line.MissingInvoiceText = isInvoiceTracked ? ydelseCostInfo.infoText : '';
                        line.UsesLaserAllocationSpread = Boolean(specialLaserCostInfo && specialLaserCostInfo.usesAllocationSpread);
                        line.LaserAllocationText = specialLaserCostInfo ? String(specialLaserCostInfo.infoText || '') : '';
                        line.HasWarning = warningMessages.length > 0;
                        line.ChildHasWarning = key === '4' && warningMessages.some(msg => msg.includes('Underliggende produktionsordre'));
                        line.WarningText = joinWarningMessages(warningMessages);
                        line.EffectiveLineCost = parseFloat(Number(effectiveLineCost || 0).toFixed(2));
                        if (line.HasWarning) hasWarnings = true;
                        if (line.UsesEstimatedOperationTime) hasEstimatedOperationTime = true;
                        lines.push(line);

                        if (line.LnNo === 1 || key === '0' || key === '3' || key === '5') continue;
                        total += (line.EffectiveLineCost || 0);
                    }

                    return {
                        lines,
                        totalCost: parseFloat(Number(total).toFixed(2)),
                        hasWarnings,
                        hasEstimatedOperationTime,
                        warningText: joinWarningMessages(lines.map(line => line.WarningText))
                    };
                })();

                productionOrderDetailsCache.set(detailsCacheKey, detailsPromise);
                return detailsPromise;
            }

            const PROD_ORDER_CONCURRENCY = 5;
            async function runProdOrderTasks(items, worker) {
                const out = new Array(items.length);
                let next = 0;
                async function pump() {
                    while (true) {
                        const i = next++;
                        if (i >= items.length) return;
                        out[i] = await worker(items[i]);
                    }
                }
                await Promise.all(Array.from({ length: Math.min(PROD_ORDER_CONCURRENCY, items.length) }, pump));
                return out;
            }
            const productionOrderResults = await runProdOrderTasks(productionLinesResult.recordset, async (prodLine) => {
                const purcNo = prodLine.PurcNo;

                const [prodOrderResult, prodDetails] = await Promise.all([
                    pool.request()
                        .input('purcNo', sql.Numeric, purcNo)
                        .query(`
                            SELECT O.OrdNo, A.Nm, O.TrTp, O.InvoAm
                            FROM Ord O
                            JOIN Actor A ON O.CustNo = A.CustNo
                            WHERE OrdNo = @purcNo
                        `),
                    loadProductionOrderDetails(purcNo, new Set(), { useSpecialLaserCost: useRouteSpecificLaserCost })
                ]);

                if (prodOrderResult.recordset.length === 0) {
                    return null;
                }
                const prodOrder = prodOrderResult.recordset[0];

                if (Array.isArray(prodDetails.lines) && prodDetails.lines.some(line => {
                    const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                    return key === '2' && isLaserLProduct(line.ProdNo);
                })) {
                    const specialLaserPricingData = await loadLaserRoutePricingData(pool, Number(purcNo)).catch(() => null);
                    if (specialLaserPricingData) {
                        let adjustedTotalCost = 0;
                        prodDetails.lines = prodDetails.lines.map(rawLine => {
                            const line = { ...rawLine };
                            const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                            if (key === '2' && isLaserLProduct(line.ProdNo)) {
                                const specialLaserCostInfo = getSpecialGr4LaserCostInfo(specialLaserPricingData, line);
                                line.UsesLaserAllocationSpread = Boolean(specialLaserCostInfo && specialLaserCostInfo.usesAllocationSpread);
                                line.LaserAllocationText = specialLaserCostInfo ? String(specialLaserCostInfo.infoText || '') : '';
                                if (specialLaserCostInfo && Number(specialLaserCostInfo.unitCost) > 0) {
                                    line.NestingCost = parseFloat(Number(specialLaserCostInfo.unitCost).toFixed(6));
                                    line.EffectiveLineCost = parseFloat(Number(
                                        specialLaserCostInfo.totalCost !== null
                                            ? specialLaserCostInfo.totalCost
                                            : (specialLaserCostInfo.unitCost * Number(line.NoFin || 0))
                                    ).toFixed(2));
                                }
                            }
                            if (!(Number(line.LnNo || 0) === 1 || key === '0' || key === '3' || key === '5')) {
                                adjustedTotalCost += Number(line.EffectiveLineCost || 0);
                            }
                            return line;
                        });
                        prodDetails.totalCost = parseFloat(Number(adjustedTotalCost).toFixed(2));
                    }
                }

                return {
                    ordNo: purcNo,
                    trTp: prodOrder.TrTp,
                    revenue: prodOrder.InvoAm || 0,
                    lines: prodDetails.lines,
                    totalCost: prodDetails.totalCost,
                    hasWarnings: Boolean(prodDetails.hasWarnings),
                    hasEstimatedOperationTime: Boolean(prodDetails.hasEstimatedOperationTime),
                    warningText: String(prodDetails.warningText || '')
                };
            });
            const productionOrders = productionOrderResults.filter(Boolean);

            const productionTotalByOrdNo = new Map(
                productionOrders.map(po => [Number(po.ordNo), Number(po.totalCost || 0)])
            );
            const productionWarningByOrdNo = new Map(
                productionOrders.map(po => [Number(po.ordNo), Boolean(po.hasWarnings)])
            );
            const productionWarningTextByOrdNo = new Map(
                productionOrders.map(po => [Number(po.ordNo), String(po.warningText || '')])
            );

            const salesOrderLinesWithProductionTotal = salesOrderLines.map(line => {
                const purcNo = line.PurcNo ? Number(line.PurcNo) : 0;
                const productionTotal = purcNo ? productionTotalByOrdNo.get(purcNo) : undefined;
                const productionHasWarning = purcNo ? Boolean(productionWarningByOrdNo.get(purcNo)) : false;
                const productionWarningText = purcNo ? String(productionWarningTextByOrdNo.get(purcNo) || '') : '';
                const warningText = joinWarningMessages([line.WarningText]);
                const hasWarning = Boolean(line.HasWarning) && warningText.length > 0;

                if (productionTotal !== undefined && !line.IsDiscountLine) {
                    const roundedTotal = parseFloat(Number(productionTotal).toFixed(2));
                    return {
                        ...line,
                        ProductionOrderTotalCost: roundedTotal,
                        EffectiveLineCost: roundedTotal,
                        LinkedProductionHasWarning: productionHasWarning,
                        LinkedProductionWarningText: productionWarningText,
                        HasWarning: hasWarning,
                        WarningText: warningText
                    };
                }

                return {
                    ...line,
                    ProductionOrderTotalCost: productionTotal !== undefined ? parseFloat(Number(productionTotal).toFixed(2)) : null,
                    LinkedProductionHasWarning: productionHasWarning,
                    LinkedProductionWarningText: productionWarningText,
                    HasWarning: hasWarning,
                    WarningText: warningText
                };
            });

            const salesNoPOLines = salesOrderLinesWithProductionTotal.filter(line => !line.PurcNo || line.PurcNo === 0);
            const salesNoPOTotalCost = salesNoPOLines.reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);

            const includedProductionOrdNos = new Set(
                salesOrderLinesWithProductionTotal
                    .filter(line => line.PurcNo && line.PurcNo !== 0 && !line.IsDiscountLine)
                    .map(line => Number(line.PurcNo))
            );

            const productionTotalCost = productionOrders.reduce((sum, ord) => {
                if (!includedProductionOrdNos.has(Number(ord.ordNo))) return sum;
                return sum + (ord.totalCost || 0);
            }, 0);
            const totalCost = salesNoPOTotalCost + productionTotalCost;
            const totalRevenue = Number(orderHeader.InvoAm || 0) + Number(orderHeader.DInvoIF || 0);
            const margin = totalRevenue - totalCost;
            const marginPercentage = totalCost > 0 ? ((totalRevenue / totalCost) * 100).toFixed(2) : 0;
            const hasWarnings = salesOrderLinesWithProductionTotal.some(line => line.HasWarning)
                || salesLines.some(line => line.HasWarning)
                || productionOrders.some(order => order.hasWarnings);

            return {
                orderHeader,
                hasWarnings,
                warningText: hasWarnings
                    ? joinWarningMessages([
                        ...salesOrderLinesWithProductionTotal.map(line => line.WarningText),
                        ...salesLines.map(line => line.WarningText),
                        ...productionOrders.map(order => order.warningText)
                    ])
                    : '',
                salesOrderLines: salesOrderLinesWithProductionTotal,
                salesLines,
                salesLinesTotalCost: parseFloat(salesLinesTotalCost.toFixed(2)),
                productionOrders,
                summary: {
                    totalRevenue: parseFloat(totalRevenue.toFixed(2)),
                    totalCost: parseFloat(totalCost.toFixed(2)),
                    margin: parseFloat(margin.toFixed(2)),
                    marginPercentage,
                    hasInvoiceWarning: productionOrders.some(ord => Array.isArray(ord.lines) && ord.lines.some(line => line.UsesMissingInvoiceFallback))
                }
            };
        } catch (err) {
            console.error('Errore in getAfterCalc:', err);
            throw err;
        }
    }

    async function fetchOrderListBase() {
        const pool = await getConnection();
        const result = await pool.request().query(`
            WITH BaseOrders AS (
                SELECT TOP ${orderListMaxRows}
                    O.SelBuy,
                    O.OrdNo,
                    O.Gr4,
                    C.Nm AS CustomerName,
                    SU.Usr AS SellerUsr,
                    O.LstInvDt,
                    O.InvoAm
                FROM Ord O
                LEFT JOIN Actor C ON C.CustNo = O.CustNo
                OUTER APPLY (
                    SELECT TOP 1 A.Usr
                    FROM Actor A
                    WHERE LTRIM(RTRIM(CONVERT(VARCHAR(50), A.EmpNo))) = LTRIM(RTRIM(CONVERT(VARCHAR(50), O.SelBuy)))
                ) SU
                WHERE O.InvoNo IS NOT NULL AND O.InvoNo <> ''
                  AND O.InvoAm > 0
                  AND O.LstInvDt >= CAST(CONVERT(VARCHAR(8), DATEADD(day, -${orderListDaysBack}, GETDATE()), 112) AS INT)
                ORDER BY O.LstInvDt DESC
            )
            SELECT
                B.SelBuy,
                B.OrdNo,
                B.Gr4,
                B.CustomerName,
                B.SellerUsr,
                B.LstInvDt,
                B.InvoAm
            FROM BaseOrders B
            ORDER BY B.LstInvDt DESC, B.OrdNo DESC
        `);

        return result.recordset;
    }

    async function getProductionSummary(ordNo, visited = new Set(), options = {}) {
        const numericOrdNo = Number(ordNo);
        if (!Number.isFinite(numericOrdNo)) {
            throw new Error('Ordrenummer ugyldigt');
        }

        if (visited.has(numericOrdNo)) {
            return {
                ordNo: numericOrdNo,
                lines: [],
                hasWarnings: false,
                totalCost: 0
            };
        }

        const nextVisited = new Set(visited);
        nextVisited.add(numericOrdNo);

        const useSpecialLaserCost = Number(options && options.orderGr4 || 0) === 3;
        const summaryCacheKey = 'prod_summary_' + numericOrdNo + (useSpecialLaserCost ? '_gr4_3' : '');
        const cachedSummary = diskCache.get(summaryCacheKey);
        if (cachedSummary && cachedSummary.cacheSchemaVersion === PRODUCTION_SUMMARY_CACHE_SCHEMA_VERSION) {
            return cachedSummary;
        }

        const pool = await getConnection();
        const linesResult = await pool.request()
            .input('ordNo', sql.Numeric, numericOrdNo)
            .query(`
                SELECT
                    OrdLn.LnNo,
                    OrdLn.ProdNo,
                    OrdLn.Descr,
                    OrdLn.NoOrg,
                    OrdLn.NoFin,
                    OrdLn.NoInvo,
                    OrdLn.NoInvoAb,
                    OrdLn.DPrice,
                    OrdLn.CCstPr,
                    OrdLn.TrInf2,
                    OrdLn.TrInf4,
                    OrdLn.ProdTp4,
                    OrdLn.PurcNo,
                    (
                        SELECT TOP 1 A.Nm
                        FROM ProdTr P
                        LEFT JOIN Actor A ON A.EmpNo = P.EmpNo
                        WHERE P.OrdNo = @ordNo
                          AND P.OrdLnNo = OrdLn.LnNo
                        ORDER BY P.FinDt DESC, P.FinTm DESC
                    ) AS HvemNm,
                    CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost,
                    (
                        SELECT SUM(CAST(n.CstPr AS DECIMAL(18,6)) * CAST(n.NoFin AS DECIMAL(18,6)))
                             / NULLIF(SUM(CAST(n.NoFin AS DECIMAL(18,6))), 0)
                        FROM OrdLn n
                        WHERE n.TrInf2 = CAST(@ordNo AS VARCHAR(20))
                          AND n.ProdNo = OrdLn.ProdNo
                    ) AS NestingCost
                FROM OrdLn
                WHERE OrdLn.OrdNo = @ordNo
                  AND (OrdLn.LnNo = 1 OR OrdLn.NoFin > 0 OR OrdLn.NoOrg > 0 OR OrdLn.NoInvo > 0)
                ORDER BY LnNo
            `);

        const hasLaserLLinesSales = (linesResult.recordset || []).some(row => {
            const key = (row.ProdTp4 === null || row.ProdTp4 === undefined) ? 'NA' : String(row.ProdTp4);
            return key === '2' && isLaserLProduct(row.ProdNo);
        });
        const specialLaserPricingData = hasLaserLLinesSales
            ? await loadLaserRoutePricingData(pool, numericOrdNo).catch(() => null)
            : null;

        const lines = linesResult.recordset
            .filter(rawLine => !isGloballyExcludedProdNo(rawLine.ProdNo))
            .map(rawLine => {
                const line = adjustOperationLinePricing({ ...rawLine });
                const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                if (prodNoKey.startsWith('R')) {
                    return null;
                }
                if (key === '1' && isExcludedOperationProdNo(line.ProdNo)) {
                    return null;
                }
                if (key === '4' && prodNoKey.startsWith('R')) {
                    return null;
                }
                const noFinValue = Number(line.NoFin || 0);
                const noOrgValue = Number(line.NoOrg || 0);
                const isTubeMaterialLine = key === '2' && prodNoKey.startsWith('3');
                const isInvoiceTracked = isInvoiceTrackedLine(line);
                const operationTimeInfo = key === '1'
                    ? getOperationTimeInfo(line)
                    : { effectiveMinutes: noFinValue, usesEstimatedMinutes: false, infoText: '' };
                const ydelseCostInfo = isInvoiceTracked
                    ? getYdelseCostInfo(line)
                    : { effectiveQuantity: noFinValue, usesNoFinFallback: false, hasInvoice: false, statusText: '', infoText: '' };
                const displayQuantity = key === '1'
                    ? operationTimeInfo.effectiveMinutes
                    : (isInvoiceTracked
                        ? ydelseCostInfo.effectiveQuantity
                        : ((isTubeMaterialLine && noFinValue === 0 && noOrgValue > 0)
                            ? noOrgValue
                            : noFinValue));
                const warningMessages = buildLineWarnings(line);
                const tubeFallbackCost = getInconsistentTubeFallbackCost(line);
                const specialLaserCostInfo = (key === '2' && isLaserLProduct(line.ProdNo) && specialLaserPricingData)
                    ? getSpecialGr4LaserCostInfo(specialLaserPricingData, line)
                    : null;
                if (specialLaserCostInfo && specialLaserCostInfo.unitCost !== null) {
                    line.NestingCost = specialLaserCostInfo.unitCost;
                }
                const recalculatedHasNestingCost = Number(line.NestingCost || 0) > 0;
                const effectiveLineCost = key === '1'
                    ? Number(operationTimeInfo.effectiveMinutes * (line.CCstPr || 0))
                    : (isInvoiceTracked
                        ? Number(ydelseCostInfo.effectiveQuantity * (line.CCstPr || 0))
                        : (key === '2' && isLaserLProduct(line.ProdNo)
                            ? ((specialLaserCostInfo && specialLaserCostInfo.totalCost !== null)
                                ? Number(specialLaserCostInfo.totalCost)
                                : (recalculatedHasNestingCost ? ((line.NestingCost || 0) * (line.NoFin || 0)) : (line.LineCost || 0)))
                            : (tubeFallbackCost !== null ? tubeFallbackCost : Number(line.LineCost || 0))));

                const roundedEffectiveLineCost = parseFloat(Number(effectiveLineCost).toFixed(2));
                const displayUnitCost = Number(displayQuantity || 0) > 0
                    ? parseFloat(Number(roundedEffectiveLineCost / displayQuantity).toFixed(2))
                    : parseFloat(Number(line.CCstPr || 0).toFixed(2));

                return {
                    ...line,
                    DisplayQuantity: parseFloat(Number(displayQuantity).toFixed(2)),
                    DisplayUnitCost: displayUnitCost,
                    EffectiveOperationMinutes: key === '1'
                        ? parseFloat(Number(operationTimeInfo.effectiveMinutes).toFixed(2))
                        : null,
                    UsesEstimatedOperationTime: Boolean(operationTimeInfo.usesEstimatedMinutes),
                    EstimatedTimeText: operationTimeInfo.infoText,
                    IsInvoiceTracked: Boolean(isInvoiceTracked),
                    HasInvoice: isInvoiceTracked ? Boolean(ydelseCostInfo.hasInvoice) : null,
                    InvoiceStatusText: isInvoiceTracked ? String(ydelseCostInfo.statusText || '') : '',
                    UsesMissingInvoiceFallback: Boolean(isInvoiceTracked && ydelseCostInfo.usesNoFinFallback),
                    MissingInvoiceText: isInvoiceTracked ? ydelseCostInfo.infoText : '',
                    UsesLaserAllocationSpread: Boolean(specialLaserCostInfo && specialLaserCostInfo.usesAllocationSpread),
                    LaserAllocationText: specialLaserCostInfo ? String(specialLaserCostInfo.infoText || '') : '',
                    HasWarning: warningMessages.length > 0,
                    WarningText: joinWarningMessages(warningMessages),
                    EffectiveLineCost: roundedEffectiveLineCost
                };
            })
            .filter(Boolean);

        for (const line of lines) {
            const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
            const childOrdNo = Number(line.PurcNo || 0);
            if (!Number.isFinite(childOrdNo) || childOrdNo <= 0 || nextVisited.has(childOrdNo)) {
                continue;
            }

            const isInvoiceTracked = isInvoiceTrackedLine(line);
            if (isInvoiceTracked) {
                const childSummary = await getProductionSummary(childOrdNo, nextVisited, options);
                const matchedChildLine = findMatchingChildYdelseLine(childSummary && childSummary.lines, line.ProdNo);
                if (matchedChildLine) {
                    const ydelseCostInfo = getYdelseCostInfo(line, matchedChildLine);
                    const matchedChildQty = Number((matchedChildLine.DisplayQuantity !== undefined && matchedChildLine.DisplayQuantity !== null)
                        ? matchedChildLine.DisplayQuantity
                        : (matchedChildLine.NoFin || 0));
                    const matchedChildUnitCost = matchedChildQty > 0 && matchedChildLine.EffectiveLineCost !== undefined && matchedChildLine.EffectiveLineCost !== null
                        ? Number(matchedChildLine.EffectiveLineCost || 0) / matchedChildQty
                        : Number(matchedChildLine.CCstPr || matchedChildLine.DPrice || matchedChildLine.DisplayUnitCost || line.CCstPr || 0);
                    line.DisplayQuantity = parseFloat(Number(ydelseCostInfo.effectiveQuantity).toFixed(2));
                    line.EffectiveLineCost = parseFloat(Number((ydelseCostInfo.effectiveQuantity || 0) * matchedChildUnitCost).toFixed(2));
                    line.DisplayUnitCost = Number(line.DisplayQuantity || 0) > 0
                        ? parseFloat(Number(line.EffectiveLineCost / line.DisplayQuantity).toFixed(2))
                        : parseFloat(Number(matchedChildUnitCost || 0).toFixed(2));
                }
                continue;
            }

            if (key !== '4') {
                continue;
            }

            const childSummary = await getProductionSummary(childOrdNo, nextVisited, options);
            if (childSummary) {
                const childTotal = parseFloat(Number(childSummary.totalCost || 0).toFixed(2));
                line.ChildProductionTotalCost = childTotal;
                line.EffectiveLineCost = childTotal;
                line.DisplayUnitCost = Number(line.DisplayQuantity || 0) > 0
                    ? parseFloat(Number(childTotal / line.DisplayQuantity).toFixed(2))
                    : parseFloat(Number(line.CCstPr || 0).toFixed(2));
            }
            if (childSummary && childSummary.hasWarnings) {
                line.HasWarning = true;
                line.ChildHasWarning = true;
                line.WarningText = joinWarningMessages([
                    line.WarningText,
                    childSummary.warningText || 'Underliggende produktionsordre har en advarsel.'
                ]);
            }
        }

        const totalCost = lines
            .filter(line => Number(line.LnNo || 0) !== 1)
            .reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);
        const roundedTotalCost = parseFloat(Number(totalCost).toFixed(2));

        const mainLine = lines.find(line => Number(line.LnNo || 0) === 1);
        if (mainLine && Number(mainLine.DisplayQuantity || 0) > 0 && Number(roundedTotalCost || 0) > 0) {
            mainLine.DisplayUnitCost = parseFloat(Number(roundedTotalCost / mainLine.DisplayQuantity).toFixed(2));
        }

        const result = {
            ordNo: numericOrdNo,
            cacheSchemaVersion: PRODUCTION_SUMMARY_CACHE_SCHEMA_VERSION,
            lines,
            hasWarnings: lines.some(line => line.HasWarning),
            warningText: joinWarningMessages(lines.map(line => line.WarningText)),
            hasEstimatedOperationTime: lines.some(line => line.UsesEstimatedOperationTime),
            totalCost: roundedTotalCost
        };
        diskCache.set(summaryCacheKey, result, cacheTtlProductionSummaryMs);
        return result;
    }

    return {
        getAfterCalc,
        fetchOrderListBase,
        getProductionSummary
    };
}

module.exports = {
    createAftercalcService
};
