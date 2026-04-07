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
    const PRODUCTION_SUMMARY_CACHE_SCHEMA_VERSION = 11;
    const laserRoutePricingCache = new Map();

    function buildLineWarnings(line, extraWarnings = []) {
        const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
        const prodNoKey = String((line && line.ProdNo) || '').trim().toUpperCase();
        const noFinValue = Number((line && line.NoFin) || 0);
        const noOrgValue = Number((line && line.NoOrg) || 0);
        const purcNoValue = Number((line && line.PurcNo) || 0);
        const noInvoValue = Number((line && line.NoInvo) || 0);
        const noInvoAbValue = Number((line && line.NoInvoAb) || 0);
        const warnings = [];

        if (prodNoKey.startsWith('3') && noFinValue === 0 && noOrgValue > 0) {
            warnings.push(key === '2'
                ? 'Inkonsekvens: materiale/tubo med NoFin=0 men NoOrg>0.'
                : 'Inkonsekvens på salgsordre: produkt/tubo med NoFin=0 men NoOrg>0.');
        }

        if (key === '6' && purcNoValue > 0 && noInvoAbValue > noInvoValue) {
            warnings.push('manglede indkøbsfaktura');
        }

        for (const warning of extraWarnings || []) {
            const text = String(warning || '').trim();
            if (text && !warnings.includes(text)) warnings.push(text);
        }

        return warnings;
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
                totalCostByProd: new Map()
            };
        }

        if (laserRoutePricingCache.has(numericOrdNo)) {
            return laserRoutePricingCache.get(numericOrdNo);
        }

        const pricingPromise = (async () => {
            const linkedFinishedRowsResult = await pool.request()
                .input('ordNo', sql.Numeric, numericOrdNo)
                .query(`
                    SELECT OrdNo, TrInf4, ProdNo, TrTp, NoFin, Free3, CstPr
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
                    totalCostByProd: new Map()
                };
            }

            const ordPlaceholders = nestingOrdNos.map((_, i) => `@nest${i}`).join(', ');
            const routeRowsRequest = pool.request();
            nestingOrdNos.forEach((ordNoValue, i) => {
                routeRowsRequest.input(`nest${i}`, sql.Numeric, ordNoValue);
            });
            const routeRowsResult = await routeRowsRequest.query(`
                SELECT OrdNo, TrInf4, ProdNo, TrTp, NoFin, Free3, CstPr
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

            const routeSheetStats = new Map();
            const prodRouteKgStats = new Map();

            for (const row of rows) {
                const routeKey = String(row.TrInf4 || '').trim().toUpperCase();
                const prodKey = String(row.ProdNo || '').trim().toUpperCase();
                if (!routeKey) continue;

                if (!routeSheetStats.has(routeKey)) {
                    routeSheetStats.set(routeKey, { sum: 0, count: 0, kgConsumati: 0, kgFiniti: 0 });
                }

                const routeStats = routeSheetStats.get(routeKey);
                if (Number(row.TrTp) === 5) {
                    const cstPr = Number(row.CstPr || 0);
                    if (cstPr > 0) {
                        routeStats.sum += cstPr;
                        routeStats.count += 1;
                    }
                    routeStats.kgConsumati += Number(row.NoFin || 0);
                    continue;
                }

                if (Number(row.TrTp) === 7 && prodKey) {
                    const noFin = Number(row.NoFin || 0);
                    const oldExpectedUnitWeight = normalizeExpectedWeight(row.Free3);
                    const oldKgProdotto = oldExpectedUnitWeight > 0 && noFin > 0
                        ? (oldExpectedUnitWeight * noFin)
                        : 0;
                    routeStats.kgFiniti += oldKgProdotto;

                    if (oldKgProdotto > 0 && noFin > 0) {
                        const key = routeKey + '|' + prodKey;
                        const stats = prodRouteKgStats.get(key) || { oldKgProdottoSum: 0, qtySum: 0 };
                        stats.oldKgProdottoSum += oldKgProdotto;
                        stats.qtySum += noFin;
                        prodRouteKgStats.set(key, stats);
                    }
                }
            }

            const avgSheetCstPrByRoute = new Map();
            for (const [routeKey, stats] of routeSheetStats.entries()) {
                if (stats.count > 0) {
                    avgSheetCstPrByRoute.set(routeKey, Number(stats.sum / stats.count));
                }
            }

            const kgPerPieceByProdRoute = new Map();
            const unitCostByProd = new Map();
            const totalCostByProd = new Map();
            const unitCostStatsByProd = new Map();
            for (const [key, stats] of prodRouteKgStats.entries()) {
                if (stats.qtySum <= 0) continue;
                const [routeKey, prodKey] = key.split('|');
                const routeStats = routeSheetStats.get(routeKey) || { kgConsumati: 0, kgFiniti: 0 };
                const kgUtilizzatiEffettivi = routeStats.kgFiniti > 0
                    ? ((stats.oldKgProdottoSum / routeStats.kgFiniti) * routeStats.kgConsumati)
                    : 0;
                const kgPerPiece = stats.qtySum > 0 ? Number(kgUtilizzatiEffettivi / stats.qtySum) : 0;
                kgPerPieceByProdRoute.set(key, kgPerPiece);

                const avgSheetCstPr = Number(avgSheetCstPrByRoute.get(routeKey) || 0);
                if (avgSheetCstPr > 0 && kgPerPiece > 0 && prodKey) {
                    const prodStats = unitCostStatsByProd.get(prodKey) || { costSum: 0, qtySum: 0 };
                    prodStats.costSum += kgPerPiece * avgSheetCstPr * stats.qtySum;
                    prodStats.qtySum += stats.qtySum;
                    unitCostStatsByProd.set(prodKey, prodStats);
                }
            }

            for (const [prodKey, stats] of unitCostStatsByProd.entries()) {
                if (stats.qtySum > 0) {
                    totalCostByProd.set(prodKey, Number(stats.costSum));
                    unitCostByProd.set(prodKey, Number(stats.costSum / stats.qtySum));
                }
            }

            return {
                avgSheetCstPrByRoute,
                kgPerPieceByProdRoute,
                unitCostByProd,
                totalCostByProd
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

        if (routeKey) {
            const avgSheetCstPr = Number(pricingData.avgSheetCstPrByRoute.get(routeKey) || 0);
            const kgPerPiece = Number(pricingData.kgPerPieceByProdRoute.get(routeKey + '|' + prodKey) || 0);
            if (avgSheetCstPr > 0 && kgPerPiece > 0) {
                const unitCost = parseFloat(Number(avgSheetCstPr * kgPerPiece).toFixed(6));
                return {
                    unitCost,
                    totalCost: qty > 0 ? parseFloat(Number(unitCost * qty).toFixed(2)) : null
                };
            }
        }

        const aggregatedTotalCost = Number(pricingData.totalCostByProd.get(prodKey) || 0);
        if (aggregatedTotalCost > 0) {
            const fallbackUnitCost = qty > 0
                ? (aggregatedTotalCost / qty)
                : Number(pricingData.unitCostByProd.get(prodKey) || 0);
            return {
                unitCost: parseFloat(Number(fallbackUnitCost).toFixed(6)),
                totalCost: parseFloat(Number(aggregatedTotalCost).toFixed(2))
            };
        }

        const aggregatedUnitCost = Number(pricingData.unitCostByProd.get(prodKey) || 0);
        if (aggregatedUnitCost > 0) {
            return {
                unitCost: parseFloat(Number(aggregatedUnitCost).toFixed(6)),
                totalCost: qty > 0 ? parseFloat(Number(aggregatedUnitCost * qty).toFixed(2)) : null
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
                        SELECT O.OrdNo, O.TrTp, O.InvoAm, O.Gr4, A.Nm as CustomerName
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
                    const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
                    const tubeFallbackCost = getInconsistentTubeFallbackCost(line);
                    const isDiscountLine = lineSalesPrice === 0 && tubeFallbackCost === null;
                    const effectiveLineCost = isDiscountLine ? 0 : (tubeFallbackCost !== null ? tubeFallbackCost : (line.LineCost || 0));
                    const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                    const warningMessages = buildLineWarnings(line);
                    const displayQuantity = Number(line.NoFin || 0) === 0 && Number(line.NoOrg || 0) > 0
                        ? Number(line.NoOrg || 0)
                        : Number(line.NoFin || 0);
                    return {
                        ...line,
                        IsDiscountLine: isDiscountLine,
                        DisplayQuantity: parseFloat(Number(displayQuantity).toFixed(2)),
                        HasWarning: warningMessages.length > 0,
                        WarningText: warningMessages.join(' '),
                        EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2)),
                        DrawingWebPg: drawingByProdNo.get(prodNoKey) || null
                    };
                });
            const salesLinesTotalCost = salesLines.reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);

            const salesOrderLines = salesOrderLinesResult.recordset
                .filter(line => !isGloballyExcludedProdNo(line.ProdNo))
                .map(line => {
                    const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
                    const tubeFallbackCost = getInconsistentTubeFallbackCost(line);
                    const isDiscountLine = lineSalesPrice === 0 && tubeFallbackCost === null;
                    const effectiveLineCost = isDiscountLine ? 0 : (tubeFallbackCost !== null ? tubeFallbackCost : (line.LineCost || 0));
                    const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                    const warningMessages = buildLineWarnings(line);
                    const displayQuantity = Number(line.NoFin || 0) === 0 && Number(line.NoOrg || 0) > 0
                        ? Number(line.NoOrg || 0)
                        : Number(line.NoFin || 0);
                    return {
                        ...line,
                        IsDiscountLine: isDiscountLine,
                        DisplayQuantity: parseFloat(Number(displayQuantity).toFixed(2)),
                        HasWarning: warningMessages.length > 0,
                        WarningText: warningMessages.join(' '),
                        EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2)),
                        DrawingWebPg: drawingByProdNo.get(prodNoKey) || null
                    };
                });

            const productionOrderDetailsCache = new Map();

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
                    const specialLaserPricingData = useSpecialLaserCost
                        ? await loadLaserRoutePricingData(pool, numericProdOrdNo).catch(() => null)
                        : null;

                    const prodLinesResult = await pool.request()
                        .input('purcNo', sql.Numeric, numericProdOrdNo)
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
                                CCstPr,
                                PurcNo,
                                TrInf2,
                                TrInf4,
                                ProdTp4,
                                (
                                    SELECT TOP 1 A.Nm
                                    FROM ProdTr P
                                    LEFT JOIN Actor A ON A.EmpNo = P.EmpNo
                                    WHERE P.OrdNo = @purcNo
                                      AND P.OrdLnNo = OrdLn.LnNo
                                    ORDER BY P.FinDt DESC, P.FinTm DESC
                                ) AS HvemNm,
                                CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost,
                                (
                                    SELECT SUM(CAST(n.CstPr AS DECIMAL(18,6)) * CAST(n.NoFin AS DECIMAL(18,6)))
                                         / NULLIF(SUM(CAST(n.NoFin AS DECIMAL(18,6))), 0)
                                    FROM OrdLn n
                                    WHERE n.TrInf2 = CAST(@purcNo AS VARCHAR(20))
                                      AND n.ProdNo = OrdLn.ProdNo
                                ) AS NestingCost
                            FROM OrdLn
                            WHERE OrdNo = @purcNo
                            ORDER BY LnNo
                        `);

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
                        const operationTimeInfo = key === '1'
                            ? getOperationTimeInfo(line)
                            : { effectiveMinutes: noFinValue, usesEstimatedMinutes: false, infoText: '' };
                        const displayQuantity = key === '1'
                            ? operationTimeInfo.effectiveMinutes
                            : ((isTubeMaterialLine && noFinValue === 0 && noOrgValue > 0)
                                ? noOrgValue
                                : noFinValue);
                        let effectiveLineCost = Number(line.LineCost || 0);
                        let childProductionTotalCost = null;
                        const warningMessages = buildLineWarnings(line);

                        if (key === '1') {
                            effectiveLineCost = Number(operationTimeInfo.effectiveMinutes * (line.CCstPr || 0));
                        } else if (key === '2' && isLaserLProduct(line.ProdNo)) {
                            const specialLaserCostInfo = useSpecialLaserCost
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
                                warningMessages.push('Underliggende produktionsordre har en advarsel.');
                            }
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
                        line.HasWarning = warningMessages.length > 0;
                        line.ChildHasWarning = key === '4' && warningMessages.some(msg => msg.includes('Underliggende produktionsordre'));
                        line.WarningText = warningMessages.join(' ');
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
                        hasEstimatedOperationTime
                    };
                })();

                productionOrderDetailsCache.set(detailsCacheKey, detailsPromise);
                return detailsPromise;
            }

            const productionOrderResults = await Promise.all(
                productionLinesResult.recordset.map(async (prodLine) => {
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

                    if (prodOrderResult.recordset.length === 0) return null;
                    const prodOrder = prodOrderResult.recordset[0];

                    if (useRouteSpecificLaserCost && Array.isArray(prodDetails.lines) && prodDetails.lines.length > 0) {
                        const specialLaserPricingData = await loadLaserRoutePricingData(pool, Number(purcNo)).catch(() => null);
                        if (specialLaserPricingData) {
                            let adjustedTotalCost = 0;
                            prodDetails.lines = prodDetails.lines.map(rawLine => {
                                const line = { ...rawLine };
                                const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                                if (key === '2' && isLaserLProduct(line.ProdNo)) {
                                    const specialLaserCostInfo = getSpecialGr4LaserCostInfo(specialLaserPricingData, line);
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
                        hasEstimatedOperationTime: Boolean(prodDetails.hasEstimatedOperationTime)
                    };
                })
            );
            const productionOrders = productionOrderResults.filter(Boolean);

            const productionTotalByOrdNo = new Map(
                productionOrders.map(po => [Number(po.ordNo), Number(po.totalCost || 0)])
            );
            const productionWarningByOrdNo = new Map(
                productionOrders.map(po => [Number(po.ordNo), Boolean(po.hasWarnings)])
            );

            const salesOrderLinesWithProductionTotal = salesOrderLines.map(line => {
                const purcNo = line.PurcNo ? Number(line.PurcNo) : 0;
                const productionTotal = purcNo ? productionTotalByOrdNo.get(purcNo) : undefined;
                const productionHasWarning = purcNo ? Boolean(productionWarningByOrdNo.get(purcNo)) : false;
                const combinedWarnings = [];
                if (line.HasWarning && line.WarningText) combinedWarnings.push(line.WarningText);
                if (productionHasWarning) combinedWarnings.push('Tilknyttet produktionsordre har en advarsel.');
                const warningText = Array.from(new Set(combinedWarnings)).join(' ');
                const hasWarning = warningText.length > 0;

                if (productionTotal !== undefined && !line.IsDiscountLine) {
                    const roundedTotal = parseFloat(Number(productionTotal).toFixed(2));
                    return {
                        ...line,
                        ProductionOrderTotalCost: roundedTotal,
                        EffectiveLineCost: roundedTotal,
                        HasWarning: hasWarning,
                        WarningText: warningText
                    };
                }

                return {
                    ...line,
                    ProductionOrderTotalCost: productionTotal !== undefined ? parseFloat(Number(productionTotal).toFixed(2)) : null,
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
            const totalRevenue = orderHeader.InvoAm || 0;
            const margin = totalRevenue - totalCost;
            const marginPercentage = totalCost > 0 ? ((totalRevenue / totalCost) * 100).toFixed(2) : 0;
            const hasWarnings = salesOrderLinesWithProductionTotal.some(line => line.HasWarning)
                || salesLines.some(line => line.HasWarning)
                || productionOrders.some(order => order.hasWarnings);

            return {
                orderHeader,
                hasWarnings,
                warningText: hasWarnings ? 'Ordren indeholder mindst én advarsel.' : '',
                salesOrderLines: salesOrderLinesWithProductionTotal,
                salesLines,
                salesLinesTotalCost: parseFloat(salesLinesTotalCost.toFixed(2)),
                productionOrders,
                summary: {
                    totalRevenue: parseFloat(totalRevenue.toFixed(2)),
                    totalCost: parseFloat(totalCost.toFixed(2)),
                    margin: parseFloat(margin.toFixed(2)),
                    marginPercentage
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

        const useSpecialLaserCost = true;
        const summaryCacheKey = 'prod_summary_' + numericOrdNo + (useSpecialLaserCost ? '_gr4_3' : '');
        const cachedSummary = diskCache.get(summaryCacheKey);
        if (cachedSummary && cachedSummary.cacheSchemaVersion === PRODUCTION_SUMMARY_CACHE_SCHEMA_VERSION) {
            return cachedSummary;
        }

        const pool = await getConnection();
        const specialLaserPricingData = useSpecialLaserCost
            ? await loadLaserRoutePricingData(pool, numericOrdNo).catch(() => null)
            : null;
        const linesResult = await pool.request()
            .input('ordNo', sql.Numeric, numericOrdNo)
            .query(`
                SELECT
                    LnNo,
                    ProdNo,
                    Descr,
                    NoOrg,
                    NoFin,
                    NoInvo,
                    NoInvoAb,
                    DPrice,
                    CCstPr,
                    TrInf2,
                    TrInf4,
                    ProdTp4,
                    PurcNo,
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
                WHERE OrdNo = @ordNo
                  AND (LnNo = 1 OR NoFin > 0 OR NoOrg > 0)
                ORDER BY LnNo
            `);

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
                const operationTimeInfo = key === '1'
                    ? getOperationTimeInfo(line)
                    : { effectiveMinutes: noFinValue, usesEstimatedMinutes: false, infoText: '' };
                const displayQuantity = key === '1'
                    ? operationTimeInfo.effectiveMinutes
                    : ((isTubeMaterialLine && noFinValue === 0 && noOrgValue > 0)
                        ? noOrgValue
                        : noFinValue);
                const warningMessages = buildLineWarnings(line);
                const tubeFallbackCost = getInconsistentTubeFallbackCost(line);
                const specialLaserCostInfo = (key === '2' && isLaserLProduct(line.ProdNo) && useSpecialLaserCost)
                    ? getSpecialGr4LaserCostInfo(specialLaserPricingData, line)
                    : null;
                if (specialLaserCostInfo && specialLaserCostInfo.unitCost !== null) {
                    line.NestingCost = specialLaserCostInfo.unitCost;
                }
                const recalculatedHasNestingCost = Number(line.NestingCost || 0) > 0;
                const effectiveLineCost = key === '1'
                    ? Number(operationTimeInfo.effectiveMinutes * (line.CCstPr || 0))
                    : (key === '2' && isLaserLProduct(line.ProdNo)
                        ? ((specialLaserCostInfo && specialLaserCostInfo.totalCost !== null)
                            ? Number(specialLaserCostInfo.totalCost)
                            : (recalculatedHasNestingCost ? ((line.NestingCost || 0) * (line.NoFin || 0)) : (line.LineCost || 0)))
                        : (tubeFallbackCost !== null ? tubeFallbackCost : Number(line.LineCost || 0)));

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
                    HasWarning: warningMessages.length > 0,
                    WarningText: warningMessages.join(' '),
                    EffectiveLineCost: roundedEffectiveLineCost
                };
            })
            .filter(Boolean);

        for (const line of lines) {
            const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
            const childOrdNo = Number(line.PurcNo || 0);
            if (key !== '4' || !Number.isFinite(childOrdNo) || childOrdNo <= 0 || nextVisited.has(childOrdNo)) {
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
                line.WarningText = [line.WarningText, 'Underliggende produktionsordre har en advarsel.']
                    .filter(Boolean)
                    .join(' ');
            }
        }

        const totalCost = lines
            .filter(line => Number(line.LnNo || 0) !== 1)
            .reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);
        const roundedTotalCost = parseFloat(Number(totalCost).toFixed(2));

        const mainLine = lines.find(line => Number(line.LnNo || 0) === 1);
        if (mainLine && Number(mainLine.DisplayQuantity || 0) > 0) {
            mainLine.DisplayUnitCost = parseFloat(Number(roundedTotalCost / mainLine.DisplayQuantity).toFixed(2));
        }

        const result = {
            ordNo: numericOrdNo,
            cacheSchemaVersion: PRODUCTION_SUMMARY_CACHE_SCHEMA_VERSION,
            lines,
            hasWarnings: lines.some(line => line.HasWarning),
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
