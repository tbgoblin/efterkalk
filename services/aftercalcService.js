function createAftercalcService({
    getConnection,
    sql,
    diskCache,
    logEvent,
    getLatestDrawingByProdNo,
    isGloballyExcludedProdNo,
    adjustOperationLinePricing,
    isLaserLProduct,
    orderListMaxRows,
    orderListDaysBack,
    cacheTtlProductionSummaryMs
}) {
    async function getAfterCalc(ordNo) {
        const pool = await getConnection();

        try {
            const [orderResult, salesOrderLinesResult, salesLinesResult, productionLinesResult] = await Promise.all([
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                        SELECT O.OrdNo, O.TrTp, O.InvoAm, A.Nm as CustomerName
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
                            NoFin,
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
                            NoFin,
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

            const prodNosForDrawings = [
                ...salesOrderLinesResult.recordset.map(r => r.ProdNo),
                ...salesLinesResult.recordset.map(r => r.ProdNo)
            ];
            const drawingByProdNo = await getLatestDrawingByProdNo(pool, prodNosForDrawings, logEvent);

            const salesLines = salesLinesResult.recordset
                .filter(line => !isGloballyExcludedProdNo(line.ProdNo))
                .map(line => {
                    const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
                    const isDiscountLine = lineSalesPrice === 0;
                    const effectiveLineCost = isDiscountLine ? 0 : (line.LineCost || 0);
                    const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                    return {
                        ...line,
                        IsDiscountLine: isDiscountLine,
                        EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2)),
                        DrawingWebPg: drawingByProdNo.get(prodNoKey) || null
                    };
                });
            const salesLinesTotalCost = salesLines.reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);

            const salesOrderLines = salesOrderLinesResult.recordset
                .filter(line => !isGloballyExcludedProdNo(line.ProdNo))
                .map(line => {
                    const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
                    const isDiscountLine = lineSalesPrice === 0;
                    const effectiveLineCost = isDiscountLine ? 0 : (line.LineCost || 0);
                    const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                    return {
                        ...line,
                        IsDiscountLine: isDiscountLine,
                        EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2)),
                        DrawingWebPg: drawingByProdNo.get(prodNoKey) || null
                    };
                });

            const productionOrderDetailsCache = new Map();

            async function loadProductionOrderDetails(prodOrdNo, visited = new Set()) {
                const numericProdOrdNo = Number(prodOrdNo);
                if (!Number.isFinite(numericProdOrdNo)) {
                    return { lines: [], totalCost: 0 };
                }

                if (productionOrderDetailsCache.has(numericProdOrdNo)) {
                    return productionOrderDetailsCache.get(numericProdOrdNo);
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
                                OrdNo, 
                                LnNo, 
                                ProdNo, 
                                Descr, 
                                DPrice,
                                NoOrg,
                                NoFin,
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

                    for (const rawLine of prodLinesResult.recordset) {
                        if (isGloballyExcludedProdNo(rawLine.ProdNo)) continue;

                        const line = adjustOperationLinePricing({ ...rawLine });
                        const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                        let effectiveLineCost = Number(line.LineCost || 0);
                        let childProductionTotalCost = null;

                        if (key === '1') {
                            const pnUp = String(line.ProdNo || '').toUpperCase();
                            if (pnUp === 'R6200') {
                                effectiveLineCost = Number((line.NoOrg || 0) * (line.CCstPr || 0));
                            }
                        } else if (key === '2' && isLaserLProduct(line.ProdNo)) {
                            const hasNestingCost = Number(line.NestingCost || 0) > 0;
                            effectiveLineCost = hasNestingCost
                                ? Number((line.NestingCost || 0) * (line.NoFin || 0))
                                : Number(line.LineCost || 0);
                        } else if (key === '4' && line.PurcNo && Number(line.PurcNo) !== 0) {
                            const childDetails = await loadProductionOrderDetails(Number(line.PurcNo), nextVisited);
                            childProductionTotalCost = Number(childDetails.totalCost || 0);
                            effectiveLineCost = childProductionTotalCost;
                        }

                        line.ChildProductionTotalCost = childProductionTotalCost === null
                            ? null
                            : parseFloat(Number(childProductionTotalCost).toFixed(2));
                        line.EffectiveLineCost = parseFloat(Number(effectiveLineCost || 0).toFixed(2));
                        lines.push(line);

                        if (line.LnNo === 1 || key === '0' || key === '3' || key === '5') continue;
                        total += (line.EffectiveLineCost || 0);
                    }

                    return {
                        lines,
                        totalCost: parseFloat(Number(total).toFixed(2))
                    };
                })();

                productionOrderDetailsCache.set(numericProdOrdNo, detailsPromise);
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
                        loadProductionOrderDetails(purcNo)
                    ]);

                    if (prodOrderResult.recordset.length === 0) return null;
                    const prodOrder = prodOrderResult.recordset[0];

                    return {
                        ordNo: purcNo,
                        trTp: prodOrder.TrTp,
                        revenue: prodOrder.InvoAm || 0,
                        lines: prodDetails.lines,
                        totalCost: prodDetails.totalCost
                    };
                })
            );
            const productionOrders = productionOrderResults.filter(Boolean);

            const productionTotalByOrdNo = new Map(
                productionOrders.map(po => [Number(po.ordNo), Number(po.totalCost || 0)])
            );

            const salesOrderLinesWithProductionTotal = salesOrderLines.map(line => {
                const purcNo = line.PurcNo ? Number(line.PurcNo) : 0;
                const productionTotal = purcNo ? productionTotalByOrdNo.get(purcNo) : undefined;

                if (productionTotal !== undefined && !line.IsDiscountLine) {
                    const roundedTotal = parseFloat(Number(productionTotal).toFixed(2));
                    return {
                        ...line,
                        ProductionOrderTotalCost: roundedTotal,
                        EffectiveLineCost: roundedTotal
                    };
                }

                return {
                    ...line,
                    ProductionOrderTotalCost: productionTotal !== undefined ? parseFloat(Number(productionTotal).toFixed(2)) : null
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

            return {
                orderHeader,
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
                B.CustomerName,
                B.SellerUsr,
                B.LstInvDt,
                B.InvoAm
            FROM BaseOrders B
            ORDER BY B.LstInvDt DESC, B.OrdNo DESC
        `);

        return result.recordset;
    }

    async function getProductionSummary(ordNo) {
        const numericOrdNo = Number(ordNo);
        if (!Number.isFinite(numericOrdNo)) {
            throw new Error('Ordrenummer ugyldigt');
        }

        const cachedSummary = diskCache.get('prod_summary_' + numericOrdNo);
        if (cachedSummary) return cachedSummary;

        const pool = await getConnection();
        const linesResult = await pool.request()
            .input('ordNo', sql.Numeric, numericOrdNo)
            .query(`
                SELECT
                    LnNo,
                    ProdNo,
                    Descr,
                    NoOrg,
                    NoFin,
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
                WHERE OrdNo = @ordNo AND NoFin > 0
                ORDER BY LnNo
            `);

        const lines = linesResult.recordset
            .filter(rawLine => !isGloballyExcludedProdNo(rawLine.ProdNo))
            .map(rawLine => {
                const line = adjustOperationLinePricing({ ...rawLine });
                const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                const hasNestingCost = Number(line.NestingCost || 0) > 0;
                const effectiveLineCost = key === '2' && isLaserLProduct(line.ProdNo)
                    ? (hasNestingCost ? ((line.NestingCost || 0) * (line.NoFin || 0)) : (line.LineCost || 0))
                    : Number(line.LineCost || 0);

                return {
                    ...line,
                    EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2))
                };
            });

        const totalCost = lines
            .filter(line => Number(line.LnNo || 0) !== 1)
            .reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);

        const result = {
            ordNo: numericOrdNo,
            lines,
            totalCost: parseFloat(Number(totalCost).toFixed(2))
        };
        diskCache.set('prod_summary_' + numericOrdNo, result, cacheTtlProductionSummaryMs);
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
