function toNumber(value) {
    if (typeof value === 'string') {
        const parsed = Number(value.replace(/\s+/g, '').replace(',', '.'));
        return Number.isFinite(parsed) ? parsed : 0;
    }
    const parsed = Number(value || 0);
    return Number.isFinite(parsed) ? parsed : 0;
}

function round(value) {
    return Math.round(toNumber(value) * 100) / 100;
}

function isPlateLine(row) {
    const prodNo = String(row && row.ProdNo || '').trim();
    const gr6 = toNumber(row && row.Gr6);
    const prodGr = toNumber(row && row.ProdGr);
    return Number(row && row.TrTp) === 5
        && prodNo.startsWith('3')
        && prodNo.charAt(2) === '1'
        && gr6 === 1
        && prodGr !== 99999
        && prodNo !== '301001';
}

function isUnregisteredSearchPlate(row) {
    if (Number(row && row.TrTp) !== 5) return false;
    const sourceInfo = String(row && row.TrInf1 || '').trim().toLocaleLowerCase('da-DK');
    return /^søg(?:\s|$|-)/u.test(sourceInfo);
}

function lineValue(row) {
    const finished = toNumber(row && row.NoFin);
    const planned = toNumber(row && row.NoOrg);
    const unitCost = toNumber(row && row.CstPr);
    if (unitCost !== 0) return round(finished * unitCost);
    return planned !== 0 ? round(toNumber(row && row.IncCst) * (finished / planned)) : round(row && row.IncCst);
}

function productProgress(row) {
    const planned = Math.abs(toNumber(row && row.NoOrg));
    const finished = Math.abs(toNumber(row && row.NoFin));
    if (planned <= 0.005) return finished > 0.005 ? 1 : 0;
    return Math.max(0, Math.min(1, finished / planned));
}

function uniqueValues(values, limit, normalize) {
    return Array.from(new Set((Array.isArray(values) ? values : []).map(normalize).filter(Boolean))).slice(0, limit);
}

function toVismaMoment(value, label) {
    const date = new Date(value);
    if (!Number.isFinite(date.getTime())) {
        const error = new Error(label + ' er ugyldig');
        error.statusCode = 400;
        throw error;
    }
    return {
        timestamp: date.getTime(),
        date: (date.getFullYear() * 10000) + ((date.getMonth() + 1) * 100) + date.getDate(),
        time: (date.getHours() * 100) + date.getMinutes()
    };
}

function buildRouteLineage(lineRows, restRows, restPrices = {}) {
    const groups = new Map();
    for (const row of Array.isArray(lineRows) ? lineRows : []) {
        const nestingOrdNo = toNumber(row.NestingOrdNo || row.OrdNo);
        const route = String(row.Route || row.TrInf4 || '').trim();
        if (!nestingOrdNo || !route) continue;
        const key = String(nestingOrdNo) + '|' + route;
        const group = groups.get(key) || {
            key,
            nestingOrdNo,
            route,
            ordDate: row.OrdDt || null,
            plates: [],
            products: [],
            estimatedRestLines: [],
            productionOrderNos: new Set(),
            salesOrderNos: new Set(),
            salesOrderReferences: new Map(),
            restPlates: []
        };
        const productionOrdNo = toNumber(row.ProductionOrdNo || row.TrInf2);
        const salesOrdNo = toNumber(row.SalesOrderNo || row.DirectSalesOrdNo);
        const salesOrderSource = String(row.SalesOrderSource || '').trim();
        if (productionOrdNo > 0) group.productionOrderNos.add(productionOrdNo);
        if (salesOrdNo > 0) {
            group.salesOrderNos.add(salesOrdNo);
            group.salesOrderReferences.set(String(salesOrdNo) + '|' + salesOrderSource, {
                orderNo: salesOrdNo,
                source: salesOrderSource || 'ukendt'
            });
        }
        if (isPlateLine(row)) {
            const normalized = {
                prodNo: String(row.ProdNo || '').trim(),
                descr: String(row.Descr || '').trim(),
                sourceInfo: String(row.TrInf1 || '').trim(),
                unregisteredRestSource: isUnregisteredSearchPlate(row),
                plannedQty: round(row.NoOrg),
                finishedQty: round(row.NoFin),
                fifoUnitCost: round(row.CstPr),
                value: lineValue(row)
            };
            if (toNumber(row.NoOrg) < 0) group.estimatedRestLines.push(normalized);
            else group.plates.push(normalized);
        } else if (Number(row.TrTp) === 7) {
            group.products.push({
                prodNo: String(row.ProdNo || '').trim(),
                descr: String(row.Descr || '').trim(),
                plannedQty: round(row.NoOrg),
                finishedQty: round(row.NoFin),
                unitCost: round(row.CstPr),
                value: lineValue(row),
                progress: productProgress(row),
                productionOrdNo,
                salesOrdNo,
                salesOrderSource,
                hasSalesReference: salesOrdNo > 0
            });
        }
        groups.set(key, group);
    }

    const groupsByOrderAndPlate = new Map();
    for (const group of groups.values()) {
        for (const plate of group.plates.concat(group.estimatedRestLines)) {
            const key = String(group.nestingOrdNo) + '|' + plate.prodNo;
            const list = groupsByOrderAndPlate.get(key) || [];
            if (!list.includes(group)) list.push(group);
            groupsByOrderAndPlate.set(key, list);
        }
    }

    const unassignedRest = [];
    for (const row of Array.isArray(restRows) ? restRows : []) {
        const prodNo = String(row.ProdNo || '').trim();
        const nestingOrdNo = toNumber(row.NestingOrdNo || row.OrdNo);
        const priceType = String(row.PriceType || prodNo.slice(0, 3)).trim();
        const weight = round(toNumber(row.Val5) * toNumber(row.Val8));
        const pricePerKg = round(restPrices[priceType]);
        const normalized = {
            prodNo,
            nestingOrdNo,
            label: String(row.Txt2 || row.Txt1 || row.Descr || '').trim(),
            weight,
            pricePerKg,
            value: round(weight * pricePerKg)
        };
        const candidates = groupsByOrderAndPlate.get(String(nestingOrdNo) + '|' + prodNo) || [];
        if (candidates.length === 1) candidates[0].restPlates.push(normalized);
        else unassignedRest.push({ ...normalized, reason: candidates.length ? 'flere mulige ruter' : 'ingen matchende rute' });
    }

    const routes = Array.from(groups.values()).map(group => {
        const progresses = group.products.map(product => product.progress);
        const progress = progresses.length ? progresses.reduce((sum, value) => sum + value, 0) / progresses.length : 0;
        const status = !progresses.length
            ? 'unknown'
            : (progresses.every(value => value >= 0.999) ? 'completed' : (progresses.every(value => value <= 0.001) ? 'not_started' : 'partial'));
        const plateValue = round(group.plates.reduce((sum, row) => sum + row.value, 0));
        const completedProductValue = round(group.products.reduce((sum, row) => sum + row.value, 0));
        const restValue = round(group.restPlates.reduce((sum, row) => sum + row.value, 0));
        return {
            key: group.key,
            nestingOrdNo: group.nestingOrdNo,
            route: group.route,
            ordDate: group.ordDate,
            status,
            progress: round(progress * 100),
            plates: group.plates,
            products: group.products,
            estimatedRestLines: group.estimatedRestLines,
            restPlates: group.restPlates,
            productionOrderNos: Array.from(group.productionOrderNos).sort((a, b) => a - b),
            salesOrderNos: Array.from(group.salesOrderNos).sort((a, b) => a - b),
            salesOrderReferences: Array.from(group.salesOrderReferences.values()).sort((a, b) => a.orderNo - b.orderNo),
            unlinkedProductCount: group.products.filter(product => !product.hasSalesReference).length,
            plateValue,
            completedProductValue,
            restValue,
            residualValue: round(plateValue - completedProductValue - restValue)
        };
    }).sort((left, right) => {
        const statusRank = { partial: 0, not_started: 1, unknown: 2, completed: 3 };
        return (statusRank[left.status] - statusRank[right.status])
            || right.nestingOrdNo - left.nestingOrdNo
            || left.route.localeCompare(right.route);
    });
    return { routes, unassignedRest };
}

function createLagerliste2Service({ getConnection, sql, getRestPrices }) {
    let cache = null;
    const cacheTtlMs = 2 * 60 * 1000;

    async function getCurrentRoutes({ forceRefresh = false } = {}) {
        if (!forceRefresh && cache && cache.expiresAt > Date.now()) return cache.payload;
        const pool = await getConnection();
        const now = new Date();
        const start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 2, 1));
        const end = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1));
        const startDate = Number(start.toISOString().slice(0, 7).replace('-', '') + '01');
        const endDate = Number(end.toISOString().slice(0, 10).replace(/-/g, ''));

        const lineResult = await pool.request()
            .input('startDate', sql.Int, startDate)
            .input('endDate', sql.Int, endDate)
            .query(`
                WITH RelevantNestingLines AS (
                    SELECT
                        N.OrdNo AS NestingOrdNo,
                        N.OrdDt,
                        N.R4 AS OrderR4,
                        L.LnNo,
                        L.R4 AS LineR4,
                        L.TrInf1,
                        L.TrInf2,
                        L.TrInf4 AS Route,
                        L.ProdNo,
                        L.Descr,
                        L.TrTp,
                        L.NoOrg,
                        L.NoFin,
                        L.CstPr,
                        L.IncCst
                    FROM Ord N WITH(NOLOCK)
                    INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = N.OrdNo
                    WHERE TRY_CONVERT(int, N.OrdDt) >= @startDate
                      AND TRY_CONVERT(int, N.OrdDt) < @endDate
                      AND TRY_CONVERT(decimal(18, 6), N.Gr3) = 2
                      AND L.TrTp IN (5, 7)
                      AND NULLIF(LTRIM(RTRIM(CONVERT(varchar(100), L.TrInf4))), '') IS NOT NULL
                ),
                ProductionAncestors AS (
                    SELECT DISTINCT
                        TRY_CONVERT(int, NULLIF(LTRIM(RTRIM(CONVERT(varchar(100), R.TrInf2))), '')) AS ProductionOrdNo,
                        P.OrdNo AS CurrentOrdNo,
                        P.OrdBasNo AS NextOrdNo,
                        CAST('|' + CONVERT(varchar(30), P.OrdNo) + '|' AS varchar(max)) AS OrderPath,
                        0 AS Depth
                    FROM RelevantNestingLines R
                    INNER JOIN Ord P WITH(NOLOCK)
                        ON P.OrdNo = TRY_CONVERT(int, NULLIF(LTRIM(RTRIM(CONVERT(varchar(100), R.TrInf2))), ''))
                    WHERE P.TrTp <> 6
                    UNION ALL
                    SELECT
                        Parent.ProductionOrdNo,
                        Base.OrdNo AS CurrentOrdNo,
                        Base.OrdBasNo AS NextOrdNo,
                        CAST(Parent.OrderPath + CONVERT(varchar(30), Base.OrdNo) + '|' AS varchar(max)) AS OrderPath,
                        Parent.Depth + 1 AS Depth
                    FROM ProductionAncestors Parent
                    INNER JOIN Ord Base WITH(NOLOCK) ON Base.OrdNo = Parent.NextOrdNo
                    WHERE Base.TrTp <> 6
                      AND Parent.Depth < 50
                      AND CHARINDEX('|' + CONVERT(varchar(30), Base.OrdNo) + '|', Parent.OrderPath) = 0
                ),
                ProductionToSales AS (
                    SELECT
                        A.ProductionOrdNo,
                        MAX(CASE WHEN O.OrdTp = 1 AND O.TrTp = 1 THEN A.CurrentOrdNo ELSE 0 END) AS SalesOrderNo
                    FROM ProductionAncestors A
                    INNER JOIN Ord O WITH(NOLOCK) ON O.OrdNo = A.CurrentOrdNo
                    GROUP BY A.ProductionOrdNo
                )
                SELECT DISTINCT
                    R.NestingOrdNo,
                    R.OrdDt,
                    COALESCE(NULLIF(TRY_CONVERT(int, R.LineR4), 0), NULLIF(TX.TransactionSalesOrdNo, 0), NULLIF(TRY_CONVERT(int, R.OrderR4), 0), 0) AS DirectSalesOrdNo,
                    R.LnNo,
                    R.TrInf1,
                    R.TrInf2,
                    R.Route,
                    R.ProdNo,
                    R.Descr,
                    R.TrTp,
                    R.NoOrg,
                    R.NoFin,
                    R.CstPr,
                    R.IncCst,
                    P.Gr6,
                    P.Gr5,
                    P.Gr9,
                    P.ProdGr,
                    COALESCE(TRY_CONVERT(int, NULLIF(LTRIM(RTRIM(CONVERT(varchar(100), R.TrInf2))), '')), 0) AS ProductionOrdNo,
                    COALESCE(TRY_CONVERT(int, R.LineR4), 0) AS LineSalesOrdNo,
                    COALESCE(TX.TransactionSalesOrdNo, 0) AS TransactionSalesOrdNo,
                    COALESCE(TRY_CONVERT(int, R.OrderR4), 0) AS OrderSalesOrdNo,
                    COALESCE(PO.SalesOrderNo, 0) AS HierarchySalesOrdNo,
                    CASE
                        WHEN NULLIF(TRY_CONVERT(int, R.LineR4), 0) IS NOT NULL THEN 'OrdLn.R4'
                        WHEN NULLIF(TX.TransactionSalesOrdNo, 0) IS NOT NULL THEN 'ProdTr.R4'
                        WHEN NULLIF(TRY_CONVERT(int, R.OrderR4), 0) IS NOT NULL THEN 'Ord.R4'
                        WHEN NULLIF(PO.SalesOrderNo, 0) IS NOT NULL THEN 'OrdBasNo'
                        ELSE ''
                    END AS SalesOrderSource,
                    COALESCE(NULLIF(TRY_CONVERT(int, R.LineR4), 0), NULLIF(TX.TransactionSalesOrdNo, 0), NULLIF(TRY_CONVERT(int, R.OrderR4), 0), NULLIF(PO.SalesOrderNo, 0), 0) AS SalesOrderNo
                FROM RelevantNestingLines R
                LEFT JOIN Prod P WITH(NOLOCK) ON P.ProdNo = R.ProdNo
                OUTER APPLY (
                    SELECT TOP 1 COALESCE(TRY_CONVERT(int, T.R4), 0) AS TransactionSalesOrdNo
                    FROM ProdTr T WITH(NOLOCK)
                    WHERE R.TrTp = 7
                      AND NULLIF(TRY_CONVERT(int, R.LineR4), 0) IS NULL
                      AND T.OrdNo = R.NestingOrdNo
                      AND T.OrdLnNo = R.LnNo
                      AND T.ProdNo = R.ProdNo
                    ORDER BY TRY_CONVERT(int, T.FinDt) DESC, COALESCE(TRY_CONVERT(int, T.FinTm), 0) DESC
                ) TX
                LEFT JOIN ProductionToSales PO
                    ON PO.ProductionOrdNo = TRY_CONVERT(int, NULLIF(LTRIM(RTRIM(CONVERT(varchar(100), R.TrInf2))), ''))
                ORDER BY R.NestingOrdNo, R.Route, R.LnNo
                OPTION (MAXRECURSION 100)
            `);

        const restResult = await pool.request()
            .input('startDate', sql.Int, startDate)
            .input('endDate', sql.Int, endDate)
            .query(`
                SELECT
                    F.ProdNo,
                    F.OrdNo AS NestingOrdNo,
                    F.Txt1,
                    F.Txt2,
                    F.Val5,
                    F.Val8,
                    LEFT(CONVERT(varchar(100), F.ProdNo), 3) AS PriceType,
                    P.Descr
                FROM FreeInf1 F WITH(NOLOCK)
                INNER JOIN Ord N WITH(NOLOCK) ON N.OrdNo = F.OrdNo
                LEFT JOIN Prod P WITH(NOLOCK) ON P.ProdNo = F.ProdNo
                WHERE F.FrInfTp = 120
                  AND COALESCE(TRY_CONVERT(decimal(18, 6), F.Gr7), 0) = 1
                  AND TRY_CONVERT(int, N.OrdDt) >= @startDate
                  AND TRY_CONVERT(int, N.OrdDt) < @endDate
            `);

        const built = buildRouteLineage(
            lineResult.recordset || [],
            restResult.recordset || [],
            typeof getRestPrices === 'function' ? getRestPrices() : {}
        );
        const payload = { generatedAt: new Date().toISOString(), ...built };
        cache = { expiresAt: Date.now() + cacheTtlMs, payload };
        return payload;
    }

    async function getMovementEvidence({ from, to, products, salesOrders, nestingOrders } = {}) {
        const fromMoment = toVismaMoment(from, 'Fra-dato');
        const toMoment = toVismaMoment(to, 'Til-dato');
        if (fromMoment.timestamp > toMoment.timestamp || toMoment.timestamp - fromMoment.timestamp > 370 * 24 * 60 * 60 * 1000) {
            const error = new Error('Afstemningsperioden er ugyldig eller længere end 370 dage');
            error.statusCode = 400;
            throw error;
        }
        const productNos = uniqueValues(products, 250, value => String(value || '').trim());
        const orderNos = uniqueValues(salesOrders, 250, value => {
            const parsed = Number(value || 0);
            return Number.isFinite(parsed) && parsed > 0 ? Math.trunc(parsed) : 0;
        });
        const nestingOrderNos = uniqueValues(nestingOrders, 250, value => {
            const parsed = Number(value || 0);
            return Number.isFinite(parsed) && parsed > 0 ? Math.trunc(parsed) : 0;
        });
        if (!productNos.length && !orderNos.length && !nestingOrderNos.length) {
            return { purchases: [], plateConsumptions: [], invoices: [], unregisteredSearchPlates: [] };
        }

        const pool = await getConnection();
        let transactionRows = [];
        if (productNos.length) {
            const request = pool.request()
                .input('fromDate', sql.Int, fromMoment.date)
                .input('fromTime', sql.Int, fromMoment.time)
                .input('toDate', sql.Int, toMoment.date)
                .input('toTime', sql.Int, toMoment.time);
            const placeholders = productNos.map((prodNo, index) => {
                const name = 'prodNo' + index;
                request.input(name, sql.VarChar(100), prodNo);
                return '@' + name;
            });
            const result = await request.query(`
                SELECT
                    LTRIM(RTRIM(CONVERT(varchar(100), ProdNo))) AS ProdNo,
                    OrdNo,
                    TrTp,
                    FinDt,
                    FinTm,
                    StcMov,
                    StcCst,
                    SupNo,
                    LTRIM(RTRIM(CONVERT(varchar(100), InvoNo))) AS InvoNo
                FROM ProdTr WITH(NOLOCK)
                WHERE ProdNo IN (${placeholders.join(', ')})
                  AND TrTp IN (5, 6)
                  AND COALESCE(TRY_CONVERT(decimal(18, 6), StcMov), 0) <> 0
                  AND (
                        TRY_CONVERT(int, FinDt) > @fromDate
                     OR (TRY_CONVERT(int, FinDt) = @fromDate AND COALESCE(TRY_CONVERT(int, FinTm), 0) >= @fromTime)
                  )
                  AND (
                        TRY_CONVERT(int, FinDt) < @toDate
                     OR (TRY_CONVERT(int, FinDt) = @toDate AND COALESCE(TRY_CONVERT(int, FinTm), 0) <= @toTime)
                  )
                ORDER BY FinDt, FinTm, ProdNo, OrdNo
            `);
            transactionRows = result.recordset || [];
        }

        let invoiceRows = [];
        if (orderNos.length) {
            const request = pool.request();
            const placeholders = orderNos.map((ordNo, index) => {
                const name = 'ordNo' + index;
                request.input(name, sql.Int, ordNo);
                return '@' + name;
            });
            const result = await request.query(`
                SELECT
                    OrdNo,
                    LTRIM(RTRIM(CONVERT(varchar(100), InvoNo))) AS InvoNo,
                    InvoAm,
                    InvoIF,
                    DInvoIF,
                    FinDt
                FROM Ord WITH(NOLOCK)
                WHERE OrdNo IN (${placeholders.join(', ')})
            `);
            invoiceRows = (result.recordset || []).filter(row => String(row.InvoNo || '').trim());
        }

        let searchPlateRows = [];
        if (nestingOrderNos.length) {
            const request = pool.request();
            const placeholders = nestingOrderNos.map((ordNo, index) => {
                const name = 'nestingOrdNo' + index;
                request.input(name, sql.Int, ordNo);
                return '@' + name;
            });
            const result = await request.query(`
                SELECT DISTINCT
                    L.OrdNo,
                    LTRIM(RTRIM(CONVERT(varchar(100), L.TrInf4))) AS Route,
                    LTRIM(RTRIM(CONVERT(varchar(100), L.ProdNo))) AS ProdNo,
                    LTRIM(RTRIM(CONVERT(nvarchar(100), L.TrInf1))) AS SourceInfo
                FROM OrdLn L WITH(NOLOCK)
                WHERE L.OrdNo IN (${placeholders.join(', ')})
                  AND L.TrTp = 5
                  AND LOWER(LTRIM(RTRIM(CONVERT(nvarchar(100), L.TrInf1)))) LIKE N'søg%'
            `);
            searchPlateRows = result.recordset || [];
        }

        const normalizeTransaction = row => ({
            prodNo: String(row.ProdNo || '').trim(),
            orderNo: toNumber(row.OrdNo),
            date: toNumber(row.FinDt),
            time: toNumber(row.FinTm),
            quantity: round(row.StcMov),
            value: round(Math.abs(toNumber(row.StcCst))),
            supplierNo: toNumber(row.SupNo),
            invoiceNo: String(row.InvoNo || '').trim()
        });
        return {
            purchases: transactionRows.filter(row => Number(row.TrTp) === 6 && toNumber(row.StcMov) > 0).map(normalizeTransaction),
            plateConsumptions: transactionRows.filter(row => Number(row.TrTp) === 5 && toNumber(row.StcMov) < 0).map(normalizeTransaction),
            invoices: invoiceRows.map(row => ({
                orderNo: toNumber(row.OrdNo),
                invoiceNo: String(row.InvoNo || '').trim(),
                invoicedAmount: round(row.InvoAm),
                remainingToInvoice: round(row.DInvoIF),
                finishedDate: toNumber(row.FinDt)
            })),
            unregisteredSearchPlates: searchPlateRows.map(row => ({
                orderNo: toNumber(row.OrdNo),
                route: String(row.Route || '').trim(),
                prodNo: String(row.ProdNo || '').trim(),
                sourceInfo: String(row.SourceInfo || '').trim()
            }))
        };
    }

    return { getCurrentRoutes, getMovementEvidence };
}

module.exports = { createLagerliste2Service, buildRouteLineage, isPlateLine, isUnregisteredSearchPlate };
