function toNumber(value) {
    if (typeof value === 'string') {
        const parsed = Number(value.replace(/\s+/g, '').replace(',', '.'));
        return Number.isFinite(parsed) ? parsed : 0;
    }
    const parsed = Number(value || 0);
    return Number.isFinite(parsed) ? parsed : 0;
}

function round(value) {
    const result = Math.round(toNumber(value) * 100) / 100;
    return Object.is(result, -0) ? 0 : result;
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

function buildOrderStates(rows) {
    const groups = new Map();
    for (const row of Array.isArray(rows) ? rows : []) {
        const orderNo = toNumber(row.OrdNo);
        if (!orderNo) continue;
        const group = groups.get(orderNo) || {
            orderNo,
            finishedDate: toNumber(row.FinDt),
            invoiceNo: String(row.InvoNo || '').trim(),
            invoicedAmount: round(row.InvoAm),
            remainingToInvoice: round(row.DInvoIF),
            orderPrintStatus: toNumber(row.OrdPrSt),
            lines: []
        };
        if (row.LnNo !== undefined && row.LnNo !== null) {
            group.lines.push({
                lineNo: toNumber(row.LnNo),
                prodNo: String(row.ProdNo || '').trim(),
                transactionType: toNumber(row.LineTrTp),
                plannedQty: round(row.NoOrg),
                finishedQty: round(row.NoFin),
                packedQty: round(row.NoPac),
                invoicedQty: round(row.NoInvo),
                invoiceableQty: round(row.NoInvoAb),
                costPrice: round(row.CCstPr)
            });
        }
        groups.set(orderNo, group);
    }

    return Array.from(groups.values()).map(group => {
        const saleLines = group.lines.filter(line => line.transactionType === 1);
        const productLines = saleLines.filter(line => /^1/.test(line.prodNo));
        const relevantLines = productLines.length ? productLines : saleLines;
        let finishedWeight = 0;
        let packedWeight = 0;
        let invoicedWeight = 0;
        for (const line of relevantLines) {
            const finished = Math.abs(toNumber(line.finishedQty));
            if (finished <= 0.005) continue;
            const unitWeight = Math.abs(toNumber(line.costPrice)) || 1;
            const lineWeight = finished * unitWeight;
            finishedWeight += lineWeight;
            packedWeight += Math.min(finished, Math.abs(toNumber(line.packedQty))) * unitWeight;
            invoicedWeight += Math.min(finished, Math.abs(toNumber(line.invoicedQty))) * unitWeight;
        }
        const packedRatio = finishedWeight > 0 ? Math.max(0, Math.min(1, packedWeight / finishedWeight)) : 0;
        const invoicedRatio = finishedWeight > 0 ? Math.max(0, Math.min(1, invoicedWeight / finishedWeight)) : 0;
        return {
            ...group,
            lineCount: relevantLines.length,
            packedRatio: round(packedRatio * 100) / 100,
            invoicedRatio: round(invoicedRatio * 100) / 100,
            hasFinishedQuantity: finishedWeight > 0.005,
            hasPackedQuantity: packedWeight > 0.005,
            fullyPacked: finishedWeight > 0.005 && packedRatio >= 0.999,
            partiallyPacked: packedRatio > 0.001 && packedRatio < 0.999,
            headerClosed: group.finishedDate > 0,
            invoiced: Boolean(group.invoiceNo)
        };
    });
}

function buildReservationSummary(sourceRows) {
    const rows = (Array.isArray(sourceRows) ? sourceRows : []).map(row => {
        const reservedQty = Math.max(0, toNumber(row.NoRsv));
        const pickedQty = Math.max(0, Math.min(reservedQty, toNumber(row.NoPic)));
        const finishedQty = Math.max(0, Math.min(reservedQty, toNumber(row.NoFin)));
        const awaitingPickQty = Math.max(0, reservedQty - pickedQty);
        const pickedNotFinishedQty = Math.max(0, pickedQty - finishedQty);
        const activeQty = awaitingPickQty + pickedNotFinishedQty;
        const costPrice = toNumber(row.CstPr);
        const salesOrderNo = toNumber(row.SalesOrderNo);
        const status = activeQty <= 0.005
            ? 'finished'
            : (awaitingPickQty > 0.005 && pickedNotFinishedQty > 0.005
                ? 'mixed'
                : (awaitingPickQty > 0.005 ? 'reserved' : 'picked'));
        return {
            prodNo: String(row.ProdNo || '').trim(),
            descr: String(row.Descr || '').trim(),
            orderNo: toNumber(row.OrdNo),
            orderLineNo: toNumber(row.OrdLnNo),
            salesOrderNo,
            linkSource: String(row.LinkSource || '').trim(),
            customerName: String(row.CustomerName || '').trim(),
            reservedQty: round(reservedQty),
            pickedQty: round(pickedQty),
            finishedQty: round(finishedQty),
            awaitingPickQty: round(awaitingPickQty),
            pickedNotFinishedQty: round(pickedNotFinishedQty),
            activeQty: round(activeQty),
            costPrice: round(costPrice),
            activeValue: round(activeQty * costPrice),
            registeredValue: round(reservedQty * costPrice),
            physicalStock: round(row.PoPhStB),
            stockReserved: round(row.ShpRsv),
            v1AvailableQty: round(toNumber(row.Bal) + toNumber(row.StcInc) - toNumber(row.ShpRsv)),
            changedDate: toNumber(row.ChDt),
            changedTime: toNumber(row.ChTm),
            status
        };
    });
    const activeRows = rows.filter(row => row.activeQty > 0.005);
    return {
        rows,
        summary: {
            rowCount: rows.length,
            activeRowCount: activeRows.length,
            finishedRowCount: rows.length - activeRows.length,
            linkedRowCount: rows.filter(row => row.salesOrderNo > 0).length,
            unlinkedRowCount: rows.filter(row => row.salesOrderNo <= 0).length,
            activeValue: round(activeRows.reduce((sum, row) => sum + row.activeValue, 0)),
            registeredValue: round(rows.reduce((sum, row) => sum + row.registeredValue, 0))
        }
    };
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
                sourceCode: String(row.TrInf2 || '').trim(),
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
    const groupsByOrderPlateAndRestCode = new Map();
    for (const group of groups.values()) {
        for (const plate of group.plates.concat(group.estimatedRestLines)) {
            const key = String(group.nestingOrdNo) + '|' + plate.prodNo;
            const list = groupsByOrderAndPlate.get(key) || [];
            if (!list.includes(group)) list.push(group);
            groupsByOrderAndPlate.set(key, list);
            const restCode = String(plate.sourceCode || '').trim().toLocaleLowerCase('da-DK');
            if (restCode && toNumber(plate.plannedQty) < 0) {
                const codeKey = key + '|' + restCode;
                const codeList = groupsByOrderPlateAndRestCode.get(codeKey) || [];
                if (!codeList.includes(group)) codeList.push(group);
                groupsByOrderPlateAndRestCode.set(codeKey, codeList);
            }
        }
    }

    const unassignedRest = [];
    for (const row of Array.isArray(restRows) ? restRows : []) {
        const prodNo = String(row.ProdNo || '').trim();
        const nestingOrdNo = toNumber(row.NestingOrdNo || row.OrdNo);
        const priceType = String(row.PriceType || prodNo.slice(0, 3)).trim();
        const preciseWeight = toNumber(row.Val5) * toNumber(row.Val8);
        const weight = round(preciseWeight);
        const pricePerKg = round(restPrices[priceType]);
        const normalized = {
            prodNo,
            nestingOrdNo,
            restCode: String(row.Txt1 || '').trim(),
            label: String(row.Txt2 || row.Txt1 || row.Descr || '').trim(),
            weight,
            pricePerKg,
            value: round(preciseWeight * pricePerKg)
        };
        const baseKey = String(nestingOrdNo) + '|' + prodNo;
        const codeKey = baseKey + '|' + normalized.restCode.toLocaleLowerCase('da-DK');
        const exactCandidates = normalized.restCode ? (groupsByOrderPlateAndRestCode.get(codeKey) || []) : [];
        const candidates = exactCandidates.length === 1 ? exactCandidates : (groupsByOrderAndPlate.get(baseKey) || []);
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
        const estimatedRestFifoValue = round(Math.abs(group.estimatedRestLines.reduce((sum, row) => sum + row.value, 0)));
        const restValue = round(group.restPlates.reduce((sum, row) => sum + row.value, 0));
        const materialAllocationResidual = round(plateValue - completedProductValue - estimatedRestFifoValue);
        const restLinesFinished = group.estimatedRestLines.length > 0
            && group.estimatedRestLines.every(row => {
                const planned = Math.abs(toNumber(row.plannedQty));
                const finished = Math.abs(toNumber(row.finishedQty));
                return planned <= 0.005 ? finished > 0.005 : finished / planned >= 0.999;
            });
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
            estimatedRestFifoValue,
            restValue,
            restWriteDown: round(estimatedRestFifoValue - restValue),
            restRegistrationStatus: group.estimatedRestLines.length
                ? (group.restPlates.length >= group.estimatedRestLines.length
                    ? 'registered'
                    : (restLinesFinished
                        ? (group.restPlates.length ? 'finished_partially_registered' : 'finished_unregistered')
                        : (status !== 'completed'
                        ? (group.restPlates.length ? 'pending_partial' : 'pending')
                        : (group.restPlates.length ? 'partial' : 'missing'))))
                : (group.restPlates.length ? 'registered_without_estimate' : 'none'),
            restLinesFinished,
            materialAllocationResidual,
            residualValue: materialAllocationResidual
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
    let reservationCache = null;
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

    async function getMovementEvidence({ from, to, products, salesOrders, nestingOrders, restCodes } = {}) {
        const fromMoment = toVismaMoment(from, 'Fra-dato');
        const toMoment = toVismaMoment(to, 'Til-dato');
        if (fromMoment.timestamp > toMoment.timestamp || toMoment.timestamp - fromMoment.timestamp > 370 * 24 * 60 * 60 * 1000) {
            const error = new Error('Afstemningsperioden er ugyldig eller længere end 370 dage');
            error.statusCode = 400;
            throw error;
        }
        const productNos = uniqueValues(products, 250, value => String(value || '').trim());
        const orderNos = uniqueValues(salesOrders, 1000, value => {
            const parsed = Number(value || 0);
            return Number.isFinite(parsed) && parsed > 0 ? Math.trunc(parsed) : 0;
        });
        const nestingOrderNos = uniqueValues(nestingOrders, 250, value => {
            const parsed = Number(value || 0);
            return Number.isFinite(parsed) && parsed > 0 ? Math.trunc(parsed) : 0;
        });
        const registeredRestCodes = uniqueValues(restCodes, 250, value => String(value || '').trim());
        if (!productNos.length && !orderNos.length && !nestingOrderNos.length && !registeredRestCodes.length) {
            return { purchases: [], plateConsumptions: [], stockConsumptions: [], reservations: [], invoices: [], orderStates: [], unregisteredSearchPlates: [], registeredRestConsumptions: [] };
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
                    R4,
                    TrTp,
                    FinDt,
                    FinTm,
                    StcMov,
                    StcCst,
                    SupNo,
                    LTRIM(RTRIM(CONVERT(varchar(100), InvoNo))) AS InvoNo
                FROM ProdTr WITH(NOLOCK)
                WHERE ProdNo IN (${placeholders.join(', ')})
                  AND TrTp IN (1, 5, 6)
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

        let reservationRows = [];
        if (productNos.length) {
            const request = pool.request()
                .input('fromDate', sql.Int, fromMoment.date)
                .input('fromTime', sql.Int, fromMoment.time)
                .input('toDate', sql.Int, toMoment.date)
                .input('toTime', sql.Int, toMoment.time);
            const placeholders = productNos.map((prodNo, index) => {
                const name = 'reservationProdNo' + index;
                request.input(name, sql.VarChar(100), prodNo);
                return '@' + name;
            });
            const result = await request.query(`
                SELECT
                    LTRIM(RTRIM(CONVERT(varchar(100), R.ProdNo))) AS ProdNo,
                    R.OrdNo,
                    R.OrdLnNo,
                    R.NoRsv,
                    R.NoPic,
                    R.NoFin,
                    R.CstPr,
                    R.ChDt,
                    R.ChTm,
                    COALESCE(
                        NULLIF(TRY_CONVERT(int, L.R4), 0),
                        NULLIF(TRY_CONVERT(int, O.R4), 0),
                        CASE WHEN O.OrdTp = 1 AND O.TrTp = 1 THEN O.OrdNo END,
                        0
                    ) AS SalesOrderNo,
                    CASE
                        WHEN NULLIF(TRY_CONVERT(int, L.R4), 0) IS NOT NULL THEN 'OrdLn.R4'
                        WHEN NULLIF(TRY_CONVERT(int, O.R4), 0) IS NOT NULL THEN 'Ord.R4'
                        WHEN O.OrdTp = 1 AND O.TrTp = 1 THEN 'OrdNo'
                        ELSE ''
                    END AS LinkSource
                FROM Rsv R WITH(NOLOCK)
                LEFT JOIN Ord O WITH(NOLOCK) ON O.OrdNo = R.OrdNo
                LEFT JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = R.OrdNo AND L.LnNo = R.OrdLnNo
                WHERE R.ProdNo IN (${placeholders.join(', ')})
                  AND COALESCE(TRY_CONVERT(decimal(18, 6), R.NoRsv), 0) > 0
                  AND (
                        TRY_CONVERT(int, R.ChDt) > @fromDate
                     OR (TRY_CONVERT(int, R.ChDt) = @fromDate AND COALESCE(TRY_CONVERT(int, R.ChTm), 0) >= @fromTime)
                  )
                  AND (
                        TRY_CONVERT(int, R.ChDt) < @toDate
                     OR (TRY_CONVERT(int, R.ChDt) = @toDate AND COALESCE(TRY_CONVERT(int, R.ChTm), 0) <= @toTime)
                  )
                ORDER BY R.ChDt, R.ChTm, R.ProdNo, R.OrdNo
            `);
            reservationRows = result.recordset || [];
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
                    O.OrdNo,
                    LTRIM(RTRIM(CONVERT(varchar(100), O.InvoNo))) AS InvoNo,
                    O.InvoAm,
                    O.InvoIF,
                    O.DInvoIF,
                    O.FinDt,
                    O.OrdPrSt,
                    L.LnNo,
                    LTRIM(RTRIM(CONVERT(varchar(100), L.ProdNo))) AS ProdNo,
                    L.TrTp AS LineTrTp,
                    L.NoOrg,
                    L.NoFin,
                    L.NoPac,
                    L.NoInvo,
                    L.NoInvoAb,
                    L.CCstPr
                FROM Ord O WITH(NOLOCK)
                LEFT JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = O.OrdNo
                WHERE O.OrdNo IN (${placeholders.join(', ')})
                ORDER BY O.OrdNo, L.LnNo
            `);
            invoiceRows = result.recordset || [];
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

        let registeredRestConsumptionRows = [];
        if (registeredRestCodes.length) {
            const request = pool.request();
            const placeholders = registeredRestCodes.map((restCode, index) => {
                const name = 'restCode' + index;
                request.input(name, sql.VarChar(100), restCode);
                return '@' + name;
            });
            const result = await request.query(`
                SELECT DISTINCT
                    L.OrdNo,
                    LTRIM(RTRIM(CONVERT(varchar(100), L.TrInf4))) AS Route,
                    LTRIM(RTRIM(CONVERT(varchar(100), L.ProdNo))) AS ProdNo,
                    LTRIM(RTRIM(CONVERT(varchar(100), L.TrInf1))) AS RestCode,
                    L.NoOrg,
                    L.NoFin,
                    L.CstPr,
                    L.IncCst
                FROM OrdLn L WITH(NOLOCK)
                WHERE L.TrTp = 5
                  AND LTRIM(RTRIM(CONVERT(varchar(100), L.TrInf1))) IN (${placeholders.join(', ')})
                ORDER BY L.OrdNo, Route, ProdNo
            `);
            registeredRestConsumptionRows = result.recordset || [];
        }

        const normalizeTransaction = row => ({
            prodNo: String(row.ProdNo || '').trim(),
            orderNo: toNumber(row.OrdNo),
            salesOrderNo: toNumber(row.R4) || toNumber(row.OrdNo),
            date: toNumber(row.FinDt),
            time: toNumber(row.FinTm),
            quantity: round(row.StcMov),
            value: round(Math.abs(toNumber(row.StcCst))),
            supplierNo: toNumber(row.SupNo),
            invoiceNo: String(row.InvoNo || '').trim()
        });
        const orderStates = buildOrderStates(invoiceRows);
        return {
            purchases: transactionRows.filter(row => Number(row.TrTp) === 6 && toNumber(row.StcMov) > 0).map(normalizeTransaction),
            plateConsumptions: transactionRows.filter(row => Number(row.TrTp) === 5 && toNumber(row.StcMov) < 0).map(normalizeTransaction),
            stockConsumptions: transactionRows.filter(row => Number(row.TrTp) === 1 && toNumber(row.StcMov) < 0).map(normalizeTransaction),
            reservations: buildReservationSummary(reservationRows).rows,
            invoices: orderStates.filter(row => row.invoiced).map(row => ({
                orderNo: row.orderNo,
                invoiceNo: row.invoiceNo,
                invoicedAmount: row.invoicedAmount,
                remainingToInvoice: row.remainingToInvoice,
                finishedDate: row.finishedDate
            })),
            orderStates,
            registeredRestConsumptions: registeredRestConsumptionRows.map(row => ({
                restCode: String(row.RestCode || '').trim(),
                nestingOrderNo: toNumber(row.OrdNo),
                route: String(row.Route || '').trim(),
                prodNo: String(row.ProdNo || '').trim(),
                plannedQty: round(row.NoOrg),
                finishedQty: round(row.NoFin),
                value: round(Math.abs(lineValue(row))),
                routeStarted: Math.abs(toNumber(row.NoFin)) > 0.005
            })),
            unregisteredSearchPlates: searchPlateRows.map(row => ({
                orderNo: toNumber(row.OrdNo),
                route: String(row.Route || '').trim(),
                prodNo: String(row.ProdNo || '').trim(),
                sourceInfo: String(row.SourceInfo || '').trim()
            }))
        };
    }

    async function getCurrentReservations({ forceRefresh = false } = {}) {
        if (!forceRefresh && reservationCache && reservationCache.expiresAt > Date.now()) return reservationCache.payload;
        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT
                LTRIM(RTRIM(CONVERT(varchar(100), R.ProdNo))) AS ProdNo,
                LTRIM(RTRIM(COALESCE(P.Descr, ''))) AS Descr,
                R.OrdNo,
                R.OrdLnNo,
                R.NoRsv,
                R.NoPic,
                R.NoFin,
                R.CstPr,
                R.ChDt,
                R.ChTm,
                B.Bal,
                B.StcInc,
                B.ShpRsv,
                B.PoPhStB,
                COALESCE(
                    NULLIF(TRY_CONVERT(int, L.R4), 0),
                    NULLIF(TRY_CONVERT(int, O.R4), 0),
                    CASE WHEN O.OrdTp = 1 AND O.TrTp = 1 THEN O.OrdNo END,
                    0
                ) AS SalesOrderNo,
                CASE
                    WHEN NULLIF(TRY_CONVERT(int, L.R4), 0) IS NOT NULL THEN 'OrdLn.R4'
                    WHEN NULLIF(TRY_CONVERT(int, O.R4), 0) IS NOT NULL THEN 'Ord.R4'
                    WHEN O.OrdTp = 1 AND O.TrTp = 1 THEN 'OrdNo'
                    ELSE ''
                END AS LinkSource,
                LTRIM(RTRIM(COALESCE(A.Nm, A2.Nm, ''))) AS CustomerName
            FROM Rsv R WITH(NOLOCK)
            LEFT JOIN Prod P WITH(NOLOCK) ON P.ProdNo = R.ProdNo
            LEFT JOIN StcBal B WITH(NOLOCK) ON B.ProdNo = R.ProdNo AND B.StcNo = 1
            LEFT JOIN Ord O WITH(NOLOCK) ON O.OrdNo = R.OrdNo
            LEFT JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = R.OrdNo AND L.LnNo = R.OrdLnNo
            LEFT JOIN Actor A WITH(NOLOCK) ON A.CustNo = O.CustNo AND COALESCE(TRY_CONVERT(decimal(18, 6), O.CustNo), 0) <> 0
            LEFT JOIN Ord SO WITH(NOLOCK) ON SO.OrdNo = COALESCE(NULLIF(TRY_CONVERT(int, L.R4), 0), NULLIF(TRY_CONVERT(int, O.R4), 0))
            LEFT JOIN Actor A2 WITH(NOLOCK) ON A2.CustNo = SO.CustNo AND COALESCE(TRY_CONVERT(decimal(18, 6), SO.CustNo), 0) <> 0
            WHERE COALESCE(TRY_CONVERT(decimal(18, 6), R.NoRsv), 0) > 0
            ORDER BY R.ChDt DESC, R.ChTm DESC, R.OrdNo DESC, R.OrdLnNo
        `);
        const built = buildReservationSummary(result.recordset || []);
        const payload = { generatedAt: new Date().toISOString(), ...built };
        reservationCache = { expiresAt: Date.now() + cacheTtlMs, payload };
        return payload;
    }

    return { getCurrentRoutes, getMovementEvidence, getCurrentReservations };
}

module.exports = { createLagerliste2Service, buildRouteLineage, buildOrderStates, buildReservationSummary, isPlateLine, isUnregisteredSearchPlate };
