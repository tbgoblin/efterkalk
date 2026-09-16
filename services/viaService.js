// ── SalgOrdre VIA ───────────────────────────────────────────────────────────
// Estratto verbatim da routes/apiRoutes.js: query per gli ordini di vendita
// produzione ricorsiva e perimetro storico di valorizzazione, minuti
// consuntivati (ProdTr), regola
// LASER EAGLE su R1100 e costo materiale (righe normali + nesting L).
function toNumber(value) {
    const parsed = Number(value || 0);
    return Number.isFinite(parsed) ? parsed : 0;
}

function round(value) {
    return Math.round(toNumber(value) * 100) / 100;
}

function normalizePurchasedPartRows(recordset) {
    return (Array.isArray(recordset) ? recordset : []).map(row => {
        let sourceDetails = [];
        try {
            sourceDetails = JSON.parse(String(row.PurchasedPartDetailsJson || '[]'));
        } catch (_) {
            sourceDetails = [];
        }
        const purchasedPartDetails = sourceDetails.map(detail => {
            const orderedQty = Math.max(0, toNumber(detail.orderedQty));
            const receivedQty = Math.max(0, toNumber(detail.receivedQty));
            const consumedQty = Math.max(0, toNumber(detail.consumedQty));
            const receivedForOrderQty = orderedQty > 0 ? Math.min(receivedQty, orderedQty) : receivedQty;
            const receivedNotConsumedQty = Math.max(0, receivedForOrderQty - consumedQty);
            const isFollowUpStock = Boolean(detail.isFollowUpStock);
            const activeReservedQty = Math.max(0, toNumber(detail.activeReservedQty));
            const physicalStockQty = Math.max(0, toNumber(detail.physicalStockQty));
            const stockTransferQty = isFollowUpStock
                ? Math.min(receivedNotConsumedQty, activeReservedQty, physicalStockQty)
                : 0;
            const directReceivedQty = isFollowUpStock ? 0 : receivedNotConsumedQty;
            const unitPrice = toNumber(detail.unitPrice);
            const stockUnitCost = toNumber(detail.stockUnitCost) || unitPrice;
            const consumedValue = consumedQty * unitPrice;
            const stockTransferValue = stockTransferQty * stockUnitCost;
            const directReceivedValue = directReceivedQty * unitPrice;
            return {
                ...detail,
                orderedQty: round(orderedQty),
                receivedQty: round(receivedQty),
                consumedQty: round(consumedQty),
                activeReservedQty: round(activeReservedQty),
                physicalStockQty: round(physicalStockQty),
                unitPrice: round(unitPrice),
                stockUnitCost: round(stockUnitCost),
                stockTransferQty: round(stockTransferQty),
                stockTransferValue: round(stockTransferValue),
                countedQty: round(consumedQty + stockTransferQty + directReceivedQty),
                countedValue: round(consumedValue + stockTransferValue + directReceivedValue)
            };
        });
        const { PurchasedPartDetailsJson, ...rest } = row;
        return {
            ...rest,
            PurchasedPartCost: round(purchasedPartDetails.reduce((sum, detail) => sum + detail.countedValue, 0)),
            PurchasedPartDetails: purchasedPartDetails
        };
    });
}

function getOpenViaOrders(flow) {
    if (!flow || !Array.isArray(flow.rows) || !Number.isFinite(flow.total?.closing)) {
        throw new Error('Ordrebeholdning kunne ikke beregnes');
    }
    const seen = new Set();
    for (const row of flow.rows) {
        if (!Number.isSafeInteger(row.ordNo) || row.ordNo <= 0 || !Number.isFinite(row.closing) || seen.has(row.ordNo)) {
            throw new Error('Ugyldige eller dublerede ordrer i Ordrebeholdning');
        }
        seen.add(row.ordNo);
    }
    return flow.rows.filter(row => row.closing > 0.01);
}

function buildViaBacklog(flow, costRows, requestedOrdNo = null) {
    const openOrders = getOpenViaOrders(flow);
    const costs = new Map();
    for (const row of costRows) {
        const ordNo = Number(row.OrdNo);
        if (costs.has(ordNo)) throw new Error('Dublerede kostdata for ordre ' + ordNo);
        costs.set(ordNo, row);
    }
    const rows = openOrders.filter(order => requestedOrdNo === null || order.ordNo === requestedOrdNo).map(order => {
        const cost = costs.get(order.ordNo);
        const costDataAvailable = !!cost && ['MaterialCost', 'StangCost', 'PurchasedPartCost', 'TimeCost'].every(key => cost[key] != null && Number.isFinite(Number(cost[key])));
        return {
            ...cost,
            OrdNo: order.ordNo,
            CustomerName: order.customerName,
            CustNo: order.custNo,
            OrderDate: order.orderDate,
            Gr12: order.gr12,
            OrdPrSt: order.ordPrSt,
            SalesValue: order.value,
            RemainingSalesValue: order.closing,
            CostDataAvailable: costDataAvailable,
            ...(!costDataAvailable ? { MaterialCost: null, StangCost: null, PurchasedPartCost: null, TimeCost: null, PurchasedPartDetails: [] } : {})
        };
    });
    return {
        scope: 'open-backlog', rows, month: flow.month, asOf: flow.asOf,
        unknownCount: flow.unknownCount || 0,
        missingCostCount: rows.filter(row => !row.CostDataAvailable).length,
        orderBacklogValueDkk: rows.reduce((sum, row) => sum + row.RemainingSalesValue, 0),
        excludedResidualDkk: flow.total.closing - openOrders.reduce((sum, row) => sum + row.closing, 0)
    };
}

async function fetchSalgordreViaRows({ getConnection, sql, requestedOrdNo = null, orderNos = null }) {
    if (requestedOrdNo !== null && (!Number.isSafeInteger(requestedOrdNo) || requestedOrdNo <= 0)) {
        throw new Error('Ordrenummer ugyldigt');
    }
    if (orderNos !== null && (!Array.isArray(orderNos) || orderNos.some(value => !Number.isSafeInteger(value) || value <= 0))) {
        throw new Error('Ordrenumre ugyldige');
    }
    const selectedOrders = orderNos === null ? null : [...new Set(orderNos)].filter(value => requestedOrdNo === null || value === requestedOrdNo);
    if (selectedOrders && selectedOrders.length === 0) return [];
    if (selectedOrders && selectedOrders.length > 1000) {
        const rows = [];
        for (let offset = 0; offset < selectedOrders.length; offset += 1000) {
            rows.push(...await fetchSalgordreViaRows({ getConnection, sql, requestedOrdNo, orderNos: selectedOrders.slice(offset, offset + 1000) }));
        }
        return rows;
    }
    const pool = await getConnection();
    const request = pool.request();
    request.timeout = 60000;
    request.input('requestedOrdNo', sql.Numeric, requestedOrdNo);
    const scopePredicate = selectedOrders
        ? 'OrdNo IN (' + selectedOrders.map((value, index) => {
            request.input('viaOrder' + index, sql.Numeric, value);
            return '@viaOrder' + index;
        }).join(',') + ')'
        : `Gr12 <> 10 AND (
            OrdPrSt & 256 = 256 OR OrdPrSt = 0 OR OrdPrSt = 402653456
            OR OrdPrSt = 134217728 OR OrdPrSt & 4194304 = 4194304
        )`;
    const result = await request.query(`
                WITH OpenSalesOrders AS (
                    SELECT OrdNo, DelDt, CreUsr, CustNo, OrdTp, TrTp, Gr12, OrdPrSt, InvoSF, InvoIF, ExRt
                    FROM Ord WITH(NOLOCK)
                                        WHERE OrdTp = 1
                                            AND TrTp = 1
                                            AND (${scopePredicate})
                                            AND (@requestedOrdNo IS NULL OR OrdNo = @requestedOrdNo)
                ),
                ProductionOrders AS (
                    SELECT DISTINCT
                        P.OrdBasNo AS SalesOrderNo,
                        P.OrdNo
                    FROM Ord P WITH(NOLOCK)
                    INNER JOIN OpenSalesOrders S ON S.OrdNo = P.OrdBasNo
                    WHERE P.TrTp <> 6
                    UNION ALL
                    SELECT
                        ProductionOrders.SalesOrderNo,
                        P.OrdNo
                    FROM ProductionOrders
                    INNER JOIN Ord P WITH(NOLOCK) ON P.OrdBasNo = ProductionOrders.OrdNo
                    WHERE P.TrTp <> 6
                ),
                ResourceMinutes AS (
                    SELECT
                        ProductionOrders.SalesOrderNo,
                        L.OrdNo,
                        L.LnNo,
                        CASE
                            WHEN R.Nm LIKE '%laser%'
                             AND ISNULL(Nesting.TotalLaserLines, 0) > 0
                             AND Nesting.FinishedLaserLines = Nesting.TotalLaserLines
                            THEN 80
                            ELSE ISNULL(L.TransGr3, 0)
                        END AS ResourceStatus,
                        ISNULL(ProdTr.FinishedMinutes, 0) AS FinishedMinutes,
                        ISNULL(L.NoOrg, 0) AS PlannedMinutes
                        ,CAST(ISNULL(L.CCstPr, 0) AS decimal(28, 6)) * CASE
                            WHEN UPPER(ISNULL(L.ProdNo, '')) = 'R1100'
                             AND UPPER(ISNULL(LastEmployee.EmployeeName, '')) LIKE '%LASER EAGLE%'
                            THEN 2
                            ELSE 1
                         END AS ResourceUnitCost
                    FROM ProductionOrders
                    INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = ProductionOrders.OrdNo
                    LEFT JOIN R7 R WITH(NOLOCK) ON R.RNo = L.R7
                    OUTER APPLY (
                        SELECT SUM(CAST(P.NoInvoAb AS decimal(28, 6))) AS FinishedMinutes
                        FROM ProdTr P WITH(NOLOCK)
                        WHERE P.OrdNo = L.OrdNo
                          AND P.OrdLnNo = L.LnNo
                    ) ProdTr
                                        OUTER APPLY (
                                                SELECT TOP 1 A.Nm AS EmployeeName
                                                FROM ProdTr P WITH(NOLOCK)
                                                LEFT JOIN Actor A WITH(NOLOCK) ON A.EmpNo = P.EmpNo
                                                WHERE P.OrdNo = L.OrdNo
                                                    AND P.OrdLnNo = L.LnNo
                                                ORDER BY P.FinDt DESC, P.FinTm DESC
                                        ) LastEmployee
                    OUTER APPLY (
                        SELECT
                            COUNT(*) AS TotalLaserLines,
                            SUM(CASE WHEN ISNULL(N.NoFin, 0) > 0 THEN 1 ELSE 0 END) AS FinishedLaserLines
                        FROM OrdLn N WITH(NOLOCK)
                        WHERE N.TrInf2 = CONVERT(varchar(20), L.OrdNo)
                          AND N.TrTp = 7
                          AND N.ProdNo LIKE '%L%'
                    ) Nesting
                    WHERE L.ProdTp4 IN (1, 3)
                      AND L.R7 <> ''
                ),
                ActiveProduction AS (
                    SELECT
                        SalesOrderNo,
                        COUNT(DISTINCT CASE WHEN ResourceStatus < 80 THEN OrdNo END) AS OpenProductionOrders,
                        COUNT(*) AS TotalResources,
                        SUM(CASE WHEN ResourceStatus = 80 THEN 1 ELSE 0 END) AS CompletedResources,
                        SUM(CASE WHEN ResourceStatus < 80 THEN 1 ELSE 0 END) AS RemainingResources,
                        SUM(FinishedMinutes) AS CompletedResourceMinutes,
                        SUM(PlannedMinutes) AS EffectiveResourceMinutes,
                        SUM(CONVERT(float, FinishedMinutes) * CONVERT(float, ResourceUnitCost)) AS TimeCost
                    FROM ResourceMinutes
                    GROUP BY SalesOrderNo
                ),
                MaterialCosts AS (
                    SELECT
                        ProductionOrders.SalesOrderNo,
                        SUM(CONVERT(float, ISNULL(L.NoFin, 0)) * CONVERT(float, ISNULL(L.CCstPr, 0))) AS MaterialCost
                    FROM ProductionOrders
                    INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = ProductionOrders.OrdNo
                    LEFT JOIN Prod P WITH(NOLOCK) ON P.ProdNo = L.ProdNo
                                        WHERE L.ProdTp4 = 2
                      AND L.ProdNo NOT LIKE '%L'
                      AND ISNULL(P.Gr6, 0) <> 2
                      AND ISNULL(L.PurcNo, 0) = 0
                    GROUP BY ProductionOrders.SalesOrderNo
                ),
                                    StangCosts AS (
                                        SELECT
                                            ProductionOrders.SalesOrderNo,
                                            SUM(CONVERT(float, ISNULL(L.NoFin, 0)) * CONVERT(float, ISNULL(L.CCstPr, 0))) AS StangCost
                                        FROM ProductionOrders
                                        INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = ProductionOrders.OrdNo
                                        INNER JOIN Prod P WITH(NOLOCK) ON P.ProdNo = L.ProdNo
                                        WHERE P.Gr6 = 2
                                          AND L.ProdNo NOT LIKE '%L'
                                          AND NOT (
                                                ISNULL(L.PurcNo, 0) <> 0
                                            AND EXISTS (
                                                    SELECT 1 FROM Rsv R WITH(NOLOCK)
                                                    WHERE R.OrdNo = L.OrdNo AND R.OrdLnNo = L.LnNo
                                                )
                                          )
                                        GROUP BY ProductionOrders.SalesOrderNo
                                    ),
                                    PurchasedPartRawLines AS (
                                        SELECT
                                            ProductionOrders.SalesOrderNo,
                                            L.OrdNo AS ProductionOrderNo,
                                            L.LnNo AS ProductionLineNo,
                                            L.ProdNo,
                                            COALESCE(PP.Descr, L.Descr, '') AS Descr,
                                            L.PurcNo AS PurchaseOrderNo,
                                            ISNULL(L.NoOrg, 0) AS OrderedQty,
                                            ISNULL(Received.ReceivedQty, 0) AS ReceivedQty,
                                            ISNULL(L.NoFin, 0) AS ConsumedQty,
                                            ISNULL(PurchasePrice.UnitPrice, ISNULL(L.CCstPr, 0)) AS UnitPrice,
                                            CASE WHEN ISNULL(PP.Gr9, 0) = 1 THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS IsFollowUpStock,
                                            ISNULL(Stock.PoPhStB, 0) AS PhysicalStockQty,
                                            ISNULL(Stock.PhCstPr, 0) AS StockUnitCost,
                                            ISNULL(Reservation.ActiveReservedQty, 0) AS ActiveReservedQty
                                        FROM ProductionOrders
                                        INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = ProductionOrders.OrdNo
                                        LEFT JOIN Prod PP WITH(NOLOCK) ON PP.ProdNo = L.ProdNo
                                        LEFT JOIN StcBal Stock WITH(NOLOCK) ON Stock.ProdNo = L.ProdNo AND Stock.StcNo = 1
                                                                                OUTER APPLY (
                                                                                        SELECT TOP 1
                                                                                                COALESCE(NULLIF(PurchaseLine.DPrice, 0), PurchaseLine.CCstPr, 0) AS UnitPrice
                                                                                        FROM Ord PurchaseOrder WITH(NOLOCK)
                                                                                        INNER JOIN OrdLn PurchaseLine WITH(NOLOCK) ON PurchaseLine.OrdNo = PurchaseOrder.OrdNo
                                                                                        WHERE PurchaseOrder.OrdNo = L.PurcNo
                                                                                            AND PurchaseOrder.TrTp = 6
                                                                                            AND PurchaseLine.ProdNo = L.ProdNo
                                                                                        ORDER BY PurchaseLine.LnNo
                                                                                ) PurchasePrice
                                                                                OUTER APPLY (
                                                                                        SELECT SUM(CASE
                                                                                                WHEN ISNULL(T.StcMov, 0) > 0 THEN ISNULL(T.StcMov, 0)
                                                                                                ELSE 0
                                                                                        END) AS ReceivedQty
                                                                                        FROM ProdTr T WITH(NOLOCK)
                                                                                        WHERE T.OrdNo = L.PurcNo
                                                                                            AND T.ProdNo = L.ProdNo
                                                                                            AND T.TrTp = 6
                                                                                ) Received
                                                                                        OUTER APPLY (
                                                                                            SELECT SUM(CASE
                                                                                                WHEN ISNULL(R.NoRsv, 0) > ISNULL(R.NoFin, 0)
                                                                                                    THEN ISNULL(R.NoRsv, 0) - ISNULL(R.NoFin, 0)
                                                                                                ELSE 0
                                                                                            END) AS ActiveReservedQty
                                                                                            FROM Rsv R WITH(NOLOCK)
                                                                                            WHERE R.OrdNo = L.OrdNo
                                                                                              AND R.OrdLnNo = L.LnNo
                                                                                              AND R.ProdNo = L.ProdNo
                                                                                        ) Reservation
                                                                                WHERE L.PurcNo IS NOT NULL
                                                                                    AND L.PurcNo <> 0
                                          AND L.ProdTp4 = 2
                                          AND L.ProdNo NOT LIKE '%L'
                                          AND (
                                                ISNULL(PP.Gr6, 0) <> 2
                                             OR EXISTS (
                                                    SELECT 1 FROM Rsv R WITH(NOLOCK)
                                                    WHERE R.OrdNo = L.OrdNo AND R.OrdLnNo = L.LnNo
                                                )
                                          )
                                    ),
                                    PurchasedPartLines AS (
                                        SELECT
                                            SalesOrderNo,
                                            MIN(ProductionOrderNo) AS ProductionOrderNo,
                                            MIN(ProductionLineNo) AS ProductionLineNo,
                                            ProdNo,
                                            MAX(Descr) AS Descr,
                                            PurchaseOrderNo,
                                            SUM(OrderedQty) AS OrderedQty,
                                            MAX(ReceivedQty) AS ReceivedQty,
                                            SUM(ConsumedQty) AS ConsumedQty,
                                            MAX(UnitPrice) AS UnitPrice,
                                            MAX(CAST(IsFollowUpStock AS int)) AS IsFollowUpStock,
                                            MAX(PhysicalStockQty) AS PhysicalStockQty,
                                            MAX(StockUnitCost) AS StockUnitCost,
                                            SUM(ActiveReservedQty) AS ActiveReservedQty
                                        FROM PurchasedPartRawLines
                                        GROUP BY SalesOrderNo, PurchaseOrderNo, ProdNo
                                    ),
                NestingMaterialCosts AS (
                    SELECT
                        ProductionOrders.SalesOrderNo,
                        SUM(CONVERT(float, ISNULL(N.NoFin, 0)) * CONVERT(float, ISNULL(N.CstPr, 0))) AS MaterialCost
                    FROM ProductionOrders
                    INNER JOIN OrdLn N WITH(NOLOCK)
                        ON N.TrInf2 = CONVERT(varchar(20), ProductionOrders.OrdNo)
                       AND N.TrTp = 7
                       AND N.ProdNo LIKE '%L'
                    GROUP BY ProductionOrders.SalesOrderNo
                )
                SELECT
                    S.OrdNo,
                    S.DelDt AS DeliveryDate,
                    S.CreUsr AS SellerUsr,
                    S.OrdTp,
                    S.TrTp,
                    S.Gr12,
                    S.OrdPrSt,
                    C.Nm AS CustomerName,
                    MainLine.ProdNo AS MainProdNo,
                    MainLine.Descr AS MainProdDescr,
                    Active.OpenProductionOrders,
                    Active.TotalResources,
                    Active.CompletedResources,
                    Active.RemainingResources,
                    Active.CompletedResourceMinutes,
                    Active.EffectiveResourceMinutes,
                    ISNULL(MaterialCosts.MaterialCost, 0) + ISNULL(NestingMaterialCosts.MaterialCost, 0) AS MaterialCost,
                    ISNULL(StangCosts.StangCost, 0) AS StangCost,
                    CAST(0 AS decimal(18, 6)) AS PurchasedPartCost,
                    ISNULL((
                        SELECT
                            D.ProductionOrderNo AS productionOrderNo,
                            D.ProductionLineNo AS productionLineNo,
                            D.ProdNo AS prodNo,
                            D.Descr AS descr,
                            D.PurchaseOrderNo AS purchaseOrderNo,
                            D.OrderedQty AS orderedQty,
                            D.ReceivedQty AS receivedQty,
                            D.ConsumedQty AS consumedQty,
                            D.UnitPrice AS unitPrice,
                            D.IsFollowUpStock AS isFollowUpStock,
                            D.PhysicalStockQty AS physicalStockQty,
                            D.StockUnitCost AS stockUnitCost,
                            D.ActiveReservedQty AS activeReservedQty
                        FROM PurchasedPartLines D
                        WHERE D.SalesOrderNo = S.OrdNo
                        ORDER BY D.ProductionOrderNo, D.ProductionLineNo
                        FOR JSON PATH
                    ), '[]') AS PurchasedPartDetailsJson,
                    ISNULL(Active.TimeCost, 0) AS TimeCost,
                    CONVERT(decimal(38, 6), (CONVERT(float, ISNULL(S.InvoSF, 0)) + CONVERT(float, ISNULL(S.InvoIF, 0))) * (CONVERT(float, ISNULL(NULLIF(S.ExRt, 0), 100)) / 100.0)) AS SalesValue,
                    CAST(NULL AS datetime) AS PlannedDate,
                    CAST(NULL AS varchar(100)) AS ResourceName
                FROM OpenSalesOrders S
                LEFT JOIN Actor C WITH(NOLOCK) ON C.CustNo = S.CustNo
                OUTER APPLY (
                    SELECT TOP 1 ML.ProdNo, ML.Descr
                    FROM OrdLn ML WITH(NOLOCK)
                    WHERE ML.OrdNo = S.OrdNo
                    ORDER BY ML.LnNo
                ) MainLine
                LEFT JOIN ActiveProduction Active ON Active.SalesOrderNo = S.OrdNo
                LEFT JOIN MaterialCosts ON MaterialCosts.SalesOrderNo = S.OrdNo
                LEFT JOIN StangCosts ON StangCosts.SalesOrderNo = S.OrdNo
                LEFT JOIN NestingMaterialCosts ON NestingMaterialCosts.SalesOrderNo = S.OrdNo
                ORDER BY
                    CASE WHEN S.DelDt > 19800101 THEN S.DelDt ELSE 99991231 END,
                    S.OrdNo
            `);
    return normalizePurchasedPartRows(result.recordset || []);
}

module.exports = { fetchSalgordreViaRows, normalizePurchasedPartRows, getOpenViaOrders, buildViaBacklog };
