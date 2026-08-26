// ── SalgOrdre VIA ───────────────────────────────────────────────────────────
// Estratto verbatim da routes/apiRoutes.js: query per gli ordini di vendita
// aperti con produzione ricorsiva, minuti consuntivati (ProdTr), regola
// LASER EAGLE su R1100 e costo materiale (righe normali + nesting L).
async function fetchSalgordreViaRows({ getConnection, sql, requestedOrdNo }) {
    const pool = await getConnection();
    const request = pool.request();
    request.timeout = 60000;
    request.input('requestedOrdNo', sql.Numeric, requestedOrdNo);
    const result = await request.query(`
                WITH OpenSalesOrders AS (
                    SELECT OrdNo, DelDt, CreUsr, CustNo, OrdTp, TrTp, Gr12, OrdPrSt, InvoSF, InvoIF, ExRt
                    FROM Ord WITH(NOLOCK)
                                        WHERE OrdTp = 1
                                            AND TrTp = 1
                                            AND Gr12 <> 10
                                            AND (
                                                    OrdPrSt & 256 = 256
                                                    OR OrdPrSt = 0
                                                    OR OrdPrSt = 402653456
                                                    OR OrdPrSt = 134217728
                                                    OR OrdPrSt & 4194304 = 4194304
                                            )
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
                        ,ISNULL(L.CCstPr, 0) * CASE
                            WHEN UPPER(ISNULL(L.ProdNo, '')) = 'R1100'
                             AND UPPER(ISNULL(LastEmployee.EmployeeName, '')) LIKE '%LASER EAGLE%'
                            THEN 2
                            ELSE 1
                         END AS ResourceUnitCost
                    FROM ProductionOrders
                    INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = ProductionOrders.OrdNo
                    LEFT JOIN R7 R WITH(NOLOCK) ON R.RNo = L.R7
                    OUTER APPLY (
                        SELECT SUM(CAST(P.NoInvoAb AS decimal(18, 6))) AS FinishedMinutes
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
                        SUM(FinishedMinutes * ResourceUnitCost) AS TimeCost
                    FROM ResourceMinutes
                    GROUP BY SalesOrderNo
                ),
                MaterialCosts AS (
                    SELECT
                        ProductionOrders.SalesOrderNo,
                        SUM(ISNULL(L.NoFin, 0) * ISNULL(L.CCstPr, 0)) AS MaterialCost
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
                                            SUM(ISNULL(L.NoFin, 0) * ISNULL(L.CCstPr, 0)) AS StangCost
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
                                    PurchasedPartCosts AS (
                                        SELECT
                                            ProductionOrders.SalesOrderNo,
                                            SUM(
                                                CASE WHEN ISNULL(L.NoFin, 0) <> 0 THEN ISNULL(L.NoFin, 0) ELSE ISNULL(L.NoOrg, 0) END
                                                * ISNULL(PurchasePrice.UnitPrice, ISNULL(L.CCstPr, 0))
                                            ) AS PurchasedPartCost
                                        FROM ProductionOrders
                                        INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = ProductionOrders.OrdNo
                                        LEFT JOIN Prod PP WITH(NOLOCK) ON PP.ProdNo = L.ProdNo
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
                                        GROUP BY ProductionOrders.SalesOrderNo
                                    ),
                NestingMaterialCosts AS (
                    SELECT
                        ProductionOrders.SalesOrderNo,
                        SUM(ISNULL(N.NoFin, 0) * ISNULL(N.CstPr, 0)) AS MaterialCost
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
                    ISNULL(PurchasedPartCosts.PurchasedPartCost, 0) AS PurchasedPartCost,
                    ISNULL(Active.TimeCost, 0) AS TimeCost,
                    (ISNULL(S.InvoSF, 0) + ISNULL(S.InvoIF, 0)) * (ISNULL(NULLIF(S.ExRt, 0), 100) / 100.0) AS SalesValue,
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
                    LEFT JOIN PurchasedPartCosts ON PurchasedPartCosts.SalesOrderNo = S.OrdNo
                LEFT JOIN NestingMaterialCosts ON NestingMaterialCosts.SalesOrderNo = S.OrdNo
                ORDER BY
                    CASE WHEN S.DelDt > 19800101 THEN S.DelDt ELSE 99991231 END,
                    S.OrdNo
            `);
    return result.recordset || [];
}

module.exports = { fetchSalgordreViaRows };
