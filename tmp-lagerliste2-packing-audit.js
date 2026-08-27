const getConnection = require('./db');

async function main() {
    const pool = await getConnection();
    const result = await pool.request().query(`
        SELECT O.OrdNo, O.OrdDt, O.ChDt, O.ChTm, O.DelDt, O.FinDt, O.OrdPrSt,
               O.InvoNo, O.InvoAm, O.InvoIF, O.DInvoIF, O.InvoSF, O.Gr12,
               L.LnNo, L.ProdNo, L.Descr, L.ProdTp4, L.TrTp, L.PurcNo,
               L.NoOrg, L.NoFin, L.NoPac, L.NoInvo, L.NoInvoAb,
               L.Price, L.DPrice, L.CCstPr, L.CstPr
        FROM Ord O WITH(NOLOCK)
        INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = O.OrdNo
        WHERE O.OrdNo IN (399912, 411492, 403745, 402154, 401645, 410972, 411494, 411611, 411616)
        ORDER BY O.OrdNo, L.LnNo;

        SELECT TOP 15 O.OrdNo, O.OrdDt, O.ChDt, O.ChTm, O.FinDt, O.OrdPrSt,
               O.InvoNo, O.InvoAm, O.InvoIF, O.DInvoIF,
               SUM(COALESCE(TRY_CONVERT(decimal(18,6), L.NoOrg),0)) AS SumNoOrg,
               SUM(COALESCE(TRY_CONVERT(decimal(18,6), L.NoFin),0)) AS SumNoFin,
               SUM(COALESCE(TRY_CONVERT(decimal(18,6), L.NoPac),0)) AS SumNoPac,
               SUM(COALESCE(TRY_CONVERT(decimal(18,6), L.NoInvo),0)) AS SumNoInvo,
               COUNT(*) AS Lines
        FROM Ord O WITH(NOLOCK)
        INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = O.OrdNo
        WHERE O.OrdTp = 1 AND O.TrTp = 1
          AND COALESCE(TRY_CONVERT(decimal(18,6), O.InvoAm),0) <> 0
          AND COALESCE(TRY_CONVERT(decimal(18,6), O.DInvoIF),0) <> 0
        GROUP BY O.OrdNo, O.OrdDt, O.ChDt, O.ChTm, O.FinDt, O.OrdPrSt,
                 O.InvoNo, O.InvoAm, O.InvoIF, O.DInvoIF
        ORDER BY TRY_CONVERT(int, O.ChDt) DESC, TRY_CONVERT(int, O.ChTm) DESC;

        WITH LineState AS (
            SELECT O.OrdNo, L.LnNo, L.ProdNo, L.ProdTp4, L.TrTp,
                   COALESCE(TRY_CONVERT(decimal(18,6), L.NoOrg),0) AS NoOrg,
                   COALESCE(TRY_CONVERT(decimal(18,6), L.NoFin),0) AS NoFin,
                   COALESCE(TRY_CONVERT(decimal(18,6), L.NoPac),0) AS NoPac,
                   COALESCE(TRY_CONVERT(decimal(18,6), L.NoInvo),0) AS NoInvo
            FROM Ord O WITH(NOLOCK)
            INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = O.OrdNo
            WHERE O.OrdTp = 1 AND O.TrTp = 1
              AND NULLIF(LTRIM(RTRIM(CONVERT(varchar(100), L.ProdNo))), '') IS NOT NULL
        )
        SELECT
            SUM(CASE WHEN NoPac < 0 OR NoFin < 0 OR NoInvo < 0 THEN 1 ELSE 0 END) AS NegativeRows,
            SUM(CASE WHEN NoPac > NoFin THEN 1 ELSE 0 END) AS PackedAboveFinished,
            SUM(CASE WHEN NoInvo > NoPac THEN 1 ELSE 0 END) AS InvoicedAbovePacked,
            SUM(CASE WHEN NoFin > NoOrg AND NoOrg > 0 THEN 1 ELSE 0 END) AS FinishedAboveOrdered,
            SUM(CASE WHEN NoPac > 0 AND NoFin = 0 THEN 1 ELSE 0 END) AS PackedWithoutFinished,
            SUM(CASE WHEN NoInvo > 0 AND NoPac = 0 THEN 1 ELSE 0 END) AS InvoicedWithoutPacked,
            COUNT(*) AS TotalRows
        FROM LineState;
    `);
    console.log(JSON.stringify({ examples: result.recordsets[0], partialOrders: result.recordsets[1], consistency: result.recordsets[2] }, null, 2));
    await pool.close();
}

main().catch(error => {
    console.error(error && error.stack || error);
    process.exitCode = 1;
});
