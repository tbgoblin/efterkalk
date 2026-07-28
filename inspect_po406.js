const getConnection = require('./db.js');

async function run() {
    const pool = await getConnection();
    const ordNo = 406548;

    const headerRes = await pool.request().query(
        'SELECT OrdNo, TrTp, InvoAm, DInvoIF FROM Ord WHERE OrdNo = ' + ordNo
    );

    const linesRes = await pool.request().query(
        'SELECT LnNo, ProdNo, ProdTp4, PurcNo, DPrice, CstPr, CCstPr, NoFin, NoOrg, ' +
        'CAST(NoFin * ISNULL(CstPr, CCstPr) AS DECIMAL(10,2)) AS CostByCstPr, ' +
        'CAST(NoFin * DPrice AS DECIMAL(10,2)) AS CostByDPrice, ' +
        'CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS CostByCCstPr ' +
        'FROM OrdLn WHERE OrdNo = ' + ordNo + ' ORDER BY LnNo'
    );

    console.log('=== Ord header ===');
    console.table(headerRes.recordset);
    console.log('=== Ord lines ===');
    console.table(linesRes.recordset.map(r => ({
        LnNo: r.LnNo,
        ProdNo: r.ProdNo,
        ProdTp4: r.ProdTp4,
        PurcNo: r.PurcNo,
        DPrice: r.DPrice,
        CstPr: r.CstPr,
        CCstPr: r.CCstPr,
        NoFin: r.NoFin,
        NoOrg: r.NoOrg,
        CostByCstPr: r.CostByCstPr,
        CostByDPrice: r.CostByDPrice,
        CostByCCstPr: r.CostByCCstPr
    })));

    process.exit(0);
}

run().catch(err => {
    console.error(err);
    process.exit(1);
});
