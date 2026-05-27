const getConnection = require('./db.js');

async function run() {
    const pool = await getConnection();
    // controlla DPrice vs CCstPr nell'ordine acquisto 400906
    const res = await pool.request().query(
        "SELECT LnNo, ProdNo, DPrice, CCstPr, NoFin, " +
        "CAST(NoFin * DPrice AS DECIMAL(10,2)) AS CostByDPrice, " +
        "CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS CostByCCstPr " +
        "FROM OrdLn WHERE OrdNo = 400906 ORDER BY LnNo"
    );
    console.log("=== Purchase order 400906 - DPrice vs CCstPr ===");
    console.table(res.recordset.map(r => ({
        LnNo: r.LnNo,
        ProdNo: r.ProdNo,
        DPrice: r.DPrice,
        CCstPr: r.CCstPr,
        NoFin: r.NoFin,
        CostByDPrice: r.CostByDPrice,
        CostByCCstPr: r.CostByCCstPr
    })));
    process.exit(0);
}
run().catch(e => { console.error(e); process.exit(1); });
