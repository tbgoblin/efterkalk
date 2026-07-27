const getConnection = require('./db.js');

async function run() {
    const pool = await getConnection();
    // controlla DPrice vs CCstPr nell'ordine allocato collegato 406548
    const res = await pool.request().query(
        "SELECT O.TrTp, L.LnNo, L.ProdNo, L.PurcNo, L.DPrice, L.CCstPr, L.NoFin, L.NoOrg, " +
        "CAST(NoFin * DPrice AS DECIMAL(10,2)) AS CostByDPrice, " +
        "CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS CostByCCstPr " +
        "FROM OrdLn L INNER JOIN Ord O ON O.OrdNo = L.OrdNo WHERE L.OrdNo = 406548 ORDER BY L.LnNo"
    );
    console.log("=== Linked order 406548 - DPrice vs CCstPr ===");
    console.table(res.recordset.map(r => ({
        TrTp: r.TrTp,
        LnNo: r.LnNo,
        ProdNo: r.ProdNo,
        PurcNo: r.PurcNo,
        DPrice: r.DPrice,
        CCstPr: r.CCstPr,
        NoFin: r.NoFin,
        NoOrg: r.NoOrg,
        CostByDPrice: r.CostByDPrice,
        CostByCCstPr: r.CostByCCstPr
    })));
    process.exit(0);
}
run().catch(e => { console.error(e); process.exit(1); });
