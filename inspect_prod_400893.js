const getConnection = require("./db.js");
const sql = require("mssql/msnodesqlv8");

const PROD_ORDNO = 400893;

function fmt(n) {
    if (n === null || n === undefined) return "-";
    const v = Number(n);
    if (!Number.isFinite(v)) return String(n);
    return v.toFixed(2);
}

async function run() {
    const pool = await getConnection();
    const prodLinesRes = await pool.request().query(`SELECT LnNo, ProdNo, Descr, ProdTp4, DPrice, NoOrg, NoFin, CCstPr FROM OrdLn WHERE OrdNo = ${PROD_ORDNO} ORDER BY LnNo`);
    const lines = prodLinesRes.recordset.map(l => ({
        LnNo: l.LnNo,
        ProdNo: l.ProdNo,
        Descr: (l.Descr || "").substring(0, 30),
        ProdTp4: l.ProdTp4,
        NoOrg: l.NoOrg,
        NoFin: l.NoFin,
        DPrice: fmt(l.DPrice),
        CCstPr: fmt(l.CCstPr),
        LineCost: fmt((l.NoFin || 0) * (l.CCstPr || 0))
    }));
    console.log(`=== OrdLn for production order ${PROD_ORDNO} ===`);
    console.table(lines);
    const totalCost = lines.reduce((sum, l) => sum + ((l.NoFin || 0) * (l.CCstPr || 0)), 0);
    console.log("Total cost:", fmt(totalCost));
    process.exit(0);
}

run().catch(err => {
    console.error(err);
    process.exit(1);
});
