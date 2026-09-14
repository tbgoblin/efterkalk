// Read-only verification against current Visma state and the last cached report.
const fs = require('fs');
const getConnection = require('./db');
const { buildOrderStates } = require('./services/lagerliste2Service');
const { allocateSharedOrders, allocateComponentStock } = require('./services/lagerlisteAllocation');

async function main() {
    const payload = JSON.parse(fs.readFileSync('C:/GantechCache/lagerliste_v29_stang_v6.json', 'utf8')).data;
    const categories = payload.categories;
    const viaOrders = new Set(categories.salgordreVia.map(row => Number(row.OrdNo)));
    const orderNos = categories.finishedNotInvoiced.map(row => Number(row.OrdNo)).filter(n => viaOrders.has(n));
    const pool = await getConnection();
    try {
        const result = await pool.request().input('orders', orderNos.join(',')).query(`
            SELECT O.OrdNo, O.InvoNo, O.InvoAm, O.DInvoIF, O.FinDt, O.OrdPrSt,
                   L.LnNo, L.ProdNo, L.TrTp AS LineTrTp, L.NoOrg, L.NoFin,
                   L.NoPac, L.NoInvo, L.NoInvoAb, L.CCstPr
            FROM Ord O LEFT JOIN OrdLn L ON L.OrdNo = O.OrdNo
            WHERE O.OrdNo IN (SELECT TRY_CAST(value AS int) FROM STRING_SPLIT(@orders, ','));
            SELECT P.ProdNo, B.PoPhStB, B.Bal, B.StcInc, B.ShpRsv, B.PhCstPr
            FROM Prod P JOIN StcBal B ON B.ProdNo = P.ProdNo AND B.StcNo = 1
            WHERE P.Gr5 = 11 AND P.Gr9 = 1 AND P.ProdNo LIKE '1%'
              AND P.ProdNo NOT LIKE '%L%' AND P.ProdGr IN (1,2);
        `);
        const allocation = allocateSharedOrders(categories.finishedNotInvoiced, categories.salgordreVia, buildOrderStates(result.recordsets[0]));
        const stock = allocateComponentStock(categories.gr5Items, categories.opfolgningvare);
        const removedValue = stock.audit.reduce((sum, row) => sum + row.removedValue, 0);
        console.log(JSON.stringify({ cachedAt: payload.generatedAt, checkedAt: new Date().toISOString(),
            componentReduction: removedValue, stockAllocations: stock.audit,
            reservationEvidence: result.recordsets[1].filter(row => stock.audit.some(a => a.prodNo === row.ProdNo)),
            sharedOrders: allocation.audit.map(row => ({ ...row,
                finished: allocation.finished.find(r => r.OrdNo === row.ordNo).Value,
                via: allocation.via.find(r => r.OrdNo === row.ordNo).Value })) }, null, 2));
    } finally { await pool.close(); }
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
