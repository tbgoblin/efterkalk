// Read-only audit: uses the same VIA query and stock transfer function as Lagerliste.
const fs = require('fs');
const getConnection = require('./db');
const sql = require('mssql/msnodesqlv8');
const { fetchSalgordreViaRows } = require('./services/viaService');
const { allocatePurchasedPartsFromFollowUpStock } = require('./services/lagerlisteAllocation');
async function main() {
  const pool = await getConnection();
  try {
    const via = await fetchSalgordreViaRows({ getConnection: async () => pool, sql });
    const result = await pool.request().query(`SELECT P.ProdNo, P.Gr9, P.ProdGr,
      B.PoPhStB, B.Bal, B.StcInc, B.ShpRsv, B.PhCstPr,
      ISNULL(B.Bal,0)+ISNULL(B.StcInc,0)-ISNULL(B.ShpRsv,0) AS Beholdning
      FROM Prod P JOIN StcBal B ON B.ProdNo=P.ProdNo AND B.StcNo=1
      WHERE P.Gr9=1 AND P.ProdNo LIKE '1%' AND P.ProdNo NOT LIKE '%L%'
      AND P.ProdGr IN (1,2) AND ISNULL(B.PoPhStB,0)<>0`);
    const allocation = allocatePurchasedPartsFromFollowUpStock(result.recordset, via);
    const products = new Map();
    const purchases = new Map();
    for (const r of via) for (const d of r.PurchasedPartDetails || []) {
      const key=String(d.prodNo).trim();
      if (d.stockTransferQty>0) {
        const p=products.get(key)||{prodNo:key,qty:0,value:0,orders:[]};
        p.qty+=d.stockTransferQty;p.value+=d.stockTransferValue;
        p.orders.push({order:r.OrdNo,qty:d.stockTransferQty,reserved:d.activeReservedQty,physical:d.physicalStockQty});products.set(key,p);
      }
      const pk=d.purchaseOrderNo+'|'+key;
      const list=purchases.get(pk)||[];list.push({order:r.OrdNo,...d});purchases.set(pk,list);
    }
    const audit=products.values();
    const transfers=[...audit].map(p=>{const a=allocation.audit.find(a=>String(a.prodNo).trim()===p.prodNo);return {...p,removedQty:a?.transferredQty||0,removedValue:a?.transferredValue||0};});
    const duplicates=via.filter((r,i)=>via.findIndex(x=>x.OrdNo===r.OrdNo)!==i).map(r=>r.OrdNo);
    const report={checkedAt:new Date().toISOString(),viaOrders:via.length,duplicateOrders:duplicates,
      transferProducts:transfers.length,transferValue:transfers.reduce((s,p)=>s+p.value,0),removedValue:allocation.audit.reduce((s,a)=>s+a.transferredValue,0),
      transferMismatches:transfers.filter(p=>Math.abs(p.qty-p.removedQty)>0.001||Math.abs(p.value-p.removedValue)>0.02),
      sharedStock:transfers.filter(p=>new Set(p.orders.map(o=>o.order)).size>1),
      sharedPurchases:[...purchases].filter(([k,v])=>new Set(v.map(d=>d.order)).size>1)};
    fs.writeFileSync('tmp-lagerliste-reservation-review.json',JSON.stringify(report,null,2));
    console.log(JSON.stringify(report,null,2));
  } finally {await pool.close();}
}
main().catch(e=>{console.error(e.message);process.exitCode=1;});
