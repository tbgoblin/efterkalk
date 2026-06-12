const getConnection = require('./db');
(async () => {
  const pool = await getConnection();
  const laser = await pool.request().query("SELECT TOP 5 FreeInf2.ProdNo, Prod.Descr, Prod.HgtU AS Tykkelse, FreeInf2.Txt1 AS Maskine, FreeInf2.Val1 AS Skaerehast, FreeInf2.Val2 AS Pircing, FreeInf2.Val3 AS Tillaeg, FreeInf2.Txt2 AS Linse FROM FreeInf2, Prod WHERE FreeInf2.ProdNo = Prod.ProdNo AND FreeInf2.FrInfTp = 100");
  const process = await pool.request().query("SELECT TOP 10 FrInfTp, Gr4, Gr5, Val1, Val2, Val3 FROM FreeInf1 WHERE FrInfTp IN (61,62) ORDER BY FrInfTp, Gr4, Gr5");
  const r21 = await pool.request().query("SELECT TOP 10 Prod.ProdNo, Prod.Descr, PrDcMat.CstPr, PrDcMat.SalePr, Prod.ProdGr FROM PrDcMat, Prod WHERE PrDcMat.ProdNo = Prod.ProdNo AND Prod.ProdGr = 3 AND Prod.ProdNo Like 'R21%'");
  console.log('laser=' + laser.recordset.length);
  console.log(JSON.stringify(laser.recordset));
  console.log('process=' + process.recordset.length);
  console.log(JSON.stringify(process.recordset));
  console.log('r21=' + r21.recordset.length);
  console.log(JSON.stringify(r21.recordset));
  process.exit(0);
})().catch(err => { console.error(err); process.exit(1); });
