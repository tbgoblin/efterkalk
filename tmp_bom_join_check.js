const getConnection = require('./db');
(async () => {
  const pool = await getConnection();
  const q1 = await pool.request().query("SELECT COUNT(*) AS n FROM Prod p JOIN Actor a ON CAST(a.CustNo AS varchar(50)) = p.Inf3 WHERE p.ProdGr <> 99999 AND a.CustNo <> 0");
  const q2 = await pool.request().query("SELECT COUNT(*) AS n FROM Prod p JOIN Actor a ON CAST(a.Gr AS varchar(50)) = p.Inf3 WHERE p.ProdGr <> 99999 AND a.CustNo <> 0");
  const q3 = await pool.request().query("SELECT TOP 20 p.Inf3, COUNT(*) AS n FROM Prod p WHERE p.ProdGr <> 99999 AND ISNULL(p.Inf3,'') <> '' GROUP BY p.Inf3 ORDER BY n DESC");
  console.log('joinCustNo=' + q1.recordset[0].n);
  console.log('joinGr=' + q2.recordset[0].n);
  console.log(JSON.stringify(q3.recordset));
  process.exit(0);
})().catch(err => { console.error(err); process.exit(1); });
