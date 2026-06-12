const getConnection = require('./db');
(async () => {
  const pool = await getConnection();
  const result = await pool.request().query("SELECT TOP 40 ProdNo, Descr FROM Prod WHERE ProdGr = 3 AND (Descr LIKE '%Svejs%' OR Descr LIKE '%Flad%' OR Descr LIKE '%Buk%' OR Descr LIKE '%Laser%') ORDER BY ProdNo");
  console.log(JSON.stringify(result.recordset));
})().catch(err => { console.error(err); process.exitCode = 1; });
