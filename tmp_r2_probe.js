const getConnection = require('./db');
(async () => {
  const pool = await getConnection();
  const result = await pool.request().query("SELECT TOP 40 ProdNo, Descr FROM Prod WHERE ProdGr = 3 AND ProdNo LIKE 'R2%' ORDER BY ProdNo");
  console.log(JSON.stringify(result.recordset));
})().catch(err => { console.error(err); process.exitCode = 1; });
