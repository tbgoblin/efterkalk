const getConnection = require('./db');
(async () => {
  const pool = await getConnection();
  const result = await pool.request().input('value', '11007').query("SELECT COUNT(*) AS n FROM Prod WHERE Inf3 = @value AND ProdGr <> 99999");
  console.log('productsFor11007=' + result.recordset[0].n);
  const rev = await pool.request().input('tgn', 'DUMMY').input('value', '11007').query("SELECT COUNT(*) AS n FROM Prod WHERE Inf3 = @value");
  console.log('inf3UsesGr works=' + rev.recordset[0].n);
  process.exit(0);
})().catch(err => { console.error(err); process.exit(1); });
