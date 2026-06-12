const getConnection = require('./db');
(async () => {
  const pool = await getConnection();
  const queries = [
    { label: 'custno', value: '20137220' },
    { label: 'gr', value: '11007' }
  ];
  for (const item of queries) {
    const result = await pool.request().input('value', item.value).query("SELECT COUNT(*) AS n FROM Prod WHERE Inf3 = @value AND ProdGr <> 99999");
    console.log(item.label + '=' + result.recordset[0].n);
  }
  process.exit(0);
})().catch(err => { console.error(err); process.exit(1); });
