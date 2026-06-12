const getConnection = require('./db');
(async () => {
  const pool = await getConnection();
  const result = await pool.request().query("SELECT Actor.Gr AS [Varenr.], Actor.CustNo AS [Kundernr.], Actor.Nm AS [Kundenavn], Actor.Ad1 AS [Adresse1], Actor.Ad2 AS [Adresse2], Actor.PNo AS [Postnr.], Actor.PArea AS [By], Actor.Phone AS [Tlf.], Actor.CustPrGr AS [Prisliste], Actor.Inf6 AS [Kunde], Actor.Shrt AS [Kortnavn] FROM Actor WHERE Actor.CustNo <> 0");
  console.log('rows=' + result.recordset.length);
  process.exit(0);
})().catch(err => {
  console.error(err);
  process.exit(1);
});
