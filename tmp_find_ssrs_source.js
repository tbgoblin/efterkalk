const getConnection = require('./db');
(async () => {
  const pool = await getConnection();
  const queries = [
    `SELECT TOP 200 o.type_desc, o.name
     FROM sys.objects o
     WHERE o.type IN ('V','P','FN','IF','TF')
       AND (o.name LIKE '%salg%' OR o.name LIKE '%ordre%' OR o.name LIKE '%tilbud%' OR o.name LIKE '%stat%')
     ORDER BY o.type_desc, o.name`,

    `SELECT TOP 120 o.type_desc, o.name
     FROM sys.sql_modules m
     JOIN sys.objects o ON o.object_id = m.object_id
     WHERE LOWER(m.definition) LIKE '%ordtp%'
       AND LOWER(m.definition) LIKE '%trtp%'
       AND (LOWER(m.definition) LIKE '%iso_week%' OR LOWER(m.definition) LIKE '%datepart(week%')
     ORDER BY o.type_desc, o.name`,

    `SELECT TOP 120 o.type_desc, o.name
     FROM sys.sql_modules m
     JOIN sys.objects o ON o.object_id = m.object_id
     WHERE LOWER(m.definition) LIKE '%totalord%'
        OR LOWER(m.definition) LIKE '%totaltilbud%'
        OR LOWER(m.definition) LIKE '%gennem%'
        OR LOWER(m.definition) LIKE '%tilbud til ordre%'
     ORDER BY o.type_desc, o.name`
  ];

  for (let i = 0; i < queries.length; i++) {
    console.log('\n=== QUERY ' + (i + 1) + ' ===');
    const r = await pool.request().query(queries[i]);
    console.table(r.recordset);
  }

  process.exit(0);
})().catch(err => { console.error(err); process.exit(1); });
