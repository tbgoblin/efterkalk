// Read-only: find old nesting routes excluded by Lagerliste1's rolling date window.
const getConnection = require('./db');
async function main() {
 const pool=await getConnection();
 try {
  const now=new Date(); const start=new Date(now);start.setDate(1);start.setMonth(start.getMonth()-2);
  const cutoff=Number(start.toISOString().slice(0,7).replace('-','')+'01');
  const result=await pool.request().input('cutoff',cutoff).query(`
   SELECT O.OrdNo,O.OrdDt,L.TrInf4 AS Route,
    SUM(CASE WHEN L.TrTp=5 AND ISNULL(L.NoOrg,0)>0 THEN 1 ELSE 0 END) AS PositivePlates,
    SUM(CASE WHEN L.TrTp=5 AND ISNULL(L.NoOrg,0)>0 AND ISNULL(L.NoFin,0)<=0 THEN 1 ELSE 0 END) AS UnfinishedPlates,
    SUM(CASE WHEN L.TrTp=7 THEN 1 ELSE 0 END) AS Products,
    SUM(CASE WHEN L.TrTp=7 AND ISNULL(L.NoFin,0)<>0 THEN 1 ELSE 0 END) AS FinishedProducts,
    SUM(CASE WHEN L.TrTp=5 AND ISNULL(L.NoOrg,0)>=0 THEN
      CASE WHEN ISNULL(L.CstPr,0)<>0 THEN CONVERT(float,L.CstPr)*ISNULL(L.NoFin,0)
      WHEN ISNULL(L.NoOrg,0)<>0 THEN CONVERT(float,ISNULL(L.IncCst,0))*ISNULL(L.NoFin,0)/L.NoOrg
      ELSE ISNULL(L.IncCst,0) END ELSE 0 END) AS CountedValue
   FROM Ord O JOIN OrdLn L ON L.OrdNo=O.OrdNo
   WHERE TRY_CONVERT(int,O.OrdDt)<@cutoff AND TRY_CONVERT(decimal(18,6),O.Gr3)=2
    AND L.TrTp IN(5,7) AND NULLIF(LTRIM(RTRIM(CONVERT(varchar(100),L.TrInf4))),'') IS NOT NULL
    AND O.OrdNo NOT IN(61423,75330,131790,140134,331368)
   GROUP BY O.OrdNo,O.OrdDt,L.TrInf4
   HAVING SUM(CASE WHEN L.TrTp=5 AND ISNULL(L.NoOrg,0)>0 THEN 1 ELSE 0 END)>0
    AND SUM(CASE WHEN L.TrTp=5 AND ISNULL(L.NoOrg,0)>0 AND ISNULL(L.NoFin,0)<=0 THEN 1 ELSE 0 END)=0
    AND SUM(CASE WHEN L.TrTp=7 THEN 1 ELSE 0 END)>0
    AND SUM(CASE WHEN L.TrTp=7 AND ISNULL(L.NoFin,0)<>0 THEN 1 ELSE 0 END)=0
   ORDER BY O.OrdDt DESC,O.OrdNo`);
  console.log(JSON.stringify({checkedAt:now.toISOString(),cutoff,excludedRoutes:result.recordset},null,2));
 }finally{await pool.close();}
}
main().catch(e=>{console.error(e.message);process.exitCode=1;});
