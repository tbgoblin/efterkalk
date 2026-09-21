// Run real source fragments with in-memory dependencies. No database/file writes.
const fs=require('fs');
const assert=require('assert/strict');
async function main(){
 const routes=fs.readFileSync('routes/apiRoutes.js','utf8');
 const start=routes.indexOf("router.post('/lagerliste/snapshots',");
 const begin=routes.indexOf('async (req, res) => {',start);
 const end=routes.indexOf('\n    });',begin);
 const pending=[],logs=[]; let response;
 const handler=new Function('setImmediate','lagerlisteService','fs','logEvent','return '+routes.slice(begin,end)+'\n}')(
  fn=>pending.push(fn),{savePointInTimeSnapshot:async()=>{throw Error('Simulated write failure');}}, {},m=>logs.push(m));
 await handler({body:{}},{json:p=>{response=p;return p;},status:()=>{throw Error('Unexpected error response');}});
 assert.equal(response.ok,true);
 pending.forEach(fn=>fn());await new Promise(resolve=>setImmediate(resolve));
 assert.equal(logs.length,1);
 const source=fs.readFileSync('services/lagerlisteService.js','utf8');
 const s=source.indexOf('    function scheduleMonthlySnapshot(');
 const e=source.indexOf('\n    return {',s);
 let existing=null,refreshes=0,saves=0,tick;
 const RealDate=Date;
 class Clock extends RealDate{constructor(...args){super(...(args.length?args:['2026-09-30T08:00:00+02:00']));}}
 const schedule=new Function('Date','fs','loadMonthlySnapshot','getCurrent','saveMonthlySnapshot','setInterval',source.slice(s,e)+';return scheduleMonthlySnapshot;')(
  Clock,{},async()=>existing,async()=>{refreshes++;},async()=>{saves++;existing={saved:true};},fn=>{tick=fn;return 1;});
 schedule();await new Promise(resolve=>setImmediate(resolve));
 assert.equal(saves,1);await tick();assert.equal(saves,1);assert.equal(refreshes,1);
 console.log(JSON.stringify({manualSnapshot:{reportedSuccess:response.ok,actualSaveFailed:logs.length===1},monthlyClose:{time:'30 September 08:00 local',savedAlready:saves===1,laterRunSkipped:refreshes===1}},null,2));
}
main().catch(e=>{console.error(e);process.exitCode=1;});
