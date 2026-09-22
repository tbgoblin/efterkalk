const test = require('node:test');
const assert = require('node:assert/strict');
const { defaults, validateDefinition, compileDefinition, evaluateRows, createDefinitionStore } = require('../services/bilancioDefinition');
const { LINES, buildReport } = require('../services/bilancioService');
const leaf = (id, accounts, sign = 1) => ({ id, name: id, type:'accounts', accounts, groups:[], exclude:[], children:false, sign });
function config() {
    return { schema:1,version:0,revenueRow:'sales',beforeTaxRow:'result',balanceTitle:'Aktiver',
        pnl:[leaf('sales',[100],-1),leaf('cost',[200],-1),{id:'result',name:'Result',type:'sum',sources:['sales','cost']}],
        balance:[leaf('deposits',[66980,66981]),leaf('receivables',[66100]),{id:'total',name:'Total',type:'sum',sources:['deposits','receivables']}] };
}
const catalog = { accounts:[100,200,66980,66981,66100].map(AcNo=>({AcNo,Nm:'Account '+AcNo,AcGr:AcNo>60000?'66':'pl'})),
    groups:[{AcGr:'pl',AgAcGr:''},{AcGr:'66',AgAcGr:'parent'},{AcGr:'parent',AgAcGr:''}] };
test('default configuration reproduces every existing income row and amount', () => {
    const definition = defaults(LINES);
    const accounts = [...new Set(LINES.flatMap(l=>l.codes||[]))].map((AcGr,i)=>({AcNo:i+1,AcGr,Nm:AcGr}));
    const records = accounts.map((a,i)=>({...a,Month:(i+1)*123.45,PriorMonth:(i+1)*100,Ytd:(i+1)*800,PriorYtd:(i+1)*500}));
    const balances = [66980,66981,61100,61120,61850,61900,66100].map(AcNo=>({AcNo,AcGr:'other',Nm:'asset'}));
    const compiled = compileDefinition(definition,{accounts:[...accounts,...balances],groups:[...accounts.map(a=>({AcGr:a.AcGr})),{AcGr:'46_Grunde_og_bygninger'},{AcGr:'47_Driftsmidler_ialt'}]});
    const actual = evaluateRows(compiled.pnl,records,4,['Month','PriorMonth','Ytd','PriorYtd']);
    const old = buildReport(records,2026,2).rows;
    assert.deepEqual(actual.map(r=>r.name),old.map(r=>r.name));
    actual.forEach((row,i)=>row.amounts.forEach((v,j)=>assert.ok(Math.abs(v-old[i].amounts[j])<0.000001)));
});
test('account sign, reference subtotal and configured percentage are evaluated exactly', () => {
    const c=config();c.pnl.push({id:'tax',name:'Tax',type:'percent',sources:['result'],rate:-22},{id:'net',name:'Net',type:'sum',sources:['result','tax']});
    const compiled=compileDefinition(c,catalog);
    const rows=evaluateRows(compiled.pnl,[{AcNo:100,Amount:-1000},{AcNo:200,Amount:400}],1,['Amount']);
    assert.deepEqual(rows.map(r=>r.amounts),[[1000],[-400],[600],[-132],[468]]);
});
test('nested groups deduplicate direct selections and exclusions remove deposits', () => {
    const c=config();c.balance[1]={...leaf('receivables',[66100]),groups:['parent','66'],children:true,exclude:[66980,66981]};
    assert.deepEqual(compileDefinition(c,catalog).balance[1].selectedAccounts,[66100]);
    c.balance[1].exclude=[];
    assert.throws(()=>compileDefinition(c,catalog),/medregnes i både/);
});
test('overlapping subtotal references cannot double count underlying rows', () => {
    const c=config();c.pnl.push({id:'bad',name:'Bad',type:'sum',sources:['result','sales']});
    assert.throws(()=>validateDefinition(c),/flere gange/);
});
test('print layout is optional, validated per orientation, and rejects rectangles that spill off the page', () => {
    const c=config();
    assert.equal(validateDefinition(c).printLayout,null);
    c.printLayout={portrait:{pnl:{x:0,y:0,width:100,height:48},balance:{x:0,y:52,width:100,height:48}},landscape:null};
    assert.deepEqual(validateDefinition(c).printLayout,{portrait:{pnl:{x:0,y:0,width:100,height:48},balance:{x:0,y:52,width:100,height:48}},landscape:null});
    const d=config();d.printLayout={portrait:{pnl:{x:60,y:0,width:60,height:48},balance:{x:0,y:52,width:100,height:48}},landscape:null};
    assert.throws(()=>validateDefinition(d),/udskriftslayout/i);
    const e=config();e.printLayout={portrait:{pnl:{x:0,y:0,width:2,height:48},balance:{x:0,y:52,width:100,height:48}},landscape:null};
    assert.throws(()=>validateDefinition(e),/udskriftslayout/i);
});
test('a Balance formula may reference a P&L row, but a P&L formula still cannot reach into Balance', () => {
    const c=config();c.balance.push({id:'check',name:'Check',type:'formula',formula:'[total] - [result]'});
    assert.doesNotThrow(()=>validateDefinition(c));
    const d=config();d.pnl.push({id:'bad',name:'Bad',type:'formula',formula:'[total]'});
    assert.throws(()=>validateDefinition(d),/ovenfor/);
});
test('a Balance formula reading a P&L row picks up its year-to-date columns as Current/Previous', () => {
    const c=config();c.balance.push({id:'check',name:'Check',type:'formula',formula:'[total] - [result]'});
    const compiled=compileDefinition(c,catalog);
    const pnlRows=evaluateRows(compiled.pnl,[{AcNo:100,Month:-10,PriorMonth:-9,Ytd:-1000,PriorYtd:-900},{AcNo:200,Month:4,PriorMonth:3,Ytd:400,PriorYtd:300}],4,['Month','PriorMonth','Ytd','PriorYtd']);
    // result = sales(Ytd 1000) + cost(Ytd -400) = 600 in both the Ytd and PriorYtd columns.
    const crossSection=new Map(pnlRows.map(r=>[r.id,[r.amounts[2],r.amounts[3]]]));
    const balanceRows=evaluateRows(compiled.balance,[{AcNo:66980,Current:100,Previous:90},{AcNo:66981,Current:200,Previous:190},{AcNo:66100,Current:900,Previous:800}],2,['Current','Previous'],new Map(),crossSection);
    // total = deposits(300/280) + receivables(900/800) = 1200/1080; check = total − result(600/600).
    assert.deepEqual(balanceRows.find(r=>r.id==='check').amounts,[600,480]);
});
test('forward and cyclic references, unknown accounts and groups fail', () => {
    const c=config();c.pnl[2].sources=['later'];assert.throws(()=>validateDefinition(c),/ovenfor/);
    const d=config();d.pnl[0].accounts=[999];assert.throws(()=>compileDefinition(d,catalog),/findes ikke/);
    const e=config();e.pnl[0].groups=['missing'];assert.throws(()=>compileDefinition(e,catalog),/findes ikke/);
});
test('labels may change without breaking revenue and before-tax references', () => {
    const c=config();c.pnl[0].name='Renamed sales';c.pnl[2].name='Renamed result';
    assert.equal(validateDefinition(c).beforeTaxRow,'result');
    c.beforeTaxRow='missing';assert.throws(()=>validateDefinition(c),/resultat før skat/);
});
test('headings carry no amount, missing or duplicate account values fail', () => {
    const c=config();c.pnl.unshift({id:'title',name:'Title',type:'heading'});
    const compiled=compileDefinition(c,catalog);
    const source=[{AcNo:100,Amount:-100},{AcNo:200,Amount:50}];
    assert.equal(evaluateRows(compiled.pnl,source,1,['Amount'])[0].type,'heading');
    assert.throws(()=>evaluateRows(compiled.pnl,[source[0]],1,['Amount']),/mangler/);
    assert.throws(()=>evaluateRows(compiled.pnl,[...source,source[0]],1,['Amount']),/Dubleret/);
});
test('GOH default loads without writes; saving uses optimistic version and author', async () => {
    let writes=0, saved;
    const gohData={getAppState:async(key,opts)=>{assert.equal(opts.strict,true);return null;},setAppState:async(key,payload,opts)=>{
        writes++;assert.equal(opts.expectedVersion,0);saved=payload;return true;
    }};
    const store=createDefinitionStore({gohData,getProfile:()=>({server:'s',database:'d'}),defaultDefinition:config()});
    const c=await store.load();assert.equal(writes,0);
    const result=await store.save(c,catalog,'admin');
    assert.equal(result.version,1);assert.equal(saved.updatedBy,'admin');assert.equal(writes,1);
});
test('conflicting write and changed database are not reported as success', async () => {
    let database='d', writes=0;
    const store=createDefinitionStore({gohData:{getAppState:async()=>null,setAppState:async()=>{writes++;return false;}},getProfile:()=>({server:'s',database}),defaultDefinition:config()});
    const c=await store.load();await assert.rejects(store.save(c,catalog,'admin'),e=>e.statusCode===409);
    database='other';await assert.rejects(store.save(c,catalog,'admin'),/Databaseprofilen/);assert.equal(writes,1);
});
test('unavailable GOH never silently substitutes the default configuration', async () => {
    const store=createDefinitionStore({gohData:{getAppState:async()=>{throw Error('offline');}},getProfile:()=>({server:'s',database:'d'}),defaultDefinition:config()});
    await assert.rejects(store.load(),/offline/);
});
