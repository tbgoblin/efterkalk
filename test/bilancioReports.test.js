const test = require('node:test');
const assert = require('node:assert/strict');
const { createDefinitionStore, validateDefinition, evaluateRows } = require('../services/bilancioDefinition');
function fixture() {
    const states=new Map();let database='db';
    const base={schema:1,version:0,balanceTitle:'Aktiver',revenueRow:'base',beforeTaxRow:'base',pnl:[{id:'base',name:'Base',type:'formula',formula:'0'}],balance:[{id:'heading',name:'Aktiver',type:'heading'}]};
    const gohData={getAppState:async key=>states.has(key)?{payload:structuredClone(states.get(key))}:null,
        getAppStateKeysByPrefix:async prefix=>[...states.keys()].filter(k=>k.startsWith(prefix)).map(key=>({key})),
        setAppState:async(key,payload,{expectedVersion})=>{if((states.get(key)?.version||0)!==expectedVersion)return false;states.set(key,structuredClone(payload));return true;}};
    return {store:createDefinitionStore({gohData,getProfile:()=>({server:'s',database}),defaultDefinition:base}),states,switchDb:()=>{database='other';}};
}
const catalog={accounts:[],groups:[]};
test('new and copied reports have isolated storage and versions; existing report remains intact',async()=>{
    const {store,states}=fixture();const original=await store.load();assert.equal(states.size,0);
    const first=await store.save({...original,reportName:'Original'},catalog,'admin');
    const copy=await store.save({...first,reportId:'copy',reportName:'Copy',version:0,showPnl:false},catalog,'admin');
    await store.save({...copy,balanceTitle:'Passiver'},catalog,'admin');
    assert.equal((await store.load()).balanceTitle,'Aktiver');assert.equal((await store.load()).version,1);
    assert.equal((await store.load('copy')).balanceTitle,'Passiver');assert.equal((await store.load('copy')).version,2);
    assert.deepEqual(await store.list(),[{reportId:'default',reportName:'Original'},{reportId:'copy',reportName:'Copy'}]);
    assert.ok([...states.keys()].every(k=>k.length<=100));
    await assert.rejects(store.save(copy,catalog,'admin'),e=>e.statusCode===409);
});
test('report ids cannot escape scope or collide through key truncation; unknown reports do not load the default',async()=>{
    const {store,switchDb}=fixture();const original=await store.load();
    for(const id of ['../default','x'.repeat(41),['default']])await assert.rejects(store.load(id),/rapport-id/);
    await assert.rejects(store.load('missing'),e=>e.statusCode===404);
    switchDb();await assert.rejects(store.save({...original,reportId:'copy'},catalog,'admin'),/Databaseprofilen/);
});
test('Aktiver and negative Passiver sum to zero, while a discrepancy and missing data remain visible',()=>{
    const rows=validateDefinition({schema:1,version:0,reportName:'Balance',showPnl:false,legacyPeriodRows:false,revenueRow:'base',beforeTaxRow:'base',pnl:[{id:'base',name:'base',type:'formula',formula:'0'}],balanceTitle:'Balance',balance:[
        {id:'aktiver',name:'Aktiver',type:'manual',period:'selected',monthValues:{}},
        {id:'passiver',name:'Passiver',type:'manual',period:'selected',monthValues:{}},
        {id:'check',name:'Kontrol',type:'formula',formula:'[aktiver] + [passiver]',checkZero:true,bold:true}
    ]}).balance;
    const result=evaluateRows(rows,[],2,[],new Map([['aktiver',[100,100]],['passiver',[-100,-90]]]));
    assert.deepEqual(result[2].amounts,[0,10]);assert.equal(result[2].checkZero,true);
    const missing=evaluateRows(rows,[],2,[],new Map([['aktiver',[100,100]],['passiver',[null,-100]]]));
    assert.deepEqual(missing[2].amounts,[null,0]);
});
