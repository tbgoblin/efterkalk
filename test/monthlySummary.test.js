const test=require('node:test');
const assert=require('node:assert/strict');
const model=require('../assets/js/monthly-summary-model');
const {compile,createMonthlySummaryService}=require('../services/monthlySummaryService');
const catalog={accounts:[
    {AcNo:12070,Nm:'Varekøb',AcGr:'12_Vareforbrug'},
    {AcNo:12085,Nm:'Internt køb',AcGr:'12_Internal'},
    {AcNo:12092,Nm:'Salg jern/skrot',AcGr:'12_Vareforbrug'}
],groups:[{AcGr:'12_Vareforbrug',AgAcGr:''},{AcGr:'12_Internal',AgAcGr:'12_Vareforbrug'}]};

test('monthly account rows use Økonomirapport groups, descendants, exclusions and per-row deduplication',()=>{
    const config=model.defaults();
    const booked=config.rows.find(row=>row.id==='booked');
    Object.assign(booked,{accounts:[12070],groups:['12_Vareforbrug'],children:true,exclude:[12092]});
    config.rows.push({...booked,id:'comparison',name:'Samme konti til sammenligning'});
    const result=compile(config,catalog);
    assert.deepEqual(result.selections.booked,[12070,12085]);
    assert.deepEqual(result.selections.comparison,[12070,12085]);
    booked.groups=['Missing'];
    assert.throws(()=>compile(config,catalog),/findes ikke/);
});

test('default formulas reconcile material, inventory and purchase gap; missing figures remain missing',()=>{
    const {definition,selections}=compile(model.defaults(),catalog);
    const context={month:'2026-10',metrics:{orders:287,invoiced:5081521.32,cost:3367998.58,dg:33.7,
        material:1980062.19,warehouse_delta:416367.11,purchase_gap:576084.06},
        selections,accounts:[{AcNo:12070,Nm:'Varekøb',Amount:1628020.74}]};
    const result=new Map(model.evaluate(definition,context).map(row=>[row.id,row.value]));
    assert.equal(Math.round(result.get('difference')*100),-35204145);
    assert.equal(Math.round(result.get('residual')*100),-19232450);
    context.metrics.warehouse_delta=null;
    assert.equal(model.evaluate(definition,context).find(row=>row.id==='residual').value,null);
    context.accounts=[];
    assert.equal(model.evaluate(definition,context).find(row=>row.id==='booked').value,null);
});

test('credit signs, manual zero, hidden dependencies and formulas retain exact amounts',()=>{
    const config={schema:1,version:0,title:'Test',rows:[
        {id:'sales',name:'Faktureret',type:'source',source:'invoiced',display:'hidden'},
        {id:'adjustment',name:'Tillæg',type:'manual',monthValues:{'2026-10':0}},
        {id:'net',name:'Netto',type:'formula',formula:'[sales] + [adjustment]'}
    ]};
    const valid=model.validate(config);
    const result=model.evaluate(valid,{month:'2026-10',metrics:{invoiced:-13.25}});
    assert.equal(result[2].value,-13.25);
    assert.equal(model.evaluate(valid,{month:'2026-09',metrics:{invoiced:100}})[2].value,null);
    config.rows[2].formula='[missing]';
    assert.throws(()=>model.validate(config),/ovenfor/);
    config.rows[2].formula='alert(1)';
    assert.throws(()=>model.validate(config),/Formel/);
});

function serviceFixture() {
    let profile={server:'Visma',database:'F0001'},state=null,writes=0,parameters={},catalogHook;
    const gohData={getAppState:async()=>state,setAppState:async(key,payload,{expectedVersion})=>{
        if((state?.payload.version || 0)!==expectedVersion)return false;
        writes++;state={payload:structuredClone(payload)};return true;
    }};
    const service=createMonthlySummaryService({gohData,getProfile:()=>profile,
        getCatalog:async()=>{if(catalogHook)await catalogHook();return catalog;},
        sql:{Int:'Int',NVarChar:()=>'',MAX:0},
        getConnection:async()=>({request:()=>({input(name,type,value){parameters[name]=value;return this;},
            async query(){return{recordset:[{AcNo:12070,Nm:'Varekøb',Amount:0}]};}})})
    });
    return{service,writes:()=>writes,parameters:()=>parameters,setProfile:value=>{profile=value;},catalogHook:fn=>{catalogHook=fn;}};
}

test('configuration is versioned, shared per database and never saved by preview or monthly reads',async()=>{
    const f=serviceFixture(),config=await f.service.load();
    const preview=await f.service.data('2026-10',config);
    assert.equal(preview.accounts[0].Amount,0);
    assert.equal(f.parameters().year,2026);
    assert.equal(f.parameters().period,4);
    assert.equal(f.writes(),0);
    const saved=await f.service.save(config,'admin');
    assert.equal(saved.version,1);
    assert.equal(f.writes(),1);
    await assert.rejects(()=>f.service.save(config,'admin'),{statusCode:409});
    f.setProfile({server:'Visma',database:'F0001_TEST'});
    await assert.rejects(()=>f.service.save(saved,'admin'),{statusCode:409});
    assert.equal(f.writes(),1);
});

test('profile switches during catalog reads and invalid months stop before saving or using mixed data',async()=>{
    const f=serviceFixture(),config=await f.service.load();
    await assert.rejects(()=>f.service.data('2026-13'),{statusCode:400});
    f.catalogHook(()=>f.setProfile({server:'Other',database:'Other'}));
    await assert.rejects(()=>f.service.save(config,'admin'),{statusCode:409});
    assert.equal(f.writes(),0);
});
