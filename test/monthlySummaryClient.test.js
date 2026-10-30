const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const model=require('../assets/js/monthly-summary-model');
const {createMonthlySummaryService}=require('../services/monthlySummaryService');

function backend() {
    let state=null,writes=0;
    const catalog={accounts:[{AcNo:12070,Nm:'Køb',AcGr:'12_Vareforbrug'},{AcNo:12085,Nm:'Internt',AcGr:'12_Vareforbrug'},{AcNo:12092,Nm:'Skrot',AcGr:'12_Vareforbrug'}],groups:[{AcGr:'12_Vareforbrug',AgAcGr:''}]};
    const amounts={12070:1500,12085:125,12092:-200};
    const service=createMonthlySummaryService({getProfile:()=>({server:'test',database:'test'}),getCatalog:async()=>catalog,
        gohData:{getAppState:async()=>state,setAppState:async(key,payload,{expectedVersion})=>{if((state?.payload.version || 0)!==expectedVersion)return false;state={payload};writes++;return true;}},
        sql:{Int:0,MAX:0,NVarChar:()=>0},getConnection:async()=>({request(){let ids=[];return{input(name,type,value){if(name==='accounts')ids=value.split(',').map(Number);return this;},async query(){return{recordset:catalog.accounts.filter(row=>ids.includes(row.AcNo)).map(row=>({...row,Amount:amounts[row.AcNo]}))};}};}})
    });
    return {service,writes:()=>writes,fetch:async(url,options={})=>{
        try {
            const body=options.body ? JSON.parse(options.body) : {};
            let payload;
            if(url.includes('/preview'))payload=await service.data(body.month,body.definition);
            else if(url.includes('/summary-definition') && options.method==='PUT')payload={definition:await service.save(body,'admin')};
            else if(url.includes('/summary-definition'))payload=await service.editor();
            else payload={...await service.data(new URL(url,'http://test').searchParams.get('month')),canEdit:true};
            return{ok:true,json:async()=>({ok:true,...payload})};
        }catch(error){return{ok:false,json:async()=>({ok:false,error:error.message})};}
    }};
}

// Minimal event/field fixture: exercises the shipped client handlers and the real
// validation/store/compiler, with no browser, UI automation or external writes.
function fixture(fetch) {
    const elements=new Map();
    const decode=value=>String(value || '').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');
    const get=id=>{
        if(elements.has(id))return elements.get(id);
        const classes=new Set();
        const element={id,value:'',checked:false,hidden:false,disabled:false,textContent:'',dataset:{},events:{},open:false,
            classList:{add:key=>classes.add(key),remove:key=>classes.delete(key),contains:key=>classes.has(key)},
            addEventListener(type,handler){this.events[type]=handler;},showModal(){this.open=true;},close(){this.open=false;},focus(){},
            setRangeText(value,start,end){this.value=this.value.slice(0,start)+value+this.value.slice(end);}
        };
        let markup='';
        Object.defineProperty(element,'innerHTML',{get:()=>markup,set:value=>{
            markup=value;
            for(const match of value.matchAll(/<input\b([^>]*\bid="([^"]+)"[^>]*)>/g)){
                const field=get(match[2]);field.value=decode(/\bvalue="([^"]*)"/.exec(match[1])?.[1]);field.checked=/\bchecked\b/.test(match[1]);
            }
            for(const match of value.matchAll(/<(select|textarea)\b[^>]*\bid="([^"]+)"[^>]*>([\s\S]*?)<\/\1>/g)){
                const field=get(match[2]);
                if(match[1]==='textarea')field.value=decode(match[3]);
                else {const options=[...match[3].matchAll(/<option\b([^>]*)>/g)];const chosen=options.find(option=>/\bselected\b/.test(option[1])) || options[0];field.value=decode(/value="([^"]*)"/.exec(chosen?.[1] || '')?.[1]);}
            }
        }});
        elements.set(id,element);return element;
    };
    const window={GohMonthlySummaryModel:model,confirm:()=>true};
    new Function('window','document','fetch',fs.readFileSync(path.join(__dirname,'../assets/js/monthly-summary.js'),'utf8'))(window,{getElementById:get},fetch);
    return{get,controller:window.GohMonthlySummary,
        click:(id,event={})=>get(id).events.click(event),
        change:field=>get('kfs-fields').events.change({target:field}),
        select:index=>get('kfs-list').events.click({target:{closest:()=>({dataset:{select:String(index)}})}})};
}

test('user can choose a group, exclude one account, preview, save and reopen the shared monthly setup',async()=>{
    const b=backend(),f=fixture(b.fetch);
    f.controller.mode(true);await f.controller.load('2026-10');
    f.controller.update({orders:2,invoiced:4000,cost:2000,dg:50,material:1600,warehouse_delta:100,purchase_gap:50});
    assert.equal(f.get('kfSummaryBuild').hidden,false);
    assert.match(f.get('kfSummaryReport').innerHTML,/1\.500,00 DKK/);
    await f.click('kfSummaryBuild');assert.equal(f.get('kfSummaryDialog').open,true);
    f.select(4);f.get('kfs-name').value='Vareforbrug faktureret';
    f.change({id:'',dataset:{kind:'groups',value:'12_Vareforbrug'},checked:true});
    f.change({id:'',dataset:{kind:'exclude',value:'12092'},checked:true});
    await f.click('kfs-preview-button');
    assert.equal(b.writes(),0);assert.match(f.get('kfs-preview').innerHTML,/1\.625,00 DKK/);
    assert.match(f.get('kfs-preview').innerHTML,/Vareforbrug faktureret/);
    assert.doesNotMatch(f.get('kfSummaryReport').innerHTML,/Vareforbrug faktureret/);
    await f.click('kfs-save');
    assert.equal(b.writes(),1);assert.equal(f.get('kfSummaryDialog').open,false);
    assert.match(f.get('kfSummaryReport').innerHTML,/Vareforbrug faktureret/);
    const fresh=fixture(b.fetch);await fresh.controller.load('2026-10');
    assert.match(fresh.get('kfSummaryReport').innerHTML,/1\.625,00 DKK/);
    await fresh.click('kfSummaryBuild');fresh.select(4);
    assert.match(fresh.get('kfs-selected').innerHTML,/12_Vareforbrug/);
    assert.match(fresh.get('kfs-selected').innerHTML,/Udelad 12092/);
});

test('late month responses cannot replace the current month, and failures restore the standard summary',async()=>{
    const pending={};
    const f=fixture(url=>new Promise(resolve=>{pending[new URL(url,'http://test').searchParams.get('month')]=resolve;}));
    const first=f.controller.load('2026-09'),second=f.controller.load('2026-10');
    const response=month=>({ok:true,json:async()=>({ok:true,month,definition:{...model.defaults(),title:month},selections:{booked:[12070]},accounts:[{AcNo:12070,Nm:'Køb',Amount:0}]})});
    pending['2026-10'](response('2026-10'));await second;
    pending['2026-09'](response('2026-09'));await first;
    assert.match(f.get('kfSummaryReport').innerHTML,/2026-10/);
    assert.doesNotMatch(f.get('kfSummaryReport').innerHTML,/2026-09/);
    const third=f.controller.load('2026-11');pending['2026-11']({ok:false,json:async()=>({ok:false,error:'GOH unavailable'})});await third;
    assert.equal(f.get('kfSummaryLegacy').classList.contains('kfs-configured'),false);
    assert.equal(f.get('kfSummaryReport').hidden,true);
    assert.match(f.get('kfSummaryStatus').textContent,/GOH unavailable/);
});

test('invalid formulas fail before any write and a conflicting save leaves the editor open',async()=>{
    const b=backend(),f=fixture(b.fetch);
    await f.controller.load('2026-10');await f.click('kfSummaryBuild');
    f.select(6);f.get('kfs-formula').value='[does_not_exist]';await f.click('kfs-save');
    assert.equal(b.writes(),0);assert.equal(f.get('kfSummaryDialog').open,true);
    assert.match(f.get('kfs-message').textContent,/ovenfor/);
    f.get('kfs-formula').value='[booked] - [material]';
    const other=await b.service.load();await b.service.save(other,'another-admin');
    await f.click('kfs-save');
    assert.equal(b.writes(),1);assert.equal(f.get('kfSummaryDialog').open,true);
    assert.match(f.get('kfs-message').textContent,/anden bruger/);
});
