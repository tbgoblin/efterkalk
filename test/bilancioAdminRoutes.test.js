const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
function fixture() {
    const source = fs.readFileSync(path.join(__dirname,'../routes/apiRoutes.js'),'utf8');
    const start = source.indexOf('    const bilancioAdminGuard =');
    const end = source.indexOf("    router.get('/bilancio-trial/data'",start);
    const routes = new Map();let writes=0,previews=0;
    const router = Object.fromEntries(['get','put','post'].map(method=>[method,(url,...handlers)=>routes.set(method+url,handlers)]));
    new Function('router','requireSuperadmin','bilancioDefinitionStore','bilancioService','getSessionUser','validateBilancioPeriod',source.slice(start,end))(
        router,(req,res)=>{if(req.role==='superadmin')return true;res.status(403).json({error:'denied'});return false;},
        {scope:()=> 'scope',load:async()=>({version:0}),save:async(config,catalog,username)=>{writes++;return {version:1,updatedBy:username};}},
        {catalog:async()=>({accounts:[],groups:[]}),report:async()=>{previews++;return {rows:[]};}},
        ()=>({username:'admin'}),()=>{}
    );
    async function request(method,url,role,body={}) {
        const result={status:200};const req={role,body};
        const res={setHeader:()=>{},status:code=>{result.status=code;return res;},json:payload=>{result.payload=payload;return res;}};
        const [guard,handler]=routes.get(method+url);let proceed=false;
        guard(req,res,()=>{proceed=true;});if(proceed)await handler(req,res);return result;
    }
    return { request,writes:()=>writes,previews:()=>previews };
}
test('all configuration endpoints reject ordinary users before reading or writing', async()=>{
    const f=fixture();for(const [method,url] of [['get','/admin/bilancio-definition'],['put','/admin/bilancio-definition'],['post','/admin/bilancio-definition/preview']]){
        assert.equal((await f.request(method,url,'user')).status,403);
    }
    assert.equal(f.writes(),0);assert.equal(f.previews(),0);
});
test('admin can save with actor attribution; preview rejects changed profile',async()=>{
    const f=fixture();assert.equal((await f.request('put','/admin/bilancio-definition','superadmin')).payload.definition.updatedBy,'admin');
    assert.equal((await f.request('post','/admin/bilancio-definition/preview','superadmin',{definition:{scope:'wrong'},year:2026,period:2})).status,409);
    assert.equal(f.previews(),0);
    assert.equal((await f.request('post','/admin/bilancio-definition/preview','superadmin',{definition:{scope:'scope'},year:2026,period:2})).status,200);
});
