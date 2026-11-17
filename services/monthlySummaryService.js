const crypto = require('crypto');
const model = require('../assets/js/monthly-summary-model');
const { compileDefinition } = require('./bilancioDefinition');
const { parseCalendarMonth } = require('./omsaetningService');

function compile(input, catalog) {
    const definition=model.validate(input);
    const selections={};
    for (const row of definition.rows.filter(row => row.type === 'accounts')) {
        // Use Økonomirapport's account/group resolver. Separate KPI rows may intentionally
        // compare overlapping accounts; deduplication still applies inside each row.
        const compiled=compileDefinition({ schema:1,version:0,reportName:'Månedsoversigt',balanceTitle:'—',
            revenueRow:row.id,beforeTaxRow:row.id,pnl:[row],balance:[{id:'unused',name:'—',type:'heading'}] },catalog);
        selections[row.id]=compiled.pnl[0].selectedAccounts;
    }
    return {definition,selections};
}
function createMonthlySummaryService({ gohData, getProfile, getConnection, sql, getCatalog }) {
    const scope = () => {
        const profile=getProfile();
        return 'efterkalk_summary_' + crypto.createHash('sha256').update(JSON.stringify([
            String(profile.server).toLowerCase(),String(profile.database).toLowerCase()
        ])).digest('hex').slice(0,32);
    };
    const error = (message,statusCode) => Object.assign(new Error(message),{statusCode});
    async function load() {
        const key=scope();
        const state=await gohData.getAppState(key,{strict:true});
        if (key !== scope()) throw error('Databaseprofilen er ændret. Genindlæs oversigten.',409);
        return { ...model.validate(state ? state.payload : model.defaults()),scope:key };
    }
    async function editor() {
        const [definition,catalog]=await Promise.all([load(),getCatalog()]);
        if (definition.scope !== scope()) throw error('Databaseprofilen er ændret. Genindlæs opsætningen.',409);
        return {definition,catalog,sources:model.sources};
    }
    async function save(input,username) {
        const key=scope();
        if (input?.scope !== key) throw error('Databaseprofilen er ændret. Genindlæs opsætningen.',409);
        const {definition}=compile(input,await getCatalog());
        if (key !== scope()) throw error('Databaseprofilen er ændret. Genindlæs opsætningen.',409);
        const payload={...definition,scope:key,version:definition.version+1,updatedBy:username,updatedAt:new Date().toISOString()};
        const saved=await gohData.setAppState(key,payload,{expectedVersion:definition.version});
        if (!saved) throw error('Opsætningen kunne ikke gemmes. En anden bruger kan have ændret den; genindlæs og prøv igen.',409);
        return payload;
    }
    async function data(month,override) {
        const meta=parseCalendarMonth(month);
        if (!meta) throw error('Vælg en gyldig måned.',400);
        const definition=override || await load();
        const key=scope();
        if (definition.scope !== key) throw error('Databaseprofilen er ændret. Genindlæs opsætningen.',409);
        const compiled=compile(definition,await getCatalog());
        if (key !== scope()) throw error('Databaseprofilen er ændret. Genindlæs oversigten.',409);
        const ids=[...new Set(Object.values(compiled.selections).flat())];
        let accounts=[];
        if (ids.length) {
            const pool=await getConnection();
            const result=await pool.request()
                .input('year',sql.Int,Math.floor(meta.fiscalPeriodKey/100))
                .input('period',sql.Int,meta.fiscalPeriodKey%100)
                .input('accounts',sql.NVarChar(sql.MAX),ids.join(','))
                .query(`SELECT a.AcNo,a.Nm,CAST(COALESCE(SUM(CAST(t.AcAm AS decimal(38,6))),0) AS decimal(38,6)) AS Amount
                        FROM Ac a LEFT JOIN AcTr t ON t.AcNo=a.AcNo AND t.AcYr=@year AND t.AcPr=@period
                        WHERE a.AcNo IN (SELECT TRY_CAST(value AS int) FROM STRING_SPLIT(@accounts,','))
                        GROUP BY a.AcNo,a.Nm ORDER BY a.AcNo`);
            accounts=result.recordset || [];
        }
        if (key !== scope()) throw error('Databaseprofilen er ændret. Genindlæs oversigten.',409);
        return {month:meta.raw,definition:{...compiled.definition,scope:key},selections:compiled.selections,accounts};
    }
    function lockKey(month) {
        if (!parseCalendarMonth(month)) throw error('Vælg en gyldig måned.',400);
        return scope() + '_lock_' + month;
    }
    async function lockState(month) {
        const key=lockKey(month);
        const state=await gohData.getAppState(key,{strict:true});
        if (key !== lockKey(month)) throw error('Databaseprofilen er ændret. Genindlæs måneden.',409);
        const value=state ? state.payload : {version:0,locked:false};
        if (!Number.isSafeInteger(value.version) || typeof value.locked !== 'boolean'
            || (value.locked && (!value.snapshot || value.snapshot.summary?.month !== month))) {
            throw error('Den gemte afstemning er ugyldig. Kontakt administratoren.',503);
        }
        return value;
    }
    async function changeLock(month,version,action,snapshot,username) {
        const key=lockKey(month);
        if (!Number.isSafeInteger(version) || version < 0 || !['lock','unlock'].includes(action)) {
            throw error('Ugyldig anmodning om låsning.',400);
        }
        const current=await lockState(month);
        if (version !== current.version || current.locked === (action === 'lock')) {
            throw error('Låsestatus er ændret. Genindlæs måneden og prøv igen.',409);
        }
        if (action === 'lock') {
            const s=snapshot?.summary, d=snapshot?.details;
            const close=(a,b) => Number.isFinite(a) && Number.isFinite(b) && Math.abs(a-b)<0.01;
            const required=['orders','invoiced','cost','material','material_only','stang','parts','subcontract',
                'warehouse_delta','purchase_gap','ub_orders','ub_material','finished_material','purchase_timing'];
            if (!s || s.month !== month || s.scope !== scope() || !Array.isArray(s.rows) || !s.rows.length
                || !s.metrics || required.some(name => !Number.isFinite(s.metrics[name]))
                || s.rows.some(row => row.display !== 'hidden' && row.type !== 'heading' && !Number.isFinite(row.value))
                || !s.signs || !d || d.month !== month || !d.varelager || !d.purchaseGap || !d.wip || !d.timing
                || !Array.isArray(snapshot.ubOrders) || !Array.isArray(snapshot.controlList)
                || !['fifo','standard'].includes(snapshot.plateMode)
                || !['registered','adjusted'].includes(snapshot.costMode)
                || d.varelager.plateMode !== snapshot.plateMode || !Array.isArray(d.varelager.categories) || !d.varelager.categories.length
                || d.varelager.categories.some(item => !Number.isFinite(item.delta))
                || !close(d.varelager.categories.reduce((sum,item) => sum+item.delta,0),s.metrics.warehouse_delta)
                || !close(d.wip.delta,s.metrics.finished_material) || !close(d.purchaseGap.totalValue,s.metrics.purchase_gap)
                || !close(d.timing.value,s.metrics.purchase_timing)
                || snapshot.ubOrders.some(row => !Number.isFinite(row.material))
                || !snapshot.orderCosts || !Array.isArray(snapshot.orderCosts.invoices)
                || !snapshot.orderCosts.margins || !snapshot.orderCosts.ubMargins) {
                throw error('Afstemningen er ikke færdigberegnet. Genindlæs og kontrollér alle tal.',400);
            }
        }
        if (key !== lockKey(month)) throw error('Databaseprofilen er ændret. Genindlæs måneden.',409);
        const now=new Date().toISOString();
        const next={version:version+1,locked:action === 'lock',snapshot:action === 'lock' ? snapshot : current.snapshot,
            lockedAt:action === 'lock' ? now : current.lockedAt,lockedBy:action === 'lock' ? username : current.lockedBy,
            updatedAt:now,updatedBy:username};
        if (!await gohData.setAppState(key,next,{expectedVersion:version})) {
            throw error('Låsestatus kunne ikke gemmes. Genindlæs måneden og prøv igen.',409);
        }
        if (key !== lockKey(month)) throw error('Databaseprofilen er ændret. Genindlæs måneden.',409);
        return next;
    }
    return {load,editor,save,data,scope,lockState,changeLock};
}
module.exports={createMonthlySummaryService,compile};
