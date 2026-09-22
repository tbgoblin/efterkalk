const test = require('node:test');
const assert = require('node:assert/strict');
const { parseFormula, evaluateFormula } = require('../services/bilancioFormula');
const { validateDefinition, evaluateRows } = require('../services/bilancioDefinition');
const { sourceMonth, lagerValue, createSourceResolver } = require('../services/bilancioSources');
const calculate = (text, refs={}) => evaluateFormula(parseFormula(text).tree, id => refs[id]);
const config = pnl => ({ schema:1,version:0,legacyPeriodRows:false,revenueRow:pnl[0].id,beforeTaxRow:pnl[0].id,balanceTitle:'Aktiver',pnl,balance:[{id:'title',name:'Aktiver',type:'heading'}] });
const formula = (id, expression) => ({id,name:id,type:'formula',formula:expression});

test('formula precedence, unary signs, decimals, percentages and stable references', () => {
    assert.equal(calculate('([sales] - [cost]) * -22%', {sales:1000,cost:400}), -132);
    assert.equal(calculate('2 + 3 * 4'), 14);
    assert.equal(calculate('-(-2) + .5'), 2.5);
    assert.equal(calculate('1500.50'), 1500.5);
    assert.deepEqual(parseFormula('[a] + [a]').refs, ['a']);
});

test('formula parser rejects code, invalid references and excessive nesting', () => {
    for (const value of ['process.exit()', '1;2', 'Math.max(1,2)', '[a].constructor', '1,50', '(2', '2 2', '', '('.repeat(50)+'1'+')'.repeat(50)]) assert.throws(()=>parseFormula(value), value);
    assert.throws(()=>validateDefinition(config([formula('first','[later]'),formula('later','1')])));
    assert.throws(()=>validateDefinition(config([formula('first','[first]')])));
});

test('missing inputs and zero division propagate through formulas and sums', () => {
    assert.equal(calculate('[missing] * 0'), null);
    assert.equal(calculate('1 / 0'), null);
    const rows=validateDefinition(config([
        {id:'sales',name:'Sales',type:'lager',metric:'sales',period:'selected',visibility:'hidden'},
        formula('margin','[sales] - 200'),
        {id:'total',name:'Total',type:'sum',sources:['margin']}
    ])).pnl;
    const values=evaluateRows(rows,[],2,[],new Map([['sales',[1000,null]]]));
    assert.deepEqual(values.map(r=>r.amounts),[[1000,null],[800,null],[800,null]]);
    assert.equal(values[0].visibility,'hidden');
});

test('periods handle fiscal boundaries, calendar current, fixed months and prior-year columns', () => {
    assert.equal(sourceMonth({period:'selected'},2026,7,0,'2026-09'),'2027-01');
    assert.equal(sourceMonth({period:'previous'},2026,7,0,'2026-09'),'2026-12');
    assert.equal(sourceMonth({period:'previous'},2026,1,1,'2026-09'),'2025-06');
    assert.equal(sourceMonth({period:'current'},2020,1,1,'2026-09'),'2025-09');
    assert.equal(sourceMonth({period:'fixed',month:'2024-08'},2026,1,1,'2026-09'),'2024-08');
});

test('manual month values preserve zero, comparison and missing values; YTD is point-in-time', async () => {
    const row={id:'adjustment',name:'Adjustment',type:'manual',period:'selected',monthValues:{'2026-08':0,'2025-08':-50}};
    validateDefinition(config([row]));
    const resolve=createSourceResolver({year:2026,period:2,today:'2026-09'});
    assert.deepEqual((await resolve([row],4)).get(row.id),[0,-50,0,-50]);
    assert.deepEqual((await resolve([{...row,period:'fixed',month:'2024-01'}],2)).get(row.id),[null,null]);
    assert.throws(()=>validateDefinition(config([{...row,monthValues:{'2026-13':1}}])));
    assert.throws(()=>validateDefinition(config([{...row,monthValues:{'2026-08':Infinity}}])));
});

test('Lagerliste sources cache reads and request price enrichment only when selected', async () => {
    const calls=[]; let live=0;
    const resolve=createSourceResolver({year:2026,period:2,today:'2026-09',fs:{},lagerlisteService:{
        loadMonthlySnapshot:async request=>{calls.push(request);return request.month==='2026-08'?{current:{totals:{finishedNotInvoicedSales:1000,finishedNotInvoiced:600}}}:null;},
        getCurrent:async()=>{live++;return {totals:{finishedNotInvoiced:800}};}
    }});
    const rows=[{id:'sales',type:'lager',metric:'sales',period:'selected',currentSales:true},{id:'cost',type:'lager',metric:'cost',period:'selected',currentSales:true}];
    const values=await resolve(rows,4);
    assert.deepEqual(values.get('sales'),[1000,null,1000,null]);
    assert.equal(calls.length,2); assert.ok(calls.every(c=>c.includeCurrentSales));assert.equal(live,0);
    assert.deepEqual((await resolve([{id:'now',type:'lager',metric:'cost',period:'current'}],2)).get('now'),[800,null]);
    assert.equal(live,1);
});

test('Lagerliste totals match overview: whole-order margin, warehouse and VIA', () => {
    const p={totals:{finishedNotInvoiced:100,finishedNotInvoicedSales:150,plates:10,restPlates:2,stang:3,opfolgningvare:4,diverse:5},categories:{gr5Items:[{FifoValue:6}],salgordreVia:[{TimeCost:20,MaterialCost:30,StangCost:40,PurchasedPartCost:50}],nestingCutting:[{CountedValue:60,Value:600},{Value:-20}]}};
    assert.equal(lagerValue(p,'margin'),50);assert.equal(lagerValue(p,'warehouse'),30);assert.equal(lagerValue(p,'via'),300);
    assert.equal(lagerValue(null,'sales'),null);
});

test('custom period rows disable the fixed block and preserve presentation through validation', () => {
    const c=validateDefinition(config([{...formula('base','1200.50'),bold:true,visibility:'period'}]));
    assert.equal(c.legacyPeriodRows,false);assert.equal(c.pnl[0].bold,true);assert.equal(c.pnl[0].visibility,'period');
    assert.equal(validateDefinition({...c,legacyPeriodRows:undefined}).legacyPeriodRows,true);
});
