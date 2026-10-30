const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

// Loads the extracted client file the same way test/viaService.test.js's clientFixture loads
// assets/js/via.js: read the source, wrap it with new Function(...), stub the page globals it
// documents at the top of the file (escapeHtml, formatNumber, pushModalStack, removeModalStack,
// searchOrderByNo), and expose just what each test needs.
function clientFixture({ fetchImpl, summary } = {}) {
    const source = fs.readFileSync(path.join(__dirname, '../assets/js/kundefaktura.js'), 'utf8');
    const els = {};
    const el = id => els[id] || (els[id] = {
        id, value: '', style: {}, textContent: '', disabled: false,
        classList: { add() {}, remove() {} }, innerHTML: '',
        querySelectorAll: () => [], querySelector: () => null,
        appendChild() {}, removeChild() {}, remove() {}
    });
    const noop = () => {};
    const document_ = { getElementById: el, querySelectorAll: () => [], querySelector: () => null, createElement: () => el('x') };
    const posts = [];
    const fetch_ = fetchImpl || (async () => ({ ok: false, json: async () => ({}) }));
    const client = new Function('document', 'window', 'fetch', 'escapeHtml', 'formatNumber', 'pushModalStack', 'removeModalStack', 'searchOrderByNo', source + `
        return {
            load: loadKundefakturaInvoices,
            setAll: v => { _kfAllCustomers = v; },
            addUb: addKfUbOrder,
            removeUb: removeKfUbOrder,
            refreshNote: refreshKfOrderClassification,
            exportOrders: exportKfOrdersCsv,
            state: () => ({ rows: _kfInvoiceRows.map(r => r.OrdNo), loadedMonth: _kfLoadedMonth, vareforbrug: _kfVareforbrugGL, loadSeq: _kfLoadSeq })
        };
    `)(document_, { requestAnimationFrame: f => f(), GohMonthlySummary:summary }, fetch_, String, String, noop, noop, noop);
    return { ...client, document: document_, posts };
}

test('loading a month fetches invoices and Vareforbrug for that exact month', async () => {
    const calls = [];
    const client = clientFixture({
        fetchImpl: async url => {
            calls.push(url);
            if (url.startsWith('/efterkalk/month-snapshot')) return { ok: false, json: async () => ({ ok: false, found: false }) };
            if (url.startsWith('/efterkalk/customer-invoices')) return { ok: true, json: async () => ({ ok: true, fromInt: 0, toInt: 0, rows: [{ OrdNo: 1, InvoAm: 1, LstInvDt: 0, OrdDt: 0, CustNo: 1 }] }) };
            if (url.startsWith('/efterkalk/vareforbrug')) return { ok: true, json: async () => ({ ok: true, month: /month=([\d-]+)/.exec(url)[1], vareforbrug: 42 }) };
            return { ok: true, json: async () => ({ totalCost: 1, costBreakdownAvailable: false, styklisteFallbackCost: 0 }) };
        }
    });
    client.setAll(true);
    client.document.getElementById('kfMonth').value = '2026-09';
    await client.load();
    assert.equal(client.state().loadedMonth, '2026-09');
    assert.deepEqual(client.state().rows, [1]);
    assert.ok(calls.some(u => u.includes('/efterkalk/vareforbrug?month=2026-09')));
    // The margin-fetch loop runs in the background (fire-and-forget, same as in the real app);
    // let it settle before the test ends so it doesn't report activity after completion.
    await new Promise(r => setTimeout(r, 20));
});

test('credit notes replace the positive-only snapshot and U/B stays outside invoice and cost totals', async () => {
    const calls = [];
    const client = clientFixture({ fetchImpl: async url => {
        calls.push(url);
        if (url.startsWith('/efterkalk/month-snapshot')) return { ok:true, json:async () => ({ ok:true, found:true, rows:[
            { OrdNo:1, Revenue:1000, Cost:100, CostComplete:true, StyklisteFallbackCost:0, OrderDate:'2026-10-01', InvoiceDate:'2026-10-05', CustNo:42, InvoiceNo:'100' }
        ] }) };
        if (url.startsWith('/efterkalk/customer-invoices')) return { ok:true, json:async () => ({ ok:true, rows:[
            { OrdNo:1, InvoAm:1000, CustNo:42 }, { OrdNo:2, InvoAm:-200, CustNo:42, InvoNo:'credit' },
            { OrdNo:3, InvoAm:500, CustNo:42, IsUB:true }
        ] }) };
        if (url.startsWith('/efterkalk/ub-orders')) return { ok:true, json:async () => ({ ok:true, rows:[] }) };
        return { ok:true, json:async () => ({ totalCost:50, costBreakdownAvailable:false }) };
    } });
    client.setAll(true);
    client.document.getElementById('kfMonth').value = '2026-10';
    await client.load();
    await new Promise(resolve => setTimeout(resolve, 20));
    assert.deepEqual(client.state().rows, [1, 2, 3]);
    assert.equal(client.document.getElementById('kfKpiTotal').textContent, '800');
    assert.equal(client.document.getElementById('kfKpiOrders').textContent, 2);
    assert.equal(client.document.getElementById('kfKpiCost').textContent, '100');
    assert.match(client.document.getElementById('kfKpiTotalSub').textContent, /1 kreditnotaer: -200/);
    assert.match(client.document.getElementById('kfTableBody').innerHTML, /Kreditnota/);
    assert.doesNotMatch(client.document.getElementById('kfTableBody').innerHTML, /searchOrderByNo\(3\)/);
    assert.ok(calls.some(url => url.includes('/efterkalk/customer-invoices?') && url.includes('includeAllAmounts=0')));

    client.refreshNote(3, { isUB:false });
    assert.equal(client.document.getElementById('kfKpiTotal').textContent, '1300');
    client.refreshNote(3, { isUB:true });
    assert.equal(client.document.getElementById('kfKpiTotal').textContent, '800');
});

test('manual U/B entry uses the loaded month, supports non-invoiced orders and refreshes after removal', async () => {
    let ubRows = [];
    const posts = [];
    const client = clientFixture({ fetchImpl: async (url, options = {}) => {
        if (url.startsWith('/efterkalk/month-snapshot')) return { ok:false, json:async () => ({}) };
        if (url.startsWith('/efterkalk/customer-invoices')) return { ok:true, json:async () => ({ ok:true, rows:[] }) };
        if (url.startsWith('/efterkalk/ub-orders')) {
            if (options.method === 'POST') {
                const body = JSON.parse(options.body); posts.push(body);
                ubRows = [{ OrdNo:99, OrdDt:20261003, CustNo:42, CustomerName:'Test <kunde>', text:body.text }];
                return { ok:true, json:async () => ({ ok:true, order:ubRows[0], note:{ isUB:true, ubMonth:body.month } }) };
            }
            if (options.method === 'DELETE') { ubRows = []; return { ok:true, json:async () => ({ ok:true, note:null }) }; }
            return { ok:true, json:async () => ({ ok:true, rows:ubRows }) };
        }
        return { ok:false, json:async () => ({}) };
    } });
    client.setAll(true);
    client.document.getElementById('kfMonth').value = '2026-10';
    await client.load();
    client.document.getElementById('kfMonth').value = '2026-11';
    client.document.getElementById('kfUbOrdNo').value = '99';
    client.document.getElementById('kfUbNote').value = 'Prøve uden beregning';
    await client.addUb();
    assert.deepEqual(posts[0], { ordNo:99, month:'2026-10', text:'Prøve uden beregning' });
    assert.match(client.document.getElementById('kfUbBody').innerHTML, /searchOrderByNo\(99\)/);
    assert.equal(client.document.getElementById('kfKpiTotal').textContent, '0');
    await client.removeUb(99);
    assert.match(client.document.getElementById('kfUbBody').innerHTML, /Ingen U\/B-ordrer/);
});

test('configurable summary receives signed invoices, separate U/B and complete cost/material sources',async()=>{
    const values=[];
    const client=clientFixture({summary:{load(){},update(value){values.push(value);}},fetchImpl:async url=>{
        if(url.startsWith('/efterkalk/month-snapshot'))return{ok:false,json:async()=>({})};
        if(url.startsWith('/efterkalk/ub-orders'))return{ok:true,json:async()=>({ok:true,rows:[{OrdNo:99}]})};
        if(url.startsWith('/efterkalk/vareforbrug'))return{ok:true,json:async()=>({ok:true,month:'2026-10',purchaseGap:{totalValue:40},varelager:{delta:10}})};
        if(url.startsWith('/efterkalk/customer-invoices'))return{ok:true,json:async()=>({ok:true,rows:[
            {OrdNo:1,InvoAm:1000},{OrdNo:2,InvoAm:-200},{OrdNo:3,InvoAm:500,IsUB:true}
        ]})};
        return{ok:true,json:async()=>({totalCost:100,costBreakdownAvailable:true,materialCost:20,stangCost:5,timeCost:30,purchasedPartCost:10,underleverandorCost:15})};
    }});
    client.setAll(true);client.document.getElementById('kfMonth').value='2026-10';
    await client.load();await new Promise(resolve=>setTimeout(resolve,20));
    assert.ok(values.some(value=>value.cost===null));
    const last=values.at(-1);
    assert.equal(last.orders,2);assert.equal(last.invoiced,800);assert.equal(last.credits,-200);
    assert.equal(last.cost,200);assert.equal(last.db,600);assert.equal(last.dg,75);
    assert.equal(last.material,100);assert.equal(last.time,60);
    assert.equal(last.purchase_gap,40);assert.equal(last.warehouse_delta,10);assert.equal(last.ub_orders,1);
});

test('a month switch while a load is in flight never leaves the older month\'s data showing', async () => {
    // Regression for the 2026-10-05 fix: _kfLoadSeq/_kfLoadedMonth must discard stale responses.
    // Without the guard, an out-of-order response from the first (September) request could
    // overwrite the second (August) request's rows after the user already switched months.
    const pending = {};
    const deferred = () => { let resolve; const p = new Promise(r => { resolve = r; }); return { p, resolve }; };
    const client = clientFixture({
        fetchImpl: async url => {
            if (url.startsWith('/efterkalk/month-snapshot')) return { ok: false, json: async () => ({ ok: false, found: false }) };
            if (url.startsWith('/efterkalk/customer-invoices')) {
                const month = /from=(\d{4}-\d{2})/.exec(decodeURIComponent(url))[1];
                const d = deferred(); pending[month] = d;
                return d.p.then(() => ({ ok: true, json: async () => ({ ok: true, fromInt: 0, toInt: 0, rows: [{ OrdNo: month === '2026-09' ? 900 : 800, InvoAm: 1, LstInvDt: 0, OrdDt: 0, CustNo: 1 }] }) }));
            }
            if (url.startsWith('/efterkalk/vareforbrug')) return { ok: true, json: async () => ({ ok: true, month: /month=([\d-]+)/.exec(url)[1], vareforbrug: 1 }) };
            return { ok: true, json: async () => ({ totalCost: 1, costBreakdownAvailable: false, styklisteFallbackCost: 0 }) };
        }
    });
    client.setAll(true);
    const tick = ms => new Promise(r => setTimeout(r, ms));

    client.document.getElementById('kfMonth').value = '2026-09';
    const sept = client.load();
    await tick(5);
    client.document.getElementById('kfMonth').value = '2026-08';
    const aug = client.load();
    await tick(5);

    // August answers first, September answers late (out of order).
    pending['2026-08'].resolve();
    await aug;
    await tick(5);
    pending['2026-09'].resolve();
    await sept;
    await tick(20);

    assert.equal(client.state().loadedMonth, '2026-08');
    assert.deepEqual(client.state().rows, [800]);
});
