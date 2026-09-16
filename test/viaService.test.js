const test = require('node:test');
const assert = require('node:assert/strict');
const { fetchSalgordreViaRows, normalizePurchasedPartRows, getOpenViaOrders, buildViaBacklog } = require('../services/viaService');
const fs = require('node:fs');
const path = require('node:path');

test('purchased order parts count received quantity when it is not follow-up stock', async () => {
    let query = '';
    const pool = {
        request() {
            return {
                input() { return this; },
                async query(statement) {
                    query = statement;
                    return { recordset: [{
                        OrdNo: 900,
                        PurchasedPartCost: 0,
                        PurchasedPartDetailsJson: JSON.stringify([{
                            productionOrderNo: 700,
                            purchaseOrderNo: 800,
                            prodNo: '100-A',
                            orderedQty: 10,
                            receivedQty: 8,
                            consumedQty: 6,
                            unitPrice: 1,
                            countedValue: 6
                        }])
                    }] };
                }
            };
        }
    };

    const rows = await fetchSalgordreViaRows({ getConnection: async () => pool, sql: { Numeric: 'Numeric' }, requestedOrdNo: null });

    assert.equal(rows[0].PurchasedPartCost, 8);
    assert.equal(rows[0].PurchasedPartDetails[0].receivedQty, 8);
    assert.equal(rows[0].PurchasedPartDetails[0].consumedQty, 6);
    assert.equal('PurchasedPartDetailsJson' in rows[0], false);
    assert.match(query, /FROM ProdTr T WITH\(NOLOCK\)/i);
    assert.match(query, /T\.TrTp\s*=\s*6/i);
    assert.match(query, /FROM Rsv R WITH\(NOLOCK\)/i);
});

test('received reserved follow-up stock moves to VIA at FIFO without using ordered-only quantity', () => {
    const rows = normalizePurchasedPartRows([{
        OrdNo: 900,
        PurchasedPartDetailsJson: JSON.stringify([{
            prodNo: 'A', orderedQty: 10, receivedQty: 8, consumedQty: 2,
            unitPrice: 10, isFollowUpStock: true, physicalStockQty: 6,
            stockUnitCost: 12, activeReservedQty: 6
        }, {
            prodNo: 'B', orderedQty: 5, receivedQty: 0, consumedQty: 0,
            unitPrice: 20, isFollowUpStock: true, physicalStockQty: 0,
            stockUnitCost: 21, activeReservedQty: 0
        }, {
            prodNo: 'C', orderedQty: 5, receivedQty: 5, consumedQty: 0,
            unitPrice: 20, isFollowUpStock: true, physicalStockQty: 5,
            stockUnitCost: 21, activeReservedQty: 0
        }])
    }]);

    assert.equal(rows[0].PurchasedPartCost, 92);
    assert.deepEqual(rows[0].PurchasedPartDetails.map(row => ({
        countedQty: row.countedQty,
        stockTransferQty: row.stockTransferQty,
        countedValue: row.countedValue
    })), [
        { countedQty: 8, stockTransferQty: 6, countedValue: 92 },
        { countedQty: 0, stockTransferQty: 0, countedValue: 0 },
        { countedQty: 0, stockTransferQty: 0, countedValue: 0 }
    ]);
});

function queryRecorder() {
    const queries = [];
    return {
        queries,
        sql: { Numeric: 'Numeric' },
        getConnection: async () => ({ request() {
            const inputs = {};
            return {
                input(name, type, value) { inputs[name] = value; return this; },
                async query(text) { queries.push({ text, inputs }); return { recordset: [] }; }
            };
        } })
    };
}

test('inventory getter retains its historical production scope and isolated single-order predicate', async () => {
    const recorder = queryRecorder();
    await fetchSalgordreViaRows({ ...recorder, requestedOrdNo: 123 });
    const { text, inputs } = recorder.queries[0];
    assert.match(text, /Gr12 <> 10/);
    assert.match(text, /OrdPrSt & 256 = 256/);
    assert.match(text, /OrdPrSt = 0 OR OrdPrSt = 402653456/);
    assert.match(text, /OrdPrSt = 134217728 OR OrdPrSt & 4194304 = 4194304/);
    assert.match(text, /AND \(@requestedOrdNo IS NULL OR OrdNo = @requestedOrdNo\)/);
    assert.doesNotMatch(text, /@firstDate|@asOfDate|LstInvDt/);
    assert.deepEqual(inputs, { requestedOrdNo: 123 });
    const source = fs.readFileSync(path.join(__dirname, '../services/lagerlisteService.js'), 'utf8');
    assert.match(source, /getSalgordreViaRows\(\{ getConnection, sql, requestedOrdNo: null \}\)/);
});

test('commercial cost enrichment selects only explicit parameterized order IDs', async () => {
    const recorder = queryRecorder();
    await fetchSalgordreViaRows({ ...recorder, requestedOrdNo: 123, orderNos: [123, 456, 123] });
    const { text, inputs } = recorder.queries[0];
    assert.deepEqual(inputs, { requestedOrdNo: 123, viaOrder0: 123 });
    assert.match(text, /OrdNo IN \(@viaOrder0\)/);
    assert.doesNotMatch(text, /Gr12 <> 10|OrdPrSt & 256/);
    assert.match(text, /AND \(@requestedOrdNo IS NULL OR OrdNo = @requestedOrdNo\)/);
    assert.deepEqual(await fetchSalgordreViaRows({ ...recorder, orderNos: [] }), []);
    assert.deepEqual(await fetchSalgordreViaRows({ ...recorder, requestedOrdNo: 999, orderNos: [123] }), []);
    await assert.rejects(fetchSalgordreViaRows({ ...recorder, orderNos: ['123) OR 1=1'] }), /ugyldige/);
    assert.equal(recorder.queries.length, 1);
});

test('large explicit selections are batched below the SQL parameter limit', async () => {
    const recorder = queryRecorder();
    await fetchSalgordreViaRows({ ...recorder, orderNos: Array.from({ length: 1001 }, (_, index) => index + 1) });
    assert.equal(recorder.queries.length, 2);
    assert.equal(Object.keys(recorder.queries[0].inputs).length, 1001);
    assert.deepEqual(recorder.queries[1].inputs, { requestedOrdNo: null, viaOrder0: 1001 });
});

function backlogFixture() {
    const flow = require('../assets/js/order-flow').build([
        { OrdNo: 1, OrderDate: 20260810, OrderValueDkk: 100, InvoicedDkk: 40, RemainingDkk: 60, Gr12: 10, OrdPrSt: 1 },
        { OrdNo: 2, OrderDate: 20260910, OrderValueDkk: 80, InvoicedDkk: 80, RemainingDkk: 0 },
        { OrdNo: 3, OrderDate: 20260910, OrderValueDkk: 25, InvoicedDkk: 0, RemainingDkk: 25, OrdPrSt: 1 },
        { OrdNo: 4, OrderDate: 20260910, OrderValueDkk: 50, InvoicedDkk: 10, RemainingDkk: 40 }
    ], [
        { OrdNo: 1, InvoicedBeforeDkk: 10, InvoicedInMonthDkk: 30, InvoicedTotalDkk: 40 },
        { OrdNo: 2, InvoicedBeforeDkk: 0, InvoicedInMonthDkk: 80, InvoicedTotalDkk: 80 }
    ], 20260901);
    return { ...flow, month: '2026-09', asOf: 20260916 };
}

function costFixture(ordNo) {
    return { OrdNo: ordNo, SalesValue: 999, MaterialCost: 10, StangCost: 2, PurchasedPartCost: 3, TimeCost: 4, PurchasedPartDetails: [] };
}

test('VIA reconciles positive closing residuals, including MultiOrdre and orders without production', () => {
    const flow = backlogFixture();
    const result = buildViaBacklog(flow, [costFixture(1), { ...costFixture(3), MaterialCost: 0, StangCost: 0, PurchasedPartCost: 0, TimeCost: 0 }]);
    assert.deepEqual(getOpenViaOrders(flow).map(row => row.ordNo), [1, 3]);
    assert.deepEqual(result.rows.map(row => [row.OrdNo, row.RemainingSalesValue, row.SalesValue]), [[1, 60, 100], [3, 25, 25]]);
    assert.equal(result.orderBacklogValueDkk, flow.total.closing);
    assert.equal(result.orderBacklogValueDkk, 85);
    assert.equal(result.unknownCount, 1);
    assert.equal(result.missingCostCount, 0);
    assert.equal(result.rows[0].PurchasedPartCost, 3);
    assert.equal(result.rows[1].CostDataAvailable, true);
    assert.equal(result.rows[1].TimeCost, 0);
    assert.equal(result.rows[0].Gr12, 10);
});

test('missing enrichment remains unknown rather than zero; duplicate and invalid sources are rejected', () => {
    const flow = backlogFixture();
    const result = buildViaBacklog(flow, [costFixture(1)]);
    assert.equal(result.rows[1].CostDataAvailable, false);
    assert.equal(result.rows[1].MaterialCost, null);
    assert.equal(result.orderBacklogValueDkk, 85);
    assert.equal(result.missingCostCount, 1);
    assert.throws(() => buildViaBacklog(flow, [costFixture(1), costFixture(1)]), /Dublerede/);
    assert.throws(() => getOpenViaOrders({ ...flow, total: { closing: null } }), /kunne ikke/);
    assert.throws(() => getOpenViaOrders({ ...flow, rows: [flow.rows[0], flow.rows[0]] }), /dublerede/);
});

test('single-order backlog results never substitute another order or a completed order', () => {
    const flow = backlogFixture();
    const result = buildViaBacklog(flow, [costFixture(1), costFixture(3)], 3);
    assert.deepEqual(result.rows.map(row => row.OrdNo), [3]);
    assert.equal(result.orderBacklogValueDkk, 25);
    assert.deepEqual(buildViaBacklog(flow, [costFixture(2)], 2).rows, []);
});

test('sub-cent residuals use the Ordreflow row threshold and remain explicit in reconciliation', () => {
    const flow = backlogFixture();
    flow.rows.push({ ordNo: 5, closing: 0.00811, value: 10 });
    flow.total.closing += 0.00811;
    const result = buildViaBacklog(flow, [costFixture(1), costFixture(3)]);
    assert.deepEqual(result.rows.map(row => row.OrdNo), [1, 3]);
    assert.equal(result.orderBacklogValueDkk, 85);
    assert.ok(Math.abs(result.excludedResidualDkk - 0.00811) < 1e-9);
    assert.equal(result.orderBacklogValueDkk + result.excludedResidualDkk, flow.total.closing);
});

function routeFixture() {
    const source = fs.readFileSync(path.join(__dirname, '../routes/apiRoutes.js'), 'utf8');
    const cache = new Map();
    const staleCache = new Map();
    const diskCache = { get: key => cache.get(key), getStale: key => staleCache.get(key), set: (key, value) => cache.set(key, value), del: key => cache.delete(key) };
    let profile = { id: 'production', server: 'server-a', database: 'database-a' };
    let flowCalls = 0;
    const costCalls = [];
    const helpers = new Function('settingsService', 'crypto', 'omsaetningService', 'getOpenViaOrders', 'buildViaBacklog', 'fetchSalgordreViaRows', 'getConnection', 'sql', 'diskCache',
        source.slice(source.indexOf('    const VIA_CACHE_KEY ='), source.indexOf('    async function warmSalgordreVia')) + '\nreturn { viaContext, getViaPayload, viaRequests };')(
        { getActiveProfile: () => profile }, require('node:crypto'), { getOrderFlow: async () => { flowCalls++; return backlogFixture(); } },
        getOpenViaOrders, buildViaBacklog, async options => { costCalls.push(options); return (options.orderNos || [99]).map(costFixture); }, null, null, diskCache
    );
    let handler;
    const start = source.indexOf("    router.get('/salgordre-via',");
    const end = source.indexOf("    router.get('/salgordre-via/reservations',", start);
    new Function('router', 'requireModulePermission', 'getSessionUser', 'viaContext', 'getViaPayload', 'viaRequests', 'diskCache', 'logEvent', source.slice(start, end))(
        { get(route, guard, callback) { handler = callback; } }, () => null, request => request.user,
        helpers.viaContext, helpers.getViaPayload, helpers.viaRequests, diskCache, () => {}
    );
    return {
        ...helpers, cache, staleCache, costCalls,
        flowCalls: () => flowCalls,
        changeProfile: () => { profile = { ...profile, id: 'other', database: 'database-b' }; },
        request: async (query = {}, user = { role: 'superadmin' }) => {
            const result = { status: 200 };
            await handler({ query, user }, { setHeader() {}, status(code) { result.status = code; return this; }, json(payload) { result.payload = payload; return this; } });
            return result;
        }
    };
}

test('cache-only requests do not query SQL; financial and production caches are isolated', async () => {
    const fixture = routeFixture();
    assert.deepEqual((await fixture.request({ cached: '1' })).payload, { notCached: true });
    assert.equal(fixture.flowCalls(), 0);
    assert.equal(fixture.costCalls.length, 0);
    const financial = await fixture.request();
    assert.equal(financial.payload.orderBacklogValueDkk, 85);
    const cached = await fixture.request({ cached: '1' });
    assert.equal(cached.payload.fresh, true);
    assert.equal(fixture.flowCalls(), 1);
    const operational = await fixture.request({}, { role: 'user', permissions: { salgordreVia: true } });
    assert.equal(operational.payload.scope, 'production');
    assert.equal(operational.payload.orderBacklogValueDkk, null);
    assert.equal(fixture.flowCalls(), 1);
    assert.equal(fixture.costCalls[1].orderNos, undefined);
    assert.equal(fixture.cache.size, 2);
});

test('forced refresh, in-flight deduplication, database scope and single-row invalidation', async () => {
    const fixture = routeFixture();
    await Promise.all([fixture.request({ force: '1' }), fixture.request({ force: '1' })]);
    assert.equal(fixture.flowCalls(), 1);
    await fixture.request({ force: '1' });
    assert.equal(fixture.flowCalls(), 2);
    const single = await fixture.request({ ordNo: '3', force: '1' });
    assert.deepEqual(single.payload.rows.map(row => row.OrdNo), [3]);
    assert.deepEqual(fixture.costCalls.at(-1).orderNos, [3]);
    assert.equal(fixture.cache.size, 0);
    assert.equal((await fixture.request({ ordNo: 'invalid' })).status, 400);
    const context = fixture.viaContext(true);
    const pending = fixture.getViaPayload(context, true);
    fixture.changeProfile();
    await assert.rejects(pending, /Databaseprofil/);
    assert.notEqual(fixture.viaContext(true).key, context.key);
    assert.equal(fixture.cache.size, 0);
    assert.equal((await fixture.request({ profile: 'production' })).status, 409);
});

function clientFixture(fetch) {
    const nodes = Object.fromEntries(['viaResults', 'viaKpis', 'viaStatus', 'viaSearchInput'].map(id => [id, { innerHTML: '', textContent: '', value: '' }]));
    const alerts = [];
    const source = fs.readFileSync(path.join(__dirname, '../assets/js/via.js'), 'utf8');
    const client = new Function('document', 'localStorage', 'fetch', 'accessGranted', 'authToken', 'loggedUsername', '_settingsActiveId', 'canAccessModule', 'formatNumber', 'escapeHtml', 'alert', source + `
        return { load: loadSalgordreVia, refresh: refreshSalgordreViaOrder, validate: validateSalgordreViaPayload,
            render: renderSalgordreVia, visible: getSalgordreViaVisibleRows,
            logout: () => { accessGranted = false; resetSalgordreVia(); },
            state: () => ({ rows: salgordreViaRows, state: salgordreViaLoadState, total: salgordreViaOrderBacklogValue, meta: salgordreViaMeta }) };
    `)({ getElementById: id => nodes[id] || null, querySelector: () => null }, { getItem: () => null }, fetch, true, 'test', 'test', 'production', () => true, String, String, value => alerts.push(value));
    return { ...client, nodes, alerts };
}

test('client restores rows with visible errors, rejects null totals, and never replaces an order with a wrong response', async () => {
    let payload = buildViaBacklog(backlogFixture(), [costFixture(1), costFixture(3)]);
    let failure = false;
    const client = clientFixture(async () => {
        if (failure) throw new Error('offline');
        return { ok: true, json: async () => payload };
    });
    await client.load(true, { loadReservations: false });
    assert.equal(client.state().total, 85);
    assert.match(client.nodes.viaResults.innerHTML, /Restsalgsværdi/);
    client.nodes.viaSearchInput.value = '3';
    client.render();
    assert.deepEqual(client.visible().map(row => row.OrdNo), [3]);
    assert.match(client.nodes.viaKpis.innerHTML, /25 DKK/);
    client.nodes.viaSearchInput.value = '';
    failure = true;
    await client.load(true, { loadReservations: false });
    assert.equal(client.state().state, 'error');
    assert.match(client.nodes.viaResults.innerHTML, /<tbody>/);
    assert.match(client.nodes.viaStatus.textContent, /viser tidligere data/);
    assert.throws(() => client.validate({ ...payload, orderBacklogValueDkk: null }), /stemmer ikke/);
    failure = false;
    payload = buildViaBacklog(backlogFixture(), [costFixture(3)], 3);
    await client.refresh(1);
    assert.match(client.alerts[0], /matcher ikke/);
    assert.deepEqual(client.state().rows.map(row => row.OrdNo), [1, 3]);
});

test('client applies metadata-only updates and discards responses after logout', async () => {
    const payload = buildViaBacklog(backlogFixture(), [costFixture(1), costFixture(3)]);
    const client = clientFixture(async url => ({ ok: true, json: async () => url.includes('cached=1') ? { ...payload, fresh: false, unknownCount: 1 } : { ...payload, unknownCount: 7 } }));
    await client.load(false, { loadReservations: false });
    assert.equal(client.state().meta.unknownCount, 7);
    assert.match(client.nodes.viaStatus.textContent, /7 ordrer udeladt/);
    let resolve;
    const late = clientFixture(() => new Promise(done => { resolve = done; }));
    const pending = late.load(true, { loadReservations: false });
    late.logout();
    resolve({ ok: true, json: async () => payload });
    await pending;
    assert.deepEqual(late.state().rows, []);
    assert.equal(late.state().state, 'idle');
});