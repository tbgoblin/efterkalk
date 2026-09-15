const test = require('node:test');
const assert = require('node:assert/strict');
const dashboard = require('../assets/js/dashboard');

test('personal themes are opt-in and survive preference normalization', () => {
    assert.equal(dashboard.normalizeConfig({}).theme, 'light');
    for (const theme of ['light', 'dark', 'system']) {
        const config = dashboard.normalizeConfig({ theme, active: 'sales', boards: [{ id: 'custom-theme', widgets: ['invoice-kpi'] }] });
        assert.equal(config.theme, theme);
        assert.deepEqual(dashboard.normalizeConfig(config), config);
        assert.equal(config.boards[0].widgets[0], 'invoice-kpi');
    }
    for (const theme of ['invalid', null, {}, true]) assert.equal(dashboard.normalizeConfig({ theme }).theme, 'light');
});

test('dark palette preserves status meaning with readable foregrounds', () => {
    const { colorRole, palette } = require('../assets/js/theme');
    assert.equal(colorRole(255, 255, 255, 'background'), 'surface');
    assert.equal(colorRole(27, 94, 32, 'text'), 'green');
    assert.equal(colorRole(139, 0, 0, 'text'), 'red');
    assert.equal(colorRole(116, 73, 0, 'text'), 'amber');
    assert.equal(colorRole(15, 53, 96, 'text'), 'text');
    assert.equal(colorRole(21, 101, 192, 'background'), null);
    const luminance = hex => {
        const values = hex.slice(1).match(/../g).map(value => parseInt(value, 16) / 255).map(value => value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4);
        return values[0] * 0.2126 + values[1] * 0.7152 + values[2] * 0.0722;
    };
    for (const foreground of ['text', 'muted', 'blue', 'green', 'red', 'amber', 'violet']) {
        for (const background of ['surface', 'raised', 'hover', 'blueSurface', 'greenSurface', 'redSurface', 'amberSurface', 'violetSurface']) {
            const ratio = (luminance(palette[foreground]) + 0.05) / (luminance(palette[background]) + 0.05);
            assert.ok(ratio >= 4.5, foreground + '/' + background + ': ' + ratio);
        }
    }
});

test('personal theme uses the serialized GOH store without losing boards', async () => {
    const writes = [];
    let loaded;
    const store = dashboard.createPreferenceStore({ schemaVersion: 3, cache() {}, onStatus() {}, onConfig(value) { loaded = value; },
        load: async () => ({ schemaVersion: 3, version: 4, config: { theme: 'dark', active: 'custom-theme', boards: [{ id: 'custom-theme', widgets: ['load-kpi'] }] } }),
        save: async (config, version) => { writes.push({ config, version }); return { version: version + 1 }; } });
    await store.load();
    assert.equal(loaded.theme, 'dark');
    await store.set({ ...loaded, theme: 'system' });
    assert.equal(writes[0].config.theme, 'system');
    assert.equal(writes[0].config.boards[0].id, 'custom-theme');
    assert.equal(writes[0].version, 4);
    store.dispose();
});

test('fiscal invoice query restricts every scope and amount mode to sales transactions', async () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const source = fs.readFileSync(path.join(__dirname, '../routes/apiRoutes.js'), 'utf8');
    const start = source.indexOf("    router.get('/efterkalk/customer-invoices',");
    const end = source.indexOf("    router.post('/efterkalk/month-snapshot',", start);
    assert.ok(start >= 0 && end > start);
    let handler;
    let queryText;
    let parameters;
    const request = {
        input(name, _type, value) { parameters[name] = value; return this; },
        async query(text) { queryText = text; return { recordset: [] }; }
    };
    new Function('router', 'getConnection', 'sql', 'logEvent', source.slice(start, end))(
        { get(_route, callback) { handler = callback; } },
        async () => ({ request: () => request }),
        { Int: 'Int', Bit: 'Bit' },
        () => {}
    );
    for (const scope of [{ scope: 'all' }, { custno: '123' }]) {
        for (const includeAllAmounts of ['0', '1']) {
            parameters = {};
            let payload;
            const response = { json(value) { payload = value; }, status() { return this; } };
            await handler({ query: { ...scope, includeAllAmounts, from: '2026-07-01', to: '2026-09-15' } }, response);
            assert.equal(payload.ok, true);
            assert.match(queryText, /WHERE\s+\(@custNo IS NULL OR O\.CustNo = @custNo\)\s+AND O\.TrTp = 1\s+AND O\.InvoNo/);
            assert.match(queryText, /AND \(@includeAllAmounts = 1 OR O\.InvoAm > 0\)/);
            assert.match(queryText, /AND O\.LstInvDt >= @fromDate\s+AND O\.LstInvDt <= @toDate/);
            assert.doesNotMatch(queryText, /SELECT\s+TOP\s+\d+\s+O\.OrdNo/i);
            assert.equal(parameters.includeAllAmounts, Number(includeAllAmounts));
            assert.equal(parameters.fromDate, 20260701);
            assert.equal(parameters.toDate, 20260915);
            assert.equal(parameters.custNo, scope.scope === 'all' ? null : 123);
        }
    }
});

test('order margin endpoint delegates to the exclusion-aware cache owner instead of returning unvalidated disk costs', async () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const source = fs.readFileSync(path.join(__dirname, '../routes/apiRoutes.js'), 'utf8');
    const start = source.indexOf("    router.get('/order-margin/:ordno',");
    const end = source.indexOf("    router.get('/production-summary/:ordno',", start);
    assert.ok(start >= 0 && end > start);
    let handler;
    let calls = 0;
    new Function('router', 'getOrComputeOrderMargin', 'diskCache', 'logEvent', source.slice(start, end))(
        { get(_path, callback) { handler = callback; } },
        async ordNo => {
            calls++;
            assert.equal(ordNo, 407940);
            return { ordNo, totalRevenue: 5718.8, totalCost: 4812.49, styklisteFallbackCost: 0 };
        },
        { get() { throw new Error('Unvalidated disk read'); }, set() { throw new Error('Cache owner metadata overwritten'); } },
        () => {}
    );
    let payload;
    let statusCode = 200;
    const response = { json(value) { payload = value; }, status(value) { statusCode = value; return this; } };
    await handler({ params: { ordno: '407940' } }, response);
    assert.equal(statusCode, 200);
    assert.equal(calls, 1);
    assert.equal(payload.totalCost, 4812.49);
    assert.ok(Math.abs(payload.totalRevenue - payload.totalCost - 906.31) < 0.001);
    await handler({ params: { ordno: 'invalid' } }, response);
    assert.equal(statusCode, 400);
    assert.equal(calls, 1);
});

test('cached margins are checked against current aftercalc before reuse without querying Visma', async () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const source = fs.readFileSync(path.join(__dirname, '../server.js'), 'utf8');
    const start = source.indexOf('async function getOrComputeOrderMargin(');
    const end = source.indexOf('\nfunction warmMarginsInBackground(', start);
    assert.ok(start >= 0 && end > start);
    const old = { ordNo: 407940, totalRevenue: 5718.8, totalCost: 7848.49, styklisteFallbackCost: 0, costExclusionFingerprint: '' };
    const memory = new Map([[407940, old]]);
    let disk = old;
    let writes = 0;
    const current = { summary: { totalRevenue: 5718.8, totalCost: 4812.49, styklisteFallbackCost: 0 }, salesOrderLines: [] };
    const getMargin = new Function('aftercalcCostExclusionsService', 'getAftercalcCacheWithFallback', 'calculateAdjustedCost',
        'orderMarginCache', 'diskCache', 'ORDER_MARGIN_CACHE_KEY_PREFIX', 'orderMarginInFlight', 'getOrComputeAftercalc',
        source.slice(start, end) + '\nreturn getOrComputeOrderMargin;')(
        { ensureHydrated: async () => {}, getFingerprintSync: () => '', getExcludedLineKeysSync: () => new Set() },
        () => current, require('../assets/js/aftercalc-cost-exclusions').calculateAdjustedCost,
        memory, { get: () => disk, getStale: () => disk, set(_key, value) { disk = value; writes++; } },
        'margin_', new Map(), async () => { throw new Error('Unexpected Visma calculation'); }
    );
    const result = await getMargin(407940);
    assert.equal(result.totalCost, 4812.49);
    assert.equal(memory.get(407940).totalCost, 4812.49);
    assert.equal(disk.totalCost, 4812.49);
    assert.equal(writes, 1);
    await getMargin(407940);
    assert.equal(writes, 1);
    current.summary.totalCost = 4500;
    current.summary.styklisteFallbackCost = 10;
    const updated = await getMargin(407940);
    assert.equal(updated.totalCost, 4500);
    assert.equal(updated.styklisteFallbackCost, 10);
});

test('generated dashboard shell keeps the login and source bridge valid JavaScript', () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const source = fs.readFileSync(path.join(__dirname, '../server.js'), 'utf8');
    const start = source.indexOf('res.send(`', source.indexOf("app.get('/',")) + 'res.send('.length;
    const end = source.indexOf('\n    `);', start) + '\n    `'.length;
    assert.ok(start > 0 && end > start);
    const html = new Function('pkgVersion', 'APP_VERSION', 'return ' + source.slice(start, end))('test', 'GOH test');
    const inline = html.slice(html.indexOf('<script>') + '<script>'.length, html.lastIndexOf('</script>'));
    assert.doesNotThrow(() => new Function(inline));
    assert.ok(inline.includes('function ensureDashboardSources('));
    assert.ok(inline.includes('async function submitAccessCode('));
    assert.ok(html.includes('/assets/js/dashboard.js?v=test-13'));
    assert.ok(html.includes('/assets/js/theme.js?v=test-1'));
    assert.ok(html.includes('name="gohTheme" value="dark"'));
    assert.ok(html.includes('/assets/js/order-flow.js?v=test-4'));
    assert.ok(inline.includes("getOrderFlowRequest(month, [])"));
    assert.ok(inline.includes("scope: flowScope, month, today, fixedMonth: true"));
    const bridge = inline.slice(inline.indexOf('function renderDashboardWidgets()'), inline.indexOf('function scheduleDashboardWidgets()'));
    assert.ok(!bridge.includes('orderListData'));
    assert.ok(!bridge.includes('marginStateByOrdNo'));
    assert.ok(bridge.includes('invoices.payload.rows'));
    assert.ok(bridge.includes('notes: orderNotesCache'));
    assert.ok(bridge.includes('economicPeriod: dashboardEconomicPeriod'));
    assert.ok(bridge.includes('ensureDashboardSources(modules, today, widgetIds, period, query, options)'));
    assert.ok(bridge.includes('refreshSources: (widgetIds, options) => refreshDashboardSources(widgetIds, today, options)'));
    assert.ok(inline.includes('payload.rows.filter(row => !window.GohDashboard.model.isCreditOrder(row, orderNotesCache))'));
    assert.ok(inline.includes("includeAllAmounts: '1'"));
    assert.ok(inline.includes("['sales-only-v1', from, today]"));
    assert.ok(inline.includes("const filters = { today, dage: days, resGr: '', ord: '', kunde: '' };"));
    assert.ok(inline.includes('days = 20'));
    assert.ok(inline.includes("['best', 'risk', 'coverage', 'sellers'].includes(id)"));
    assert.ok(inline.includes('synchronizeDashboardOrderMargin(detailOrdNo, marginStateByOrdNo[detailOrdNo])'));
});

test('interactive order flow uses a separate refresh source and economic permissions', async () => {
    assert.equal(dashboard.allowedWidgets(['order-flow'], () => false).length, 0);
    assert.equal(dashboard.allowedWidgets(['order-flow'], key => key === 'omsaetning').length, 1);
    const calls = [];
    await dashboard.refreshVisibleSources(['order-flow', 'invoice-kpi', 'order-flow'], key => key === 'omsaetning', async key => calls.push(key));
    assert.deepEqual(calls, ['order-flow', 'omsaetning']);
    const board = dashboard.normalizeConfig({ boards: [{ id: 'custom-flow', name: 'Flow', widgets: ['order-flow'] }], active: 'custom-flow' });
    assert.deepEqual(board.boards[0].widgets, ['order-flow']);
});

test('calendar source requests and manual refresh use January dates and separate cache keys', async () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const source = fs.readFileSync(path.join(__dirname, '../server.js'), 'utf8');
    const start = source.indexOf('            function getDashboardSourceRequest(');
    const end = source.indexOf('            function ensureDashboardSources(', start);
    assert.ok(start >= 0 && end > start);
    const requests = [];
    const summaries = [];
    const requestSource = new Function('window', 'dashboardSourceKey', 'fetch', 'OMSAETNING_SSRS_DEFAULT_ACCOUNTS',
        'buildOmsaetningSummaryCacheKey', 'fetchOmsaetningSummaryCached', 'belastningSummaryKey', 'requestBelastningSummary', 'getDashboardSourceCache',
        'const dashboardEconomicPeriod = "all"; const dashboardLoadQuery = "";\n' + source.slice(start, end) + '\nreturn getDashboardSourceRequest;')(
        { GohDashboard: { model: dashboard } }, (moduleKey, values) => JSON.stringify([moduleKey, values]),
        async url => { requests.push(url); return { ok: true, json: async () => ({ ok: true, rows: [] }) }; },
        new Set(['11012']), (...values) => JSON.stringify(values),
        async (...values) => { summaries.push(values); return { ok: true, rows: [] }; },
        filters => JSON.stringify(filters), async filters => filters, () => ({ peek: () => null })
    );
    const today = '2026-09-15';
    const fiscalOrders = requestSource('efterkalk', today, false, 'all');
    const calendarOrders = requestSource('efterkalk', today, false, 'year');
    assert.notEqual(calendarOrders.key, fiscalOrders.key);
    await calendarOrders.loader();
    const url = new URL(requests[0], 'http://localhost');
    assert.equal(url.searchParams.get('from'), '2026-01-01');
    assert.equal(url.searchParams.get('to'), today);
    assert.equal(url.searchParams.get('scope'), 'all');
    const calendarSummary = requestSource('omsaetning', today, true, 'year');
    assert.notEqual(calendarSummary.key, requestSource('omsaetning', today, false, 'all').key);
    await calendarSummary.loader();
    assert.deepEqual(summaries[0], ['202507', '202607', ['11012'], [], { forceRefresh: true }]);
    const load = requestSource('belastning', today, false, 'year');
    assert.equal(load.key, requestSource('belastning', today, false, 'all').key);
    assert.equal((await load.loader()).dage, 20);
    const weeklyLoad = requestSource('belastning', today, false, 'year', 7);
    assert.notEqual(weeklyLoad.key, load.key);
    assert.equal((await weeklyLoad.loader()).dage, 7);
});

test('quarter KPI months reconcile signed accounting revenue and do not invent future zeros', () => {
    const result = dashboard.buildData({ today: '2026-09-15', period: 'quarter', revenue: { rows: [
        { date: '2026-07-01', revenueMio: 4, custNo: 1, customerName: 'A' },
        { date: '2026-08-01', revenueMio: 5, custNo: 1, customerName: 'A' },
        { date: '2026-08-01', revenueMio: -0.2, custNo: 2, customerName: 'B' },
        { date: '2026-09-01', revenueMio: 3, custNo: 1, customerName: 'A' }
    ] } }).widgets['invoice-kpi'];
    assert.deepEqual(result.slice(1, 4).map(row => row.label), ['juli', 'august', 'september']);
    assert.equal(result.slice(1, 4).reduce((total, row) => total + row.value, 0), result[0].value);
    assert.equal(result[0].value, 11800000);
    const future = dashboard.buildData({ today: '2026-07-15', period: 'quarter' }).widgets['invoice-kpi'];
    assert.equal(future[1].value, 0);
    assert.equal(future[2].value, null);
    assert.equal(future[3].value, null);
});

test('customer shares group the remainder, preserve signed balances and count distinct invoice orders', async () => {
    const revenue = [100, 80, 60, 40, -10].map((value, index) => ({ customerId: String(index), customer: 'Customer ' + index, revenue: value }));
    const share = dashboard.customerShares(revenue, [], 'revenue', 3);
    assert.equal(share.segments.length, 4);
    assert.equal(share.segments.at(-1).label, 'Andre');
    assert.equal(share.segments.at(-1).value, 40);
    assert.equal(share.segments.reduce((total, row) => total + row.value, 0), share.positiveTotal);
    assert.equal(share.positiveTotal, 280);
    assert.equal(share.negativeTotal, -10);
    assert.equal(share.total, 270);
    const orders = dashboard.customerShares([], [{ order: 1, customerId: '1', customer: 'A' }, { order: 1, customerId: '1', customer: 'A' }, { order: 2, customerId: '1', customer: 'A' }], 'orders');
    assert.equal(orders.total, 2);
    const sources = [];
    await dashboard.refreshVisibleSources(['customer-share'], () => true, async key => sources.push(key), { 'customer-share': { metric: 'orders' } });
    assert.deepEqual(sources, ['efterkalk']);
    assert.equal(dashboard.allowedWidgets(['recent-orders'], key => key === 'omsaetning').length, 0);
    assert.equal(dashboard.allowedWidgets(['recent-orders'], key => ['omsaetning', 'ordreindgang'].includes(key)).length, 1);
});

test('per-widget options survive normalization and independently bound workload horizons and backlog', () => {
    const config = dashboard.normalizeConfig({ active: 'custom-options', boards: [{ id: 'custom-options', widgets: ['load-kpi', 'load-days'], options: {
        'load-kpi': { days: 7, includeBacklog: false }, 'load-days': { days: 999, title: 'Plan' }, unknown: { days: 5 }
    } }] });
    assert.deepEqual(dashboard.loadHorizons(config.boards[0].widgets, config.boards[0].options), [7, 90]);
    assert.equal(config.boards[0].options.unknown, undefined);
    assert.equal(config.boards[0].options['load-days'].title, 'Plan');
    assert.deepEqual(dashboard.normalizeConfig(config), config);
    const result = dashboard.buildData({ today: '2026-09-15', options: { includeBacklog: false }, load: { odd: { rows: [
        { ResGr: '21', Nm: 'Buk', Dato: null, Resv: 600, Kap: 0 }, { ResGr: '21', Nm: 'Buk', Dato: 20260916, Resv: 120, Kap: 480 }
    ] } } });
    assert.equal(result.widgets['load-kpi'][0].value, 2);
    assert.equal(result.widgets['load-kpi'][1].value, 8);
});

test('recent orders use order dates, stable latest sorting, local filters and credit exclusions', () => {
    const rows = dashboard.buildData({ today: '2026-09-15', widgetId: 'recent-orders', options: { days: 7, search: 'A' }, recentOrders: [
        { OrdNo: 1, OrderDate: 20260914, CustomerName: 'A', OrderValueDkk: 200 },
        { OrdNo: 2, OrderDate: 20260915, CustomerName: 'A', OrderValueDkk: 100 },
        { OrdNo: 3, OrderDate: 20260801, CustomerName: 'A', OrderValueDkk: 100 },
        { OrdNo: 4, OrderDate: 20260915, CustomerName: 'A', OrderValueDkk: -20 },
        { OrdNo: 5, OrderDate: 20260915, CustomerName: 'B', OrderValueDkk: 100 },
        { OrdNo: 6, OrderDate: 20260915, CustomerName: 'A', OrderValueDkk: 0 }
    ], notes: { 6: { isCreditNote: true } } }).widgets['recent-orders'];
    assert.deepEqual(rows.map(row => row.order), [2, 1]);
});

test('recent order service makes one bounded sales-only query against the active database', async () => {
    const { createOrdreindgangService } = require('../services/ordreindgangService');
    let queryText;
    const parameters = {};
    let queries = 0;
    const request = {
        input(name, _type, value) { parameters[name] = value; return this; },
        async query(value) { queries++; queryText = value; return { recordset: [{ OrdNo: 10, OrderDate: 20260915, OrderValueDkk: 150 }] }; }
    };
    const service = createOrdreindgangService({ getConnection: async () => ({ request: () => request }), sql: { Int: 'Int' } });
    const result = await service.getRecentOrders({ from: '2026-09-01', to: '2026-09-15' });
    assert.equal(result.rows.length, 1);
    assert.deepEqual(parameters, { fromDate: 20260901, toDate: 20260915 });
    assert.match(queryText, /O\.TrTp = 1 AND O\.OrdTp = 1/);
    assert.match(queryText, /O\.OrdDt >= @fromDate AND O\.OrdDt <= @toDate/);
    assert.match(queryText, /\(O\.InvoSF \+ O\.InvoIF\) \* \(O\.ExRt \/ 100\.0\)/);
    assert.match(queryText, /MAX\(LTRIM\(RTRIM\(A\.Nm\)\)\)/);
    assert.match(queryText, /O\.CustNo > 0/);
    assert.doesNotMatch(queryText, /F0001|LstInvDt|SELECT TOP/);
    await assert.rejects(service.getRecentOrders({ from: '2026-01-01', to: '2026-09-15' }), { statusCode: 400 });
    await assert.rejects(service.getRecentOrders({ from: '2026-02-30', to: '2026-03-01' }), { statusCode: 400 });
    assert.equal(queries, 1);
});

test('widget preferences cannot be silently saved by a backend without the required schema', async () => {
    let status;
    let writes = 0;
    const store = dashboard.createPreferenceStore({ schemaVersion: 2, cache() {}, onConfig() {}, onStatus(value) { status = value; },
        load: async () => ({ version: 1, config: { active: 'sales' } }), save: async () => { writes++; return { version: 2 }; } });
    await store.load();
    await store.set({ active: 'sales' });
    assert.equal(status, 'error');
    assert.equal(writes, 0);
    store.dispose();
});

test('closing an order returns to its origin without changing report back navigation', () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const source = fs.readFileSync(path.join(__dirname, '../server.js'), 'utf8');
    const openStart = source.indexOf('            function openDashboardOrder(');
    const closeStart = source.indexOf('            function closeOrderDetailModal(');
    const openSource = source.slice(openStart, source.indexOf('\n            //', openStart));
    const closeSource = source.slice(closeStart, source.indexOf('            function buildStandaloneReportPrintCss(', closeStart));
    const elements = {
        orderInput: { value: '' }, mainWorkspace: { style: { display: 'none' } },
        orderDetailModal: { style: { display: 'flex' } }, orderDetailModalBody: { innerHTML: 'order' }
    };
    const destinations = [];
    const selected = [];
    let allowed = true;
    let ready = true;
    let inReport = false;
    let restored = 0;
    const navigation = new Function('document', 'canAccessModule', 'openModule', 'selectOrder', 'goToDashboard', 'goBackToList',
        'isOrderDetailReportViewActive', 'restoreOrderDetailFromReport', 'updateOrderDetailModalBackButton', 'removeModalStack',
        'let orderDetailReturnModule = null; let reportOriginState = null;\n' + openSource + closeSource +
        '\nreturn { openDashboardOrder, closeOrderDetailModal, origin: () => orderDetailReturnModule, setOrigin: value => { orderDetailReturnModule = value; }, setReport: value => { reportOriginState = value; } };')(
        { getElementById: id => elements[id], body: { classList: { remove() {} } } },
        () => allowed,
        moduleKey => { destinations.push(moduleKey); if (ready && moduleKey === 'efterkalk') elements.mainWorkspace.style.display = 'block'; },
        ordNo => selected.push(ordNo), () => destinations.push('dashboard'), () => destinations.push('list'),
        () => inReport, () => { restored++; }, () => {}, () => {}
    );
    navigation.openDashboardOrder(412385);
    assert.equal(navigation.origin(), 'dashboard');
    assert.deepEqual(selected, [412385]);
    navigation.closeOrderDetailModal({ target: { id: 'orderDetailModalBody' } });
    assert.equal(navigation.origin(), 'dashboard');
    inReport = true;
    navigation.setReport({});
    navigation.closeOrderDetailModal();
    assert.equal(restored, 1);
    assert.equal(navigation.origin(), 'dashboard');
    inReport = false;
    navigation.closeOrderDetailModal();
    assert.equal(destinations.at(-1), 'dashboard');
    assert.equal(navigation.origin(), null);
    assert.equal(elements.orderDetailModal.style.display, 'none');
    navigation.closeOrderDetailModal();
    assert.equal(destinations.at(-1), 'list');
    navigation.setOrigin('salgordre-via');
    navigation.closeOrderDetailModal({ target: { id: 'orderDetailModal' } });
    assert.equal(destinations.at(-1), 'salgordre-via');
    assert.equal(navigation.origin(), null);
    elements.mainWorkspace.style.display = 'none';
    ready = false;
    navigation.openDashboardOrder(100);
    assert.equal(navigation.origin(), null);
    allowed = false;
    ready = true;
    navigation.openDashboardOrder(200);
    assert.deepEqual(selected, [412385]);
});

test('manual dashboard refresh loads only visible authorized sources once and reports partial failures', async () => {
    const calls = [];
    const failed = await dashboard.refreshVisibleSources(['best', 'risk', 'trend', 'load-days', 'load-resources'], () => true, async moduleKey => {
        calls.push(moduleKey);
        if (moduleKey === 'omsaetning') throw new Error('Offline');
    });
    assert.deepEqual(calls, ['efterkalk', 'omsaetning', 'belastning']);
    assert.deepEqual(failed, ['omsaetning']);
    calls.length = 0;
    assert.deepEqual(await dashboard.refreshVisibleSources(['best', 'trend', 'load-days', 'via-due'], moduleKey => moduleKey !== 'omsaetning', async moduleKey => calls.push(moduleKey)), []);
    assert.deepEqual(calls, ['belastning', 'salgordre-via']);
    calls.length = 0;
    await dashboard.refreshVisibleSources([], () => true, async moduleKey => calls.push(moduleKey));
    assert.deepEqual(calls, []);
});

test('manual refresh bypasses fresh cache and retry backoff while sharing pending requests and retaining stale data', async () => {
    const cache = dashboard.createSourceCache();
    cache.put('orders', { rows: [1] });
    cache.put('other-user', { rows: [9] });
    let complete;
    let calls = 0;
    const loader = () => { calls++; return new Promise(resolve => { complete = resolve; }); };
    const first = cache.load('orders', loader, true);
    const second = cache.load('orders', loader, true);
    assert.equal(first, second);
    await Promise.resolve();
    assert.equal(calls, 1);
    assert.deepEqual(cache.peek('orders').payload.rows, [1]);
    complete({ rows: [2] });
    await first;
    await assert.rejects(cache.load('orders', async () => { throw new Error('Offline'); }, true));
    assert.deepEqual(cache.peek('orders').payload.rows, [2]);
    await cache.load('orders', async () => ({ rows: [3] }), true);
    assert.deepEqual(cache.peek('orders').payload.rows, [3]);
    assert.deepEqual(cache.peek('other-user').payload.rows, [9]);
});

test('all fiscal orders contribute to coverage and rankings beyond the old 150-row limit', () => {
    const orders = Array.from({ length: 527 }, (_, index) => ({ OrdNo: index + 1, LstInvDt: 20260706, InvoAm: 1000, TotalCost: 950, SellerUsr: 'MW' }));
    orders[526].TotalCost = 0;
    orders[525].InvoAm = -200;
    orders[524].InvoAm = 0;
    const result = dashboard.buildData({ today: '2026-09-15', orders });
    assert.equal(result.orders.length, 526);
    assert.equal(result.widgets.coverage[0].value, 526);
    assert.equal(result.widgets.best[0].order, 527);
    assert.equal(result.widgets.risk[0].order, 525);
    assert.equal(result.widgets.sellers[0].value, 525000);
    assert.equal(result.widgets.latest.length, 5);
});

test('dashboard excludes negative and marked credit orders but retains genuine sales losses and accounting credits', () => {
    const orders = [
        { OrdNo: 1, LstInvDt: 20260911, InvoAm: -5779, TotalCost: 0, SellerUsr: 'TB' },
        { OrdNo: 2, LstInvDt: 20260911, InvoAm: 100, TotalCost: 150, SellerUsr: 'TB' },
        { OrdNo: 3, LstInvDt: 20260911, InvoAm: 200, TotalCost: 300, SellerUsr: 'TB' },
        { OrdNo: 4, LstInvDt: 20260911, InvoAm: 0, TotalCost: 0, SellerUsr: 'TB' },
        { OrdNo: 5, LstInvDt: 20260911, InvoAm: 0, TotalCost: 801.92, SellerUsr: 'MW' }
    ];
    const notes = { 3: { isCreditNote: true }, 5: { isCreditNote: true } };
    const before = JSON.stringify(orders);
    const result = dashboard.buildData({ today: '2026-09-15', orders, notes, revenue: { rows: [
        { date: '2026-09-01', revenueMio: -0.001 }
    ] } });
    assert.deepEqual(result.orders.map(row => row.order), [2, 4]);
    assert.deepEqual(result.widgets.latest.map(row => row.order), [4, 2]);
    assert.equal(result.widgets.risk[0].order, 2);
    assert.equal(result.widgets.risk[0].value, -50);
    assert.equal(result.widgets.coverage[0].value, 2);
    assert.equal(result.widgets.sellers[0].value, 100);
    assert.equal(result.widgets['invoice-kpi'][0].value, -1000);
    assert.equal(JSON.stringify(orders), before);
    notes[3].isCreditNote = false;
    assert.deepEqual(dashboard.buildData({ today: '2026-09-15', orders, notes }).orders.map(row => row.order), [2, 3, 4]);
});

test('seller DB percentages are revenue-weighted, preserve losses and disclose missing costs', () => {
    const orders = [
        { OrdNo: 1, LstInvDt: 20260901, SellerUsr: 'TB', InvoAm: 100, TotalCost: 0 },
        { OrdNo: 2, LstInvDt: 20260901, SellerUsr: 'TB', InvoAm: 900, TotalCost: 900 },
        { OrdNo: 3, LstInvDt: 20260901, SellerUsr: 'TB', InvoAm: 1000, TotalCost: null },
        { OrdNo: 4, LstInvDt: 20260901, SellerUsr: 'MW', InvoAm: 200, TotalCost: 300 },
        { OrdNo: 5, LstInvDt: 20260901, SellerUsr: 'ZERO', InvoAm: 0, TotalCost: 25 },
        { OrdNo: 6, LstInvDt: 20260901, SellerUsr: 'PENDING', InvoAm: 100, TotalCost: null },
        { OrdNo: 7, LstInvDt: 20260901, SellerUsr: 'TB', InvoAm: -100, TotalCost: 0 },
        { OrdNo: 8, LstInvDt: 20260901, SellerUsr: 'TB', InvoAm: 10000, TotalCost: 0 }
    ];
    const input = { today: '2026-09-15', orders, notes: { 8: { isCreditNote: true } } };
    const result = dashboard.buildData(input);
    const seller = result.widgets.sellers[0];
    assert.equal(seller.label, 'TB');
    assert.equal(seller.value, 2000);
    assert.equal(seller.dbPercent, 10);
    assert.equal(seller.costCount, 2);
    assert.equal(seller.count, 3);
    assert.equal(result.widgets.sellers.find(row => row.label === 'MW').dbPercent, -50);
    assert.equal(result.widgets.sellers.find(row => row.label === 'ZERO').dbPercent, null);
    assert.equal(result.widgets.sellers.find(row => row.label === 'PENDING').dbPercent, null);
    const completed = dashboard.buildData({ ...input, margins: { 3: { status: 'success', totalCost: 500 } } });
    assert.equal(completed.widgets.sellers[0].dbPercent, 30);
    assert.equal(completed.widgets.sellers[0].costCount, 3);
    assert.equal(dashboard.buildData({ ...input, query: 'MW' }).widgets.sellers[0].dbPercent, -50);
});

test('fiscal order costs use at most three workers and preserve known zero and missing costs', async () => {
    const rows = Array.from({ length: 527 }, (_, index) => ({ OrdNo: index + 1, TotalCost: index === 0 ? 0 : null }));
    let active = 0;
    let maximum = 0;
    let calls = 0;
    await dashboard.hydrateOrderCosts(rows, async ordNo => {
        active++; calls++; maximum = Math.max(maximum, active);
        await Promise.resolve();
        active--;
        if (ordNo === 526) throw new Error('Offline');
        return { totalCost: ordNo === 527 ? null : ordNo, styklisteFallbackCost: 10 };
    }, () => true, () => {});
    assert.equal(maximum, 3);
    assert.equal(calls, 526);
    assert.equal(rows[0].TotalCost, 0);
    assert.equal(rows[1].TotalCost, 12);
    assert.equal(rows[525].TotalCost, null);
    assert.equal(rows[526].TotalCost, null);
});

test('a late dashboard cost response cannot overwrite a newer order-detail cost', async () => {
    const rows = [{ OrdNo: 407940 }];
    await dashboard.hydrateOrderCosts(rows, async () => {
        rows[0].TotalCost = 4812.49;
        return { totalCost: 7848.49, styklisteFallbackCost: 0 };
    }, () => true, () => {});
    assert.equal(rows[0].TotalCost, 4812.49);
    const result = dashboard.buildData({ today: '2026-09-15', orders: [{ ...rows[0], LstInvDt: 20260813, InvoAm: 5718.8 }] });
    assert.ok(Math.abs(result.widgets.risk[0].value - 906.31) < 0.001);
});

test('fiscal cost hydration stops and ignores responses after leaving its authorized scope', async () => {
    const rows = [{ OrdNo: 1 }, { OrdNo: 2 }, { OrdNo: 3 }, { OrdNo: 4 }];
    let active = true;
    let updates = 0;
    await dashboard.hydrateOrderCosts(rows, async () => {
        active = false;
        return { totalCost: 0 };
    }, () => active, () => { updates++; });
    assert.ok(rows.every(row => row.TotalCost === undefined));
    assert.equal(updates, 0);
});

test('calendar year includes January through June while fiscal year still starts in July', () => {
    const input = { today: '2026-09-15', orders: [
        { OrdNo: 1, LstInvDt: 20251231, InvoAm: 100, TotalCost: 0 },
        { OrdNo: 2, LstInvDt: 20260101, InvoAm: 200, TotalCost: 0 },
        { OrdNo: 3, LstInvDt: 20260630, InvoAm: 300, TotalCost: 0 },
        { OrdNo: 4, LstInvDt: 20260701, InvoAm: 400, TotalCost: 0 },
        { OrdNo: 5, LstInvDt: 20260916, InvoAm: 500, TotalCost: 0 }
    ], revenue: { rows: [
        { date: '2025-12-01', revenueMio: 8 },
        { date: '2026-01-01', revenueMio: 1 },
        { date: '2026-06-01', revenueMio: 2 },
        { date: '2026-07-01', revenueMio: 3 },
        { date: '2026-10-01', revenueMio: 9 }
    ] } };
    const calendar = dashboard.buildData({ ...input, period: 'year' });
    const fiscal = dashboard.buildData({ ...input, period: 'all' });
    assert.deepEqual(dashboard.economicRange(input.today, 'year'), { from: '2026-01-01', fra: '202507', til: '202607' });
    assert.deepEqual(dashboard.economicRange(input.today, 'all'), dashboard.fiscalRange(input.today));
    assert.equal(calendar.periodFrom, '2026-01-01');
    assert.deepEqual(calendar.orders.map(row => row.order), [2, 3, 4]);
    assert.deepEqual(fiscal.orders.map(row => row.order), [4]);
    assert.equal(calendar.widgets['invoice-kpi'][0].value, 6000000);
    assert.equal(fiscal.widgets['invoice-kpi'][0].value, 3000000);
    assert.equal(calendar.widgets.trend.length, 9);
    assert.equal(fiscal.widgets.trend.length, 3);
    const january = dashboard.buildData({ ...input, today: '2027-01-15', period: 'year' });
    assert.equal(january.periodFrom, '2027-01-01');
    assert.equal(january.orders.length, 0);
    assert.equal(january.widgets.trend.length, 1);
    assert.deepEqual(dashboard.economicRange('2027-01-15', 'year'), { from: '2027-01-01', fra: '202607', til: '202707' });
});

test('monthly revenue uses the complete fiscal Omsaetning source rather than the limited order list', () => {
    const result = dashboard.buildData({ today: '2026-09-15', orders: [{ OrdNo: 1, LstInvDt: 20260901, InvoAm: 999 }], revenue: { rows: [
        { date: '2026-06-01', revenueMio: 8 }, { date: '2026-07-01', revenueMio: 2 },
        { date: '2026-08-01', revenueMio: 3 }, { date: '2026-09-01', revenueMio: -0.1 }
    ] } });
    assert.equal(result.fiscalFrom, '2026-07-01');
    assert.deepEqual(result.widgets.trend.map(row => row.value), [2000000, 3000000, -100000]);
    assert.equal(result.widgets['invoice-kpi'][0].value, 4900000);
    assert.deepEqual(dashboard.fiscalRange('2026-09-15'), { from: '2026-07-01', fra: '202601', til: '202701' });
    assert.equal(dashboard.buildData({ today: '2027-01-15' }).fiscalFrom, '2026-07-01');
});

test('fiscal revenue includes zero months, filters customers and never crosses into a previous fiscal year', () => {
    const revenue = { rows: [
        { date: '2026-07-01T00:00:00.000Z', revenueMio: 1, custNo: 1, customerName: 'Alpha' },
        { date: '2026-07-01', revenueMio: 2, custNo: 2, customerName: 'Beta' }
    ] };
    const result = dashboard.buildData({ today: '2026-09-15', revenue, query: 'alpha', period: 'year' });
    assert.deepEqual(result.widgets.trend.map(row => row.value), [1000000, 0, 0]);
    assert.equal(result.widgets.customers[0].label, 'Alpha');
    assert.equal(result.widgets['invoice-kpi'][2].value, 1);
    assert.equal(dashboard.fiscalRange('2027-06-30').fra, '202601');
    assert.equal(dashboard.fiscalRange('2027-07-01').fra, '202701');
    const aftercalc = dashboard.buildData({ today: '2026-09-15', period: 'all', orders: [
        { OrdNo: 1, LstInvDt: 20260630, InvoAm: 500, TotalCost: 0 },
        { OrdNo: 2, LstInvDt: 20260701, InvoAm: 100, TotalCost: 0 }
    ] });
    assert.deepEqual(aftercalc.widgets.best.map(row => row.order), [2]);
});

test('Belastning widgets convert minutes once and separate backlog and evening from daily capacity', () => {
    const load = { toDay: '2026-09-15', dage: 30, odd: { rows: [
        { ResGr: '11', Nm: 'Laser', Dato: '2026-09-15', Resv: 120, Kap: 60, Aften: 30 },
        { ResGr: '11', Nm: 'Laser', Dato: null, Resv: 60, Kap: 0, Aften: 15 }
    ] }, even: { rows: [
        { ResGr: '21', Nm: 'Buk', Dato: '2026-09-15', Resv: 60, Kap: 0, Aften: 0 }
    ] } };
    const before = JSON.stringify(load);
    const result = dashboard.buildData({ today: '2026-09-15', period: 'month', load });
    assert.equal(result.widgets['load-kpi'][0].value, 4);
    assert.equal(result.widgets['load-kpi'][1].value, 1);
    assert.equal(result.widgets['load-kpi'][2].value, 0.75);
    assert.equal(result.widgets['load-kpi'][3].value, 2);
    assert.equal(result.widgets['load-days'][0].value, 3);
    assert.equal(result.widgets['load-days'][0].capacity, 1);
    assert.equal(result.widgets['load-backlog'][0].value, 1.25);
    assert.equal(dashboard.buildData({ today: '2026-09-15', load, query: 'laser' }).widgets['load-kpi'][0].value, 3);
    assert.equal(JSON.stringify(load), before);
});

test('production users cannot see financial templates or monetary widgets even with Efterkalk and VIA access', () => {
    const canAccess = moduleKey => moduleKey !== 'omsaetning';
    const visible = dashboard.allowedWidgets(dashboard.catalog.map(widget => widget.id), canAccess);
    assert.ok(visible.every(widget => widget.module === 'belastning' || ['via-due', 'via-next'].includes(widget.id)));
    assert.deepEqual(dashboard.allowedTemplates(canAccess).map(template => template.id), ['production']);
    assert.deepEqual(dashboard.allowedWidgets(['best', 'risk', 'via-cost', 'trend', 'invoice-kpi'], canAccess), []);
    assert.deepEqual(dashboard.allowedTemplates(() => false), []);
});

test('dashboard sources deduplicate requests, reuse fresh data and isolate scopes', async () => {
    let now = 0;
    let calls = 0;
    const cache = dashboard.createSourceCache(() => now);
    const loader = async () => { calls++; return { ok: true, rows: [] }; };
    await Promise.all([cache.load('MW:production', loader), cache.load('MW:production', loader)]);
    await cache.load('MW:production', loader);
    assert.equal(calls, 1);
    await cache.load('TB:production', loader);
    await cache.load('MW:test', loader);
    assert.equal(calls, 3);
    now = 900001;
    await cache.load('MW:production', loader);
    assert.equal(calls, 4);
    cache.clear();
    assert.equal(cache.peek('MW:production'), undefined);
});

test('dashboard failures retain cached data and back off automatic retries', async () => {
    let now = 0;
    let calls = 0;
    const cache = dashboard.createSourceCache(() => now);
    cache.put('revenue', { ok: true, rows: [1] });
    const loader = async () => { calls++; throw new Error('Offline'); };
    await assert.rejects(cache.load('revenue', loader, true), /Offline/);
    await assert.rejects(cache.load('revenue', loader), /Offline/);
    assert.equal(calls, 1);
    assert.deepEqual(cache.peek('revenue').payload.rows, [1]);
    now = 60001;
    await assert.rejects(cache.load('revenue', loader), /Offline/);
    assert.equal(calls, 2);
});

const input = {
    today: '2026-09-15', period: 'month',
    orders: [
        { OrdNo: 1, LstInvDt: 20260901, InvoAm: 100, TotalCost: null, CustomerName: 'Alpha', SellerUsr: 'MW' },
        { OrdNo: 2, LstInvDt: 20260912, InvoAm: 200, TotalCost: 50, CustomerName: 'Alpha', SellerUsr: 'TB' },
        { OrdNo: 3, LstInvDt: 20260913, InvoAm: -20, TotalCost: 0, CustomerName: 'Beta', SellerUsr: 'MW' },
        { OrdNo: 4, LstInvDt: 20260801, InvoAm: 900, TotalCost: 1, CustomerName: 'Beta' }
    ],
    via: [
        { OrdNo: 5, DeliveryDate: 20260914, MaterialCost: 20, StangCost: 5, PurchasedPartCost: 10, TimeCost: 15 },
        { OrdNo: 6, DeliveryDate: 20260915, TimeCost: 10 },
        { OrdNo: 7, DeliveryDate: 0, MaterialCost: 30 }
    ]
};

test('dashboard DB excludes missing costs and credit notes', () => {
    const result = dashboard.buildData(input);
    assert.equal(result.orders.reduce((total, row) => total + row.revenue, 0), 300);
    assert.equal(result.widgets.best.length, 1);
    assert.equal(result.widgets.best[0].value, 150);
    assert.equal(result.widgets.risk[0].value, 150);
    assert.equal(result.widgets.coverage[1].value, 1);
});

test('dashboard known zero is not missing cost and completed margins override old cache', () => {
    const result = dashboard.buildData({ ...input, margins: { 1: { status: 'success', totalCost: 0 }, 2: { status: 'success', totalCost: 250 } } });
    assert.equal(result.valued.length, 2);
    assert.equal(result.widgets.best[0].value, 100);
    assert.equal(result.widgets.risk[0].value, -50);
});

test('dashboard VIA sums cost categories once and does not inherit invoice period', () => {
    const result = dashboard.buildData(input);
    assert.equal(result.widgets['via-kpi'][0].value, 90);
    assert.equal(result.widgets['via-cost'].reduce((total, row) => total + row.value, 0), 90);
    assert.deepEqual(result.widgets['via-due'].map(row => row.order), [5]);
    assert.deepEqual(result.widgets['via-next'].map(row => row.order), [6]);
    assert.equal(result.via.length, 3);
});

test('dashboard dates reject invalid calendar days and support SQL compact and ISO dates', () => {
    assert.equal(dashboard.dateKey(20260230), '');
    assert.equal(dashboard.dateKey(0), '');
    assert.equal(dashboard.dateKey('2026-09-15T12:00:00'), '2026-09-15');
    assert.equal(dashboard.dateKey(20260915), '2026-09-15');
});

test('dashboard filters Aftercalc orders without mutating source data', () => {
    const before = JSON.stringify(input);
    const result = dashboard.buildData({ ...input, query: 'mw' });
    assert.equal(result.orders.length, 1);
    assert.equal(result.orders.reduce((total, row) => total + row.revenue, 0), 100);
    assert.equal(JSON.stringify(input), before);
});

test('dashboard quarter and year filters have calendar boundaries', () => {
    assert.equal(dashboard.buildData({ ...input, period: 'quarter' }).orders.length, 3);
    assert.equal(dashboard.buildData({ ...input, period: 'year' }).from, '2026-01-01');
    assert.equal(dashboard.buildData({ ...input, period: 'all' }).from, '');
});

test('Belastning search routes customers and orders to the server and keeps resource search local', () => {
    const load = { toDay: '2026-09-15', odd: { rows: [{ ResGr: '21', Nm: 'Buk', Dato: '2026-09-15', Resv: 120, Kap: 480, Aften: 30 }] }, even: { rows: [] } };
    assert.deepEqual(dashboard.loadSearchFilters(' logi ', load), { ord: '', kunde: 'logi', resourceQuery: '' });
    assert.deepEqual(dashboard.loadSearchFilters('409780', load), { ord: '409780', kunde: '', resourceQuery: '' });
    assert.deepEqual(dashboard.loadSearchFilters('BUK', load), { ord: '', kunde: '', resourceQuery: 'BUK' });
    assert.deepEqual(dashboard.loadSearchFilters('', load), { ord: '', kunde: '', resourceQuery: '' });
    const result = dashboard.buildData({ today: load.toDay, query: 'logi', load: { ...load, kunde: 'logi' } });
    assert.equal(result.loadRows.length, 1);
    assert.equal(result.widgets['load-kpi'][0].value, 2);
    assert.equal(result.widgets['load-kpi'][1].value, 8);
    assert.equal(result.widgets['load-kpi'][2].value, 0.5);
    assert.equal(dashboard.buildData({ today: load.toDay, query: '409780', load: { ...load, ord: '409780' } }).loadRows.length, 1);
    assert.equal(dashboard.buildData({ today: load.toDay, query: 'other', load: { ...load, kunde: 'logi' } }).loadRows.length, 0);
});

test('default dashboard packing fills the space below short widgets beside order flow', () => {
    const ids = ['invoice-kpi', 'order-flow', 'best', 'risk', 'coverage'];
    const result = dashboard.arrangeWidgets(ids);
    assert.deepEqual(result['invoice-kpi'], { x: 0, y: 0, w: 4, h: 12 });
    assert.deepEqual(result['order-flow'], { x: 4, y: 0, w: 8, h: 22 });
    assert.deepEqual(result.best, { x: 0, y: 12, w: 4, h: 12 });
    assert.equal(result.risk.y, 22);
});

test('drag and resize preserve the pinned position and compact other widgets without overlaps', () => {
    const ids = dashboard.catalog.map(widget => widget.id);
    let layout = dashboard.arrangeWidgets(ids);
    const before = JSON.stringify(layout);
    const moved = dashboard.arrangeWidgets(ids, { ...layout, best: { x: 4, y: 2, w: 5, h: 10 } }, [], 'best', true);
    assert.equal(JSON.stringify(layout), before);
    assert.deepEqual(moved.best, { x: 4, y: 2, w: 5, h: 10 });
    for (let iteration = 0; iteration < 24; iteration++) {
        const id = ids[iteration % ids.length];
        layout = dashboard.arrangeWidgets(ids, { ...layout, [id]: { x: iteration % 9, y: iteration % 14, w: 3 + iteration % 10, h: 8 + iteration % 25 } }, [], id, true);
        const rects = Object.values(layout);
        for (const [index, rect] of rects.entries()) {
            assert.ok(rect.x >= 0 && rect.x + rect.w <= 12 && rect.y >= 0);
            for (const other of rects.slice(index + 1)) assert.ok(!(rect.x < other.x + other.w && rect.x + rect.w > other.x && rect.y < other.y + other.h && rect.y + rect.h > other.y));
        }
    }
});

test('stored geometry is bounded, permission-neutral and survives legacy wide layouts', () => {
    const normalized = dashboard.normalizeConfig({ active: 'custom-layout', boards: [{ id: 'custom-layout', widgets: ['best', 'order-flow'], wide: ['best'], layout: {
        best: { x: -2, y: -1, w: 50, h: 0 }, 'order-flow': { x: 11, y: 3, w: 1, h: 100 }, unknown: { x: 0, y: 0, w: 4, h: 12 }
    } }] });
    assert.deepEqual(normalized.boards[0].layout.best, { x: 0, y: 0, w: 12, h: 8 });
    assert.deepEqual(normalized.boards[0].layout['order-flow'], { x: 6, y: 3, w: 6, h: 32 });
    assert.equal(normalized.boards[0].layout.unknown, undefined);
    assert.deepEqual(dashboard.normalizeLayout({ best: { x: NaN, y: 0, w: 4, h: 12 } }, ['best']), {});
    assert.equal(dashboard.arrangeWidgets(['best'], {}, ['best']).best.w, 8);
});

test('templates never expose widgets beyond module permissions', () => {
    const allIds = dashboard.catalog.map(widget => widget.id);
    assert.equal(dashboard.allowedWidgets(allIds, () => false).length, 0);
    assert.ok(dashboard.allowedWidgets(allIds, moduleKey => moduleKey === 'salgordre-via').every(widget => widget.module === 'salgordre-via'));
    for (const template of dashboard.templates) assert.equal(dashboard.allowedWidgets(template.widgets, () => true).length, template.widgets.length);
});

test('dashboard preferences normalize damaged input and isolate users and database profiles', () => {
    assert.notEqual(dashboard.storageKey('MW', 'production'), dashboard.storageKey('TB', 'production'));
    assert.notEqual(dashboard.storageKey('MW', 'production'), dashboard.storageKey('MW', 'test'));
    assert.equal(dashboard.storageKey('MW', 'production'), dashboard.storageKey('mw', 'production'));
    const config = dashboard.normalizeConfig({ active: 'invalid', period: 'bad', limit: 500, boards: [null, { id: 'custom-abc', name: 'My dashboard', widgets: ['best', 'best', 'unknown'], wide: ['best'] }, { id: 'custom-abc', widgets: [] }] });
    assert.equal(config.active, 'management');
    assert.equal(config.period, 'all');
    assert.equal(config.limit, 5);
    assert.equal(config.boards.length, 1);
    assert.deepEqual(config.boards[0].widgets, ['best']);
    assert.deepEqual(config.boards[0].wide, ['best']);
});

test('GOH preferences win over legacy local layouts and missing remote profiles migrate once', async () => {
    for (const remoteExists of [true, false]) {
        const writes = [];
        let selected;
        const store = dashboard.createPreferenceStore({
            legacy: { active: 'sales', query: '  logi  ', flow: { month: '2026-08', cohort: 'new' } },
            initial: { active: 'production' }, cache() {}, onStatus() {}, onConfig(value) { selected = value; },
            load: async () => ({ version: remoteExists ? 4 : 0, config: remoteExists ? { active: 'economy' } : null }),
            save: async (config, version) => { writes.push({ config, version }); return { version: version + 1 }; }
        });
        await store.load();
        assert.equal(selected.active, remoteExists ? 'economy' : 'sales');
        assert.equal(writes.length, remoteExists ? 0 : 1);
        if (!remoteExists) {
            assert.equal(writes[0].version, 0);
            assert.equal(writes[0].config.query, 'logi');
            assert.equal(writes[0].config.flow.month, '2026-08');
        }
        store.dispose();
    }
});

test('preference writes serialize and retain edits made during an in-flight save', async () => {
    let release, cached;
    const writes = [];
    const store = dashboard.createPreferenceStore({
        cache(value) { cached = value; }, onConfig() {}, onStatus() {},
        load: async () => ({ version: 6, config: { active: 'production' } }),
        save: async (config, version) => {
            writes.push({ config, version });
            if (writes.length === 1) await new Promise(resolve => { release = resolve; });
            return { version: version + 1 };
        }
    });
    await store.load();
    const saving = store.set({ active: 'sales' });
    store.set({ active: 'economy', limit: 10 });
    assert.equal(writes.length, 1);
    release();
    await saving;
    assert.deepEqual(writes.map(write => write.version), [6, 7]);
    assert.equal(cached.config.active, 'economy');
    assert.equal(cached.config.limit, 10);
    assert.equal(cached.version, 8);
    assert.equal(cached.pending, false);
    store.dispose();
});

test('offline edits survive retry but cannot overwrite a newer remote profile', async () => {
    let status, cached, selected;
    let offline = true;
    let writes = 0;
    const store = dashboard.createPreferenceStore({
        cached: { config: { active: 'production' }, version: 3, pending: false },
        cache(value) { cached = value; }, onConfig(value) { selected = value; }, onStatus(value) { status = value; },
        load: async () => { if (offline) throw new Error('offline'); return { version: 4, config: { active: 'economy' } }; },
        save: async () => { writes++; return { version: 5 }; }
    });
    await store.load();
    assert.equal(status, 'error');
    assert.equal(selected.active, 'production');
    await store.set({ active: 'sales' });
    assert.equal(cached.pending, true);
    offline = false;
    await store.load();
    assert.equal(status, 'conflict');
    assert.equal(writes, 0);
    assert.equal(selected.active, 'sales');
    await store.load(true);
    assert.equal(selected.active, 'economy');
    assert.equal(cached.pending, false);
    store.dispose();
});

test('failed GOH saves stay pending and retry with the original version', async () => {
    let status, cached;
    let fail = true;
    const versions = [];
    const store = dashboard.createPreferenceStore({
        cache(value) { cached = value; }, onConfig() {}, onStatus(value) { status = value; },
        load: async () => ({ version: 2, config: { active: 'production' } }),
        save: async (_config, version) => { versions.push(version); if (fail) throw new Error('offline'); return { version: version + 1 }; }
    });
    await store.load();
    await store.set({ active: 'sales' });
    assert.equal(status, 'error');
    assert.equal(cached.pending, true);
    fail = false;
    await store.load();
    assert.equal(status, 'saved');
    assert.equal(cached.pending, false);
    assert.deepEqual(versions, [2, 2]);
    store.dispose();
});

test('late preference responses cannot change a logged-out or switched user', async () => {
    let finish;
    let callbacks = 0;
    const store = dashboard.createPreferenceStore({
        cache() { callbacks++; }, onConfig() { callbacks++; }, onStatus() {},
        load: () => new Promise(resolve => { finish = resolve; }), save: async () => ({ version: 1 })
    });
    const loading = store.load();
    store.dispose();
    finish({ version: 1, config: { active: 'sales' } });
    await loading;
    assert.equal(callbacks, 0);
});

test('dashboard API authenticates, scopes by canonical user and database, and rejects stale writes', async () => {
    const source = require('node:fs').readFileSync(require('node:path').join(__dirname, '../routes/apiRoutes.js'), 'utf8');
    const start = source.indexOf('    function dashboardPreferenceKey(');
    const end = source.indexOf("    router.get('/admin/users',", start);
    assert.ok(start >= 0 && end > start);
    const handlers = new Map();
    const records = new Map();
    let profile = { id: 'production', server: 'TEST-SERVER', database: 'FactoryOne' };
    let offline = false;
    const gohData = {
        isEnabled: () => !offline,
        getAppState: async (key, options) => { assert.equal(options.strict, true); if (offline) throw new Error('offline'); return records.has(key) ? { payload: records.get(key) } : null; },
        setAppState: async (key, payload, options) => {
            const previous = records.get(key);
            if (offline || (previous ? options.createOnly || options.expectedVersion !== previous.version : options.expectedVersion !== 0)) return false;
            records.set(key, payload);
            return true;
        }
    };
    new Function('router', 'settingsService', 'getSessionUser', 'requireAuthenticated', 'crypto', 'gohData', 'require', 'express', 'logEvent', source.slice(start, end))(
        { get(route, ...callbacks) { handlers.set('GET ' + route, callbacks); }, put(route, ...callbacks) { handlers.set('PUT ' + route, callbacks); } },
        { getActiveProfile: () => profile }, request => request.sessionUser,
        (request, response, next) => request.sessionUser ? next() : response.status(401).json({ ok: false }),
        require('node:crypto'), gohData, require, { json: () => (_request, _response, next) => next() }, () => {}
    );
    async function invoke(method, username, body, query = { profile: 'production' }) {
        const result = { status: 200, body: null };
        const response = { setHeader() {}, status(value) { result.status = value; return this; }, json(value) { result.body = value; return this; } };
        const request = { query, body, sessionUser: username ? { username } : null };
        for (const callback of handlers.get(method + ' /dashboard/preferences')) {
            let next = false;
            await callback(request, response, () => { next = true; });
            if (!next) break;
        }
        return result;
    }
    assert.equal((await invoke('GET')).status, 401);
    assert.equal((await invoke('PUT', null, {})).status, 401);
    assert.equal((await invoke('PUT', 'MW', { version: 0, username: 'TB', config: { active: 'sales', query: ' logi ', unknown: 'ignored' } })).status, 200);
    assert.equal((await invoke('GET', 'mw')).body.config.query, 'logi');
    assert.equal((await invoke('GET', 'TB')).body.config, null);
    assert.equal(records.size, 1);
    assert.ok([...records.keys()][0].length <= 100);
    assert.equal((await invoke('PUT', 'MW', { version: 0, config: { active: 'economy' } })).status, 409);
    assert.equal((await invoke('GET', 'MW')).body.config.active, 'sales');
    assert.equal((await invoke('GET', 'MW', null, { profile: 'wrong' })).status, 409);
    assert.equal((await invoke('PUT', 'MW', { version: -1, config: {} })).status, 400);
    profile = { ...profile, database: 'FactoryTwo' };
    assert.equal((await invoke('GET', 'MW')).body.config, null);
    offline = true;
    assert.equal((await invoke('GET', 'MW')).status, 503);
    assert.equal((await invoke('PUT', 'MW', { version: 0, config: {} })).status, 503);
});

test('strict AppState reads distinguish missing data from GOH failure without changing fail-soft callers', async () => {
    const source = require('node:fs').readFileSync(require('node:path').join(__dirname, '../services/gohDataService.js'), 'utf8');
    const start = source.indexOf('async function getAppState(');
    const end = source.indexOf('async function getAppStatesByPrefix(', start);
    let enabled = false;
    let failed = false;
    const request = { input() { return this; }, async query() { if (failed) throw new Error('query failed'); return { recordset: [] }; } };
    const getState = new Function('sql', 'isEnabled', 'getPool', 'markUnavailable', source.slice(start, end) + '; return getAppState;')(
        { NVarChar: () => 'nvarchar' }, () => enabled, async () => ({ request: () => request }), () => {}
    );
    assert.equal(await getState('dashboard'), null);
    await assert.rejects(getState('dashboard', { strict: true }), /GOH/);
    enabled = true;
    assert.equal(await getState('dashboard', { strict: true }), null);
    failed = true;
    await assert.rejects(getState('dashboard', { strict: true }), /query failed/);
    assert.equal(await getState('dashboard'), null);
    assert.match(source, /MERGE dbo\.AppState WITH \(HOLDLOCK\)/);
    assert.match(source, /expectedVersion = -1/);
    assert.match(source, /WHEN NOT MATCHED AND @expectedVersion IN \(-1, 0\)/);
});