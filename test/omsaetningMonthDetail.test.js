const test = require('node:test');
const assert = require('node:assert/strict');
const {
    createOmsaetningService,
    parseCalendarMonth,
    groupMonthDetailRows
} = require('../services/omsaetningService');
const orderFlow = require('../assets/js/order-flow');

test('order flow reconciles carryover, new orders, partial invoicing and same-month completion', () => {
    const result = orderFlow.build([
        { OrdNo: 1, OrderDate: 20260801, OrderValueDkk: 1000, InvoicedDkk: 600, RemainingDkk: 400 },
        { OrdNo: 2, OrderDate: 20260901, OrderValueDkk: 500, InvoicedDkk: 500, RemainingDkk: 0 },
        { OrdNo: 3, OrderDate: 20260915, OrderValueDkk: 200, InvoicedDkk: 0, RemainingDkk: 200 }
    ], [
        { OrdNo: 1, InvoicedBeforeDkk: 200, InvoicedInMonthDkk: 400, InvoicedTotalDkk: 600 },
        { OrdNo: 2, InvoicedBeforeDkk: 0, InvoicedInMonthDkk: 500, InvoicedTotalDkk: 500 }
    ], 20260901);
    assert.equal(result.total.opening, 800);
    assert.equal(result.total.incoming, 700);
    assert.equal(result.prior.invoiced, 400);
    assert.equal(result.received.invoiced, 500);
    assert.equal(result.received.completed, 1);
    assert.equal(result.total.closing, 600);
    assert.equal(result.total.current, 600);
    assert.equal(result.unknownCount, 0);
    assert.equal(result.total.opening + result.total.incoming - result.total.invoiced + result.total.adjustment, result.total.closing);
});

test('order flow separates current rest from historical rest and never invents missing invoice history', () => {
    const result = orderFlow.build([
        { OrdNo: 1, OrderDate: 20260701, OrderValueDkk: 1000, InvoicedDkk: 1000, RemainingDkk: 0 },
        { OrdNo: 2, OrderDate: 20260801, OrderValueDkk: 600, InvoicedDkk: 400, RemainingDkk: 200 },
        { OrdNo: 3, OrderDate: 20260801, OrderValueDkk: 600, InvoicedDkk: 400, RemainingDkk: 200 },
        { OrdNo: 4, OrderDate: 20260801, OrderValueDkk: -600, InvoicedDkk: -600, RemainingDkk: 0 }
    ], [
        { OrdNo: 1, InvoicedBeforeDkk: 0, InvoicedInMonthDkk: 300, InvoicedTotalDkk: 1000 },
        { OrdNo: 3, InvoicedBeforeDkk: 0, InvoicedInMonthDkk: 100, InvoicedTotalDkk: 100 }
    ], 20260801);
    assert.equal(result.unknownCount, 2);
    assert.equal(result.rows.length, 1);
    assert.equal(result.total.closing, 700);
    assert.equal(result.total.current, 0);
});

test('order flow retains signed invoice corrections and makes balancing adjustments explicit', () => {
    const result = orderFlow.build([
        { OrdNo: 1, OrderDate: 20260701, OrderValueDkk: 100, InvoicedDkk: 150, RemainingDkk: 0 },
        { OrdNo: 2, OrderDate: 20260701, OrderValueDkk: 100, InvoicedDkk: 50, RemainingDkk: 50 }
    ], [
        { OrdNo: 1, InvoicedBeforeDkk: 0, InvoicedInMonthDkk: 150, InvoicedTotalDkk: 150 },
        { OrdNo: 2, InvoicedBeforeDkk: 100, InvoicedInMonthDkk: -50, InvoicedTotalDkk: 50 }
    ], 20260801);
    assert.equal(result.total.adjustment, 50);
    assert.equal(result.rows[1].invoiced, -50);
    assert.equal(result.total.opening + result.total.incoming - result.total.invoiced + result.total.adjustment, result.total.closing);
});

test('order flow fetches sales-only carryover and unique invoices in one customer-filtered batch', async () => {
    const inputs = new Map();
    let queries = 0;
    let sqlText = '';
    const request = {
        input(name, _type, value) { inputs.set(name, value); return this; },
        async query(text) {
            queries++;
            sqlText = text;
            return { recordsets: [[{ OrdNo: 1, OrderDate: 20250115, OrderValueDkk: 100, InvoicedDkk: 0, RemainingDkk: 100 }], []] };
        }
    };
    const service = createOmsaetningService({ getConnection: async () => ({ request: () => request }), sql: { Int: 'Int', MAX: 'MAX', NVarChar: () => 'text' } });
    const result = await service.getOrderFlow({ month: '2025-02', customerCsv: '42' });
    assert.equal(queries, 1);
    assert.equal(inputs.get('customerCsv'), '42');
    assert.equal(inputs.get('firstDate'), 20250201);
    assert.equal(inputs.get('asOf'), 20250228);
    assert.equal(inputs.get('period'), 202408);
    assert.match(sqlText, /o\.TrTp = 1 AND o\.OrdTp = 1/);
    assert.match(sqlText, /o\.CustNo > 0 AND c\.CustNo = o\.CustNo/);
    assert.match(sqlText, /posting\.Cust = invoice\.CustNo/);
    assert.match(sqlText, /HAVING COUNT_BIG\(\*\) = 1/);
    assert.match(sqlText, /o\.InvoIF > 0 OR o\.LstInvDt >= @firstDate/);
    assert.doesNotMatch(sqlText, /\bTOP\b/);
    assert.equal(result.prior.opening, 100);
    await assert.rejects(() => service.getOrderFlow({ month: 'invalid' }), { statusCode: 400 });
    await assert.rejects(() => service.getOrderFlow({ month: '2200-01' }), { statusCode: 400 });
    assert.equal(queries, 1);
});

test('calendar months map to Visma fiscal periods and exact date limits', () => {
    assert.deepEqual(parseCalendarMonth('2026-08'), {
        raw: '2026-08',
        year: 2026,
        month: 8,
        fiscalPeriodKey: 202602,
        firstDateInt: 20260801,
        lastDateInt: 20260831
    });
    assert.equal(parseCalendarMonth('2026-01').fiscalPeriodKey, 202507);
    assert.equal(parseCalendarMonth('2024-02').lastDateInt, 20240229);
    assert.equal(parseCalendarMonth('2026-13'), null);
    assert.equal(parseCalendarMonth('202608'), null);
});

test('monthly accounting rows are grouped by invoice without losing corrections', () => {
    const result = groupMonthDetailRows([
        {
            InvoNo: '1042048', VoNo: 1042048, VoDt: 20260827,
            CustNo: 20742710, CustomerName: 'ACJ Maskiner ApS',
            AcNo: 11012, AccountName: 'Salg', RevenueDkk: 555295,
            MatchedOrdNo: 398383, OrderMatchCount: 1, OrderDate: 20260812
        },
        {
            InvoNo: '1042048', VoNo: 1042048, VoDt: 20260827,
            CustNo: 20742710, CustomerName: 'ACJ Maskiner ApS',
            AcNo: 11012, AccountName: 'Salg', RevenueDkk: -253684,
            MatchedOrdNo: 398383, OrderMatchCount: 1
        },
        {
            InvoNo: '', VoNo: 900, VoDt: 20260828,
            CustNo: 0, CustomerName: '', AcNo: 11040, AccountName: 'Regulering',
            RevenueDkk: 125, MatchedOrdNo: null, OrderMatchCount: 0
        }
    ]);

    const order = result.rows.find(row => row.invoiceNo === '1042048');
    assert.equal(order.ordNo, 398383);
    assert.equal(order.orderDate, 20260812);
    assert.equal(order.linkStatus, 'matched');
    assert.equal(order.revenueDkk, 301611);
    assert.equal(order.accounts.length, 2);
    assert.equal(result.totalRevenueDkk, 301736);
    assert.equal(result.linkedRevenueDkk, 301611);
    assert.equal(result.unresolvedRevenueDkk, 125);
    assert.equal(result.unresolvedCount, 1);
});

test('ambiguous invoice matches remain visible but never choose an order', () => {
    const result = groupMonthDetailRows([{
        InvoNo: '42', VoNo: 42, VoDt: 20260801,
        AcNo: 11012, AccountName: 'Salg', RevenueDkk: 100,
        MatchedOrdNo: 400001, OrderMatchCount: 2
    }]);

    assert.equal(result.rows[0].ordNo, null);
    assert.equal(result.rows[0].linkStatus, 'ambiguous');
    assert.equal(result.unresolvedRevenueDkk, 100);
});

test('month detail query uses the same filters and returns exact related week keys', async () => {
    const inputs = new Map();
    let sqlText = '';
    const request = {
        input(name, _type, value) {
            inputs.set(name, value);
            return this;
        },
        async query(text) {
            sqlText = text;
            return {
                recordsets: [
                    [{
                        InvoNo: '100', VoNo: 100, VoDt: 20260810,
                        CustNo: 1, CustomerName: 'Kunde', AcNo: 11012,
                        AccountName: 'Salg', RevenueDkk: 500,
                        MatchedOrdNo: 400100, OrderMatchCount: 1
                    }],
                    [{ WeekKey: 202632 }, { WeekKey: 202633 }],
                    [{
                        OrdNo: 400100, OrderDate: 20260803, WeekKey: 202632, CustNo: 1,
                        CustomerName: 'Kunde', InvoNo: '100', InvoiceDate: 20260810,
                        InvoicedDkk: 500, RemainingDkk: 100, OrderValueDkk: 600
                    }, {
                        OrdNo: 400101, OrderDate: 20260820, WeekKey: 202633, CustNo: 1,
                        CustomerName: 'Kunde', InvoNo: '101', InvoiceDate: 20260905,
                        InvoicedDkk: 250, RemainingDkk: 0, OrderValueDkk: 250
                    }],
                    [{ OrdNo: 400100, InvoicedThroughMonthDkk: 500 }]
                ]
            };
        }
    };
    const fakeSql = {
        Int: 'Int',
        Bit: 'Bit',
        MAX: 'MAX',
        NVarChar(value) { return 'NVarChar(' + value + ')'; }
    };
    const service = createOmsaetningService({
        getConnection: async () => ({ request: () => request }),
        sql: fakeSql
    });

    const result = await service.getMonthDetail({
        month: '2026-08',
        accountCsv: '11012,11015',
        customerCsv: '1'
    });

    assert.equal(inputs.get('period'), 202602);
    assert.equal(inputs.get('firstDate'), 20260801);
    assert.equal(inputs.get('lastDate'), 20260831);
    assert.equal(inputs.get('accountCsv'), '11012,11015');
    assert.equal(inputs.get('customerCsv'), '1');
    assert.match(sqlText, /t\.AcYrPr = @period/);
    assert.match(sqlText, /o\.InvoNo/);
    assert.match(sqlText, /CustTr customerTransaction/);
    assert.match(sqlText, /o\.InvoSF/);
    assert.match(sqlText, /o\.InvoIF/);
    assert.match(sqlText, /orderCalendar\.Val8/);
    assert.deepEqual(result.weekKeys, ['202632', '202633']);
    assert.equal(result.rows[0].ordNo, 400100);
    assert.equal(result.receivedSummary.receivedCount, 2);
    assert.equal(result.receivedSummary.receivedValueDkk, 850);
    assert.equal(result.receivedSummary.invoicedThisMonthCount, 1);
    assert.equal(result.receivedSummary.invoicedThisMonthDkk, 500);
    assert.equal(result.receivedOrders[0].invoicedInMonthDkk, 500);
    assert.equal(result.receivedOrders[0].invoicedThroughMonthDkk, 500);
    assert.equal(result.receivedOrders[0].historicalRemainingDkk, 100);
    assert.equal(result.receivedOrders[0].weekKey, '202632');
    assert.equal(result.receivedOrders[1].completedAfterMonth, true);
    assert.equal(result.receivedOrders[1].complete, false);
    assert.equal(result.receivedOrders[1].historicalRemainingDkk, 250);
    assert.deepEqual(result.weeklyOrderRows, [
        { weekKey: '202632', totalOrd: 0.6, totalTilbud: 0 },
        { weekKey: '202633', totalOrd: 0.25, totalTilbud: 0 }
    ]);
    assert.equal(result.receivedSummary.openCount, 2);
    assert.equal(result.receivedSummary.remainingDkk, 350);
});
