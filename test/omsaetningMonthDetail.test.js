const test = require('node:test');
const assert = require('node:assert/strict');
const {
    createOmsaetningService,
    parseCalendarMonth,
    groupMonthDetailRows
} = require('../services/omsaetningService');

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
            MatchedOrdNo: 398383, OrderMatchCount: 1
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
                    [{ WeekKey: 202632 }, { WeekKey: 202633 }]
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
    assert.deepEqual(result.weekKeys, ['202632', '202633']);
    assert.equal(result.rows[0].ordNo, 400100);
});
