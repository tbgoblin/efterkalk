const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { allocatePeriodRevenue } = require('../services/invoicePeriodService');

function fixture(periodRevenue, records) {
    const source = fs.readFileSync(path.join(__dirname, '../routes/apiRoutes.js'), 'utf8');
    const start = source.indexOf("    router.get('/efterkalk/customer-invoices',");
    const end = source.indexOf("    router.post('/efterkalk/month-snapshot',", start);
    let handler;
    const parameters = {};
    const request = {
        input(name, type, value) { parameters[name] = value; return this; },
        async query(query) {
            // The cumulative-amount filter must apply only to the last-invoice-date branch.
            assert.match(query, /AND \(\(O\.LstInvDt >= @fromDate AND O\.LstInvDt <= @toDate\s+AND \(@includeAllAmounts = 1 OR O\.InvoAm <> 0\)\)\s+OR O\.OrdNo IN/);
            return { recordset: records };
        }
    };
    const pool = { request: () => request };
    const sql = { Int: 'int', Bit: 'bit', MAX: 'max', NVarChar: () => 'nvarchar' };
    new Function('router', 'getConnection', 'sql', 'periodRevenueForOrders', 'orderNotesService', 'logEvent',
        source.slice(start, end))(
        { get(url, callback) { handler = callback; } }, async () => pool, sql,
        async options => {
            assert.equal(options.pool, pool);
            assert.equal(options.custNo, parameters.expectedCustomer);
            return periodRevenue;
        },
        { getNote: () => null }, message => assert.fail(message)
    );
    return {
        parameters,
        async run(query) {
            parameters.expectedCustomer = query.scope === 'all' ? null : Number(query.custno);
            let body;
            await handler({ query }, { json(value) { body = value; }, status(code) { assert.fail(`Unexpected HTTP ${code}`); } });
            return body;
        }
    };
}

for (const [month, amount, share] of [['10', 12959.84, 1], ['11', -12959.84, 0]]) {
    test(`customer invoices retain net-zero orders with posted revenue in month ${month}`, async () => {
        const periodRevenue = allocatePeriodRevenue([
            { invoNo: 'F', inPeriod: month === '10', revenue: 12959.84 },
            { invoNo: 'K', inPeriod: month === '11', revenue: -12959.84 }
        ], [
            { ordNo: 1, invoNo: 'F', amount: 12959.84 },
            { ordNo: 1, invoNo: 'K', amount: -12959.84 }
        ]);
        for (const scope of [{ scope: 'all' }, { custno: '123' }]) {
            const f = fixture(periodRevenue, [{ OrdNo: 1, InvoAm: 0, LstInvDt: 20261105 }]);
            const result = await f.run({ ...scope, from: `2026-${month}-01`, to: `2026-${month}-30`, includeAllAmounts: '0' });
            assert.equal(result.ok, true);
            assert.equal(result.count, 1);
            assert.equal(result.totalInvoAm, amount);
            assert.equal(result.rows[0].InvoAm, amount);
            assert.equal(result.rows[0].InvoAmTotal, 0);
            assert.equal(result.rows[0].InvoiceShare, share);
            assert.equal(f.parameters.periodOrders, '1');
            assert.equal(f.parameters.includeAllAmounts, 0);
        }
    });
}

test('ordinary invoices keep their cumulative amounts and includeAllAmounts remains explicit', async () => {
    for (const includeAllAmounts of ['0', '1']) {
        const records = [{ OrdNo: 2, InvoAm: 100 }];
        if (includeAllAmounts === '1') records.push({ OrdNo: 3, InvoAm: 0 });
        const f = fixture(new Map(), records);
        const result = await f.run({ scope: 'all', from: '2026-10-01', to: '2026-10-31', includeAllAmounts });
        assert.equal(result.totalInvoAm, 100);
        assert.equal(result.count, records.length);
        assert.equal(Object.hasOwn(result.rows[0], 'InvoiceShare'), false);
        assert.equal(f.parameters.periodOrders, '');
        assert.equal(f.parameters.includeAllAmounts, Number(includeAllAmounts));
    }
});
