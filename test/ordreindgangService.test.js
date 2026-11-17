const test = require('node:test');
const assert = require('node:assert/strict');
const { createOrdreindgangService } = require('../services/ordreindgangService');

test('monthly incoming order summary groups sales orders by order month and customer', async () => {
    const parameters = {};
    let queryText = '';
    const request = {
        input(name, _type, value) { parameters[name] = value; return this; },
        async query(text) {
            queryText = text;
            return { recordset: [
                { MonthKey: 202607, CustNo: 12, CustomerName: 'Kunde A', OrderCount: 2, OrderValueDkk: 125000 },
                { MonthKey: 202608, CustNo: null, CustomerName: 'Ukendt kunde', OrderCount: 1, OrderValueDkk: 50000 }
            ] };
        }
    };
    const service = createOrdreindgangService({
        getConnection: async () => ({ request: () => request }),
        sql: { Int: 'Int' }
    });

    const rows = await service.getMonthlyCustomerSummary({ fromDate: 20260701, exclusiveToDate: 20260901 });

    assert.deepEqual(parameters, { fromDate: 20260701, exclusiveToDate: 20260901 });
    assert.match(queryText, /O\.TrTp = 1\s+AND O\.OrdTp = 1/);
    assert.match(queryText, /O\.OrdDt >= @fromDate\s+AND O\.OrdDt < @exclusiveToDate/);
    assert.match(queryText, /COUNT\(DISTINCT OrdNo\)/);
    assert.match(queryText, /SUM\(OrderValueDkk\)/);
    assert.deepEqual(rows, [
        { month: '202607', custNo: 12, customerName: 'Kunde A', orderCount: 2, orderValueDkk: 125000 },
        { month: '202608', custNo: null, customerName: 'Ukendt kunde', orderCount: 1, orderValueDkk: 50000 }
    ]);
});
