const test = require('node:test');
const assert = require('node:assert/strict');
const { allocatePeriodRevenue } = require('../services/invoicePeriodService');

test('an order invoiced in two months gets only this month\'s revenue and a positive-invoice cost share', () => {
    const invoices = [
        { invoNo: 'I1', inPeriod: false, revenue: 5400 }, { invoNo: 'I2', inPeriod: false, revenue: 3000 }, { invoNo: 'I3', inPeriod: true, revenue: 1500 }
    ];
    const lines = [{ ordNo: 1, invoNo: 'I1', amount: 5400 }, { ordNo: 1, invoNo: 'I2', amount: 3000 }, { ordNo: 1, invoNo: 'I3', amount: 1500 }];
    const order = allocatePeriodRevenue(invoices, lines).get(1);
    assert.equal(order.revenueInPeriod, 1500);
    assert.equal(order.revenueTotal, 9900);
    assert.equal(Math.round(order.share * 10000), 1515);
});

test('a later credit note does not move material: the invoice month keeps the full cost, the credit month none', () => {
    const lines = [{ ordNo: 2, invoNo: 'F', amount: 11326.4 }, { ordNo: 2, invoNo: 'K', amount: -2000 }];
    const october = allocatePeriodRevenue([{ invoNo: 'F', inPeriod: true, revenue: 11326.4 }, { invoNo: 'K', inPeriod: false, revenue: -2000 }], lines).get(2);
    assert.deepEqual([october.revenueInPeriod, october.share], [11326.4, 1]);
    const november = allocatePeriodRevenue([{ invoNo: 'F', inPeriod: false, revenue: 11326.4 }, { invoNo: 'K', inPeriod: true, revenue: -2000 }], lines).get(2);
    assert.deepEqual([november.revenueInPeriod, november.share], [-2000, 0]);
});

test('orders with every invoice inside the period are left alone, and a shared invoice is split by line amounts', () => {
    const inside = allocatePeriodRevenue([{ invoNo: 'A', inPeriod: true, revenue: 100 }, { invoNo: 'B', inPeriod: true, revenue: 50 }],
        [{ ordNo: 3, invoNo: 'A', amount: 100 }, { ordNo: 3, invoNo: 'B', amount: 50 }]);
    assert.equal(inside.size, 0);
    const shared = allocatePeriodRevenue([{ invoNo: 'S', inPeriod: true, revenue: 1000 }, { invoNo: 'O', inPeriod: false, revenue: 400 }],
        [{ ordNo: 4, invoNo: 'S', amount: 300 }, { ordNo: 5, invoNo: 'S', amount: 700 }, { ordNo: 4, invoNo: 'O', amount: 400 }]);
    assert.equal(shared.get(4).revenueInPeriod, 300);
    assert.equal(shared.has(5), false);
});
