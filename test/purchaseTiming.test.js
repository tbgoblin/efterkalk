const test = require('node:test');
const assert = require('node:assert/strict');
const { classifyPurchaseTiming } = require('../services/bilancioService');

const lines = [
    { invoNo: 'A', receiptMonth: '2026-08', value: 100 },
    { invoNo: 'B', receiptMonth: '2026-10', value: 300 },
    { invoNo: 'B', receiptMonth: '2026-11', value: 100 },
    { invoNo: 'C', receiptMonth: '2026-09', value: 50 },
    { invoNo: 'D', receiptMonth: '2026-10', value: 70 }
];

test('bookings for other months are taken out and bookings elsewhere for this month are put back', () => {
    const postings = [
        { month: '2026-10', amount: 1000, invoNo: 'A', text: 'Indkøbs faktura' },     // received in August → out
        { month: '2026-10', amount: 400, invoNo: 'B', text: 'Indkøbs faktura' },      // 300 Oct (stays), 100 Nov → out
        { month: '2026-09', amount: 70, invoNo: 'D', text: 'Indkøbs faktura' },       // invoiced before October receipt → back in
        { month: '2026-10', amount: 5000, invoNo: 'X', text: 'Paller' }                // no purchase order → ignored
    ];
    const result = classifyPurchaseTiming('2026-10', postings, lines);
    assert.equal(result.bookedForOtherMonths, 1100);
    assert.equal(result.bookedElsewhereForThisMonth, 70);
    assert.equal(result.value, 1030);
    assert.equal(result.items[0].invoNo, 'A');
});

test('a credit note without its own order follows the invoice it refers to', () => {
    const postings = [{ month: '2026-10', amount: -40, invoNo: 'KN1', text: 'KN, Inox, vedr. fa. C' }];
    assert.equal(classifyPurchaseTiming('2026-10', postings, lines).value, 0);
    const withNumber = [{ month: '2026-10', amount: -40, invoNo: 'KN2', text: 'KN, Inox, vedr. fa. 30609690' }];
    const result = classifyPurchaseTiming('2026-10', withNumber, [{ invoNo: '30609690', receiptMonth: '2026-09', value: 80 }]);
    assert.equal(result.value, -40);
    assert.equal(result.creditNotes, -40);
    assert.equal(result.items[0].referencedInvoNo, '30609690');
    // Seen from September, the same credit note is booked elsewhere for September's goods.
    assert.equal(classifyPurchaseTiming('2026-09', withNumber, [{ invoNo: '30609690', receiptMonth: '2026-09', value: 80 }]).value, 40);
});
