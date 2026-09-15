const test = require('node:test');
const assert = require('node:assert/strict');
const { fetchSalgordreViaRows, normalizePurchasedPartRows } = require('../services/viaService');

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