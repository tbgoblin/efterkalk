const test = require('node:test');
const assert = require('node:assert/strict');
const { createLagerlisteService } = require('../services/lagerlisteService');

test('product lookup resolves reservations through an existing line or header sales order', async () => {
    const queries = [];
    const pool = {
        request() {
            return {
                input() { return this; },
                async query(statement) {
                    queries.push(statement);
                    if (/FROM Prod P WITH\(NOLOCK\)/i.test(statement)) {
                        return { recordset: [{ ProdNo: '100-A', Descr: 'Component' }] };
                    }
                    return { recordset: [] };
                }
            };
        }
    };
    const service = createLagerlisteService({
        getConnection: async () => pool,
        sql: { NVarChar: 'NVarChar' }
    });

    await service.lookupProduct('100-A');

    const reservationQuery = queries.find(statement => /FROM Rsv R WITH\(NOLOCK\)/i.test(statement));
    assert.ok(reservationQuery);
    assert.match(reservationQuery, /LEFT JOIN OrdLn L[^\n]+L\.LnNo\s*=\s*R\.OrdLnNo/i);
    assert.match(reservationQuery, /LEFT JOIN Ord LineSO[^\n]+LineSO\.OrdNo\s*=\s*NULLIF\(TRY_CONVERT\(int, L\.R4\), 0\)/i);
    assert.match(reservationQuery, /LEFT JOIN Ord HeaderSO[^\n]+HeaderSO\.OrdNo\s*=\s*NULLIF\(TRY_CONVERT\(int, O\.R4\), 0\)/i);
    assert.match(reservationQuery, /COALESCE\(\s*LineSO\.OrdNo,\s*HeaderSO\.OrdNo,/i);
    assert.doesNotMatch(reservationQuery, /COALESCE\(TRY_CONVERT\(int, O\.R4\), 0\) AS SalesOrdNo/i);
});