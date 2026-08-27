const test = require('node:test');
const assert = require('node:assert/strict');
const { createLagerliste2Service, buildRouteLineage, isUnregisteredSearchPlate } = require('../services/lagerliste2Service');
const { reconcileMovements } = require('../assets/js/lagerliste2-engine');

function routeLine(overrides = {}) {
    return {
        NestingOrdNo: 500,
        Route: '01',
        OrdDt: 20260827,
        TrTp: 5,
        ProdNo: '30100200235123',
        Descr: 'Plate',
        Gr6: 1,
        ProdGr: 2,
        NoOrg: 100,
        NoFin: 100,
        CstPr: 10,
        IncCst: 1000,
        ...overrides
    };
}

test('route lineage keeps special L2/L3 products because membership comes from TrTp=7', () => {
    const result = buildRouteLineage([
        routeLine(),
        routeLine({ TrTp: 7, ProdNo: '1081400025-2L2', Descr: 'Special laser item', Gr6: 0, NoOrg: 2, NoFin: 1, CstPr: 300, TrInf2: '700', SalesOrderNo: 900 })
    ], [], {});

    assert.equal(result.routes.length, 1);
    assert.equal(result.routes[0].products[0].prodNo, '1081400025-2L2');
    assert.equal(result.routes[0].status, 'partial');
    assert.equal(result.routes[0].progress, 50);
});

test('registered rest is linked only to one exact nesting and plate', () => {
    const result = buildRouteLineage([
        routeLine(),
        routeLine({ TrTp: 7, ProdNo: '1000000001L', Gr6: 0, NoOrg: 1, NoFin: 0 })
    ], [{ ProdNo: '30100200235123', NestingOrdNo: 500, Val5: 20, Val8: 1, PriceType: '301', Txt2: 'Rest' }], { 301: 3 });

    assert.equal(result.routes[0].restPlates.length, 1);
    assert.equal(result.routes[0].restValue, 60);
    assert.equal(result.unassignedRest.length, 0);
});

test('ambiguous registered rest remains unassigned instead of guessing a route', () => {
    const result = buildRouteLineage([
        routeLine({ Route: '01' }),
        routeLine({ Route: '02' })
    ], [{ ProdNo: '30100200235123', NestingOrdNo: 500, Val5: 20, Val8: 1 }], { 301: 3 });

    assert.equal(result.routes[0].restPlates.length + result.routes[1].restPlates.length, 0);
    assert.equal(result.unassignedRest.length, 1);
    assert.equal(result.unassignedRest[0].reason, 'flere mulige ruter');
});

test('completed and open route states are calculated per product line', () => {
    const open = buildRouteLineage([routeLine(), routeLine({ TrTp: 7, ProdNo: '100L', Gr6: 0, NoOrg: 4, NoFin: 0 })], [], {}).routes[0];
    const completed = buildRouteLineage([routeLine(), routeLine({ TrTp: 7, ProdNo: '100L', Gr6: 0, NoOrg: 4, NoFin: 4 })], [], {}).routes[0];
    assert.equal(open.status, 'not_started');
    assert.equal(completed.status, 'completed');
});

test('Søg plate is marked as an unregistered rest source without confusing registered scrap codes', () => {
    assert.equal(isUnregisteredSearchPlate(routeLine({ TrInf1: 'Søg - står ved LG' })), true);
    assert.equal(isUnregisteredSearchPlate(routeLine({ TrInf1: 'søg' })), true);
    assert.equal(isUnregisteredSearchPlate(routeLine({ TrInf1: '301003_SCR0' })), false);

    const route = buildRouteLineage([
        routeLine({ TrInf1: 'Søg - står ved Eagle' }),
        routeLine({ TrTp: 7, ProdNo: '100L', Gr6: 0, NoOrg: 1, NoFin: 0, SalesOrderNo: 0 })
    ], [], {}).routes[0];
    assert.equal(route.plates[0].unregisteredRestSource, true);
    assert.equal(route.plates[0].sourceInfo, 'Søg - står ved Eagle');
    assert.equal(route.unlinkedProductCount, 1);
});

test('route exposes the selected sales-order source', () => {
    const route = buildRouteLineage([
        routeLine(),
        routeLine({
            TrTp: 7,
            ProdNo: '100L',
            Gr6: 0,
            NoOrg: 1,
            NoFin: 0,
            SalesOrderNo: 900,
            SalesOrderSource: 'OrdLn.R4'
        })
    ], [], {}).routes[0];
    assert.deepEqual(route.salesOrderReferences, [{ orderNo: 900, source: 'OrdLn.R4' }]);
});

test('route service performs only two SELECT queries and caches the read result', async () => {
    const queries = [];
    const recordsets = [[
        routeLine(),
        routeLine({ TrTp: 7, ProdNo: '100L', Gr6: 0, NoOrg: 1, NoFin: 0 })
    ], []];
    const pool = {
        request() {
            return {
                input() { return this; },
                async query(statement) {
                    queries.push(statement);
                    return { recordset: recordsets[queries.length - 1] || [] };
                }
            };
        }
    };
    const service = createLagerliste2Service({
        getConnection: async () => pool,
        sql: { Int: 'Int' },
        getRestPrices: () => ({})
    });

    const first = await service.getCurrentRoutes();
    const second = await service.getCurrentRoutes();
    assert.equal(first.routes.length, 1);
    assert.equal(second, first);
    assert.equal(queries.length, 2);
    assert.equal(queries.every(statement => /^\s*(WITH|SELECT)\b/i.test(statement)), true);
    assert.equal(queries.some(statement => /\b(INSERT|UPDATE|DELETE|MERGE|EXEC(?:UTE)?)\b/i.test(statement)), false);
    assert.match(queries[0], /CHARINDEX\s*\(/i);
    assert.match(queries[0], /Parent\.Depth\s*<\s*50/i);
    assert.match(queries[0], /L\.R4\s+AS\s+LineR4/i);
    assert.match(queries[0], /NULLIF\(TRY_CONVERT\(int, R\.LineR4\), 0\) IS NULL/i);
    assert.match(queries[0], /T\.ProdNo\s*=\s*R\.ProdNo/i);
    assert.match(queries[0], /ORDER BY TRY_CONVERT\(int, T\.FinDt\) DESC/i);
});

test('movement evidence reads purchase, nesting, invoice and Søg state without writes', async () => {
    const queries = [];
    const recordsets = [[
        { ProdNo: '311A', OrdNo: 700, TrTp: 6, FinDt: 20260827, FinTm: 1000, StcMov: 10, StcCst: 500, SupNo: 12 },
        { ProdNo: '311B', OrdNo: 701, TrTp: 5, FinDt: 20260827, FinTm: 1010, StcMov: -5, StcCst: -250 }
    ], [{ OrdNo: 900, InvoNo: '1042000', InvoAm: 1000, DInvoIF: 0, FinDt: 20260827 }], [
        { OrdNo: 500, Route: '01', ProdNo: '301X', SourceInfo: 'Søg - står ved LG' }
    ]];
    const pool = {
        request() {
            return {
                input() { return this; },
                async query(statement) {
                    queries.push(statement);
                    return { recordset: recordsets[queries.length - 1] || [] };
                }
            };
        }
    };
    const service = createLagerliste2Service({
        getConnection: async () => pool,
        sql: { Int: 'Int', VarChar: () => 'VarChar' },
        getRestPrices: () => ({})
    });

    const evidence = await service.getMovementEvidence({
        from: '2026-08-27T08:00:00.000Z',
        to: '2026-08-27T12:00:00.000Z',
        products: ['311A', '311B'],
        salesOrders: [900],
        nestingOrders: [500]
    });
    assert.equal(evidence.purchases[0].orderNo, 700);
    assert.equal(evidence.plateConsumptions[0].orderNo, 701);
    assert.equal(evidence.invoices[0].invoiceNo, '1042000');
    assert.deepEqual(evidence.unregisteredSearchPlates[0], { orderNo: 500, route: '01', prodNo: '301X', sourceInfo: 'Søg - står ved LG' });
    assert.equal(queries.length, 3);
    assert.equal(queries.some(statement => /\b(INSERT|UPDATE|DELETE|MERGE|EXEC(?:UTE)?)\b/i.test(statement)), false);
});

function payload(categories) {
    return { categories: { plates: [], restPlates: [], nestingCutting: [], stang: [], gr5Items: [], opfolgningvare: [], salgordreVia: [], finishedNotInvoiced: [], ...categories } };
}

test('period reconciliation nets plate transformation into VIA and registered rest', () => {
    const before = payload({ plates: [{ ProdNo: '301X', FifoValue: 1000 }] });
    const after = payload({
        nestingCutting: [{ OrdNo: 500, Route: '01', ProdNo: '301X', CountedValue: 800, SalesOrdNo: 900, Products: '100L' }],
        restPlates: [{ OrdNo: 500, ProdNo: '301X', Value: 200 }]
    });

    const result = reconcileMovements(before, after);
    assert.equal(result.reconciledValue, 1000);
    assert.equal(result.unresolved.length, 0);
});

test('period reconciliation preserves a real residual instead of forcing zero', () => {
    const before = payload({ plates: [{ ProdNo: '301X', FifoValue: 1000 }] });
    const after = payload({ nestingCutting: [{ OrdNo: 500, Route: '01', ProdNo: '301X', CountedValue: 750 }] });

    const result = reconcileMovements(before, after);
    assert.equal(result.reconciledValue, 750);
    assert.equal(result.unresolved.length, 1);
    assert.equal(result.unresolved[0].remaining, -250);
});

test('VIA Laser is linked to finished SO by sales order', () => {
    const before = payload({ salgordreVia: [{ OrdNo: 900, MaterialCost: 800 }] });
    const after = payload({ finishedNotInvoiced: [{ OrdNo: 900, Value: 800 }] });
    const result = reconcileMovements(before, after);
    assert.equal(result.transfers[0].flow, 'VIA → Færdige SO');
    assert.equal(result.unresolved.length, 0);
});

test('positive VIA time is generated order value and never a warehouse withdrawal', () => {
    const before = payload({});
    const after = payload({ salgordreVia: [{ OrdNo: 900, TimeCost: 425 }] });

    const result = reconcileMovements(before, after);
    assert.equal(result.transfers.length, 1);
    assert.equal(result.transfers[0].kind, 'addition');
    assert.equal(result.transfers[0].flow, 'Registreret arbejdstid → VIA Tid');
    assert.equal(result.transfers[0].sourceCategory, 'Registreret arbejdstid');
    assert.equal(result.transfers[0].netAmount, 425);
    assert.equal(result.addedValue, 425);
    assert.equal(result.unresolved.length, 0);
});

test('purchase receipt explains a positive plate-stock movement', () => {
    const before = payload({});
    const after = payload({ plates: [{ ProdNo: '311A', FifoValue: 14720 }] });
    const result = reconcileMovements(before, after, {
        evidence: { purchases: [{ prodNo: '311A', orderNo: 411993, value: 14720 }] }
    });

    assert.equal(result.transfers[0].flow, 'Indkøb → Pladelager');
    assert.equal(result.transfers[0].sourceKey, '411993');
    assert.equal(result.transfers[0].netAmount, 14720);
    assert.equal(result.unresolved.length, 0);
});

test('purchase immediately consumed by nesting explains direct value entering VIA plates', () => {
    const before = payload({});
    const after = payload({ nestingCutting: [{ OrdNo: 411992, Route: '01', ProdNo: '311A', CountedValue: 14720 }] });
    const result = reconcileMovements(before, after, {
        evidence: {
            purchases: [{ prodNo: '311A', orderNo: 411993, value: 14720 }],
            plateConsumptions: [{ prodNo: '311A', orderNo: 411992, value: 14720 }]
        }
    });

    assert.equal(result.transfers[0].flow, 'Indkøb → Pladelager → VIA Plader');
    assert.equal(result.transfers[0].sourceKey, '411993');
    assert.equal(result.transfers[0].targetKey, '411992/01/311A');
    assert.equal(result.unresolved.length, 0);
});

test('unregistered Søg rest explains value entering VIA without inventing a warehouse withdrawal', () => {
    const before = payload({});
    const after = payload({ nestingCutting: [{ OrdNo: 500, Route: '01', ProdNo: '301X', CountedValue: 250 }] });
    const result = reconcileMovements(before, after, {
        evidence: { unregisteredSearchPlates: [{ orderNo: 500, route: '01', prodNo: '301X', sourceInfo: 'Søg - står ved LG' }] }
    });

    assert.equal(result.transfers[0].flow, 'Uregistreret REST (Søg) → VIA Plader');
    assert.equal(result.transfers[0].kind, 'addition');
    assert.equal(result.transfers[0].netAmount, 250);
    assert.equal(result.unresolved.length, 0);
});

test('rest from another nesting is not matched by product number alone', () => {
    const before = payload({ plates: [{ ProdNo: '311A', FifoValue: 1000 }] });
    const after = payload({ restPlates: [{ OrdNo: 600, ProdNo: '311A', Value: 200 }] });
    const result = reconcileMovements(before, after);

    assert.equal(result.transfers.length, 0);
    assert.equal(result.unresolved.length, 2);
});

test('invoiced finished order is a legitimate exit, not an unexplained error', () => {
    const before = payload({ finishedNotInvoiced: [{ OrdNo: 409571, Value: 7049.93 }] });
    const after = payload({});
    const result = reconcileMovements(before, after, {
        evidence: { invoices: [{ orderNo: 409571, invoiceNo: '1042039' }] }
    });

    assert.equal(result.transfers[0].flow, 'Færdige SO → Faktureret');
    assert.equal(result.transfers[0].targetKey, '1042039');
    assert.equal(result.transfers[0].netAmount, -7049.93);
    assert.equal(result.unresolved.length, 0);
});

test('registered plate consumption explains stock leaving for a nesting order', () => {
    const before = payload({ plates: [{ ProdNo: '311A', FifoValue: 9307.98 }] });
    const after = payload({});
    const result = reconcileMovements(before, after, {
        evidence: { plateConsumptions: [{ prodNo: '311A', orderNo: 411747, value: 9307.98 }] }
    });

    assert.equal(result.transfers[0].flow, 'Pladelager → VIA Plader (registreret nesting)');
    assert.equal(result.transfers[0].targetKey, '411747');
    assert.equal(result.unresolved.length, 0);
});

test('completed nesting route explains remaining VIA plate value moving to production orders', () => {
    const before = payload({ nestingCutting: [{ OrdNo: 411618, Route: '07', ProdNo: '301A', CountedValue: 2286 }] });
    const after = payload({ restPlates: [{ OrdNo: 411618, ProdNo: '301A', Value: 659.86 }] });
    const result = reconcileMovements(before, after, {
        routes: [{ nestingOrdNo: 411618, route: '07', status: 'completed', productionOrderNos: [411582, 411584], salesOrderNos: [411581] }]
    });

    assert.equal(result.transfers[0].flow, 'VIA Plader → Rest plader');
    assert.equal(result.transfers[0].amount, 659.86);
    assert.equal(result.transfers[1].flow, 'VIA Plader → produktions-/salgsordre');
    assert.equal(result.transfers[1].amount, 1626.14);
    assert.equal(result.unresolved.length, 0);
});

test('completed nesting product can move into FIFO component stock', () => {
    const before = payload({
        nestingCutting: [{ OrdNo: 500, Route: '01', ProdNo: '301X', CountedValue: 800, Products: '1081400025-2L2' }]
    });
    const after = payload({
        gr5Items: [{ ProdNo: '1081400025-2L2', FifoValue: 800 }]
    });

    const result = reconcileMovements(before, after);
    assert.equal(result.transfers[0].flow, 'VIA Plader → lagerført produkt');
    assert.equal(result.transfers[0].targetCategory, 'Lager Komponenter (FIFO)');
    assert.equal(result.unresolved.length, 0);
});
