const test = require('node:test');
const assert = require('node:assert/strict');
const { createLagerliste2Service, buildRouteLineage, buildOrderStates, buildReservationSummary, isUnregisteredSearchPlate } = require('../services/lagerliste2Service');
const { buildMovements, canonicalValueSummary, reconcileMovements } = require('../assets/js/lagerliste2-engine');

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

test('registered REST is linked by the exact SCR code when one plate is used on multiple routes', () => {
    const result = buildRouteLineage([
        routeLine({ Route: '20' }),
        routeLine({ Route: '20', NoOrg: -67.44, NoFin: -67.44, TrInf2: 'REST_A_SCR0' }),
        routeLine({ Route: '20', TrTp: 7, ProdNo: '100A-L', Gr6: 0, NoOrg: 1, NoFin: 1, CstPr: 555.89 }),
        routeLine({ Route: '21' }),
        routeLine({ Route: '21', NoOrg: -49, NoFin: -49, TrInf2: 'REST_B_SCR0' }),
        routeLine({ Route: '21', TrTp: 7, ProdNo: '100B-L', Gr6: 0, NoOrg: 1, NoFin: 1, CstPr: 510 })
    ], [
        { ProdNo: '30100200235123', NestingOrdNo: 500, Txt1: 'REST_A_SCR0', Val5: 67.44, Val8: 1, PriceType: '301' },
        { ProdNo: '30100200235123', NestingOrdNo: 500, Txt1: 'REST_B_SCR0', Val5: 49, Val8: 1, PriceType: '301' }
    ], { 301: 3 });

    assert.equal(result.unassignedRest.length, 0);
    assert.equal(result.routes.find(route => route.route === '20').restPlates[0].restCode, 'REST_A_SCR0');
    assert.equal(result.routes.find(route => route.route === '21').restPlates[0].restCode, 'REST_B_SCR0');
});

test('route separates material allocation from REST write-down', () => {
    const route = buildRouteLineage([
        routeLine({ NoOrg: 150, NoFin: 150, CstPr: 6.585379 }),
        routeLine({ NoOrg: -67.4388, NoFin: -67.4388, CstPr: 6.585379, TrInf2: 'REST_A_SCR0' }),
        routeLine({ TrTp: 7, ProdNo: '100L', Gr6: 0, NoOrg: 5, NoFin: 5, CstPr: 108.74 })
    ], [{ ProdNo: '30100200235123', NestingOrdNo: 500, Txt1: 'REST_A_SCR0', Val5: 67.524, Val8: 1, PriceType: '301' }], { 301: 3 }).routes[0];

    assert.equal(route.estimatedRestFifoValue, 444.11);
    assert.equal(route.restValue, 202.57);
    assert.equal(route.materialAllocationResidual, 0);
    assert.equal(route.restWriteDown, 241.54);
    assert.equal(route.restRegistrationStatus, 'registered');
});

test('a REST line is recognized as færdigmeldt independently of unfinished products', () => {
    const partial = buildRouteLineage([
        routeLine({ NoOrg: 288, NoFin: 288, CstPr: 6.55 }),
        routeLine({ NoOrg: -89.472, NoFin: -89.472, CstPr: 6.45, TrInf2: 'BAWQKHFQ_SCR0' }),
        routeLine({ TrTp: 7, ProdNo: '100L', Gr6: 0, NoOrg: 10, NoFin: 8, CstPr: 54.36 })
    ], [], { 301: 3 }).routes[0];
    const completed = buildRouteLineage([
        routeLine({ NoOrg: 288, NoFin: 288, CstPr: 6.55 }),
        routeLine({ NoOrg: -89.472, NoFin: -89.472, CstPr: 6.45, TrInf2: 'BAWQKHFQ_SCR0' }),
        routeLine({ TrTp: 7, ProdNo: '100L', Gr6: 0, NoOrg: 10, NoFin: 10, CstPr: 54.36 })
    ], [], { 301: 3 }).routes[0];
    const unfinishedRest = buildRouteLineage([
        routeLine({ NoOrg: 288, NoFin: 288, CstPr: 6.55 }),
        routeLine({ NoOrg: -89.472, NoFin: 0, CstPr: 6.45, TrInf2: 'BAWQKHFQ_SCR0' }),
        routeLine({ TrTp: 7, ProdNo: '100L', Gr6: 0, NoOrg: 10, NoFin: 8, CstPr: 54.36 })
    ], [], { 301: 3 }).routes[0];

    assert.equal(partial.status, 'partial');
    assert.equal(partial.restLinesFinished, true);
    assert.equal(partial.restRegistrationStatus, 'finished_unregistered');
    assert.equal(completed.restRegistrationStatus, 'finished_unregistered');
    assert.equal(unfinishedRest.restRegistrationStatus, 'pending');
});

test('NoPac allocates only packed main-product quantities to finished stock', () => {
    const state = buildOrderStates([
        { OrdNo: 900, FinDt: 20260827, LnNo: 1, ProdNo: '111A', LineTrTp: 1, NoFin: 10, NoPac: 4, NoInvo: 0, CCstPr: 100 },
        { OrdNo: 900, FinDt: 20260827, LnNo: 2, ProdNo: '520', LineTrTp: 1, NoFin: 1, NoPac: 1, NoInvo: 0, CCstPr: 50 }
    ])[0];

    assert.equal(state.packedRatio, 0.4);
    assert.equal(state.partiallyPacked, true);
    assert.equal(state.headerClosed, true);
});

test('Rsv distinguishes still reserved, picked and finished quantities without adding a second value', () => {
    const result = buildReservationSummary([
        { ProdNo: '100-A', OrdNo: 700, SalesOrderNo: 900, NoRsv: 10, NoPic: 4, NoFin: 1, CstPr: 25 },
        { ProdNo: '100-B', OrdNo: 901, SalesOrderNo: 901, NoRsv: 2, NoPic: 2, NoFin: 2, CstPr: 100 }
    ]);

    assert.equal(result.rows[0].awaitingPickQty, 6);
    assert.equal(result.rows[0].pickedNotFinishedQty, 3);
    assert.equal(result.rows[0].activeValue, 225);
    assert.equal(result.rows[1].status, 'finished');
    assert.equal(result.summary.activeValue, 225);
    assert.equal(result.summary.registeredValue, 450);
    assert.equal(result.summary.finishedRowCount, 1);
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

test('current reservations are read-only, linked and cached', async () => {
    const queries = [];
    const pool = {
        request() {
            return {
                async query(statement) {
                    queries.push(statement);
                    return { recordset: [{
                        ProdNo: '100-A', Descr: 'Component', OrdNo: 700, OrdLnNo: 1,
                        SalesOrderNo: 900, LinkSource: 'OrdLn.R4', NoRsv: 10, NoPic: 4, NoFin: 1,
                        CstPr: 25, PoPhStB: 20, ShpRsv: 10, Bal: 20, StcInc: 0
                    }] };
                }
            };
        }
    };
    const service = createLagerliste2Service({ getConnection: async () => pool, sql: {}, getRestPrices: () => ({}) });

    const first = await service.getCurrentReservations();
    const second = await service.getCurrentReservations();
    assert.equal(first.rows[0].salesOrderNo, 900);
    assert.equal(first.rows[0].activeQty, 9);
    assert.equal(first.summary.activeValue, 225);
    assert.equal(second, first);
    assert.equal(queries.length, 1);
    assert.equal(/\b(INSERT|UPDATE|DELETE|MERGE|EXEC(?:UTE)?)\b/i.test(queries[0]), false);
});

test('movement evidence reads purchase, nesting, invoice and Søg state without writes', async () => {
    const queries = [];
    const recordsets = [[
        { ProdNo: '311A', OrdNo: 700, TrTp: 6, FinDt: 20260827, FinTm: 1000, StcMov: 10, StcCst: 500, SupNo: 12 },
        { ProdNo: '311B', OrdNo: 701, TrTp: 5, FinDt: 20260827, FinTm: 1010, StcMov: -5, StcCst: -250 },
        { ProdNo: '100-C', OrdNo: 900, R4: 900, TrTp: 1, FinDt: 20260827, FinTm: 1020, StcMov: -2, StcCst: -125 }
    ], [{ ProdNo: '100-C', OrdNo: 900, OrdLnNo: 1, SalesOrderNo: 900, LinkSource: 'OrdLn.R4', NoRsv: 2, NoPic: 2, NoFin: 2, CstPr: 62.5 }],
    [{ OrdNo: 900, InvoNo: '1042000', InvoAm: 1000, DInvoIF: 0, FinDt: 20260827 }], [
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
        products: ['311A', '311B', '100-C'],
        salesOrders: [900],
        nestingOrders: [500]
    });
    assert.equal(evidence.purchases[0].orderNo, 700);
    assert.equal(evidence.plateConsumptions[0].orderNo, 701);
    assert.equal(evidence.stockConsumptions[0].salesOrderNo, 900);
    assert.equal(evidence.stockConsumptions[0].value, 125);
    assert.equal(evidence.reservations[0].salesOrderNo, 900);
    assert.equal(evidence.reservations[0].status, 'finished');
    assert.equal(evidence.invoices[0].invoiceNo, '1042000');
    assert.deepEqual(evidence.unregisteredSearchPlates[0], { orderNo: 500, route: '01', prodNo: '301X', sourceInfo: 'Søg - står ved LG' });
    assert.equal(queries.length, 4);
    assert.match(queries[0], /TrTp IN \(1, 5, 6\)/i);
    assert.match(queries[1], /FROM Rsv R/i);
    assert.equal(queries.some(statement => /\b(INSERT|UPDATE|DELETE|MERGE|EXEC(?:UTE)?)\b/i.test(statement)), false);
});

function payload(categories) {
    return { categories: { plates: [], restPlates: [], nestingCutting: [], stang: [], gr5Items: [], opfolgningvare: [], salgordreVia: [], finishedNotInvoiced: [], ...categories } };
}

test('current V2 valuation preserves V1 rules and removes only documented VIA/finished overlap', () => {
    const current = payload({
        salgordreVia: [{ OrdNo: 900, MaterialCost: 800, StangCost: 100, PurchasedPartCost: 50, TimeCost: 50 }],
        finishedNotInvoiced: [{ OrdNo: 900, Value: 1000 }]
    });
    current.totals = { plates: 5000, restPlates: 200, stang: 300, opfolgningvare: 400, salgordreVia: 1000, finishedNotInvoiced: 1000, diverse: 0, total: 7900 };
    const result = canonicalValueSummary(current, {
        evidence: { orderStates: [{ orderNo: 900, packedRatio: 1, fullyPacked: true }] }
    });

    assert.equal(result.rawV1Total, 7900);
    assert.equal(result.total, 6900);
    assert.equal(result.duplicateReduction, 1000);
    assert.equal(result.categories.finishedNotInvoiced, 1000);
    assert.equal(result.categories.viaLaser, 0);
});

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
    assert.equal(result.transfers[0].flow, 'VIA → Færdige SO (ordre færdig/pakket)');
    assert.equal(result.unresolved.length, 0);
});

test('registered component consumption links Opfølgningsvarer to the exact finished sales order', () => {
    const before = payload({ opfolgningvare: [{ ProdNo: '1073400623-24', Value: 3135 }] });
    const after = payload({ finishedNotInvoiced: [{ OrdNo: 410168, Value: 3135 }] });
    const result = reconcileMovements(before, after, {
        evidence: {
            // ProdTr.StcCst may have been recalculated after the historical snapshot.
            stockConsumptions: [{ prodNo: '1073400623-24', orderNo: 410168, salesOrderNo: 410168, value: 3117.56 }]
        }
    });

    assert.equal(result.transfers[0].flow, 'Opfølgningsvarer → Færdige SO (registreret forbrug)');
    assert.equal(result.transfers[0].salesOrderNo, 410168);
    assert.equal(result.transfers[0].amount, 3135);
    assert.equal(result.unresolved.length, 0);
});

test('registered component consumption remains documented when VIA already consumed the finished target', () => {
    const before = payload({
        opfolgningvare: [{ ProdNo: '1073400623-24', Value: 3135 }],
        salgordreVia: [{ OrdNo: 410168, MaterialCost: 5000 }]
    });
    const after = payload({ finishedNotInvoiced: [{ OrdNo: 410168, Value: 5000 }] });
    const result = reconcileMovements(before, after, {
        evidence: {
            stockConsumptions: [{ prodNo: '1073400623-24', orderNo: 410168, salesOrderNo: 410168, value: 3117.56 }]
        }
    });

    assert.equal(result.transfers.length, 2);
    assert.equal(result.transfers[0].flow, 'VIA → Færdige SO (ordre færdig/pakket)');
    assert.equal(result.transfers[1].flow, 'Opfølgningsvarer → ordre/VIA (registreret forbrug)');
    assert.equal(result.transfers[1].salesOrderNo, 410168);
    assert.equal(result.transfers[1].amount, 3135);
    assert.equal(result.transfers[1].netAmount, 0);
    assert.equal(result.unresolved.length, 0);
});

test('an exact active Rsv link moves a component to reserved order stock without double counting', () => {
    const before = payload({ opfolgningvare: [{ ProdNo: '100-A', Value: 250 }] });
    const result = reconcileMovements(before, payload({}), {
        evidence: {
            reservations: [{ prodNo: '100-A', orderNo: 700, salesOrderNo: 900, activeQty: 10, activeValue: 250 }]
        }
    });

    assert.equal(result.transfers.length, 1);
    assert.equal(result.transfers[0].flow, 'Opfølgningsvarer → reserveret til ordre (Rsv)');
    assert.equal(result.transfers[0].targetCategory, 'Reserveret lager');
    assert.equal(result.transfers[0].salesOrderNo, 900);
    assert.equal(result.transfers[0].netAmount, 0);
    assert.equal(result.unresolved.length, 0);
});

test('ambiguous Rsv destinations are not guessed', () => {
    const before = payload({ opfolgningvare: [{ ProdNo: '100-A', Value: 250 }] });
    const result = reconcileMovements(before, payload({}), {
        evidence: {
            reservations: [
                { prodNo: '100-A', orderNo: 700, salesOrderNo: 900, activeQty: 5 },
                { prodNo: '100-A', orderNo: 701, salesOrderNo: 901, activeQty: 5 }
            ]
        }
    });

    assert.equal(result.transfers.length, 0);
    assert.equal(result.unresolved.length, 1);
});

test('Opfølgningsvarer are not linked to the wrong finished order', () => {
    const before = payload({ opfolgningvare: [{ ProdNo: '1073400623-24', Value: 3135 }] });
    const after = payload({ finishedNotInvoiced: [{ OrdNo: 410086, Value: 3135 }] });
    const result = reconcileMovements(before, after, {
        evidence: {
            stockConsumptions: [{ prodNo: '1073400623-24', orderNo: 410168, salesOrderNo: 410168, value: 3135 }]
        }
    });

    assert.equal(result.transfers.length, 1);
    assert.equal(result.transfers[0].flow, 'Opfølgningsvarer → ordre/VIA (registreret forbrug)');
    assert.equal(result.transfers[0].salesOrderNo, 410168);
    assert.equal(result.unresolved.length, 1);
    assert.equal(result.unresolved[0].category, 'Færdige SO');
    assert.equal(result.unresolved[0].salesOrderNo, 410086);
});

test('ambiguous registered component destinations remain unresolved', () => {
    const before = payload({ opfolgningvare: [{ ProdNo: '1073400623-24', Value: 3135 }] });
    const result = reconcileMovements(before, payload({}), {
        evidence: {
            stockConsumptions: [
                { prodNo: '1073400623-24', salesOrderNo: 410168, value: 2000 },
                { prodNo: '1073400623-24', salesOrderNo: 410731, value: 1135 }
            ]
        }
    });

    assert.equal(result.transfers.length, 0);
    assert.equal(result.unresolved.length, 1);
    assert.equal(result.unresolved[0].category, 'Opfølgningsvarer');
});

test('a fully packed overlapping order is represented by Færdige SO only', () => {
    const before = payload({
        salgordreVia: [{ OrdNo: 900, MaterialCost: 800 }],
        finishedNotInvoiced: [{ OrdNo: 900, Value: 800 }]
    });
    const after = payload({ finishedNotInvoiced: [{ OrdNo: 900, Value: 800 }] });
    const context = { evidence: { orderStates: [{ orderNo: 900, packedRatio: 1, fullyPacked: true, headerClosed: true }] } };
    const result = reconcileMovements(before, after, context);

    assert.equal(result.transfers.length, 0);
    assert.equal(result.movements.length, 0);
    assert.equal(result.unresolved.length, 0);
    assert.equal(result.orderAllocations.length, 1);
});

test('a newly fully packed overlap moves once from VIA to Færdige SO', () => {
    const before = payload({ salgordreVia: [{ OrdNo: 900, MaterialCost: 800 }] });
    const after = payload({
        salgordreVia: [{ OrdNo: 900, MaterialCost: 800 }],
        finishedNotInvoiced: [{ OrdNo: 900, Value: 800 }]
    });
    const context = { evidence: { orderStates: [{ orderNo: 900, packedRatio: 1, fullyPacked: true, headerClosed: true }] } };
    const result = reconcileMovements(before, after, context);

    assert.equal(result.transfers.length, 1);
    assert.equal(result.transfers[0].flow, 'VIA → Færdige SO (ordre færdig/pakket)');
    assert.equal(result.transfers[0].amount, 800);
    assert.equal(result.unresolved.length, 0);
});

test('partially packed finished order is split by NoPac without duplicating its value', () => {
    const after = payload({ finishedNotInvoiced: [{ OrdNo: 900, Value: 1000 }] });
    const movements = buildMovements(payload({}), after, {
        evidence: { orderStates: [{ orderNo: 900, packedRatio: 0.4, partiallyPacked: true }] }
    });
    const finished = movements.find(row => row.category === 'Færdige SO');
    const unpacked = movements.find(row => row.category === 'VIA ikke pakket');

    assert.equal(finished.diff, 400);
    assert.equal(unpacked.diff, 600);
    assert.equal(finished.diff + unpacked.diff, 1000);
});

test('current route lineage supplies a missing historical sales reference for plate-to-laser transfer', () => {
    const before = payload({ nestingCutting: [{ OrdNo: 500, Route: '01', ProdNo: '301X', CountedValue: 800 }] });
    const after = payload({ salgordreVia: [{ OrdNo: 900, MaterialCost: 800 }] });
    const result = reconcileMovements(before, after, {
        routes: [{ nestingOrdNo: 500, route: '01', salesOrderNos: [900] }]
    });

    assert.equal(result.transfers[0].flow, 'VIA Plader → VIA Laser');
    assert.equal(result.unresolved.length, 0);
});

test('an invoiced fully packed overlap exits once from Færdige SO', () => {
    const before = payload({
        salgordreVia: [{ OrdNo: 900, MaterialCost: 800 }],
        finishedNotInvoiced: [{ OrdNo: 900, Value: 800 }]
    });
    const result = reconcileMovements(before, payload({}), {
        evidence: {
            orderStates: [{ orderNo: 900, packedRatio: 1, fullyPacked: true, headerClosed: true, invoiced: true }],
            invoices: [{ orderNo: 900, invoiceNo: '1042000' }]
        }
    });

    assert.equal(result.transfers[0].flow, 'Færdige SO → Faktureret');
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

test('registered REST consumption is linked across nesting orders by exact SCR code', () => {
    const before = payload({
        restPlates: [{ OrdNo: 411727, ProdNo: '30100800355483', Txt1: 'BFTDOTKW_SCR0', Value: 96.6 }]
    });
    const after = payload({
        nestingCutting: [{ OrdNo: 412075, Route: '01', ProdNo: '30100800355483', CountedValue: 210.91 }]
    });
    const result = reconcileMovements(before, after, {
        evidence: {
            registeredRestConsumptions: [{
                restCode: 'BFTDOTKW_SCR0', nestingOrderNo: 412075, route: '01',
                prodNo: '30100800355483', value: 210.91, routeStarted: true
            }]
        }
    });

    assert.equal(result.transfers[0].flow, 'REST Plader → VIA Plader (genbrugt REST)');
    assert.equal(result.transfers[0].amount, 96.6);
    assert.equal(result.transfers[1].flow, 'Genindvundet værdi ved REST-forbrug → VIA Plader');
    assert.equal(result.transfers[1].amount, 114.31);
    assert.equal(result.unresolved.length, 0);
});

test('registered REST reserved by an open nesting is not left as a red disappearance', () => {
    const before = payload({
        restPlates: [{ OrdNo: 410744, ProdNo: '30100500355484', Txt1: 'USNQMAEZ_SCR0', Value: 50.03 }]
    });
    const result = reconcileMovements(before, payload({}), {
        evidence: {
            registeredRestConsumptions: [{
                restCode: 'USNQMAEZ_SCR0', nestingOrderNo: 412128, route: '15',
                prodNo: '30100500355484', value: 0, routeStarted: false
            }]
        }
    });

    assert.equal(result.transfers[0].flow, 'REST Plader → nesting (genbrugt REST)');
    assert.equal(result.transfers[0].targetCategory, 'Åben nesting');
    assert.equal(result.unresolved.length, 0);
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
