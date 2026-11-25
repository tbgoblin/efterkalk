const test = require('node:test');
const assert = require('node:assert/strict');
const { createAftercalcService } = require('../services/aftercalcService');

const salesOrderLine = {
    OrdNo: 418405, LnNo: 1, ProdNo: '1005401472', Descr: 'Multi Diameter Storage Frame',
    DPrice: 0, NoOrg: 100, NoFin: 100, NoInvo: 0, NoInvoAb: 0,
    CstPr: 2722.1571, CCstPr: 2722.1571, ProdTp4: 0, PurcNo: 418897, LineCost: 0
};

const productionOrderLines = [
    {
        OrdNo: 418897, LnNo: 1, ProdNo: '1005401472', Descr: 'Multi Diameter Storage Frame',
        DPrice: 0, NoOrg: 100, NoFin: 100, NoInvo: 0, NoInvoAb: 0,
        CstPr: 2722.1571, CCstPr: 2722.1571, PurcNo: 0, ProdTp4: 0, LineCost: 272215.71
    },
    {
        OrdNo: 418897, LnNo: 2, ProdNo: '1005401474', Descr: 'Rail for storage system. El-galv',
        DPrice: 0, NoOrg: 100, NoFin: 100, NoInvo: 0, NoInvoAb: 0,
        CstPr: 446.655166, CCstPr: 43325.329766, PurcNo: 0, ProdTp4: 2, LineCost: 4332532.98
    }
];

function createService(extraSalesOrderLines = []) {
    const pool = {
        request() {
            const inputs = {};
            return {
                input(name, _type, value) { inputs[name] = value; return this; },
                async query(statement) {
                    const query = statement.toLowerCase().replace(/\s+/g, ' ');
                    if (query.includes('select o.ordno, o.trtp, o.invoam')) {
                        return { recordset: [{ OrdNo: 418405, TrTp: 1, InvoAm: 0, DInvoIF: 0, Gr4: 0, InvoNo: '', CustomerName: 'Test' }] };
                    }
                    if (query.includes('select distinct purcno from ordln')) {
                        return { recordset: [{ PurcNo: 418897 }] };
                    }
                    if (query.includes('select ordln.ordno') && query.includes('where ordln.ordno = @purcno')) {
                        return { recordset: productionOrderLines };
                    }
                    if (query.includes('select o.ordno, a.nm, o.trtp, o.invoam')) {
                        return { recordset: [{ OrdNo: 418897, Nm: 'Test', TrTp: 7, InvoAm: 0 }] };
                    }
                    if (query.includes('select trtp from ord')) {
                        return { recordset: [{ TrTp: 7 }] };
                    }
                    if (query.includes('select ordln.lnno') && query.includes('ordln.cstpr')) {
                        return { recordset: productionOrderLines };
                    }
                    if (query.includes('select ordno, lnno, prodno, descr, dprice')) {
                        return { recordset: query.includes('and purcno is null') ? [] : [salesOrderLine, ...extraSalesOrderLines] };
                    }
                    throw new Error(`Unexpected SQL for order ${inputs.ordNo || inputs.purcNo || 'unknown'}`);
                }
            };
        }
    };
    return createAftercalcService({
        getConnection: async () => pool,
        sql: { Numeric: 'Numeric' },
        diskCache: { get: () => null, set: () => {} },
        logEvent: () => {},
        getLatestDrawingByProdNo: async () => new Map(),
        isGloballyExcludedProdNo: () => false,
        isExcludedOperationProdNo: () => false,
        isEstimatedOperationMinutesFallback: () => false,
        getEffectiveOperationMinutes: () => 0,
        adjustOperationLinePricing: line => line,
        isLaserLProduct: () => false,
        orderListMaxRows: 100,
        orderListDaysBack: 365,
        cacheTtlProductionSummaryMs: 1000
    });
}

test('unlinked stock materials use Visma CstPr in order and production summaries', async () => {
    const service = createService();
    const aftercalc = await service.getAfterCalc(418405);
    const productionLine = aftercalc.productionOrders[0].lines.find(line => line.ProdNo === '1005401474');

    assert.equal(productionLine.PurcNo, 0);
    assert.equal(productionLine.EffectiveLineCost, 44665.52);
    assert.equal(aftercalc.productionOrders[0].totalCost, 44665.52);

    const summary = await service.getProductionSummary(418897, new Set(), { forceRefresh: true });
    const summaryLine = summary.lines.find(line => line.ProdNo === '1005401474');

    assert.equal(summaryLine.CstPr, 446.655166);
    assert.equal(summaryLine.DisplayUnitCost, 446.66);
    assert.equal(summaryLine.EffectiveLineCost, 44665.52);
    assert.equal(summary.totalCost, 44665.52);
});

test('direct unlinked sales-order stock items use Visma CstPr while linked production cost stays intact', async () => {
    const directLine = {
        OrdNo: 418405, LnNo: 2, ProdNo: '1038400544', Descr: 'Lock plate',
        DPrice: 20.999045, NoOrg: 200, NoFin: 200, NoInvo: 0, NoInvoAb: 0,
        CstPr: 12.84391, CCstPr: 778.29211, ProdTp4: 0, PurcNo: 0,
        LineCost: 155658.42
    };
    const service = createService([directLine]);
    const aftercalc = await service.getAfterCalc(418405);
    const directSalesLine = aftercalc.salesOrderLines.find(line => line.LnNo === 2);
    const linkedSalesLine = aftercalc.salesOrderLines.find(line => line.LnNo === 1);

    assert.equal(directSalesLine.CstPr, 12.84391);
    assert.equal(directSalesLine.DisplayUnitCost, 12.84);
    assert.equal(directSalesLine.EffectiveLineCost, 2568.78);
    assert.equal(linkedSalesLine.EffectiveLineCost, 44665.52);
    assert.equal(aftercalc.summary.totalCost, 47234.3);
});

test('MultiOrdre laser cost includes all linked nesting pieces when the production line quantity differs', async () => {
    const sourceLines = [
        {
            LnNo: 1, ProdNo: '1108101132-3', Descr: '10mm',
            NoOrg: 1, NoFin: 1, NoInvo: 0, NoInvoAb: 0,
            DPrice: 0, CstPr: 1216.5, CCstPr: 1216.5,
            TrInf2: '', TrInf4: '0', ProdTp4: 0, PurcNo: 0,
            LineCost: 1216.5, ProdTrMinutes: 0
        },
        {
            LnNo: 2, ProdNo: '1108101132-3L', Descr: 'Laser 10mm',
            NoOrg: 12, NoFin: 12, NoInvo: 0, NoInvoAb: 0,
            DPrice: 0, CstPr: 101.375, CCstPr: 101.375,
            TrInf2: '', TrInf4: '0', ProdTp4: 2, PurcNo: 0,
            LineCost: 1216.5, NestingCost: null, ProdTrMinutes: 0
        }
    ];
    const finishedRow = {
        OrdNo: 418521, TrInf4: '03', ProdNo: '1108101132-3L', TrTp: 7,
        NoFin: 66, Free3: 1, CstPr: 170.6812, IncCst: 11264.96
    };
    const nestingRows = [
        {
            OrdNo: 418521, TrInf4: '03', ProdNo: '30101000355483', TrTp: 5,
            NoFin: 100, Free3: 0, CstPr: 112.6496, IncCst: 11264.96
        },
        finishedRow
    ];
    const salesOrderLine = {
        OrdNo: 418481, LnNo: 1, ProdNo: '1108101132-3', Descr: '10mm',
        DPrice: 0, NoOrg: 1, NoFin: 1, NoInvo: 0, NoInvoAb: 0,
        CstPr: 1216.5, CCstPr: 1216.5, ProdTp4: 4, PurcNo: 418491,
        LineCost: 1216.5
    };
    const pool = {
        request() {
            return {
                input() { return this; },
                async query(statement) {
                    const query = statement.toLowerCase().replace(/\s+/g, ' ');
                    if (query.includes('select o.ordno, o.trtp, o.invoam')) {
                        return {
                            recordset: [{
                                OrdNo: 418481, TrTp: 1, InvoAm: 0, DInvoIF: 0,
                                Gr4: 3, InvoNo: '', CustomerName: 'Test'
                            }]
                        };
                    }
                    if (query.includes('select distinct purcno from ordln')) {
                        return { recordset: [{ PurcNo: 418491 }] };
                    }
                    if (query.includes('select o.ordno, a.nm, o.trtp, o.invoam')) {
                        return { recordset: [{ OrdNo: 418491, Nm: 'Test', TrTp: 7, InvoAm: 0 }] };
                    }
                    if (query.includes('select trtp from ord')) {
                        return { recordset: [{ TrTp: 7 }] };
                    }
                    if (query.includes('select ordln.ordno') && query.includes('where ordln.ordno = @purcno')) {
                        return { recordset: sourceLines };
                    }
                    if (query.includes('select ordno, lnno, prodno, descr, dprice')) {
                        return { recordset: query.includes('and purcno is null') ? [] : [salesOrderLine] };
                    }
                    if (query.includes('where trinf2 = cast(@ordno as varchar(20))') && query.includes('and trtp = 7')) {
                        return { recordset: [finishedRow] };
                    }
                    if (query.includes('where ordno in (@nest0)') && query.includes('and trtp in (5, 7)')) {
                        return { recordset: nestingRows };
                    }
                    if (query.includes('from struct')) {
                        return { recordset: [] };
                    }
                    if (query.includes('select ordln.lnno')) {
                        return { recordset: sourceLines };
                    }
                    throw new Error(`Unexpected SQL in nesting cost test: ${query}`);
                }
            };
        }
    };
    const service = createAftercalcService({
        getConnection: async () => pool,
        sql: { Numeric: 'Numeric', VarChar: 'VarChar' },
        diskCache: { get: () => null, set: () => {} },
        logEvent: () => {},
        getLatestDrawingByProdNo: async () => new Map(),
        isGloballyExcludedProdNo: () => false,
        isExcludedOperationProdNo: () => false,
        isEstimatedOperationMinutesFallback: () => false,
        getEffectiveOperationMinutes: () => 0,
        adjustOperationLinePricing: line => line,
        isLaserLProduct: prodNo => String(prodNo || '').endsWith('L'),
        orderListMaxRows: 100,
        orderListDaysBack: 365,
        cacheTtlProductionSummaryMs: 1000
    });

    const aftercalc = await service.getAfterCalc(418481);
    const productionOrder = aftercalc.productionOrders[0];
    const nestedLaserLine = productionOrder.lines.find(line => line.ProdNo === '1108101132-3L');
    const summary = await service.getProductionSummary(418491, new Set(), { orderGr4: 3, forceRefresh: true });
    const laserLine = summary.lines.find(line => line.ProdNo === '1108101132-3L');

    assert.equal(nestedLaserLine.DisplayQuantity, 66);
    assert.equal(nestedLaserLine.EffectiveLineCost, 11264.96);
    assert.equal(productionOrder.totalCost, 11264.96);
    assert.equal(aftercalc.salesOrderLines[0].ProductionOrderTotalCost, 11264.96);
    assert.equal(aftercalc.summary.totalCost, 11264.96);
    assert.equal(laserLine.UsesLaserAllocationSpread, true);
    assert.equal(laserLine.DisplayQuantity, 66);
    assert.equal(laserLine.DisplayUnitCost, 170.68);
    assert.equal(laserLine.EffectiveLineCost, 11264.96);
    assert.equal(summary.totalCost, 11264.96);
    assert.match(laserLine.LaserAllocationText, /66 stk/);
});
