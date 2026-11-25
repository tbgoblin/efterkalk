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

function createService() {
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
                        return { recordset: query.includes('and purcno is null') ? [] : [salesOrderLine] };
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
