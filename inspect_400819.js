const getConnection = require("./db.js");
const sql = require("mssql/msnodesqlv8");
const diskCache = require("./diskCache.js");
const { createAftercalcService } = require("./services/aftercalcService.js");
const {
    isGloballyExcludedProdNo,
    isExcludedOperationProdNo,
    isEstimatedOperationMinutesFallback,
    getEffectiveOperationMinutes,
    adjustOperationLinePricing,
    isLaserLProduct
} = require("./utils/productRules.js");

const ORDNO = Number(process.argv[2] || 400819);

async function getLatestDrawingByProdNo() { return new Map(); }

function fmt(n) {
    if (n === null || n === undefined) return "-";
    const v = Number(n);
    if (!Number.isFinite(v)) return String(n);
    return v.toFixed(2);
}

async function run() {
    const pool = await getConnection();
    const { getAfterCalc } = createAftercalcService({
        getConnection: async () => pool,
        sql,
        diskCache,
        logEvent: () => {},
        getLatestDrawingByProdNo,
        isGloballyExcludedProdNo,
        isExcludedOperationProdNo,
        isEstimatedOperationMinutesFallback,
        getEffectiveOperationMinutes,
        adjustOperationLinePricing,
        isLaserLProduct,
        orderListMaxRows: 500,
        orderListDaysBack: 365,
        cacheTtlProductionSummaryMs: 30 * 60 * 1000
    });

    const data = await getAfterCalc(ORDNO);

    if (!data || data.error) {
        console.log("ERROR:", data && data.error);
        process.exit(1);
    }

    console.log("=== Order header (ordre " + ORDNO + ") ===");
    console.log({
        OrdNo: data.orderHeader.OrdNo,
        InvoAm: data.orderHeader.InvoAm,
        DInvoIF: data.orderHeader.DInvoIF
    });

    console.log("\n=== Salgsordrelinjer ===");
    console.table((data.salesOrderLines || []).map(l => ({
        LnNo: l.LnNo,
        ProdNo: l.ProdNo,
        Descr: (l.Descr || "").substring(0, 28),
        NoFin: l.NoFin,
        DPrice: fmt(l.DPrice),
        Salgspris: fmt((l.DPrice || 0) * (l.NoFin || 0)),
        EffCost: fmt(l.EffectiveLineCost),
        ProdTotal: fmt(l.ProductionOrderTotalCost),
        PurcNo: l.PurcNo || "-",
        IsDisc: l.IsDiscountLine
    })));

    console.log("\n=== Ekstra salgslinjer (uden PurcNo) ===");
    console.table((data.salesLines || []).map(l => ({
        LnNo: l.LnNo,
        ProdNo: l.ProdNo,
        Descr: (l.Descr || "").substring(0, 28),
        NoFin: l.NoFin,
        DPrice: fmt(l.DPrice),
        Salgspris: fmt((l.DPrice || 0) * (l.NoFin || 0)),
        EffCost: fmt(l.EffectiveLineCost),
        IsDisc: l.IsDiscountLine
    })));

    console.log("\n=== Produktionsordrer ===");
    console.table((data.productionOrders || []).map(p => ({
        OrdNo: p.ordNo,
        TrTp: p.trTp,
        Revenue: fmt(p.revenue),
        TotalCost: fmt(p.totalCost),
        Lines: (p.lines || []).length
    })));

    console.log("\n=== Totaler ===");
    const s = data.summary || {};
    console.log({
        salesLinesTotalCost: fmt(data.salesLinesTotalCost),
        totalCost: fmt(s.totalCost),
        totalRevenue: fmt(s.totalRevenue),
        margin: fmt(s.margin),
        marginPercentage: s.marginPercentage
    });

    process.exit(0);
}

run().catch(err => {
    console.error(err);
    process.exit(1);
});
