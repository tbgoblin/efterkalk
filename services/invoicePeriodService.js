// Visma keeps only the last invoice date (Ord.LstInvDt) and the cumulative amount (Ord.InvoAm) per order.
// For an order invoiced in several periods, the revenue of a period is taken from the revenue postings of
// its invoices dated in that period, and its cost is scaled by the same share of the order's total revenue.
// Orders with every invoice inside the period are left untouched.

const REVENUE_GROUP = '10_Omsætning';

// invoices: [{ invoNo, inPeriod, revenue }] (revenue posted on the invoice, DKK, positive for sales)
// lines:    [{ ordNo, invoNo, amount }]     (order-line amounts per invoice, used to split shared invoices)
// Returns Map ordNo -> { revenueInPeriod, revenueTotal, share } for orders with invoices both inside and outside the period.
// share (for cost and material) counts positive invoices only: a credit note lowers revenue but returns no material.
function allocatePeriodRevenue(invoices, lines) {
    const invoiceTotals = new Map();
    for (const line of lines) invoiceTotals.set(line.invoNo, (invoiceTotals.get(line.invoNo) || 0) + Math.abs(line.amount));
    const invoicesByNo = new Map(invoices.map(invoice => [invoice.invoNo, invoice]));
    const orders = new Map();
    for (const line of lines) {
        const invoice = invoicesByNo.get(line.invoNo);
        if (!invoice) continue;
        const total = invoiceTotals.get(line.invoNo) || 0;
        const part = total > 0 ? invoice.revenue * Math.abs(line.amount) / total : 0;
        const order = orders.get(line.ordNo) || { revenueInPeriod: 0, revenueTotal: 0, positiveInPeriod: 0, positiveTotal: 0, outside: false };
        order.revenueTotal += part;
        if (part > 0) order.positiveTotal += part;
        if (invoice.inPeriod) {
            order.revenueInPeriod += part;
            if (part > 0) order.positiveInPeriod += part;
        } else order.outside = true;
        orders.set(line.ordNo, order);
    }
    const result = new Map();
    for (const [ordNo, order] of orders) {
        if (!order.outside || Math.abs(order.revenueInPeriod) < 0.005) continue;
        result.set(ordNo, {
            revenueInPeriod: Math.round(order.revenueInPeriod * 100) / 100,
            revenueTotal: Math.round(order.revenueTotal * 100) / 100,
            share: order.positiveTotal > 0 ? order.positiveInPeriod / order.positiveTotal : 0
        });
    }
    return result;
}

async function periodRevenueForOrders({ pool, sql, fromInt, toInt, custNo = null }) {
    // Orders with more than one invoice and at least one revenue posting dated in the period.
    const candidates = (await pool.request()
        .input('fromDate', sql.Int, fromInt).input('toDate', sql.Int, toInt).input('custNo', sql.Int, custNo).input('group', sql.NVarChar(40), REVENUE_GROUP)
        .query(`SELECT DISTINCT L.OrdNo FROM AcTr T WITH(NOLOCK)
                INNER JOIN Ac A WITH(NOLOCK) ON A.AcNo = T.AcNo AND A.AcGr = @group
                INNER JOIN OrdLn L WITH(NOLOCK) ON LTRIM(RTRIM(L.InvoNo)) = LTRIM(RTRIM(T.InvoNo))
                INNER JOIN Ord O WITH(NOLOCK) ON O.OrdNo = L.OrdNo AND O.TrTp = 1 AND ISNULL(O.InvoCnt, 0) > 1
                WHERE T.VoDt BETWEEN @fromDate AND @toDate AND ISNULL(T.InvoNo, '') <> ''
                  AND (@custNo IS NULL OR O.CustNo = @custNo)`)).recordset || [];
    if (!candidates.length) return new Map();
    const orderList = candidates.map(row => Number(row.OrdNo)).join(',');
    const lines = ((await pool.request().input('orders', sql.NVarChar(sql.MAX), orderList)
        .query(`SELECT L.OrdNo, LTRIM(RTRIM(L.InvoNo)) AS InvoNo,
                    SUM(CONVERT(float, ISNULL(L.InvoSF, 0)) + CONVERT(float, ISNULL(L.InvoIF, 0))) AS Amount
                FROM OrdLn L WITH(NOLOCK)
                WHERE L.InvoNo IN (SELECT DISTINCT L2.InvoNo FROM OrdLn L2 WITH(NOLOCK)
                                   WHERE L2.OrdNo IN (SELECT TRY_CONVERT(int, value) FROM STRING_SPLIT(@orders, ',')) AND ISNULL(L2.InvoNo, '') <> '')
                GROUP BY L.OrdNo, LTRIM(RTRIM(L.InvoNo))`)).recordset || [])
        .map(row => ({ ordNo: Number(row.OrdNo), invoNo: String(row.InvoNo), amount: Number(row.Amount || 0) }));
    const invoiceList = [...new Set(lines.map(line => line.invoNo))].join(',');
    const invoices = ((await pool.request().input('invoices', sql.NVarChar(sql.MAX), invoiceList).input('group', sql.NVarChar(40), REVENUE_GROUP)
        .query(`SELECT LTRIM(RTRIM(T.InvoNo)) AS InvoNo, MIN(T.VoDt) AS VoDt, SUM(-CONVERT(float, T.AcAm)) AS Revenue
                FROM AcTr T WITH(NOLOCK) INNER JOIN Ac A WITH(NOLOCK) ON A.AcNo = T.AcNo AND A.AcGr = @group
                WHERE LTRIM(RTRIM(T.InvoNo)) IN (SELECT LTRIM(RTRIM(value)) FROM STRING_SPLIT(@invoices, ','))
                GROUP BY LTRIM(RTRIM(T.InvoNo))`)).recordset || [])
        .map(row => ({ invoNo: String(row.InvoNo), revenue: Number(row.Revenue || 0), inPeriod: Number(row.VoDt) >= fromInt && Number(row.VoDt) <= toInt }));
    const candidateSet = new Set(candidates.map(row => Number(row.OrdNo)));
    const allocated = allocatePeriodRevenue(invoices, lines);
    return new Map([...allocated].filter(([ordNo]) => candidateSet.has(ordNo)));
}

module.exports = { allocatePeriodRevenue, periodRevenueForOrders };
