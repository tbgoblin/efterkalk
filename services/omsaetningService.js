function parseCalendarMonth(value) {
    const raw = String(value || '').trim();
    const match = raw.match(/^(\d{4})-(\d{2})$/);
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    if (!Number.isInteger(year) || year < 1900 || year > 2200 || month < 1 || month > 12) return null;

    const fiscalYear = month >= 7 ? year : year - 1;
    const fiscalPeriod = month >= 7 ? month - 6 : month + 6;
    const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
    return {
        raw,
        year,
        month,
        fiscalPeriodKey: fiscalYear * 100 + fiscalPeriod,
        firstDateInt: year * 10000 + month * 100 + 1,
        lastDateInt: year * 10000 + month * 100 + lastDay
    };
}

function groupMonthDetailRows(rawRows) {
    const groups = new Map();
    for (const rawRow of (Array.isArray(rawRows) ? rawRows : [])) {
        const invoiceNo = String(rawRow.InvoNo || '').trim();
        const voucherNo = Number(rawRow.VoNo || 0);
        const invoiceDate = Number(rawRow.VoDt || 0);
        const orderDate = Number(rawRow.OrderDate || 0);
        const custNo = Number(rawRow.CustNo || 0);
        const matchCount = Number(rawRow.OrderMatchCount || 0);
        const matchedOrdNo = matchCount === 1 ? Number(rawRow.MatchedOrdNo || 0) : null;
        const groupKey = invoiceNo
            ? 'invoice:' + invoiceNo
            : ['voucher', voucherNo, invoiceDate, custNo].join(':');

        if (!groups.has(groupKey)) {
            groups.set(groupKey, {
                invoiceNo,
                voucherNo,
                invoiceDate,
                orderDate: orderDate > 0 ? orderDate : null,
                custNo: custNo > 0 ? custNo : null,
                customerName: String(rawRow.CustomerName || '').trim(),
                ordNo: matchedOrdNo && matchedOrdNo > 0 ? matchedOrdNo : null,
                linkStatus: matchCount === 1 ? 'matched' : (matchCount > 1 ? 'ambiguous' : 'unmatched'),
                orderMatchCount: matchCount,
                revenueDkk: 0,
                accounts: []
            });
        }

        const group = groups.get(groupKey);
        const revenueDkk = Number(rawRow.RevenueDkk || 0);
        group.revenueDkk += revenueDkk;
        group.accounts.push({
            acNo: Number(rawRow.AcNo || 0),
            name: String(rawRow.AccountName || '').trim(),
            revenueDkk
        });
    }

    const rows = Array.from(groups.values()).sort((a, b) => {
        if (b.invoiceDate !== a.invoiceDate) return b.invoiceDate - a.invoiceDate;
        return String(b.invoiceNo || b.voucherNo).localeCompare(String(a.invoiceNo || a.voucherNo), 'da', { numeric: true });
    });
    const totalRevenueDkk = rows.reduce((sum, row) => sum + row.revenueDkk, 0);
    const linkedRevenueDkk = rows
        .filter(row => row.linkStatus === 'matched')
        .reduce((sum, row) => sum + row.revenueDkk, 0);

    return {
        rows,
        totalRevenueDkk,
        linkedRevenueDkk,
        unresolvedRevenueDkk: totalRevenueDkk - linkedRevenueDkk,
        linkedCount: rows.filter(row => row.linkStatus === 'matched').length,
        unresolvedCount: rows.filter(row => row.linkStatus !== 'matched').length
    };
}

function createOmsaetningService({ getConnection, sql }) {
    function isValidPeriod(value) {
        const raw = String(value || '').trim();
        const match = raw.match(/^(\d{4})(\d{2})$/);
        if (!match) return false;
        const month = Number(match[2]);
        return month >= 1 && month <= 12;
    }

    async function getAccounts() {
        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT AcNo, Nm
            FROM Ac
            WHERE AcGr = '10_Omsætning'
            ORDER BY AcNo
        `);

        return (result.recordset || []).map(row => ({
            acNo: Number(row.AcNo),
            name: String(row.Nm || '').trim()
        }));
    }

    async function searchCustomers({ queryText, limit = 20 }) {
        const pool = await getConnection();
        const normalizedLimit = Math.min(100, Math.max(1, Number(limit) || 20));
        const query = String(queryText || '').trim();
        const likePrefix = query + '%';

        const result = await pool.request()
            .input('query', sql.NVarChar(200), query)
            .input('likePrefix', sql.NVarChar(202), likePrefix)
            .input('limit', sql.Int, normalizedLimit)
            .query(`
                SELECT TOP (@limit)
                    a.CustNo,
                    a.Nm
                FROM Actor a
                WHERE a.CustNo > 0
                  AND LTRIM(RTRIM(ISNULL(a.Nm, ''))) <> ''
                  AND (
                      @query = ''
                      OR a.Nm LIKE @likePrefix
                      OR CONVERT(varchar(30), a.CustNo) LIKE @likePrefix
                  )
                  AND EXISTS (
                      SELECT 1
                      FROM AcTr t
                      WHERE t.Cust = a.CustNo
                        AND (t.SrcTp = 9 OR t.SrcTp = 1)
                  )
                ORDER BY a.Nm ASC, a.CustNo ASC
            `);

        return (result.recordset || []).map(row => ({
            custNo: Number(row.CustNo),
            name: String(row.Nm || '').trim()
        }));
    }

    async function getSummary({ fra, til, accountCsv, customerCsv }) {
        if (!isValidPeriod(fra) || !isValidPeriod(til)) {
            const error = new Error('Ugyldig periode. Brug format YYYYMM.');
            error.statusCode = 400;
            throw error;
        }

        const pool = await getConnection();
        const request = pool.request()
            .input('fra', sql.Int, Number(fra))
            .input('til', sql.Int, Number(til))
            .input('accountCsv', sql.NVarChar(sql.MAX), accountCsv)
            .input('hasAccounts', sql.Bit, accountCsv.length > 0 ? 1 : 0)
            .input('customerCsv', sql.NVarChar(sql.MAX), customerCsv)
            .input('hasCustomers', sql.Bit, customerCsv.length > 0 ? 1 : 0);

        const result = await request.query(`
            SELECT
                t.AcNo,
                a.Nm,
                CASE WHEN t.Cust > 0 THEN t.Cust ELSE NULL END AS CustNo,
                CASE WHEN t.Cust > 0 THEN c.Nm ELSE NULL END AS CustNm,
                (t.AcYrPr / 100) AS Yr,
                (t.AcYrPr % 100) AS Pr,
                DATEFROMPARTS(
                    CASE
                        WHEN (t.AcYrPr % 100) BETWEEN 1 AND 6 THEN (t.AcYrPr / 100)
                        ELSE (t.AcYrPr / 100) + 1
                    END,
                    CASE
                        WHEN (t.AcYrPr % 100) BETWEEN 1 AND 6 THEN (t.AcYrPr % 100) + 6
                        ELSE (t.AcYrPr % 100) - 6
                    END,
                    1
                ) AS FrDtConverted,
                CAST(SUM(CAST(t.AcAm AS decimal(38, 6))) / 1000000.0 * -1.0 AS decimal(38, 6)) AS RevenueMio
            FROM AcTr t
            INNER JOIN Ac a
                ON t.AcNo = a.AcNo
            LEFT JOIN Actor c
                ON t.Cust > 0
               AND t.Cust = c.CustNo
            WHERE
                (t.SrcTp = 9 OR t.SrcTp = 1)
                AND t.AcYrPr >= @fra
                AND t.AcYrPr < @til
                AND a.AcGr = '10_Omsætning'
                AND (
                    @hasAccounts = 0
                    OR EXISTS (
                        SELECT 1
                        FROM STRING_SPLIT(@accountCsv, ',') s
                        WHERE TRY_CAST(LTRIM(RTRIM(s.value)) AS int) = t.AcNo
                    )
                )
                AND (
                    @hasCustomers = 0
                    OR EXISTS (
                        SELECT 1
                        FROM STRING_SPLIT(@customerCsv, ',') s
                        WHERE TRY_CAST(LTRIM(RTRIM(s.value)) AS int) = t.Cust
                    )
                )
            GROUP BY t.AcNo, a.Nm, CASE WHEN t.Cust > 0 THEN t.Cust ELSE NULL END, CASE WHEN t.Cust > 0 THEN c.Nm ELSE NULL END, t.AcYrPr
            ORDER BY FrDtConverted ASC, t.AcNo ASC
        `);

        const rows = (result.recordset || []).map(row => ({
            acNo: Number(row.AcNo),
            name: String(row.Nm || '').trim(),
            custNo: row.CustNo === null || row.CustNo === undefined ? null : Number(row.CustNo),
            customerName: String(row.CustNm || '').trim(),
            year: Number(row.Yr),
            period: Number(row.Pr),
            date: row.FrDtConverted,
            revenueMio: Number(row.RevenueMio || 0)
        }));

        const totalRevenueMio = rows.reduce((sum, row) => sum + Number(row.revenueMio || 0), 0);

        return {
            filters: {
                fra,
                til,
                accounts: accountCsv
                    .split(',')
                    .map(v => String(v || '').trim())
                    .filter(Boolean),
                customers: customerCsv
                    .split(',')
                    .map(v => String(v || '').trim())
                    .filter(Boolean)
            },
            totalRevenueMio,
            rows
        };
    }

    async function getMonthDetail({ month, accountCsv = '', customerCsv = '' }) {
        const monthMeta = parseCalendarMonth(month);
        if (!monthMeta) {
            const error = new Error('Ugyldig måned. Brug format YYYY-MM.');
            error.statusCode = 400;
            throw error;
        }

        const normalizedAccountCsv = String(accountCsv || '').trim();
        const normalizedCustomerCsv = String(customerCsv || '').trim();
        const pool = await getConnection();
        const request = pool.request()
            .input('period', sql.Int, monthMeta.fiscalPeriodKey)
            .input('firstDate', sql.Int, monthMeta.firstDateInt)
            .input('lastDate', sql.Int, monthMeta.lastDateInt)
            .input('accountCsv', sql.NVarChar(sql.MAX), normalizedAccountCsv)
            .input('hasAccounts', sql.Bit, normalizedAccountCsv.length > 0 ? 1 : 0)
            .input('customerCsv', sql.NVarChar(sql.MAX), normalizedCustomerCsv)
            .input('hasCustomers', sql.Bit, normalizedCustomerCsv.length > 0 ? 1 : 0);

        const result = await request.query(`
            ;WITH FilteredRevenue AS (
                SELECT
                    t.InvoNo,
                    t.VoNo,
                    t.VoDt,
                    CASE WHEN t.Cust > 0 THEN t.Cust ELSE NULL END AS CustNo,
                    CASE WHEN t.Cust > 0 THEN c.Nm ELSE NULL END AS CustomerName,
                    t.AcNo,
                    a.Nm AS AccountName,
                    CAST(SUM(CAST(t.AcAm AS decimal(38, 6))) * -1.0 AS decimal(38, 6)) AS RevenueDkk
                FROM AcTr t
                INNER JOIN Ac a
                    ON t.AcNo = a.AcNo
                LEFT JOIN Actor c
                    ON t.Cust > 0
                   AND t.Cust = c.CustNo
                WHERE
                    t.SrcTp IN (1, 9)
                    AND t.AcYrPr = @period
                    AND a.AcGr = '10_Omsætning'
                    AND (
                        @hasAccounts = 0
                        OR EXISTS (
                            SELECT 1
                            FROM STRING_SPLIT(@accountCsv, ',') s
                            WHERE TRY_CAST(LTRIM(RTRIM(s.value)) AS int) = t.AcNo
                        )
                    )
                    AND (
                        @hasCustomers = 0
                        OR EXISTS (
                            SELECT 1
                            FROM STRING_SPLIT(@customerCsv, ',') s
                            WHERE TRY_CAST(LTRIM(RTRIM(s.value)) AS int) = t.Cust
                        )
                    )
                GROUP BY
                    t.InvoNo,
                    t.VoNo,
                    t.VoDt,
                    CASE WHEN t.Cust > 0 THEN t.Cust ELSE NULL END,
                    CASE WHEN t.Cust > 0 THEN c.Nm ELSE NULL END,
                    t.AcNo,
                    a.Nm
            )
            SELECT
                revenue.*,
                orderMatch.OrdNo AS MatchedOrdNo,
                orderMatch.MatchCount AS OrderMatchCount,
                matchedOrder.OrdDt AS OrderDate
            FROM FilteredRevenue revenue
            OUTER APPLY (
                SELECT
                    MIN(candidate.OrdNo) AS OrdNo,
                    COUNT_BIG(*) AS MatchCount
                FROM (
                    SELECT o.OrdNo
                    FROM Ord o
                    WHERE revenue.InvoNo IS NOT NULL
                      AND revenue.InvoNo <> ''
                      AND o.InvoNo = revenue.InvoNo
                    UNION
                    SELECT customerTransaction.OrdNo
                    FROM CustTr customerTransaction
                    WHERE customerTransaction.OrdNo > 0
                      AND revenue.InvoNo IS NOT NULL
                      AND revenue.InvoNo <> ''
                      AND customerTransaction.InvoNo = revenue.InvoNo
                ) candidate
            ) orderMatch
            LEFT JOIN Ord matchedOrder
              ON orderMatch.MatchCount = 1
             AND matchedOrder.OrdNo = orderMatch.OrdNo;

            SELECT DISTINCT CONVERT(int, Val8) AS WeekKey
            FROM FreeInf2
            WHERE FrInfTp = 550
              AND Dt1 >= @firstDate
              AND Dt1 <= @lastDate
              AND Val8 BETWEEN 190001 AND 220053
            ORDER BY WeekKey;

            SELECT
                o.OrdNo,
                o.OrdDt AS OrderDate,
                (
                    SELECT MIN(CONVERT(int, orderCalendar.Val8))
                    FROM FreeInf2 orderCalendar
                    WHERE orderCalendar.FrInfTp = 550
                      AND orderCalendar.Dt1 = o.OrdDt
                ) AS WeekKey,
                o.CustNo,
                c.Nm AS CustomerName,
                o.InvoNo,
                o.LstInvDt AS InvoiceDate,
                CAST(ISNULL(o.InvoSF, 0) * (ISNULL(o.ExRt, 100) / 100.0) AS decimal(38, 6)) AS InvoicedDkk,
                CAST(ISNULL(o.InvoIF, 0) * (ISNULL(o.ExRt, 100) / 100.0) AS decimal(38, 6)) AS RemainingDkk,
                CAST((ISNULL(o.InvoSF, 0) + ISNULL(o.InvoIF, 0)) * (ISNULL(o.ExRt, 100) / 100.0) AS decimal(38, 6)) AS OrderValueDkk
            FROM Ord o
            LEFT JOIN Actor c ON c.CustNo = o.CustNo
            WHERE o.TrTp = 1
              AND o.OrdTp = 1
              AND o.OrdDt >= @firstDate
              AND o.OrdDt <= @lastDate
              AND (
                  @hasCustomers = 0
                  OR EXISTS (
                      SELECT 1 FROM STRING_SPLIT(@customerCsv, ',') s
                      WHERE TRY_CAST(LTRIM(RTRIM(s.value)) AS int) = o.CustNo
                  )
              )
            ORDER BY o.OrdDt DESC, o.OrdNo DESC;

            ;WITH InvoiceOrderCandidates AS (
                SELECT DISTINCT LTRIM(RTRIM(o.InvoNo)) AS InvoNo, o.OrdNo
                FROM Ord o
                WHERE o.TrTp = 1 AND o.OrdTp = 1
                  AND o.OrdDt >= @firstDate AND o.OrdDt <= @lastDate
                  AND o.InvoNo IS NOT NULL AND LTRIM(RTRIM(o.InvoNo)) <> ''
                UNION
                SELECT DISTINCT LTRIM(RTRIM(ct.InvoNo)) AS InvoNo, ct.OrdNo
                FROM CustTr ct
                INNER JOIN Ord targetOrder ON targetOrder.OrdNo = ct.OrdNo
                WHERE targetOrder.TrTp = 1 AND targetOrder.OrdTp = 1
                  AND targetOrder.OrdDt >= @firstDate AND targetOrder.OrdDt <= @lastDate
                  AND ct.OrdNo > 0 AND ct.InvoNo IS NOT NULL AND LTRIM(RTRIM(ct.InvoNo)) <> ''
            ), UniqueInvoiceOrders AS (
                SELECT InvoNo, MIN(OrdNo) AS OrdNo
                FROM InvoiceOrderCandidates
                GROUP BY InvoNo
                HAVING COUNT_BIG(*) = 1
            )
            SELECT
                invoiceOrder.OrdNo,
                CAST(SUM(CAST(transactionRow.AcAm AS decimal(38, 6))) * -1.0 AS decimal(38, 6)) AS InvoicedThroughMonthDkk
            FROM AcTr transactionRow
            INNER JOIN Ac accountRow
                ON transactionRow.AcNo = accountRow.AcNo
               AND accountRow.AcGr = '10_Omsætning'
            INNER JOIN UniqueInvoiceOrders invoiceOrder
                ON invoiceOrder.InvoNo = LTRIM(RTRIM(transactionRow.InvoNo))
            WHERE transactionRow.SrcTp IN (1, 9)
              AND transactionRow.VoDt <= @lastDate
            GROUP BY invoiceOrder.OrdNo;
        `);

        const detail = groupMonthDetailRows((result.recordsets && result.recordsets[0]) || []);
        const invoicedInMonthByOrder = new Map();
        detail.rows.forEach(row => {
            const ordNo = Number(row.ordNo || 0);
            if (ordNo <= 0 || row.linkStatus !== 'matched') return;
            invoicedInMonthByOrder.set(ordNo, (invoicedInMonthByOrder.get(ordNo) || 0) + Number(row.revenueDkk || 0));
        });
        const weekKeys = ((result.recordsets && result.recordsets[1]) || [])
            .map(row => String(Number(row.WeekKey || 0)).padStart(6, '0'))
            .filter(value => /^\d{6}$/.test(value));
        const invoicedThroughMonthByOrder = new Map(
            ((result.recordsets && result.recordsets[3]) || []).map(row => [
                Number(row.OrdNo || 0),
                Number(row.InvoicedThroughMonthDkk || 0)
            ])
        );
        const receivedOrders = ((result.recordsets && result.recordsets[2]) || []).map(row => {
            const invoiceDate = Number(row.InvoiceDate || 0);
            const invoicedDkk = Number(row.InvoicedDkk || 0);
            const remainingDkk = Number(row.RemainingDkk || 0);
            const ordNo = Number(row.OrdNo || 0);
            const invoicedInMonthDkk = Number(invoicedInMonthByOrder.get(ordNo) || 0);
            const invoicedThisMonth = invoicedInMonthByOrder.has(ordNo);
            const completedAfterMonth = invoiceDate > monthMeta.lastDateInt;
            const orderValueDkk = Number(row.OrderValueDkk || 0);
            const hasFinancialHistory = invoicedThroughMonthByOrder.has(ordNo);
            const invoicedThroughMonthDkk = Number(invoicedThroughMonthByOrder.get(ordNo) || 0);
            const historicalRemainingDkk = hasFinancialHistory
                ? Math.max(0, orderValueDkk - invoicedThroughMonthDkk)
                : (completedAfterMonth ? orderValueDkk : remainingDkk);
            return {
                ordNo, orderDate:Number(row.OrderDate || 0), weekKey:String(Number(row.WeekKey || 0)).padStart(6, '0'),
                custNo:Number(row.CustNo || 0), customerName:String(row.CustomerName || '').trim(),
                invoiceNo:String(row.InvoNo || '').trim(), invoiceDate,
                orderValueDkk, invoicedDkk, invoicedInMonthDkk, invoicedThroughMonthDkk, remainingDkk,
                historicalRemainingDkk, completedAfterMonth, hasFinancialHistory,
                invoicedThisMonth, complete:historicalRemainingDkk <= 0
            };
        });
        const receivedSummary = {
            receivedCount:receivedOrders.length,
            receivedValueDkk:receivedOrders.reduce((sum, row) => sum + row.orderValueDkk, 0),
            invoicedThisMonthCount:receivedOrders.filter(row => row.invoicedThisMonth).length,
            invoicedThisMonthDkk:receivedOrders.reduce((sum, row) => sum + row.invoicedInMonthDkk, 0),
            openCount:receivedOrders.filter(row => !row.complete).length,
            remainingDkk:receivedOrders.filter(row => !row.complete).reduce((sum, row) => sum + Math.max(0, row.historicalRemainingDkk), 0)
        };
        const weeklyOrderRows = weekKeys.map(weekKey => ({
            weekKey,
            totalOrd:receivedOrders
                .filter(row => row.weekKey === weekKey)
                .reduce((sum, row) => sum + row.orderValueDkk, 0) / 1000,
            totalTilbud:0
        }));

        return {
            month: monthMeta.raw,
            fiscalPeriod: monthMeta.fiscalPeriodKey,
            filters: {
                accounts: normalizedAccountCsv.split(',').map(value => value.trim()).filter(Boolean),
                customers: normalizedCustomerCsv.split(',').map(value => value.trim()).filter(Boolean)
            },
            weekKeys,
            weeklyOrderRows,
            receivedOrders,
            receivedSummary,
            ...detail
        };
    }

    return {
        getAccounts,
        searchCustomers,
        getSummary,
        getMonthDetail
    };
}

module.exports = {
    createOmsaetningService,
    parseCalendarMonth,
    groupMonthDetailRows
};
