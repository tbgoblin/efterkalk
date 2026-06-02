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

    async function getSummary({ fra, til, accountCsv }) {
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
            .input('hasAccounts', sql.Bit, accountCsv.length > 0 ? 1 : 0);

        const result = await request.query(`
            SELECT
                t.AcNo,
                a.Nm,
                p.Yr,
                p.Pr,
                CONVERT(date, CONVERT(varchar(8), p.FrDt)) AS FrDtConverted,
                CAST(SUM(CAST(t.AcAm AS decimal(38, 6))) / 1000000.0 * -1.0 AS decimal(38, 6)) AS RevenueMio
            FROM AcTr t
            INNER JOIN AcPr p
                ON t.AcYr = p.Yr
               AND t.AcPr = p.Pr
            INNER JOIN Ac a
                ON t.AcNo = a.AcNo
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
            GROUP BY t.AcNo, a.Nm, p.Yr, p.Pr, p.FrDt
            ORDER BY FrDtConverted ASC, t.AcNo ASC
        `);

        const rows = (result.recordset || []).map(row => ({
            acNo: Number(row.AcNo),
            name: String(row.Nm || '').trim(),
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
                    .filter(Boolean)
            },
            totalRevenueMio,
            rows
        };
    }

    return {
        getAccounts,
        getSummary
    };
}

module.exports = {
    createOmsaetningService
};
